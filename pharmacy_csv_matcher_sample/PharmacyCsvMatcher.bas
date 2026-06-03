Attribute VB_Name = "PharmacyCsvMatcher"
Option Explicit

' Pharmacy CSV Matcher
' 公開用サンプルCSVに合わせた、採用品CSV・在庫CSV・対象品目CSVの突合ツールです。
' 商品コードはCSV化の過程で指数表記・丸め・先頭ゼロ消失が起きやすいため、
' 突合キーは「品名 + 規格容量 + メーカー」を主キーにしています。

Public Sub RunPharmacyCsvMatcher()
    Dim adoptedPath As String, inventoryPath As String, targetPath As String

    adoptedPath = PickCsvFile("採用品CSVを選択してください")
    If adoptedPath = "" Then Exit Sub

    inventoryPath = PickCsvFile("在庫CSVを選択してください")
    If inventoryPath = "" Then Exit Sub

    targetPath = PickCsvFile("対象品目CSVを選択してください")
    If targetPath = "" Then Exit Sub

    MatchPharmacyCsvFiles adoptedPath, inventoryPath, targetPath
End Sub

Private Sub MatchPharmacyCsvFiles(ByVal adoptedPath As String, ByVal inventoryPath As String, ByVal targetPath As String)
    Dim adoptedData As Variant, inventoryData As Variant, targetData As Variant
    Dim adoptedHeader As Object, inventoryHeader As Object, targetHeader As Object
    Dim adoptedDict As Object, targetDict As Object
    Dim i As Long, outRow As Long
    Dim key As String, stockQty As Double
    Dim resultWs As Worksheet, logWs As Worksheet
    Dim totalInventory As Long, stockPositiveCount As Long
    Dim adoptedMatchedCount As Long, targetMatchedCount As Long, outputCount As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    adoptedData = ReadCsvAsArray(adoptedPath)
    inventoryData = ReadCsvAsArray(inventoryPath)
    targetData = ReadCsvAsArray(targetPath)

    Set adoptedHeader = HeaderMap(adoptedData)
    Set inventoryHeader = HeaderMap(inventoryData)
    Set targetHeader = HeaderMap(targetData)

    ValidateHeaders adoptedHeader, Array("品名", "規格容量", "メーカー"), "採用品CSV"
    ValidateHeaders inventoryHeader, Array("品名", "規格容量", "メーカー", "在庫数"), "在庫CSV"
    ValidateHeaders targetHeader, Array("品名", "規格容量", "メーカー"), "対象品目CSV"

    Set adoptedDict = CreateObject("Scripting.Dictionary")
    Set targetDict = CreateObject("Scripting.Dictionary")

    For i = 2 To UBound(adoptedData, 1)
        key = MakeKey( _
            adoptedData(i, adoptedHeader("品名")), _
            adoptedData(i, adoptedHeader("規格容量")), _
            adoptedData(i, adoptedHeader("メーカー")) _
        )
        If key <> "" Then
            If Not adoptedDict.Exists(key) Then adoptedDict.Add key, i
        End If
    Next i

    For i = 2 To UBound(targetData, 1)
        key = MakeKey( _
            targetData(i, targetHeader("品名")), _
            targetData(i, targetHeader("規格容量")), _
            targetData(i, targetHeader("メーカー")) _
        )
        If key <> "" Then
            If Not targetDict.Exists(key) Then targetDict.Add key, i
        End If
    Next i

    Set resultWs = RecreateSheet("突合結果")
    Set logWs = RecreateSheet("突合ログ")

    With resultWs
        .Range("A1:N1").Value = Array( _
            "店舗", "品名", "規格容量", "メーカー", "在庫数", "単位", _
            "採用品一致", "対象品目一致", "対象区分", "先発品名", _
            "在庫側商品コード", "突合キー", "対象備考", "判定" _
        )
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(221, 235, 247)
    End With

    outRow = 2

    For i = 2 To UBound(inventoryData, 1)
        totalInventory = totalInventory + 1
        stockQty = ToNumber(inventoryData(i, inventoryHeader("在庫数")))

        If stockQty > 0 Then
            stockPositiveCount = stockPositiveCount + 1

            key = MakeKey( _
                inventoryData(i, inventoryHeader("品名")), _
                inventoryData(i, inventoryHeader("規格容量")), _
                inventoryData(i, inventoryHeader("メーカー")) _
            )

            If adoptedDict.Exists(key) Then
                adoptedMatchedCount = adoptedMatchedCount + 1

                If targetDict.Exists(key) Then
                    targetMatchedCount = targetMatchedCount + 1
                    outputCount = outputCount + 1

                    WriteResultRow resultWs, outRow, inventoryData, inventoryHeader, i, targetData, targetHeader, targetDict(key), key
                    outRow = outRow + 1
                End If
            End If
        End If
    Next i

    With resultWs
        .Columns("A:N").EntireColumn.AutoFit
        .Range("A1:N1").AutoFilter
        .Range("A1").Select
    End With

    With logWs
        .Range("A1:B1").Value = Array("項目", "件数")
        .Range("A2:B7").Value = Array( _
            Array("在庫CSV明細件数", totalInventory), _
            Array("在庫数が1以上の件数", stockPositiveCount), _
            Array("採用品CSVと一致した件数", adoptedMatchedCount), _
            Array("対象品目CSVと一致した件数", targetMatchedCount), _
            Array("突合結果への出力件数", outputCount), _
            Array("主キー", "品名 + 規格容量 + メーカー") _
        )
        .Range("A9").Value = "注意"
        .Range("A10").Value = "商品コード/JANは指数表記・丸め・先頭ゼロ消失が起きやすいため、判定には使用していません。"
        .Columns("A:B").EntireColumn.AutoFit
        .Rows(1).Font.Bold = True
        .Range("A9").Font.Bold = True
    End With

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "突合が完了しました。" & vbCrLf & _
           "出力件数: " & outputCount & "件" & vbCrLf & _
           "詳細は「突合結果」「突合ログ」を確認してください。", vbInformation
End Sub

Private Sub WriteResultRow(ByVal ws As Worksheet, ByVal outRow As Long, _
                           ByRef inventoryData As Variant, ByVal inventoryHeader As Object, ByVal inventoryRow As Long, _
                           ByRef targetData As Variant, ByVal targetHeader As Object, _
                           ByVal targetRow As Long, ByVal key As String)
    ws.Cells(outRow, 1).Value = GetValueIfExists(inventoryData, inventoryHeader, inventoryRow, "店舗")
    ws.Cells(outRow, 2).Value = inventoryData(inventoryRow, inventoryHeader("品名"))
    ws.Cells(outRow, 3).Value = inventoryData(inventoryRow, inventoryHeader("規格容量"))
    ws.Cells(outRow, 4).Value = inventoryData(inventoryRow, inventoryHeader("メーカー"))
    ws.Cells(outRow, 5).Value = inventoryData(inventoryRow, inventoryHeader("在庫数"))
    ws.Cells(outRow, 6).Value = GetValueIfExists(inventoryData, inventoryHeader, inventoryRow, "単位")
    ws.Cells(outRow, 7).Value = "○"
    ws.Cells(outRow, 8).Value = "○"
    ws.Cells(outRow, 9).Value = GetValueIfExists(targetData, targetHeader, targetRow, "対象区分")
    ws.Cells(outRow, 10).Value = GetValueIfExists(targetData, targetHeader, targetRow, "先発品名")
    ws.Cells(outRow, 11).Value = GetValueIfExists(inventoryData, inventoryHeader, inventoryRow, "商品コード")
    ws.Cells(outRow, 12).Value = key
    ws.Cells(outRow, 13).Value = GetValueIfExists(targetData, targetHeader, targetRow, "備考")
    ws.Cells(outRow, 14).Value = "在庫あり・採用品・対象品目"
End Sub

Private Function PickCsvFile(ByVal titleText As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = titleText
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            PickCsvFile = ""
        Else
            PickCsvFile = .SelectedItems(1)
        End If
    End With
End Function

Private Function ReadCsvAsArray(ByVal csvPath As String) As Variant
    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=csvPath, Local:=True)
    ReadCsvAsArray = wb.Worksheets(1).UsedRange.Value
    wb.Close SaveChanges:=False
End Function

Private Function HeaderMap(ByRef data As Variant) As Object
    Dim dict As Object
    Dim c As Long, headerName As String

    Set dict = CreateObject("Scripting.Dictionary")
    For c = 1 To UBound(data, 2)
        headerName = Trim(CStr(data(1, c)))
        If headerName <> "" Then
            If Not dict.Exists(headerName) Then dict.Add headerName, c
        End If
    Next c

    Set HeaderMap = dict
End Function

Private Sub ValidateHeaders(ByVal header As Object, ByVal requiredHeaders As Variant, ByVal csvName As String)
    Dim i As Long, h As String, missing As String

    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        h = CStr(requiredHeaders(i))
        If Not header.Exists(h) Then missing = missing & "・" & h & vbCrLf
    Next i

    If missing <> "" Then
        Err.Raise vbObjectError + 101, , csvName & "に必要な列がありません。" & vbCrLf & missing
    End If
End Sub

Private Function MakeKey(ByVal drugName As Variant, ByVal packageSize As Variant, ByVal maker As Variant) As String
    Dim n As String, p As String, m As String

    n = NormalizeText(drugName)
    p = NormalizeText(packageSize)
    m = NormalizeMaker(maker)

    If n = "" Or p = "" Or m = "" Then
        MakeKey = ""
    Else
        MakeKey = n & "|" & p & "|" & m
    End If
End Function

Private Function NormalizeText(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, "　", "")
    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, "（", "(")
    s = Replace(s, "）", ")")
    s = Replace(s, "【", "[")
    s = Replace(s, "】", "]")
    s = Replace(s, "「", """")
    s = Replace(s, "」", """")
    NormalizeText = UCase$(Trim$(s))
End Function

Private Function NormalizeMaker(ByVal v As Variant) As String
    Dim s As String
    s = NormalizeText(v)
    s = Replace(s, "株式会社", "")
    s = Replace(s, "(株)", "")
    s = Replace(s, "㈱", "")
    s = Replace(s, "有限会社", "")
    s = Replace(s, "(有)", "")
    NormalizeMaker = s
End Function

Private Function ToNumber(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        ToNumber = CDbl(v)
    Else
        ToNumber = 0
    End If
End Function

Private Function GetValueIfExists(ByRef data As Variant, ByVal header As Object, ByVal rowIndex As Long, ByVal headerName As String) As String
    If header.Exists(headerName) Then
        GetValueIfExists = CStr(data(rowIndex, header(headerName)))
    Else
        GetValueIfExists = ""
    End If
End Function

Private Function RecreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0

    Set RecreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    RecreateSheet.Name = sheetName
End Function
