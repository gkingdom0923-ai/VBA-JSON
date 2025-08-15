Option Explicit

' 需先在 Excel 參考 "Microsoft Scripting Runtime"
' 可使用 VBA-JSON 模組解析 JSON: https://github.com/VBA-tools/VBA-JSON
' 將 JSON.bas 模組導入到專案

'--------------------------------------
' 一鍵按鈕呼叫
'--------------------------------------
Sub 一鍵回測_JSON()
    Call 回測_TwelveData_JSON
End Sub

'--------------------------------------
' 主程式
'--------------------------------------
Sub 回測_TwelveData_JSON()
    
    Dim wsParam As Worksheet, wsData As Worksheet, wsCalc As Worksheet, wsWithdraw As Worksheet
    Dim stockCode As String, investPerMonth As Double, investYears As Long
    Dim withdrawRate As Double, retireAnnualReturn As Double
    Dim lastRow As Long, i As Long, yr As Long
    Dim capital As Double, monthlyWithdraw As Double, mRet As Double
    Dim chartObj As ChartObject
    Dim savePath As String, baseName As String, pdfFile As String, copyXlsm As String
    Dim URL As String, jsonText As String
    Dim startDate As Date, endDate As Date
    Dim apiKey As String: apiKey = "YOUR_API_KEY" ' ←請替換為你的 Twelve Data API Key
    
    Dim http As Object, JSON As Object, tsData As Object
    Dim item As Object, cumUnits As Double, price As Double, units As Double
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '---------------------------
    ' 輸入參數工作表
    '---------------------------
    On Error Resume Next
    Set wsParam = Sheets("輸入參數")
    If wsParam Is Nothing Then
        Set wsParam = Sheets.Add
        wsParam.Name = "輸入參數"
        wsParam.Range("A1:B6").Value = Array( _
            Array("參數", "數值"), _
            Array("股票代碼", "2330"), _
            Array("每月投入金額", 10000), _
            Array("投資年數", 20), _
            Array("提領率(年)", 0.04), _
            Array("退休期年化報酬", 0.03))
    End If
    wsParam.Columns("A:B").AutoFit
    On Error GoTo 0
    
    stockCode = wsParam.Range("B2").Value
    investPerMonth = wsParam.Range("B3").Value
    investYears = wsParam.Range("B4").Value
    withdrawRate = wsParam.Range("B5").Value
    retireAnnualReturn = wsParam.Range("B6").Value
    
    ' 台股自動加 .TW
    If IsNumeric(stockCode) Then stockCode = stockCode & ".TW"
    
    startDate = DateSerial(2000, 1, 1)
    endDate = Date
    
    URL = "https://api.twelvedata.com/time_series?symbol=" & stockCode & _
          "&interval=1month&start_date=" & Format(startDate, "yyyy-mm-dd") & _
          "&end_date=" & Format(endDate, "yyyy-mm-dd") & "&apikey=" & apiKey
    
    '---------------------------
    ' 下載 JSON
    '---------------------------
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", URL, False
    http.send
    If http.Status <> 200 Then
        MsgBox "下載資料失敗，請檢查股票代碼與 API Key", vbCritical
        Exit Sub
    End If
    jsonText = http.responseText
    
    '---------------------------
    ' 解析 JSON
    '---------------------------
    Set JSON = JsonConverter.ParseJson(jsonText)
    If Not JSON.Exists("values") Then
        MsgBox "JSON 格式錯誤或無資料", vbCritical
        Exit Sub
    End If
    Set tsData = JSON("values")
    
    '---------------------------
    ' 刪除舊工作表
    '---------------------------
    On Error Resume Next
    Sheets("StockData").Delete
    Sheets("Calculation").Delete
    Sheets("提領模擬").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    '---------------------------
    ' 建立 StockData 工作表
    '---------------------------
    Set wsData = Sheets.Add
    wsData.Name = "StockData"
    wsData.Range("A1:D1").Value = Array("日期", "開盤", "收盤", "調整收盤")
    
    lastRow = 2
    For Each item In tsData
        wsData.Cells(lastRow, 1).Value = item("datetime")
        wsData.Cells(lastRow, 2).Value = CDbl(item("open"))
        wsData.Cells(lastRow, 3).Value = CDbl(item("close"))
        wsData.Cells(lastRow, 4).Value = CDbl(item("close")) ' 可替換為 adj_close 若 API 回傳
        lastRow = lastRow + 1
    Next item
    wsData.Columns("A:D").AutoFit
    
    '---------------------------
    ' 建立 Calculation 工作表
    '---------------------------
    Set wsCalc = Sheets.Add
    wsCalc.Name = "Calculation"
    wsCalc.Range("A1:E1").Value = Array("日期", "Adj Close", "買進股數", "累積股數", "總資產")
    
    cumUnits = 0
    For i = 2 To wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
        price = wsData.Cells(i, 4).Value
        If price > 0 Then units = investPerMonth / price Else units = 0
        cumUnits = cumUnits + units
        wsCalc.Cells(i, 1).Value = wsData.Cells(i, 1).Value
        wsCalc.Cells(i, 2).Value = price
        wsCalc.Cells(i, 3).Value = units
        wsCalc.Cells(i, 4).Value = cumUnits
        wsCalc.Cells(i, 5).Value = cumUnits * price
    Next i
    wsCalc.Columns("A:E").AutoFit
    
    '---------------------------
    ' 資產成長圖
    '---------------------------
    Set chartObj = wsCalc.ChartObjects.Add(Left:=350, Top:=20, Width:=520, Height:=300)
    With chartObj.Chart
        .ChartType = xlLineMarkers
        .SetSourceData wsCalc.Range("A1:E" & wsCalc.Cells(wsCalc.Rows.Count, "A").End(xlUp).Row)
        .HasTitle = True
        .ChartTitle.Text = stockCode & " 定期定額回測"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "日期"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "資產(元)"
    End With
    
    '---------------------------
    ' 提領模擬
    '---------------------------
    Set wsWithdraw = Sheets.Add
    wsWithdraw.Name = "提領模擬"
    wsWithdraw.Range("A1:D1").Value = Array("年份", "期初資產", "年提領", "期末資產")
    
    capital = wsCalc.Cells(wsCalc.Cells(Rows.Count, "E").End(xlUp).Row, "E").Value
    monthlyWithdraw = (capital * withdrawRate) / 12
    mRet = (1 + retireAnnualReturn) ^ (1 / 12) - 1
    yr = 0
    Do While capital > 0
        Dim startCap As Double
        startCap = capital
        For i = 1 To 12
            capital = capital * (1 + mRet) - monthlyWithdraw
            If capital <= 0 Then Exit For
        Next i
        yr = yr + 1
        wsWithdraw.Cells(yr + 1, 1).Value = yr
        wsWithdraw.Cells(yr + 1, 2).Value = startCap
        wsWithdraw.Cells(yr + 1, 3).Value = monthlyWithdraw * 12
        wsWithdraw.Cells(yr + 1, 4).Value = Application.Max(capital, 0)
        If capital <= 0 Then Exit Do
        If yr > 120 Then Exit Do
    Loop
    wsWithdraw.Columns("A:D").AutoFit
    
    '---------------------------
    ' 提領圖
    '---------------------------
    Set chartObj = wsWithdraw.ChartObjects.Add(Left:=350, Top:=20, Width:=520, Height:=300)
    With chartObj.Chart
        .ChartType = xlLineMarkers
        .SetSourceData wsWithdraw.Range("A1:D" & yr + 1)
        .HasTitle = True
        .ChartTitle.Text = "提領模擬 (" & Format(withdrawRate, "0.0%") & " 年提領)"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "年份"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "資產(元)"
    End With
    
    '---------------------------
    ' 匯出 PDF + 保存專案
    '---------------------------
    savePath = ThisWorkbook.Path
    If Len(savePath) = 0 Then savePath = Application.DefaultFilePath
    baseName = stockCode & "_回測報告_" & Format(Now, "yyyymmdd_HHMM")
    pdfFile = savePath & "\" & baseName & ".pdf"
    copyXlsm = savePath & "\" & baseName & ".xlsm"
    
    Worksheets(Array("Calculation", "提領模擬")).Select
    Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFile, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Worksheets("Calculation").Select
    ActiveWorkbook.SaveCopyAs copyXlsm
    
    Application.ScreenUpdating = True
    
    MsgBox stockCode & " 回測完成！" & vbCrLf & _
           "投入 " & investYears & " 年，期末資產：" & Format(wsCalc.Cells(wsCalc.Cells(Rows.Count, "E").End(xlUp).Row, "E").Value, "#,##0") & vbCrLf & _
           "可維持約 " & yr & " 年（至資產歸零）" & vbCrLf & _
           "PDF 報告：" & pdfFile & vbCrLf & _
           "專案檔：" & copyXlsm, vbInformation

End Sub

