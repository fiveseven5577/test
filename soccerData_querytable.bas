Attribute VB_Name = "soccerData_querytable"
Dim connectString As String
Dim tableRows, tableColumns As Byte
Dim dataInQuery()
Dim URLlist_XMLHttp As Variant
Dim URLCounts As Byte
Dim seasonYears As Variant
Sub CallModules()
'This is MainEntry

Dim i, j, k As Byte
Dim seasonYear As String

Call XmlHttpData

Worksheets(1).Activate
Worksheets(1).Cells.Clear

For i = 1 To URLCounts
    seasonYear = seasonYears(i)
    connectString = URLlist_XMLHttp(i)
    Debug.Print connectString
    Call dataInOkooo(connectString, seasonYear)
    
    Application.StatusBar = GetProgress(i, URLCounts)
    
Next i

End Sub
Sub dataInOkooo(constr As String, whichYear As String)
Dim i, j, k As Integer
Dim webtableFour As Object
On Error Resume Next

Set shFirstQtr = Workbooks(1).Worksheets(1)
    
    'Set webtableFour = Worksheets(1).QueryTables.Add(Connection:="url;" & "http://www.okooo.com/soccer/league/17/", _
     '               Destination:=ThisWorkbook.Sheets(1).Range("A1048576").End(xlUp).Offset(1, 0))
     
     Set webtableFour = Worksheets(1).QueryTables.Add(Connection:="url;" & constr, _
                    Destination:=Worksheets(1).Range("A1048576").End(xlUp).Offset(1, 1))
     
    With webtableFour
    
        .WebTables = "4"
        .FieldNames = True
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .Refresh BackgroundQuery:=False
        .ResultRange.Cells(1, 1).EntireRow.Delete
        tableRows = .ResultRange.rows.Count
        tableColumns = .ResultRange.Columns.Count

    End With
                Do
                    DoEvents
                Loop While (webtableFour.Refreshing)
    
                For i = 1 To tableRows          'resize the cells
                
                        webtableFour.ResultRange.NumberFormat = "General"
                        webtableFour.ResultRange.VerticalAlignment = xlCenter
                        webtableFour.ResultRange.HorizontalAlignment = xlCenter
                        webtableFour.ResultRange.Cells(i, 1).Offset(0, -1) = whichYear

                Next i
        
                For i = 1 To tableColumns       'contents fitted with format of cells
                
                        webtableFour.ResultRange.Columns.AutoFit
                        webtableFour.ResultRange.Cells(i, 1).Offset(0, -1).AutoFit
                        
                Next i
                
                webtableFour.ResultRange.Cells(1, 1).Offset(0, -1).ClearContents
                webtableFour.ResultRange.Cells(1, 1).Offset(0, -1).EntireRow.Interior.ColorIndex = 37
    
                ReDim dataInQuery(1 To tableRows - 2, 1 To tableColumns - 2) As Variant
        
                'Create a datatale(dataInQuery) to store the data from the webtable4,
                'cause data in array will run in a faster speed than directly run in the cells
    
                For i = 1 To UBound(dataInQuery, 1)
        
                    For j = 1 To UBound(dataInQuery, 2)

                        dataInQuery(i, j) = webtableFour.ResultRange.Cells(i + 2, j + 3).Value

                    Next j
                    
                Next i
                
                Set webtableFour = Nothing
                
End Sub
Function GetProgress(curValue, maxValue)   'display the progress with alternative black and white square
    Dim blackSquare, whiteSquare As String
    blackSquare = Application.Rept("��", curValue)
    whiteSquare = Application.Rept("��", maxValue - curValue)
    GetProgress = blackSquare & whiteSquare & FormatNumber(curValue / maxValue * 100, 2) & "%"
End Function
Function SuccessRate(won As Integer, draw As Integer, lose As Integer)

        SuccessRate = (won + draw * 0.5) / lose
    
End Function

Sub XmlHttpData()
    
    Dim xmlHttp, hrefURL, HTML As Object
    Dim LinkString As String
    Dim oDom, oDom1 As Object
    Dim i As Byte
    Dim temp As String
    
    LinkString = "http://www.okooo.com/soccer/league/8/schedule/14011/"
    
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")

    xmlHttp.Open "GET", LinkString, False
    xmlHttp.send

        Do While xmlHttp.readyState <> 4
            DoEvents
        Loop
        
    Set HTML = CreateObject("htmlfile")
    
    HTML.body.innerHTML = StrConv(xmlHttp.responseBody, vbUnicode)
    Set oDom = HTML.body.getElementsByTagName("ul")(2)   'Item(2)
    Set oDom1 = oDom.getElementsByTagName("a")

    ReDim URLlist_XMLHttp(1 To oDom1.length)
    ReDim seasonYears(1 To oDom1.length)
    
     URLCounts = oDom1.length
    
    For i = 1 To URLCounts
        
        URLlist_XMLHttp(i) = oDom1.Item(i - 1)
        seasonYears(i) = oDom1.Item(i - 1).innerText
        temp = URLlist_XMLHttp(i)
        temp = "http://www.okooo.com" & Replace(temp, "about:", "") & "/1"  'turn RelativePath to AbsolutePath
        URLlist_XMLHttp(i) = temp
        Debug.Print URLlist_XMLHttp(i)
        seasonYear = oDom1.Item(i - 1).innerText
    Next i

End Sub
    
Sub FourDigitYear(twoDigitYears As Variant) '��δ���

    Dim beyondYearReg, beyondYearMatches As Object
    Dim i As Byte
    Dim objTwoDigitYear, objTwoDigitYears As Object
    ReDim twoDigitYears(1 To URLCounts)
    
    For i = 1 To URLCounts
        twoDigitYears(i) = seasonYears(i)
    Next i
    
        Set YearsFormat = CreateObject("vbscript.regexp")
            With YearsFormat
                .Pattern = "\d{2,}/\d{2,}"
                .Global = True
                If .test(twoDigitYear) Then
                    Set objTwoDigitYear = .Execute(twoDigitYear)
                    
                    
                    
                    beyondYear = True
                    If Left(beyondYearMatches.Item(0).Value, 2) <= 50 Then
                        previousYear = Left(beyondYearMatches.Item(0).Value, 2) + 2000
                    Else
                        previousYear = Left(beyondYearMatches.Item(0).Value, 2) + 1900
                    End If
                    
                    If Right(beyondYearMatches.Item(0).Value, 2) <= 50 Then
                        nextYear = Right(beyondYearMatches.Item(0).Value, 2) + 2000
                    Else
                        nextYear = Right(beyondYearMatches.Item(0).Value, 2) + 1900
                    End If
                    
                Else
                    .Pattern = "\d{4,}"
                    Set beyondYearMatches = .Execute(seasonInTitle)
                    beyondYear = False
                    seasonYear = beyondYearMatches.Item(0).Value
                End If
            End With

End Sub
