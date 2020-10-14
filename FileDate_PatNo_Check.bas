Attribute VB_Name = "FileDate_PatNo_Check"
Option Explicit
Dim appNoRng As Range
Dim appNoRng_Arr
Const incorrectPatentNo = 77
Const incorrectDateFormat = 44
Const correctDateFormat = 66
Dim fileDateCheckTab As Byte
Dim filedaterng As Range
Dim incorrectPatNoValue()  '建立变体型数组，存储错误的专利申请号
Dim blankCounts As Integer

Sub Main_AppNoDetection()
'主函数，程序入口

    Dim i_interation, j_interation, k, l, m, n As Integer
    Dim tempResult As Byte  '用于存储CheckBit_FarRightAppNo函数的临时结果

    Call CreateAppNoRng
    Call CreateText
    j_interation = 1
    ReDim incorrectPatNoValue(1 To UBound(appNoRng_Arr))  'incorrectPatNoValue数组的元素个数不可能比appNoRng_Arr还多，所以用UBound(appNoRng_Arr)的上限足够
    
    For i_interation = 1 To UBound(appNoRng_Arr)
    
        appNoRng.Cells(i_interation + 2, 1).Select
        tempResult = CheckBit_FarRightAppNo(appNoRng_Arr(i_interation))
        If tempResult = incorrectPatentNo Then
        
            Call SetPatAppCellFormat(i_interation + 2)
            
            incorrectPatNoValue(j_interation) = appNoRng_Arr(i_interation)
            

            Call AppendFile(incorrectPatNoValue(j_interation), appNoRng.Cells(i_interation + 2, 1).Address, appNoRng.Cells(i_interation + 2, 1).Offset(0, -4))
            
            j_interation = j_interation + 1
            
        End If
        
    Next i_interation
    
    Call BlankValueInPatentNumbers
    
    
    'Call FiledateCheck  '调用申请日“专列”区域的测试函数
    
End Sub
Sub BlankValueInPatentNumbers()
'找到表格中专利列的空值并返回行号
    Dim singleCell, appNoRng1 As Range
    Dim appNoRng_RowCounts, blankCounts As Integer
    
    appNoRng_RowCounts = Worksheets("专利").Range("K1048576").End(xlUp).Offset(1, 0).Row
    '建立一个区域对象，该区域对象是手工输入的专利申请号列
    '与其他sub中的名称相同，但作用域不同，互不干扰。
    
    Set appNoRng1 = Worksheets("专利").Range("L1", "L" & CStr(appNoRng_RowCounts))
    'Range对象中的参数ReferenceStyle:=xlA1时，参数都必须是文本格式，所以需要用到CStr函数转换数据类型
    
        For Each singleCell In appNoRng1
        
            If singleCell = "" Then
                blankCounts = blankCounts + 1
            End If
            
        Next
        
        Debug.Print "There are " & blankCounts & " empty cells in PatentNoColumns"
        
        blankCounts = 0
    
End Sub

Sub SetPatAppCellFormat(cellRowIndex)

    appNoRng.Cells(cellRowIndex, 1).Interior.Color = RGB(255, 0, 0)

End Sub
Function TheoreticalCheckBit(applicationNumber)
'在输入值正确的前提下，计算专利申请号的理论正确值

    
    Dim LengthOfAppNumber As Byte
    Dim ArrAppNo()
    Dim sum, i As Integer
    Dim CheckBitSequence()
    
    i = 1
    sum = 0
    applicationNumber = Replace(Replace(applicationNumber, " ", ""), ".", "")  'delete ""  "/" in applicationNumber
    
    CheckBitSequence = Array(2, 3, 4, 5, 6, 7, 8, 9, 2, 3, 4, 5)
    
    LengthOfAppNumber = Len(applicationNumber) - 1
    
    ReDim ArrAppNo(1 To LengthOfAppNumber)
    
    For i = 1 To LengthOfAppNumber Step 1
        
        ArrAppNo(i) = CInt(Mid(applicationNumber, i, 1))
        sum = sum + ArrAppNo(i) * CheckBitSequence(i - 1)
        
    Next i
    
    '专利申请号校验位算法:
    '比如 :200710308494.X
    '从第1位到第12位数字依次以下列变量代表: X4 , X3, X2, X1, Y, Z7, Z6, Z5, Z4, Z3, Z2, Z1?
    '校验位的最终计算方法为:
    '(X4*2+X3*3+X2*4+X1*5+Y*6+Z7*7+Z6*8+Z5*9+Z4*2+Z3*3+Z2*4+Z1*5)MOD(11)
    '其中: 余数10对应的为校验位为X
    
    TheoreticalCheckBit = sum Mod 11
    
End Function

Sub CreateAppNoRng()    ''建立一个区域对象，该区域对象是手工输入的专利申请号列，未知输入值是否正确

    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    Set appNoRng = Worksheets("专利").Range("L:L")  '生成专利申请号的区域对象，在本表格中存储专利申请号的区域为Range("K:K")
    
    Dim appNoRng_Row As Integer
    Dim checkBitidentifier As Byte
    
    appNoRng_Row = Worksheets("专利").Range("L1048576").End(xlUp).Offset(1, 0).Row
    
    ReDim appNoRng_Arr(appNoRng_Row)
    
    For i = 1 To appNoRng_Row
        appNoRng_Arr(i) = appNoRng.Cells(i + 2, 1)
        'Debug.Print appNoRng_Arr(i)
    Next i
    
    Application.ScreenUpdating = True
    
End Sub
Function CheckBit_FarRightAppNo(unmodifiedPattentAppNo) '输入的数据未知正确与否，判断其校验位是否符合校验位规则
                                                        '把可能存在的错误情况穷举一下
    Dim checkBitValue As Byte
    Dim FarRightAppNo As Variant
    Dim appNoLength As Integer
    Dim Right12 As Variant
    
    CheckBit_FarRightAppNo = 0
    
    appNoLength = Len(unmodifiedPattentAppNo)
    
    If (appNoLength > 15 Or appNoLength < 8) Then   '专利申请号的位数为8位（2013年以前）或13位，考虑到有句点.的存在，最多也就是9位或14位
                                                    '所以小于8位或大于9位的数据不可能是专利申请号，可以直接标记为77并退出本函数,都不用去计算理论校验位
            CheckBit_FarRightAppNo = incorrectPatentNo
            Exit Function
            
    End If
    
    '13位专利申请号的前12位均为数字（最后一位校验位可能为X），12位的数据类型为Double，如果不是Double数据类型，那么肯定不是专利申请号
    
    If Not (IsNumeric(Left(unmodifiedPattentAppNo, 12))) Then
        CheckBit_FarRightAppNo = incorrectPatentNo
        Exit Function
    End If
    
    
    FarRightAppNo = Right(unmodifiedPattentAppNo, 1) '获取专利申请号真实的校验位，用于和checkBitValue
    
    unmodifiedPattentAppNo = Replace(Replace(unmodifiedPattentAppNo, ".", ""), " ", "")
    
    'unmodifiedPattentAppNo = Replace(unmodifiedPattentAppNo, " ", "")
    'unmodifiedPattentAppNo means 在专利申请号中可能存在句点、空格等错误文本
    
    checkBitValue = TheoreticalCheckBit(unmodifiedPattentAppNo) '获取专利申请号的理论上正确的校验位
    
    If FarRightAppNo = "X" Then
        
        If Not (checkBitValue = 10) Then
            
            CheckBit_FarRightAppNo = incorrectPatentNo
            'CheckBit_FarRightAppNo设置为inCorrectPatentNo=77(魔法数），表示专利申请号的最后一位不是X，相当于CheckBit_FarRightAppNo来标记错误的校验位
            'CheckBit_FarRightAppNo也就是传说中的Magic Number
            Exit Function
        End If
    Else
        
        If Not (FarRightAppNo = checkBitValue) Then
            
            CheckBit_FarRightAppNo = incorrectPatentNo
            
        End If
        
    End If

        'FarRightAppNo = CInt(FarRightAppNo) 'farRightAppNo变量的数据类型是变体型，这里强制转换成整型，以便在if中进行判断

End Function
Sub FiledateCheck() '检查申请日“专列”区域内的单元格数值的数据类型是否为日期型
    
    Set filedaterng = Worksheets("专利").Range("J:J")  '建立申请日“专列”区域
    
    Dim i_fileDateRng_ArrRow, rowBottomCount As Integer
    Dim tempCellsValue As Variant
    
    rowBottomCount = Worksheets("专利").Range("J1048576").End(xlUp).Offset(1, 0).Row
    
    For i_fileDateRng_ArrRow = 3 To rowBottomCount Step 1
        
        filedaterng.Cells(i_fileDateRng_ArrRow, 1).Select
        tempCellsValue = filedaterng.Cells(i_fileDateRng_ArrRow, 1)
        
        Call isDateType(tempCellsValue)
        
        If fileDateCheckTab = incorrectDateFormat Then
            
            Call SetFileDateCellFormat(i_fileDateRng_ArrRow)
            
        End If
        
        fileDateCheckTab = 0
        
    Next i_fileDateRng_ArrRow
    
End Sub
Sub isDateType(applicationDate)     '判断数据是否为日期类型
    
    If IsDate(applicationDate) Then     'correctDateFormat=66, incorrectDateFormat=44
        
        fileDateCheckTab = correctDateFormat
        
        Else: fileDateCheckTab = incorrectDateFormat
        
    End If
    
End Sub
Sub SetFileDateCellFormat(cellRowIndex)

    filedaterng.Cells(cellRowIndex, 1).Interior.Color = RGB(169, 169, 169)

End Sub

Sub CreateText()
On Error Resume Next
    
    Dim intNum As Integer
    intNum = FreeFile
    'Open ThisWorkbook.Path & "\test.txt" For Output As #intNum
    Open ThisWorkbook.Path & "\" & ThisWorkbook.Name For Output As #intNum
    Write #intNum, "Incorrect patent application numbers :"
    
    Close intNum
End Sub
Sub AppendFile(PatentAppNo, CellAddress, PatentName)
    Dim intNum As Integer
    intNum = FreeFile
    Open ThisWorkbook.Path & "\test.txt" For Append As #intNum
    Print #intNum, "*************************"
    Print #intNum, PatentAppNo; "  "; CellAddress; "  "; PatentName
    Close intNum
End Sub
