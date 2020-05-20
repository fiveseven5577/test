Attribute VB_Name = "模块1"
Option Explicit
Dim appNoRng As Range
Dim appNoRng_Arr

Sub Main_AppNoDetection()

    Dim i, j, k, l, m, n As Integer
    Dim tempResult As Byte
    
    On Error Resume Next

    Call CreateAppNoRng
    
    For i = 1 To UBound(appNoRng_Arr)
    
        appNoRng.Cells(i + 2, 1).Select
    
        tempResult = CheckBit_FarRightAppNo(appNoRng_Arr(i))
        
        If tempResult = 77 Then
            Call SetCellFormat(i + 2)
        End If
        
    Next i
        
End Sub
Sub SetCellFormat(index)
            
            appNoRng.Cells(index, 1).Interior.Color = RGB(255, 0, 0)

End Sub
Function PatentAppNoCheckBit(applicationNumber)     '计算专利申请号的数据的正确校验位
    
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
    
    PatentAppNoCheckBit = sum Mod 11
    
End Function

Sub CreateAppNoRng()    '创建一个区域对象，该区域中有专利申请号

    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    Set appNoRng = Worksheets("专利").Range("L:L")  '生成专利申请号的区域对象，在本表格中存储专利申请号的区域为Range("K:K")
    
    Dim appNoRng_Row As Integer
    Dim checkBitidentifier As Byte
    
    appNoRng_Row = Worksheets("专利").Range("K1048576").End(xlUp).Offset(1, 0).Row
    
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
            CheckBit_FarRightAppNo = 77
            Exit Function
            
    End If
    
    '13位专利申请号的前12位均为数字（最后一位校验位可能为X），12位的数据类型为Double，如果不是Double数据类型，那么肯定不是专利申请号
    
    If Not (IsNumeric(Left(unmodifiedPattentAppNo, 12))) Then
        CheckBit_FarRightAppNo = 77
        Exit Function
    End If
    
    
    FarRightAppNo = Right(unmodifiedPattentAppNo, 1) '获取专利申请号真实的校验位，用于和checkBitValue
    
    unmodifiedPattentAppNo = Replace(Replace(unmodifiedPattentAppNo, ".", ""), " ", "")
    
    'unmodifiedPattentAppNo = Replace(unmodifiedPattentAppNo, " ", "")
    'unmodifiedPattentAppNo means 在专利申请号中可能存在句点、空格等错误文本
    
    checkBitValue = PatentAppNoCheckBit(unmodifiedPattentAppNo) '获取专利申请号的理论上正确的校验位
    
    If FarRightAppNo = "X" Then
        
        If Not (checkBitValue = 10) Then
            
            CheckBit_FarRightAppNo = 77  'CheckBit_FarRightAppNo设置为77，表示专利申请号的最后一位不是X，相当于CheckBit_FarRightAppNo来标记错误的校验位
            'CheckBit_FarRightAppNo也就是传说中的Magic Number
            Exit Function
        End If
    Else
        
        If Not (FarRightAppNo = checkBitValue) Then
            
            CheckBit_FarRightAppNo = 77
            
        End If
        
    End If


        'FarRightAppNo = CInt(FarRightAppNo) 'farRightAppNo变量的数据类型是变体形，这里强制转换成整型，以便在if中进行判断

End Function
Sub CheckFileDate()   '检查申请日的日期格式是否正确
    
    Set fileDateRng = Worksheets("专利").Range("J:J")
    
    Dim i, rowBottomCount As Integer
    Dim tempCellsValue As Variant
    
    rowBottomCount = Worksheets("专利").Cells(1048576, 10).End(xlUp).Row
    
    For i = 3 To rowBottomCount Step 1
        
        tempCellsValue = fileDateRng.Cells(i, 1)
        
        Call isDateType(tempCellsValue)
        
        If dateTypeCheck = False Then
            
            'fileDateRng.Cells(i, 1).Interior.Color = RGB(192, 192, 192)
            
        End If
        
    Next i
    
End Sub
Sub isDateType(applicationDate)     '判断数据是否为日期类型
    
    Dim dateTypeCheck As Boolean
    
    
    dateTypeCheck = False
    
    If IsDate(applicationDate) Then
        
        dateTypeCheck = True
        
        Else: dateTypeCheck = False
        
    End If
    
End Sub
Sub test09909()
Myvar1 = "45"
MyVar2 = "55"
MyVar = Myvar1 + MyVar2
myCheck = IsNumeric(MyVar)
Debug.Print myCheck
End Sub
