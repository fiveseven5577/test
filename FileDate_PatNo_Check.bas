Attribute VB_Name = "FileDate_PatNo_Check"
Option Explicit
Dim appNoRng As Range
Dim appNoRng_Arr
Const incorrectPatentNo = 77
Const incorrectDateFormat = 44
Const correctDateFormat = 66
Dim fileDateCheckTab As Byte
Dim filedaterng As Range
Dim incorrectPatNoValue()  '�������������飬�洢�����ר�������
Dim blankCounts As Integer

Sub Main_AppNoDetection()
'���������������

    Dim i_interation, j_interation, k, l, m, n As Integer
    Dim tempResult As Byte  '���ڴ洢CheckBit_FarRightAppNo��������ʱ���

    Call CreateAppNoRng
    Call CreateText
    j_interation = 1
    ReDim incorrectPatNoValue(1 To UBound(appNoRng_Arr))  'incorrectPatNoValue�����Ԫ�ظ��������ܱ�appNoRng_Arr���࣬������UBound(appNoRng_Arr)�������㹻
    
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
    
    
    'Call FiledateCheck  '���������ա�ר�С�����Ĳ��Ժ���
    
End Sub
Sub BlankValueInPatentNumbers()
'�ҵ������ר���еĿ�ֵ�������к�
    Dim singleCell, appNoRng1 As Range
    Dim appNoRng_RowCounts, blankCounts As Integer
    
    appNoRng_RowCounts = Worksheets("ר��").Range("K1048576").End(xlUp).Offset(1, 0).Row
    '����һ��������󣬸�����������ֹ������ר���������
    '������sub�е�������ͬ����������ͬ���������š�
    
    Set appNoRng1 = Worksheets("ר��").Range("L1", "L" & CStr(appNoRng_RowCounts))
    'Range�����еĲ���ReferenceStyle:=xlA1ʱ���������������ı���ʽ��������Ҫ�õ�CStr����ת����������
    
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
'������ֵ��ȷ��ǰ���£�����ר������ŵ�������ȷֵ

    
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
    
    'ר�������У��λ�㷨:
    '���� :200710308494.X
    '�ӵ�1λ����12λ�������������б�������: X4 , X3, X2, X1, Y, Z7, Z6, Z5, Z4, Z3, Z2, Z1?
    'У��λ�����ռ��㷽��Ϊ:
    '(X4*2+X3*3+X2*4+X1*5+Y*6+Z7*7+Z6*8+Z5*9+Z4*2+Z3*3+Z2*4+Z1*5)MOD(11)
    '����: ����10��Ӧ��ΪУ��λΪX
    
    TheoreticalCheckBit = sum Mod 11
    
End Function

Sub CreateAppNoRng()    ''����һ��������󣬸�����������ֹ������ר��������У�δ֪����ֵ�Ƿ���ȷ

    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    Set appNoRng = Worksheets("ר��").Range("L:L")  '����ר������ŵ���������ڱ�����д洢ר������ŵ�����ΪRange("K:K")
    
    Dim appNoRng_Row As Integer
    Dim checkBitidentifier As Byte
    
    appNoRng_Row = Worksheets("ר��").Range("L1048576").End(xlUp).Offset(1, 0).Row
    
    ReDim appNoRng_Arr(appNoRng_Row)
    
    For i = 1 To appNoRng_Row
        appNoRng_Arr(i) = appNoRng.Cells(i + 2, 1)
        'Debug.Print appNoRng_Arr(i)
    Next i
    
    Application.ScreenUpdating = True
    
End Sub
Function CheckBit_FarRightAppNo(unmodifiedPattentAppNo) '���������δ֪��ȷ����ж���У��λ�Ƿ����У��λ����
                                                        '�ѿ��ܴ��ڵĴ���������һ��
    Dim checkBitValue As Byte
    Dim FarRightAppNo As Variant
    Dim appNoLength As Integer
    Dim Right12 As Variant
    
    CheckBit_FarRightAppNo = 0
    
    appNoLength = Len(unmodifiedPattentAppNo)
    
    If (appNoLength > 15 Or appNoLength < 8) Then   'ר������ŵ�λ��Ϊ8λ��2013����ǰ����13λ�����ǵ��о��.�Ĵ��ڣ����Ҳ����9λ��14λ
                                                    '����С��8λ�����9λ�����ݲ�������ר������ţ�����ֱ�ӱ��Ϊ77���˳�������,������ȥ��������У��λ
            CheckBit_FarRightAppNo = incorrectPatentNo
            Exit Function
            
    End If
    
    '13λר������ŵ�ǰ12λ��Ϊ���֣����һλУ��λ����ΪX����12λ����������ΪDouble���������Double�������ͣ���ô�϶�����ר�������
    
    If Not (IsNumeric(Left(unmodifiedPattentAppNo, 12))) Then
        CheckBit_FarRightAppNo = incorrectPatentNo
        Exit Function
    End If
    
    
    FarRightAppNo = Right(unmodifiedPattentAppNo, 1) '��ȡר���������ʵ��У��λ�����ں�checkBitValue
    
    unmodifiedPattentAppNo = Replace(Replace(unmodifiedPattentAppNo, ".", ""), " ", "")
    
    'unmodifiedPattentAppNo = Replace(unmodifiedPattentAppNo, " ", "")
    'unmodifiedPattentAppNo means ��ר��������п��ܴ��ھ�㡢�ո�ȴ����ı�
    
    checkBitValue = TheoreticalCheckBit(unmodifiedPattentAppNo) '��ȡר������ŵ���������ȷ��У��λ
    
    If FarRightAppNo = "X" Then
        
        If Not (checkBitValue = 10) Then
            
            CheckBit_FarRightAppNo = incorrectPatentNo
            'CheckBit_FarRightAppNo����ΪinCorrectPatentNo=77(ħ����������ʾר������ŵ����һλ����X���൱��CheckBit_FarRightAppNo����Ǵ����У��λ
            'CheckBit_FarRightAppNoҲ���Ǵ�˵�е�Magic Number
            Exit Function
        End If
    Else
        
        If Not (FarRightAppNo = checkBitValue) Then
            
            CheckBit_FarRightAppNo = incorrectPatentNo
            
        End If
        
    End If

        'FarRightAppNo = CInt(FarRightAppNo) 'farRightAppNo���������������Ǳ����ͣ�����ǿ��ת�������ͣ��Ա���if�н����ж�

End Function
Sub FiledateCheck() '��������ա�ר�С������ڵĵ�Ԫ����ֵ�����������Ƿ�Ϊ������
    
    Set filedaterng = Worksheets("ר��").Range("J:J")  '���������ա�ר�С�����
    
    Dim i_fileDateRng_ArrRow, rowBottomCount As Integer
    Dim tempCellsValue As Variant
    
    rowBottomCount = Worksheets("ר��").Range("J1048576").End(xlUp).Offset(1, 0).Row
    
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
Sub isDateType(applicationDate)     '�ж������Ƿ�Ϊ��������
    
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
