Attribute VB_Name = "ģ��1"
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
Function PatentAppNoCheckBit(applicationNumber)     '����ר������ŵ����ݵ���ȷУ��λ
    
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
    
    PatentAppNoCheckBit = sum Mod 11
    
End Function

Sub CreateAppNoRng()    '����һ��������󣬸���������ר�������

    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    Set appNoRng = Worksheets("ר��").Range("L:L")  '����ר������ŵ���������ڱ�����д洢ר������ŵ�����ΪRange("K:K")
    
    Dim appNoRng_Row As Integer
    Dim checkBitidentifier As Byte
    
    appNoRng_Row = Worksheets("ר��").Range("K1048576").End(xlUp).Offset(1, 0).Row
    
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
            CheckBit_FarRightAppNo = 77
            Exit Function
            
    End If
    
    '13λר������ŵ�ǰ12λ��Ϊ���֣����һλУ��λ����ΪX����12λ����������ΪDouble���������Double�������ͣ���ô�϶�����ר�������
    
    If Not (IsNumeric(Left(unmodifiedPattentAppNo, 12))) Then
        CheckBit_FarRightAppNo = 77
        Exit Function
    End If
    
    
    FarRightAppNo = Right(unmodifiedPattentAppNo, 1) '��ȡר���������ʵ��У��λ�����ں�checkBitValue
    
    unmodifiedPattentAppNo = Replace(Replace(unmodifiedPattentAppNo, ".", ""), " ", "")
    
    'unmodifiedPattentAppNo = Replace(unmodifiedPattentAppNo, " ", "")
    'unmodifiedPattentAppNo means ��ר��������п��ܴ��ھ�㡢�ո�ȴ����ı�
    
    checkBitValue = PatentAppNoCheckBit(unmodifiedPattentAppNo) '��ȡר������ŵ���������ȷ��У��λ
    
    If FarRightAppNo = "X" Then
        
        If Not (checkBitValue = 10) Then
            
            CheckBit_FarRightAppNo = 77  'CheckBit_FarRightAppNo����Ϊ77����ʾר������ŵ����һλ����X���൱��CheckBit_FarRightAppNo����Ǵ����У��λ
            'CheckBit_FarRightAppNoҲ���Ǵ�˵�е�Magic Number
            Exit Function
        End If
    Else
        
        If Not (FarRightAppNo = checkBitValue) Then
            
            CheckBit_FarRightAppNo = 77
            
        End If
        
    End If


        'FarRightAppNo = CInt(FarRightAppNo) 'farRightAppNo���������������Ǳ����Σ�����ǿ��ת�������ͣ��Ա���if�н����ж�

End Function
Sub CheckFileDate()   '��������յ����ڸ�ʽ�Ƿ���ȷ
    
    Set fileDateRng = Worksheets("ר��").Range("J:J")
    
    Dim i, rowBottomCount As Integer
    Dim tempCellsValue As Variant
    
    rowBottomCount = Worksheets("ר��").Cells(1048576, 10).End(xlUp).Row
    
    For i = 3 To rowBottomCount Step 1
        
        tempCellsValue = fileDateRng.Cells(i, 1)
        
        Call isDateType(tempCellsValue)
        
        If dateTypeCheck = False Then
            
            'fileDateRng.Cells(i, 1).Interior.Color = RGB(192, 192, 192)
            
        End If
        
    Next i
    
End Sub
Sub isDateType(applicationDate)     '�ж������Ƿ�Ϊ��������
    
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
