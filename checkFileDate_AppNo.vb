
Dim dateTypeCheck As Boolean
Dim appNoColumnBottom As Range
Dim appFileDateColumnBottom As Range

Sub running()

    Call CheckFileDate
    Call CheckAppNo
    
End Sub

Sub CheckFileDate()

Set fileDateRng = Worksheets("×¨Àû").Range("J:J")

Dim i, rowBottomCount As Integer
Dim tempCellsValue As Variant
    
    rowBottomCount = Worksheets("×¨Àû").Cells(1048576, 10).End(xlUp).Row
        
For i = 3 To rowBottomCount Step 1

    tempCellsValue = fileDateRng.Cells(i, 1)

    Call isDateType(tempCellsValue)
    
    If dateTypeCheck = False Then
    
        fileDateRng.Cells(i, 1).Interior.Color = RGB(96, 96, 96)
        
    End If
    
Next i

End Sub
Sub isDateType(applicationDate)

    dateTypeCheck = False

    If IsDate(applicationDate) Then

        dateTypeCheck = True
        
    Else: dateTypeCheck = False
    
    End If
    
End Sub

Private Function PatentAppNoCheckBit(applicationNumber)

    Dim LengthOfAppNumber As Byte
    Dim ArrAppNo()
    Dim sum, i As Integer
    
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
            
        PatentAppNoCheckBit = sum Mod 11
        
End Function

Private Sub CheckAppNo()

On Error Resume Next

Application.ScreenUpdating = False
                            
Set appNoRng = Worksheets("×¨Àû").Range("K:K")

Dim checkBitValue As Byte
Dim farRightAppNo As Variant
Dim appNoColumnRowCount As Integer


appNoColumnRowCount = Worksheets("×¨Àû").Range("K1048576").End(xlUp).Offset(1, 0).Row

For i = 3 To appNoColumnRowCount Step 1

    checkBitValue = PatentAppNoCheckBit(appNoRng.Cells(i, 1))
    
    appNoRng.Cells(i, 1) = Replace(appNoRng.Cells(i, 1), " ", "")
    
    farRightAppNo = Right(appNoRng.Cells(i, 1), 1)
    
    If farRightAppNo = "X" Then
    
        If checkBitValue < 10 Then
        
            appNoRng.Cells(i, 1).Interior.Color = RGB(96, 96, 96)
        
         End If
        
    Else
    
        farRightAppNo = CInt(farRightAppNo)
        
        If Not (farRightAppNo = checkBitValue) Then
        
            appNoRng.Cells(i, 1).Interior.Color = RGB(96, 96, 96)
            
        End If
    
    End If
    
Next i

Application.ScreenUpdating = True

End Sub
