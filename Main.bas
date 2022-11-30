Attribute VB_Name = "Main"
'//**
'*��ĺ���ڰ��ݎ��x�񍐗p�\�쐬
'*�쐬��:Yusaku Suzuki(2022/05/16)
'**//
Option Explicit
Const STANDARD_COL_THIS_YEAR As Long = 11  '// �����񌎂́u���сv�̗�ԍ�
Const STANDARD_COL_LAST_YEAR As Long = 9   '// �O���񌎂́u���сv�̗�ԍ�
Const INTERVAL As Long = 5                 '// �e���̊Ԃ̗�

'/**
 '* �����̃f�[�^��\�ɓ���
'**/
Public Sub importThisYearData()

    If MsgBox("�e�V�[�g��[����]�ɒl����͂��܂��B" & vbLf & "��낵���ł���?", vbQuestion + vbYesNo, ThisWorkbook.Name) = vbNo Then
        Exit Sub
    End If

    Call ImportData(STANDARD_COL_THIS_YEAR)

End Sub

'/**
 '* �O���̃f�[�^��\�ɓ���
'**/
Public Sub importLastYearData()

    If MsgBox("�e�V�[�g��[�O�N�x]�ɒl����͂��܂��B" & vbLf & "��낵���ł���?", vbQuestion + vbYesNo, ThisWorkbook.Name) = vbNo Then
        Exit Sub
    End If

    Call ImportData(STANDARD_COL_LAST_YEAR)

End Sub
'/**
 '* ���C�����[�`��
 '* Money One����o�͂���CSV�f�[�^�̊e�l��\�ɔ��f������
'**/
Sub ImportData(ByVal standardColumn As Long)

    '// �\��t�����\�����������̂��m�F
    If validateFile = False Then: Exit Sub

    '// 1)������i�[�����z����쐬
    Dim arrDiv As Variant: arrDiv = CreateDivArray
    
    '// 2)�V�[�g"���[�N"�̒l�����ꂼ��̃V�[�g�̃Z���ɓ���
    
    Dim lastRow As Long: lastRow = Sheets("���[�N").Cells(Rows.Count, 2).End(xlUp).Row
    Dim lastColumn As Long: lastColumn = Sheets("���[�N").Cells(2, Columns.Count).End(xlToLeft).Column
    Dim counter As Long
    Dim i As Long
    Dim j As Long: j = 2
    Dim k As Long
    
    Dim dicCode As Dictionary
    
    For i = 0 To UBound(arrDiv)
        
        '// 2-1)�R�[�h���i�[�����z����쐬[�R�[�h�ԍ��ˍs�ԍ�]
        Set dicCode = CreateDicCode(arrDiv(i))
        
        '// 2-2)�V�[�g"���[�N"�̒l��Ή�����Z���ɓ���
       '/**
        '* �@ ���喼(�V�[�g"���[�N"��cells(j,1)�̒l)���ς��܂Ń��[�v�������s��
        '* �A �Ȗں���(�V�[�g"���[�N"��cells(j,2)�̒l��dicCode�̃L�[�ɑ��݂�����B�̏������s��
        '* �B ���񌎂���w����Ԃ̍ŏI���܂ł̒l��Ή�����Z���ɓ���
       '**//
        
        Do While Sheets("���[�N").Cells(j, 1).Value = arrDiv(i) '// �@
            
            If dicCode.Exists(Sheets("���[�N").Cells(j, 2).Value) = False Then '// �A
                GoTo Continue
            End If
            
            counter = 0
            
            For k = 6 To lastColumn '// �B
                Sheets(arrDiv(i)).Cells(dicCode(Sheets("���[�N").Cells(j, 2).Value), standardColumn + INTERVAL * counter).Value = Sheets("���[�N").Cells(j, k).Value
                counter = counter + 1
            Next
Continue:
            j = j + 1
        Loop

    Next
    
    Set dicCode = Nothing
    
    MsgBox "���͂��������܂����B", Title:=ThisWorkbook.Name
    
End Sub

'/**
 '* �\��t�����\���K�؂��m�F
 Private Function validateFile() As Boolean

    With Sheets("���[�N").Cells(1, 1)
        If .Value = "����" _
        And .Offset(, 1).Value = "�R�[�h" _
        And .Offset(, 2).Value = "����Ȗ�" _
        And .Offset(, 3).Value = "���ԗ݌v" Then
        
            validateFile = True
            Exit Function
        End If
    End With
    
    MsgBox "�V�[�g�u���[�N�v�ɓ\��t�����\���K�؂ł͂���܂���B", vbExclamation, ThisWorkbook.Name
    
    validateFile = False
        
 End Function
 
 

'//**
'*  ���喼���i�[�����z����쐬
'**//
Private Function CreateDivArray() As Variant

    '// ���喼���i�[����z��
    Dim arrDiv() As String
    
    '//�z��arrDiv�Ɋ��ɒl���o�^����Ă��邩���ׂ�ۂɎg�p����
    Dim arrTarget As Variant
    
    '//�z��̃T�C�Y
    Dim lastRow As Long: lastRow = Sheets("���[�N").Cells(Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    
    '// ���喼�̔z���1�߂̍��ڂ̓V�[�g�u���[�N�v��A2�̒l
    ReDim arrDiv(0)
    arrDiv(0) = Sheets("���[�N").Cells(2, 1).Value
            
    '// ���傪�z��ɂȂ���Γo�^
    For i = 3 To lastRow
        
        arrTarget = Filter(arrDiv, Sheets("���[�N").Cells(i, 1).Value)
        
        If UBound(arrTarget) = -1 Then
            ReDim Preserve arrDiv(UBound(arrDiv) + 1)
            arrDiv(UBound(arrDiv)) = Sheets("���[�N").Cells(i, 1).Value
        End If
    Next

    CreateDivArray = arrDiv
    
End Function

'//**
'*  �e�V�[�g�̉ȖڃR�[�h���i�[�����A�z�z����쐬
'*
'* @param sheetName as String�F�z����쐬����ۂɎQ�Ƃ���V�[�g��
'**//
Private Function CreateDicCode(ByVal sheetName As String) As Dictionary
      
    Dim lastRow As Long: lastRow = Sheets(sheetName).Cells(Rows.Count, 2).End(xlUp).Row
    
    '//�ȖڃR�[�h���i�[����z��[�L�[�F�R�[�h�ԍ��A�l�F��ԍ�]
    Dim dicCode As Dictionary: Set dicCode = New Dictionary
    Dim i As Long
    
    For i = 4 To lastRow
        
        '// �Z���̒l���󔒁A�������͐����łȂ���Ύ��̃��[�v��
        If IsNumeric(Sheets(sheetName).Cells(i, 2).Value) = False Or Sheets(sheetName).Cells(i, 2).Value = "" Then
            GoTo Continue
        End If
        
        '// �Z���̒l���z��̃L�[�ɑ��݂��Ȃ���Δz��ɒǉ�
        If dicCode.Exists(Sheets(sheetName).Cells(i, 2).Value) = False Then
            dicCode.Add Sheets(sheetName).Cells(i, 2).Value, i
        End If
    
Continue:
    Next

    Set CreateDicCode = dicCode
    Set dicCode = Nothing

End Function

