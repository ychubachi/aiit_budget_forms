Attribute VB_Name = "Module2"
Sub ������{�^��_Click()
    Worksheets("������").Activate
End Sub

Sub ����{�^��_Click()
    Worksheets("����").Activate
End Sub

Sub �o���{�^��_Click()
    Worksheets("�o��").Activate
End Sub

Sub �`�[����{�^��_Click()
    Dim myNum As Variant
    Dim oval As Excel.Shape
    Dim ovalX, ovalY, ovalW, ovalH As Integer
    Dim oval2 As Excel.Shape
    Dim oval2X, oval2Y, oval2W, oval2H As Integer

    myNum = Application.InputBox("�������`�[��No����͂��Ă�������")
    
    If myNum <> False Then
        ' �`�[��No��]�L����
        Worksheets("�������p").Range("�`�[No").Value = myNum
        
        ' ����̎�ނ𔻒肷��
        Dim �`�[��� As Variant
        �`�[��� = Worksheets("�������p").Range("�`�[���").Value
        
        
        Select Case �`�[���
        Case "����"
            ' ���֕����F�͂̏ꍇ
            Worksheets("���֕����F��").Activate
            With ActiveSheet
                ' �\������
                .Visible = True
                ' �������ʂ��܂�ň͂�
                With Worksheets("�������p")
                    ovalX = CInt(.Range("���֗p������敪���W").Item(1).Value)
                    ovalY = CInt(.Range("���֗p������敪���W").Item(2).Value)
                    ovalW = CInt(.Range("���֗p������敪���W").Item(3).Value)
                    ovalH = CInt(.Range("���֗p������敪���W").Item(4).Value)
                End With
                            
                Set oval = .Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, ovalX, ovalY, ovalW, ovalH)
                oval.Fill.Transparency = 1#
                
                ' ���֕����R���܂�ň͂�
                With Worksheets("�������p")
                    
                    oval2X = CInt(.Range("���֗p���R�敪���W").Item(1).Value)
                    oval2Y = CInt(.Range("���֗p���R�敪���W").Item(2).Value)
                    oval2W = CInt(.Range("���֗p���R�敪���W").Item(3).Value)
                    oval2H = CInt(.Range("���֗p���R�敪���W").Item(4).Value)
                End With
                            
                Set oval2 = .Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, oval2X, oval2Y, oval2W, oval2H)
                oval2.Fill.Transparency = 1#
                
                ' PDF�t�@�C���ɏo�͂���
                PDF�t�@�C���o�� "No" & myNum & "-���֕����F��.pdf"
                
                ' �܂���폜����
                oval.Delete
                oval2.Delete
                
                ' ��\���ɂ��ǂ�
                .Visible = False
            End With
            ' ����V�[�g�ɖ߂�
            Worksheets("����").Activate
        Case "����"
            ' �������ʒm���̏ꍇ
            Worksheets("������񓙒ʒm��").Activate
            With ActiveSheet
                ' �\������
                .Visible = True
                ' �������ʂ��܂�ň͂�
                With Worksheets("�������p")
                    ovalX = CInt(.Range("�����p������敪���W").Item(1).Value)
                    ovalY = CInt(.Range("�����p������敪���W").Item(2).Value)
                    ovalW = CInt(.Range("�����p������敪���W").Item(3).Value)
                    ovalH = CInt(.Range("�����p������敪���W").Item(4).Value)
                End With
                            
                Set oval = .Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, ovalX, ovalY, ovalW, ovalH)
                oval.Fill.Transparency = 1#
                
                ' PDF�t�@�C���ɏo�͂���
                PDF�t�@�C���o�� "No" & myNum & "-�������ʒm��.pdf"
                
                
                ' �܂���폜����
                oval.Delete
                
                ' ��\���ɂ��ǂ�
                .Visible = False
            End With
            ' ����V�[�g�ɖ߂�
            Worksheets("����").Activate
        Case "����"
            ' �o���̏ꍇ�A�o���V�[�g�ɑJ�ڂ���
            MsgBox "�o���V�[�g���������Ă�������"
            Worksheets("�o��").Activate
        Case Else
            MsgBox "�w�肳�ꂽ�`�[�̏o�͂ɂ͑Ή����Ă���܂���"
        End Select
           
    End If
End Sub


Sub ���ߕ���󏑃{�^��_Click()
    Dim myNum As Variant
    Dim area As Variant
    Dim oval As Excel.Shape
    Dim ovalX, ovalY, ovalW, ovalH As Integer

    myNum = Application.InputBox("�������`�[��No����͂��Ă�������")
    ' myNum = 1
    
    If myNum <> False Then
        With Worksheets("�������p�i����j")
            ' �`�[��No���V�[�g�ɓ]�L���܂�
            .Range("���sNo").Value = myNum
        
            ' �������C�O�����擾���܂�
            area = .Range("���O").Value
        End With
        
        Select Case area
        Case "����"
            ' ���s���ߕ�
            Worksheets("���s���ߕ�").Activate
            With ActiveSheet
                ' �\������
                .Visible = True
    
                ' PDF�t�@�C���ɏo�͂���
                PDF�t�@�C���o�� "No" & myNum & "-���s���ߕ�.pdf"
    
                ' ��\���ɂ��ǂ�
                .Visible = False
            End With
            
            ' ����v�Z����
            Worksheets("����v�Z����").Activate
            With ActiveSheet
                ' �\������
                .Visible = True
                
                ' �������ʂ��܂�ň͂�
                With Worksheets("�������p�i����j")
                    ovalX = CInt(.Range("���s�敪���W").Item(1).Value)
                    ovalY = CInt(.Range("���s�敪���W").Item(2).Value)
                    ovalW = CInt(.Range("���s�敪���W").Item(3).Value)
                    ovalH = CInt(.Range("���s�敪���W").Item(4).Value)
                End With
                            
                Set oval = .Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, ovalX, ovalY, ovalW, ovalH)
                oval.Fill.Transparency = 1#
                    
                ' PDF�t�@�C���ɏo�͂���
                PDF�t�@�C���o�� "No" & myNum & "-����v�Z����.pdf"
                    
                ' �܂���폜����
                oval.Delete
                    
                ' ��\���ɂ��ǂ�
                .Visible = False
            End With
        Case "�C�O"
            With Worksheets("�l���P�i���s�\�����j")
                .Activate
                .Visible = True
                PDF�t�@�C���o�� "No" & myNum & "-�l���P�i���s�\�����j.pdf"
                .Visible = False
            End With
            With Worksheets("�l���Q�b�i���s���ߕ�j")
                .Activate
                .Visible = True
                PDF�t�@�C���o�� "No" & myNum & "-�l���Q�b�i���s���ߕ�j.pdf"
                .Visible = False
            End With
            With Worksheets("�l���Q���i���s�����\�j")
                .Activate
                .Visible = True
                PDF�t�@�C���o�� "No" & myNum & "-�l���Q���i���s�����\�j.pdf"
                .Visible = False
            End With
            
            
        Case Else
            MsgBox "�����܂��͊C�O�̂ǂ��炩��I�����Ă��������B"
        End Select
        
        
        ' �o���V�[�g�ɖ߂�
        Worksheets("�o��").Activate
    End If
End Sub

Sub �o���������{�^��_Click()
    Dim myNum As Variant

    myNum = Application.InputBox("�������`�[��No����͂��Ă�������")
    ' myNum = 1
    
    If myNum <> False Then
        ' �`�[��No��]�L����
        Worksheets("�������p�i����j").Range("B7").Value = myNum
        
        ' ���s���ߕ�
        Worksheets("�o��������").Activate
        With ActiveSheet
            ' �\������
            .Visible = True

            ' ����v���r���[��\������
            PDF�t�@�C���o�� "No" & myNum & "�o��������.pdf"
            ' ��\���ɂ��ǂ�
            .Visible = False
        End With
    End If
End Sub


Sub PDF�t�@�C���o��(ByVal myFileName As String)
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=myFileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    MsgBox "�}�C�h�L�������g��PDF�t�@�C���u" & myFileName & "�v���쐬���܂���"
End Sub

