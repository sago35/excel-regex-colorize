VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColorPalletEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ColorPalletImage As MSForms.image
Attribute ColorPalletImage.VB_VarHelpID = -1
Public WithEvents ColorPalletLabel As MSForms.Label
Attribute ColorPalletLabel.VB_VarHelpID = -1

Public index As Integer



'�R���X�g���N�^
Public Sub NewClass(ByVal pallet_control As MSForms.Control, pallet_index As Integer, pallet_id As Integer)

    If TypeOf pallet_control Is MSForms.image Then
    
        Set ColorPalletImage = pallet_control
        
    ElseIf TypeOf pallet_control Is MSForms.Label Then
    
        Set ColorPalletLabel = pallet_control

    End If
    
    index = pallet_index

End Sub



'�摜�N���b�N���̓���
Private Sub ColorPalletImage_Click()
    
    With add_font_color_re_helper.ComboBox1
        .ListIndex = index
    End With
    
    '�N���b�N�ʒu�̌���摜��\��
    add_font_color_re_helper.set_select_image (index)
    
End Sub



'�摜�_�u���N���b�N���̓���
Private Sub ColorPalletImage_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Call add_font_color_re_helper.textcolor_execute
    
End Sub



'�摜�Ƀ}�E�X�J�[�\�����������Ƃ��̏���
Private Sub ColorPalletImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    '�}�E�X�J�[�\���ʒu�̌��摜��\��
    add_font_color_re_helper.set_candidate_image (index)
    
End Sub



'���x���Ƀ}�E�X�J�[�\�����������Ƃ��̏���
Private Sub ColorPalletLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    '�}�E�X�J�[�\���ʒu�̌��摜��\��
    add_font_color_re_helper.set_candidate_image (index)
    
End Sub

