VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} add_font_color_re_helper2 
   Caption         =   "���K�\���ŕ����F������ "
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   OleObjectBlob   =   "add_font_color_re_helper2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "add_font_color_re_helper2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim cmd_execute As Boolean



''''''''''''''''''''''''''''''''''''''
'�J���[�p���b�g���J��
Private Sub ToggleButton1_Click()
    Dim top, left
    
    'add_font_color_re_helper2���B��
    add_font_color_re_helper2.Hide
    
    'add_font_color_re_helper2�̃E�B���h�E�ʒu���擾
    top = add_font_color_re_helper2.top
    left = add_font_color_re_helper2.left
    
    'add_font_color_re_helper2��RegExp��ݒ�
    add_font_color_re_helper.ComboBox2.Value = ComboBox2.Value
    
    'add_font_color_re_helper2�̃t�H���g�T�C�Y��ݒ�
    add_font_color_re_helper.ComboBox3.Value = ComboBox3.Value
    
    'add_font_color_re_helper2��IgnoreCase��ݒ�
    add_font_color_re_helper.CheckBox1.Value = CheckBox1.Value
    
    'add_font_color_re_helper2��Bold��ݒ�
    add_font_color_re_helper.CheckBox2.Value = CheckBox2.Value
    
    'add_font_color_re_helper�̐F��ݒ�
    add_font_color_re_helper.ComboBox1.ListIndex = ComboBox1.ListIndex
    
    'add_font_color_re_helper�̃E�B���h�E�ʒu��ݒ�
    add_font_color_re_helper.StartUpPosition = 0
    add_font_color_re_helper.top = top
    add_font_color_re_helper.left = left
    
    'add_font_color_re_helper��\������
    add_font_color_re_helper.Show
    
    'add_font_color_re_helper2���I������
    Unload add_font_color_re_helper2
    
    '����摜��\��
    add_font_color_re_helper.set_select_image (ComboBox1.ListIndex)
    
    '���摜��\��
    add_font_color_re_helper.set_candidate_image (ComboBox1.ListIndex)
            
    '����������X�V����
    Call regex_colorize.SerchKeyWord
    
    '�g�O���{�^���Ƀt�H�[�J�X�����킹��
    add_font_color_re_helper.ToggleButton1.SetFocus
    
    '�C�x���g��L���ɂ���
    Call ThisWorkbook.set_sht
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'�L�[���͐���
Private Sub cmd_execute_judge(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        cmd_execute = True
    End If
    
    'ESC���̓t�H�[���I��
    If ((KeyCode = vbKeyEscape) Or (KeyCode = 219 And Shift = 2)) Then
        ' KeyCode : 219 = [
        ' Shift   :   2 = Ctrl
        Unload add_font_color_re_helper2
    End If
    
    'Ctrl����Excel�Ƀt�H�[�J�X���ړ�
    'If (KeyCode = vbKeyControl) Then
        'AppActivate Excel.Application
    'End If
End Sub


Private Sub CommandButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call cmd_execute_judge(KeyCode, Shift)
End Sub


Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call cmd_execute_judge(KeyCode, Shift)
End Sub


Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call cmd_execute_judge(KeyCode, Shift)
End Sub


Private Sub ComboBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call cmd_execute_judge(KeyCode, Shift)
End Sub


Private Sub ComboBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call cmd_execute_judge(KeyCode, Shift)
End Sub


Private Sub CheckBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call cmd_execute_judge(KeyCode, Shift)
    
    If (cmd_execute) Then
        cmd_execute = False
        
        If (CheckBox1.Value) Then
            CheckBox1.Value = False
        Else
            CheckBox1.Value = True
        End If
    End If
End Sub


Private Sub CheckBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call cmd_execute_judge(KeyCode, Shift)
    
    If (cmd_execute) Then
        cmd_execute = False
        
        If (CheckBox2.Value) Then
            CheckBox2.Value = False
        Else
            CheckBox2.Value = True
        End If
    End If
End Sub


Private Sub ToggleButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call cmd_execute_judge(KeyCode, Shift)
    
    If (cmd_execute) Then
        cmd_execute = False
        
        ToggleButton1_Click
    End If
End Sub


Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If (cmd_execute) Then
        cmd_execute = False
        Cancel = True
        Call CommandButton1_Click
    End If
End Sub


Private Sub ComboBox1_Click()
    Dim idx
    idx = ComboBox1.ListIndex
End Sub


Private Sub ComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If (cmd_execute) Then
        cmd_execute = False
        Cancel = True
        Call CommandButton1_Click
        
        AppActivate Excel.Application 'Excel�Ƀt�H�[�J�X���ړ�
    End If
End Sub


Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If (cmd_execute) Then
        cmd_execute = False
        Cancel = True
        Call CommandButton1_Click
        
        AppActivate Excel.Application 'Excel�Ƀt�H�[�J�X���ړ�
    End If
End Sub


Private Sub ComboBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If (cmd_execute) Then
        cmd_execute = False
        Cancel = True
        Call CommandButton1_Click
        
        AppActivate Excel.Application 'Excel�Ƀt�H�[�J�X���ړ�
    End If
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'RegExp�̃t�H�[�J�X���e�L�X�g�S�I����Ԃł��킹��
Public Sub FocusRegExp()

    ComboBox1.SetFocus '�F�R���{�{�b�N�X�ɂ�������t�H�[�J�X�����킹��
    
     With ComboBox2 'RegExp�R���{�{�b�N�X�Ƀt�H�[�J�X�����킹��
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'�F�t��
Private Sub CommandButton1_Click()
    Call textcolor_execute
End Sub


Public Sub textcolor_execute()
    Dim idx
    idx = ComboBox1.ListIndex
    If (idx < 0) Then
        idx = 0
    End If
    
    '�Z���I����
    If TypeName(Selection) = "Range" Then
        If (ComboBox2.Value = "\-") Then
            Call colorize_strike(ComboBox1.list(idx, 1), Selection, CheckBox2.Value, ComboBox3.Value)
        Else
            Call add_font_color_re(ComboBox2.Value, ComboBox1.list(idx, 1), Selection, CheckBox1.Value, CheckBox2.Value, ComboBox3.Value)
        End If
    End If
    
    '�I�[�g�V�F�C�v�I����
    If TypeName(Selection) = "TextBox" Or TypeName(Selection) = "Rectangle" Or TypeName(Selection) = "Oval" Then
        If (ComboBox2.Value = "\-") Then
            Call colorize_strike_shape(ComboBox1.list(idx, 1), Selection, CheckBox2.Value, ComboBox3.Value)
        Else
            Call add_font_color_re_shape(ComboBox2.Value, ComboBox1.list(idx, 1), Selection, CheckBox1.Value, CheckBox2.Value, ComboBox3.Value)
        End If
    End If
    
    '�I�[�g�V�F�C�v�����I����
    If TypeName(Selection) = "DrawingObjects" Then
        If (ComboBox2.Value = "\-") Then
            Call colorize_strike_shapes(ComboBox1.list(idx, 1), Selection, CheckBox2.Value, ComboBox3.Value)
        Else
            Call add_font_color_re_shapes(ComboBox2.Value, ComboBox1.list(idx, 1), Selection, CheckBox1.Value, CheckBox2.Value, ComboBox3.Value)
        End If
    End If
    
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'���[�U�[�t�H�[����������
Private Sub UserForm_Initialize()

    'Caption��ݒ�
    add_font_color_re_helper2.Caption = "���K�\���ŐF������ " & regex_colorize.VERSION & " "
    
    '�J���[�p���b�g�̃��X�g��ݒ�
    Dim clr() As String
    clr = regex_colorize.colorlist
    
    With ComboBox1
        .list = clr
    End With
    
    '�t�H���g�T�C�Y�̃��X�g�ݒ�
    Dim fontsizelists() As Single
    fontsizelists = regex_colorize.fontsizelist
    
    With ComboBox3
        .list = fontsizelists
    End With
    
    cmd_execute = False
    
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'���[�U�[�t�H�[���I�����O
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim IntFlNo As Integer
    Dim SettingFile As String
    
    ' �ݒ�t�@�C����������
    SettingFile = Environ("APPDATA") & "\" & regex_colorize.MSADDIN_PATH_SETTING_FILE
    IntFlNo = FreeFile
    Open SettingFile For Output As #IntFlNo
    Write #IntFlNo, ComboBox2.Value     'Regex
    Write #IntFlNo, ComboBox1.ListIndex 'Color
    Write #IntFlNo, CheckBox1.Value     'IgnoreCase
    Write #IntFlNo, CheckBox2.Value     'Bold
    Write #IntFlNo, 1                   'Window
    Write #IntFlNo, add_font_color_re_helper2.top  'Top
    Write #IntFlNo, add_font_color_re_helper2.left 'Left
    Write #IntFlNo, ThisWorkbook.shortcutkey 'ShortCutKey
    Close #IntFlNo
    
    Call ThisWorkbook.reset_sht
End Sub
''''''''''''''''''''''''''''''''''''''



