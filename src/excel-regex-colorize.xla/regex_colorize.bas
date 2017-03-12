Attribute VB_Name = "regex_colorize"
Option Explicit

Public Const VERSION = "Ver1.3.0"
Public Texts() As String

Public Const MSADDIN_PATH_SETTING_FILE = "Microsoft\AddIns" & "\" & "textcolor.txt"

'�ݒ�t�@�C�����
Public Type SettingFileInfo
    regex As String
    color As Long
    icase As Boolean
    bold As Boolean
    window As Integer
    top_pos As Double
    left_pos As Double
    shortcutkey As String
End Type

Public EXCEL_MAX_ROW As Long 'Excel���F�����Ă���ŏI��(������)
Public EXCEL_MAX_COLUMN As Long  'Excel���F�����Ă���ŏI�s(�E����)
Public Const MAX_SELECTION_COUNT = 50

'�ݒ藓�N��
Public Sub setting()
    setting.Show
End Sub

'�t�H�[���̋N��
Sub regex_highlight()
Attribute regex_highlight.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim SettingFile As String
    Dim LoadInfo As SettingFileInfo
    
    EXCEL_MAX_ROW = ActiveCell.SpecialCells(xlLastCell).Row
    EXCEL_MAX_COLUMN = ActiveCell.SpecialCells(xlLastCell).Column
    
    '�t�H�[�����J���Ă��Ȃ��ꍇ�͕\������
    If ((Not add_font_color_re_helper.Visible) And (Not add_font_color_re_helper2.Visible)) Then
    
        'Excel��SelectionChange�C�x���g��L���ɂ���
        Call ThisWorkbook.set_sht
        
        ' �ݒ�t�@�C����ǂݍ���
        SettingFile = Environ("APPDATA") & "\" & MSADDIN_PATH_SETTING_FILE
        LoadInfo = LoadSettingFile(SettingFile)
    
        '�J���[�p���b�g����
        If (LoadInfo.window = 0) Then
            add_font_color_re_helper.Show
            
            add_font_color_re_helper.ComboBox2.Value = LoadInfo.regex
            add_font_color_re_helper.ComboBox1.ListIndex = LoadInfo.color
            add_font_color_re_helper.CheckBox1.Value = LoadInfo.icase
            add_font_color_re_helper.CheckBox2.Value = LoadInfo.bold
            
            If Dir(SettingFile) <> "" Then
                add_font_color_re_helper.StartUpPosition = 0
                add_font_color_re_helper.top = LoadInfo.top_pos
                add_font_color_re_helper.left = LoadInfo.left_pos
            End If
            
            '����摜��\��
            add_font_color_re_helper.set_select_image (LoadInfo.color)
    
            '���摜��\��
            add_font_color_re_helper.set_candidate_image (LoadInfo.color)
            
            '����������X�V����
            Call regex_colorize.SerchKeyWord
            
            'RegExp�Ƀt�H�[�J�X�����킹��
            add_font_color_re_helper.FocusRegExp
            
        '�J���[�p���b�g�Ȃ�
        Else
            add_font_color_re_helper2.Show
            
            add_font_color_re_helper2.ComboBox2.Value = LoadInfo.regex
            add_font_color_re_helper2.ComboBox1.ListIndex = LoadInfo.color
            add_font_color_re_helper2.CheckBox1.Value = LoadInfo.icase
            add_font_color_re_helper2.CheckBox2.Value = LoadInfo.bold
            
            If Dir(SettingFile) <> "" Then
                add_font_color_re_helper2.StartUpPosition = 0
                add_font_color_re_helper2.top = LoadInfo.top_pos
                add_font_color_re_helper2.left = LoadInfo.left_pos
            End If
            
            '����������X�V����
            Call regex_colorize.SerchKeyWord
            
            'RegExp�Ƀt�H�[�J�X�����킹��
            add_font_color_re_helper2.FocusRegExp
            
        End If
    
    '���ɃE�B���h�E�������オ���Ă����ꍇ�̓t�H�[�J�X���ړ�����
    ElseIf (add_font_color_re_helper.Visible) Then
        AppActivate add_font_color_re_helper.Caption
        
        '����������X�V����
        Call regex_colorize.SerchKeyWord
        
        'RegExp�Ƀt�H�[�J�X�����킹��
        add_font_color_re_helper.FocusRegExp
        
    ElseIf (add_font_color_re_helper2.Visible) Then
        AppActivate add_font_color_re_helper2.Caption
        
        '����������X�V����
        Call regex_colorize.SerchKeyWord
        
        'RegExp�Ƀt�H�[�J�X�����킹��
        add_font_color_re_helper2.FocusRegExp
        
    End If
    
End Sub


'�I��͈͂̌��������擾����&�t�H���g�T�C�Y���X�V����
Sub SerchKeyWord()

    Dim r As Range
    Dim re, m, cnt
    Dim fontsize As Single
    Dim index As Long
    Dim pre_value As String
    Dim selectioncnt As Long
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "\w+|[�-�]+|[�O-�X]+|[�`-�y]+|[��-��]+|[�@-���[]+|[��-�]+"
    re.ignoreCase = True
    re.Global = True
    re.MultiLine = True

    cnt = 0
    selectioncnt = 0
    
    ReDim Texts(1)
    
    '�t�H���g�T�C�Y�̃f�t�H���g�l
    fontsize = 11
    
    '�Z���I����
    If TypeName(Selection) = "Range" Then
        '�I��͈͂̕�������擾����
        With ActiveSheet
            For Each r In Selection
                selectioncnt = selectioncnt + 1
                
                '�񐔂�Excel���F�����Ă���ŏI��𒴂�����I��
                If r.Row > EXCEL_MAX_ROW Then
                    Exit For
                End If
            
                '�s����Excel���F�����Ă���ŏI�s�𒴂�����X�L�b�v
                If r.Column > EXCEL_MAX_COLUMN Then
                    GoTo ContinueLabel1
                End If
                
                '�������ƂɃT�C�Y���قȂ�ꍇ
                If IsNull(r.Font.size) Then
                    For index = 1 To Len(r.Text)
                        If r.Characters(index, 1).Font.size <> fontsize Then
                            fontsize = r.Characters(index, 1).Font.size
                        End If
                    Next
                Else
                    fontsize = r.Font.size
                End If
                
                '�F�Â���������擾
                For Each m In re.Execute(r)
                    If (m.Value <> "") Then
                        cnt = cnt + 1
                        ReDim Preserve Texts(cnt)
                        Texts(cnt) = m.Value
                    End If
                Next
                
                '�����Z��������
                If selectioncnt > MAX_SELECTION_COUNT Then
                    Exit For
                End If
                
ContinueLabel1:
            Next
        End With
    End If
    
    
    '�I�[�g�V�F�C�v�I����
    If TypeName(Selection) = "TextBox" Or TypeName(Selection) = "Rectangle" Or TypeName(Selection) = "Oval" Then
        
        On Error GoTo myError
        '�������ƂɃT�C�Y���قȂ�ꍇ
        If IsNull(Selection.Font.size) Then
            For index = 1 To Len(r.Text)
                If r.Characters(index, 1).Font.size <> fontsize Then
                    fontsize = r.Characters(index, 1).Font.size
                End If
            Next
        Else
            fontsize = Selection.Font.size
        End If
        
        '�F�Â���������擾
        For Each m In re.Execute(Selection.Text)
            If (m.Value <> "") Then
                cnt = cnt + 1
                ReDim Preserve Texts(cnt)
                Texts(cnt) = m.Value
            End If
        Next
        
myError:
'�R�l�N�^�I�����̃G���[���
    End If
    
    '�I�[�g�V�F�C�v�����I����
    If TypeName(Selection) = "DrawingObjects" Then
        
        Dim shape As Variant
        
        On Error Resume Next
        For Each shape In Selection
            '�������ƂɃT�C�Y���قȂ�ꍇ
            If IsNull(shape.Font.size) Then
                For index = 1 To Len(shape.Text)
                    If shape.Characters(index, 1).Font.size <> fontsize Then
                        fontsize = shape.Characters(index, 1).Font.size
                    End If
                Next
            Else
                fontsize = shape.Font.size
            End If
            
            '�F�Â���������擾
            For Each m In re.Execute(shape.Text)
                If (m.Value <> "") Then
                    cnt = cnt + 1
                    ReDim Preserve Texts(cnt)
                    Texts(cnt) = m.Value
                End If
            Next
        Next
        
    End If
    
    
    '�����Ƀ\�[�g���ďd�����폜����
    Call Q_Sort(Texts, 1, UBound(Texts))
    Call Unique(Texts)
    
    
    'add_font_color_re_helper���J���Ă���ꍇ
    If (add_font_color_re_helper.Visible) Then
        pre_value = add_font_color_re_helper.ComboBox2.Value
        
        With add_font_color_re_helper.ComboBox2
            .ColumnCount = 1
            .list = Texts
            .ListIndex = 0
            .Value = pre_value
        End With
        
    'add_font_color_re_helper2���J���Ă���ꍇ
    ElseIf (add_font_color_re_helper2.Visible) Then
        pre_value = add_font_color_re_helper2.ComboBox2.Value
        
        With add_font_color_re_helper2.ComboBox2
            .ColumnCount = 1
            .list = Texts
            .ListIndex = 0
            .Value = pre_value
        End With
    End If
    
    'add_font_color_re_helper���J���Ă���ꍇ
    If (add_font_color_re_helper.Visible) Then
        With add_font_color_re_helper.ComboBox3
            .Value = fontsize
        End With
    
    'add_font_color_re_helper2���J���Ă���ꍇ
    ElseIf (add_font_color_re_helper2.Visible) Then
        With add_font_color_re_helper2.ComboBox3
            .Value = fontsize
        End With
    End If
    
End Sub



'�\�[�g
Sub Q_Sort(ByRef myData() As String, ByVal L As Long, ByVal U As Long)
    Dim i As Long
    Dim j As Long
    Dim S As Variant
    Dim tmp As Variant
        S = myData(Int((L + U) / 2))
        i = L
        j = U
        Do
            Do While myData(i) < S
                i = i + 1
            Loop
            Do While myData(j) > S
                j = j - 1
            Loop
            If i >= j Then Exit Do
            tmp = myData(i)
            myData(i) = myData(j)
            myData(j) = tmp
            i = i + 1
            j = j - 1
      Loop
      If (L < i - 1) Then Q_Sort myData, L, i - 1
      If (U > j + 1) Then Q_Sort myData, j + 1, U
End Sub



'���j�[�N
Sub Unique(ByRef myData() As String)
    Dim i As Long
    Dim cnt As Long
    Dim tmp() As String
    Dim pre As String
    
    '���j�[�N�ȗv�f�𒊏o
    pre = ""
    cnt = 0
    ReDim tmp(1)
    For i = 1 To UBound(Texts)
        If (Texts(i) <> pre) Then
            cnt = cnt + 1
            ReDim Preserve tmp(cnt)
            tmp(cnt) = Texts(i)
            pre = Texts(i)
        End If
    Next
    
    'tmp�ɏ]���Č��ʂ��Ē�`
    ReDim myData(UBound(tmp))
    
    For i = 1 To UBound(tmp)
        myData(i) = tmp(i)
    Next
End Sub


'�F�t��(�Z���p)
Sub add_font_color_re(ptn As String, clr As Long, rng As Range, _
                    Optional icase As Boolean = False, Optional bold As Boolean = False, Optional size As Single = 11 _
                    )
    Dim r As Range, i As Long, colInd As Integer
    Dim ptr, re, m
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = ptn
    re.ignoreCase = icase
    re.Global = True
    re.MultiLine = True
    
    On Error GoTo ErrorHandler
    With ActiveSheet
        For Each r In rng
            '�񐔂�Excel���F�����Ă���ŏI��𒴂�����I��
            If r.Row > EXCEL_MAX_ROW Then
                Exit For
            End If
            
            '�s����Excel���F�����Ă���ŏI�s�𒴂�����X�L�b�v
            If r.Column > EXCEL_MAX_COLUMN Then
                GoTo ContinueLabel2
            End If
                
            For Each m In re.Execute(r)
                With r.Characters(m.FirstIndex + 1, m.Length).Font
                    .colorindex = clr
                    .bold = bold
                    .size = size
                End With
            Next
            
ContinueLabel2:
        Next
    End With
    
ErrorHandler:
        'MsgBox err.Description
End Sub


'�F�t��(�I�[�g�V�F�C�v�p)
Sub add_font_color_re_shape(ptn As String, clr As Long, shape As Variant, _
                    Optional icase As Boolean = False, Optional bold As Boolean = False, Optional size As Single = 11 _
                    )
                    
    Dim re, m
    
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = ptn
    re.ignoreCase = icase
    re.Global = True
    re.MultiLine = True
    
    On Error GoTo ErrorHandler

    For Each m In re.Execute(shape.Text)
        With shape.Characters(m.FirstIndex + 1, m.Length).Font
            .colorindex = clr
            .bold = bold
            .size = size
        End With
    Next
        
ErrorHandler:
    
End Sub


'�F�t��(�I�[�g�V�F�C�v�����p)
Sub add_font_color_re_shapes(ptn As String, clr As Long, shapes As Variant, _
                    Optional icase As Boolean = False, Optional bold As Boolean = False, Optional size As Single = 11 _
                    )
                    
    Dim re, m
    
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = ptn
    re.ignoreCase = icase
    re.Global = True
    re.MultiLine = True
    
    Dim shape As Variant
    
    On Error Resume Next
    For Each shape In shapes
        For Each m In re.Execute(shape.Text)
            With shape.Characters(m.FirstIndex + 1, m.Length).Font
                .colorindex = clr
                .bold = bold
                .size = size
            End With
        Next
    Next
    
End Sub


'�������t���̕����F��ς���(�Z���p)
Sub colorize_strike(clr As Long, rng As Range, Optional bold As Boolean = False, Optional size As Single = 11)
    Dim r, i, c
    
    For Each r In rng
        '�񐔂�Excel���F�����Ă���ŏI��𒴂�����I��
        If r.Row > EXCEL_MAX_ROW Then
            Exit For
        End If
        
        '�s����Excel���F�����Ă���ŏI�s�𒴂�����X�L�b�v
        If r.Column > EXCEL_MAX_COLUMN Then
            GoTo ContinueLabel3
        End If
            
        For i = 1 To Len(r)
            Set c = r.Characters(i, 1).Font
            If (c.Strikethrough) Then
                c.colorindex = clr
                c.bold = bold
                c.size = size
            End If
        Next i
        
ContinueLabel3:
    Next
    
End Sub


'�������t���̕����F��ς���(�I�[�g�V�F�C�v�p)
Sub colorize_strike_shape(clr As Long, shape As Variant, Optional bold As Boolean = False, Optional size As Single = 11)
    Dim i, c

    On Error GoTo ErrorHandler
    For i = 1 To Len(shape.Text)
        Set c = shape.Characters(i, 1).Font
        If (c.Strikethrough) Then
            c.colorindex = clr
            c.bold = bold
            c.size = size
        End If
    Next i
ErrorHandler:
   
End Sub


'�������t���̕����F��ς���(�I�[�g�V�F�C�v�����p)
Sub colorize_strike_shapes(clr As Long, shapes As Variant, Optional bold As Boolean = False, Optional size As Single = 11)
    Dim i, c
    Dim shape As Variant
    
    On Error Resume Next
    For Each shape In shapes
        For i = 1 To Len(shape.Text)
            Set c = shape.Characters(i, 1).Font
            If (c.Strikethrough) Then
                c.colorindex = clr
                c.bold = bold
                c.size = size
            End If
        Next i
    Next
    
End Sub


'�J���[�p���b�g�̐F�ݒ�
Function colorlist() As String()

    Dim clr(0 To 39, 0 To 1) As String
    
    clr(0, 0) = "BLACK"
    clr(0, 1) = 1
    clr(1, 0) = "BROWN"
    clr(1, 1) = 53
    clr(2, 0) = "OLIVE"
    clr(2, 1) = 52
    clr(3, 0) = "DARKGREEN"
    clr(3, 1) = 51
    clr(4, 0) = "DARKBLUEGREEN"
    clr(4, 1) = 49
    clr(5, 0) = "DARKBLUE"
    clr(5, 1) = 11
    clr(6, 0) = "INDIGO"
    clr(6, 1) = 55
    clr(7, 0) = "GRAY80"
    clr(7, 1) = 56
    clr(8, 0) = "DARKRED"
    clr(8, 1) = 9
    clr(9, 0) = "ORANGE"
    clr(9, 1) = 46
    clr(10, 0) = "DARKYELLOW"
    clr(10, 1) = 12
    clr(11, 0) = "GREEN"
    clr(11, 1) = 10
    clr(12, 0) = "BLUEGREEN"
    clr(12, 1) = 14
    clr(13, 0) = "BLUE"
    clr(13, 1) = 5
    clr(14, 0) = "BLUEGRAY"
    clr(14, 1) = 47
    clr(15, 0) = "GRAY50"
    clr(15, 1) = 16
    clr(16, 0) = "RED"
    clr(16, 1) = 3
    clr(17, 0) = "PALEORANGE"
    clr(17, 1) = 45
    clr(18, 0) = "LIME"
    clr(18, 1) = 43
    clr(19, 0) = "SEAGREEN"
    clr(19, 1) = 50
    clr(20, 0) = "AQUA"
    clr(20, 1) = 42
    clr(21, 0) = "LIGHTBLUE"
    clr(21, 1) = 41
    clr(22, 0) = "PURPLE"
    clr(22, 1) = 13
    clr(23, 0) = "GRAY40"
    clr(23, 1) = 48
    clr(24, 0) = "PINK"
    clr(24, 1) = 7
    clr(25, 0) = "GOLD"
    clr(25, 1) = 44
    clr(26, 0) = "YELLOW"
    clr(26, 1) = 6
    clr(27, 0) = "LIGHTGREEN"
    clr(27, 1) = 4
    clr(28, 0) = "WATER"
    clr(28, 1) = 8
    clr(29, 0) = "SKYBLUE"
    clr(29, 1) = 33
    clr(30, 0) = "PRAM"
    clr(30, 1) = 54
    clr(31, 0) = "GRAY25"
    clr(31, 1) = 15
    clr(32, 0) = "ROSE"
    clr(32, 1) = 38
    clr(33, 0) = "BEIGE"
    clr(33, 1) = 40
    clr(34, 0) = "PALEYELLOW"
    clr(34, 1) = 36
    clr(35, 0) = "PALEGREEN"
    clr(35, 1) = 35
    clr(36, 0) = "PALEWATER"
    clr(36, 1) = 34
    clr(37, 0) = "PALEBLUE"
    clr(37, 1) = 37
    clr(38, 0) = "LAVENDER"
    clr(38, 1) = 39
    clr(39, 0) = "WHITE"
    clr(39, 1) = 2
    
    colorlist = clr
    
End Function

'�t�H���g�T�C�Y�̐ݒ�
Function fontsizelist() As Single()

    Dim list(0 To 16) As Single
    
    list(0) = 6
    list(1) = 8
    list(2) = 9
    list(3) = 10
    list(4) = 11
    list(5) = 12
    list(6) = 14
    list(7) = 16
    list(8) = 18
    list(9) = 20
    list(10) = 22
    list(11) = 24
    list(12) = 26
    list(13) = 28
    list(14) = 36
    list(15) = 48
    list(16) = 72
    
    fontsizelist = list
    
End Function

'�ݒ�t�@�C���̓ǂݍ���
Function LoadSettingFile(ByVal SettingFile As String) As SettingFileInfo

    Dim IntFlNo As Integer
    Dim buf
    
    ' �����l��ݒ�
    With LoadSettingFile
        .regex = ""
        .color = 16
        .icase = False
        .bold = True
        .window = 0
        .top_pos = 0#
        .left_pos = 0#
        .shortcutkey = "t"
    End With
    
    ' �ݒ�t�@�C��������ꍇ
    If Dir(SettingFile) <> "" Then
        IntFlNo = FreeFile
        Open SettingFile For Input As #IntFlNo
        
        'Regex�擾
        Line Input #IntFlNo, buf
        LoadSettingFile.regex = Mid(buf, 2, Len(buf) - 2)
        
        'Color�擾
        Line Input #IntFlNo, buf
        LoadSettingFile.color = buf
        
        'IgnoreCase�擾
        Line Input #IntFlNo, buf
        LoadSettingFile.icase = buf
        
        'Bold�擾
        Line Input #IntFlNo, buf
        LoadSettingFile.bold = buf
        
        'Window�擾
        Line Input #IntFlNo, buf
        LoadSettingFile.window = buf
        
        'Top�擾
        Line Input #IntFlNo, buf
        LoadSettingFile.top_pos = buf
        
        'Left�擾
        Line Input #IntFlNo, buf
        LoadSettingFile.left_pos = buf
        
        '�V���[�g�J�b�g�L�[�擾
        Line Input #IntFlNo, buf
        LoadSettingFile.shortcutkey = Mid(buf, 2, Len(buf) - 2)
        
        Close #IntFlNo
    End If
    
End Function



