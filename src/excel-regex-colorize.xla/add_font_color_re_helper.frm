VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} add_font_color_re_helper 
   Caption         =   "正規表現で文字色をつける"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   OleObjectBlob   =   "add_font_color_re_helper.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "add_font_color_re_helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Dim cmd_execute As Boolean

'クラスを格納する変数を定義
Private ColorPallet(0 To 159) As New clsColorPalletEvent



''''''''''''''''''''''''''''''''''''''
'全ての決定画像を非表示にする
Public Sub clear_select_image()
    Dim i As Integer

    For i = 40 To 79
        ColorPallet(i).ColorPalletImage.Visible = False
    Next
    
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'全ての候補画像を非表示にする
Public Sub clear_candidate_image()
    Dim i As Integer

    For i = 80 To 119
        ColorPallet(i).ColorPalletImage.Visible = False
    Next
    
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'決定画像を表示する
Public Sub set_select_image(color As Long)

    '決定画像を全て非表示
    Call add_font_color_re_helper.clear_select_image
    
    'クリック位置の決定画像を表示
    add_font_color_re_helper.Controls("Image" & color + 41).Visible = True
    
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'候補画像を表示する
Public Sub set_candidate_image(color As Long)

    '候補画像を全て非表示
    Call add_font_color_re_helper.clear_candidate_image
    
    'マウスカーソル位置の候補画像を表示
    add_font_color_re_helper.Controls("Image" & color + 81).Visible = True
    
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'フォームにカーソルが合った場合
' ・候補画像を消す
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    '候補画像を消す
    Call clear_candidate_image
    
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'カラーパレットを閉じる
Private Sub ToggleButton1_Click()
    Dim top, left
    
    'add_font_color_re_helperを隠す
    add_font_color_re_helper.Hide
    
    'add_font_color_re_helperのウィンドウ位置を取得
    top = add_font_color_re_helper.top
    left = add_font_color_re_helper.left
    
    'add_font_color_re_helper2のRegExpを設定
    add_font_color_re_helper2.ComboBox2.Value = ComboBox2.Value
    
    'add_font_color_re_helper2のフォントサイズを設定
    add_font_color_re_helper2.ComboBox3.Value = ComboBox3.Value
    
    'add_font_color_re_helper2のIgnoreCaseを設定
    add_font_color_re_helper2.CheckBox1.Value = CheckBox1.Value
    
    'add_font_color_re_helper2のBoldを設定
    add_font_color_re_helper2.CheckBox2.Value = CheckBox2.Value
    
    'add_font_color_re_helper2の色を設定
    add_font_color_re_helper2.ComboBox1.ListIndex = ComboBox1.ListIndex
    
    'add_font_color_re_helper2のウィンドウ位置を設定
    add_font_color_re_helper2.StartUpPosition = 0
    add_font_color_re_helper2.top = top
    add_font_color_re_helper2.left = left
    
    'add_font_color_re_helper2を表示する
    add_font_color_re_helper2.Show
    
    'add_font_color_re_helperを終了する
    Unload add_font_color_re_helper
    
    '文字列候補を更新する
    Call regex_colorize.SerchKeyWord
    
    'トグルボタンにフォーカスをあわせる
    add_font_color_re_helper2.ToggleButton1.SetFocus
    
    'イベントを有効にする
    Call ThisWorkbook.set_sht
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'キー入力制御
Private Sub cmd_execute_judge(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        cmd_execute = True
    End If
    
    'ESC時はフォーム終了
    If ((KeyCode = vbKeyEscape) Or (KeyCode = 219 And Shift = 2)) Then
        ' KeyCode : 219 = [
        ' Shift   :   2 = Ctrl
        Unload add_font_color_re_helper
    End If
    
    'Ctrl時はExcelにフォーカスを移動
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


Private Sub ComboBox1_Click()
    Dim idx
    idx = ComboBox1.ListIndex
    
    'クリック位置の決定画像を表示
    add_font_color_re_helper.set_select_image (idx)
    
    '候補画像を表示
    add_font_color_re_helper.set_candidate_image (idx)
    
End Sub


Private Sub ComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If (cmd_execute) Then
        cmd_execute = False
        Cancel = True
        Call CommandButton1_Click
        
        AppActivate Excel.Application 'Excelにフォーカスを移動
    End If
End Sub


Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If (cmd_execute) Then
        cmd_execute = False
        Cancel = True
        Call CommandButton1_Click
        
        AppActivate Excel.Application 'Excelにフォーカスを移動
    End If
End Sub


Private Sub ComboBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If (cmd_execute) Then
        cmd_execute = False
        Cancel = True
        Call CommandButton1_Click
        
        AppActivate Excel.Application 'Excelにフォーカスを移動
    End If
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'RegExpのフォーカスをテキスト全選択状態であわせる
Public Sub FocusRegExp()

    ComboBox1.SetFocus '色コンボボックスにいったんフォーカスを合わせる
    
     With ComboBox2 'RegExpコンボボックスにフォーカスを合わせる
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'色付け
Private Sub CommandButton1_Click()
    Call textcolor_execute
End Sub


Public Sub textcolor_execute()
    Dim idx
    idx = ComboBox1.ListIndex
    If (idx < 0) Then
        idx = 0
    End If
    
    'セル選択時
    If TypeName(Selection) = "Range" Then
        If (ComboBox2.Value = "\-") Then
            Call colorize_strike(ComboBox1.list(idx, 1), Selection, CheckBox2.Value, ComboBox3.Value)
        Else
            Call add_font_color_re(ComboBox2.Value, ComboBox1.list(idx, 1), Selection, CheckBox1.Value, CheckBox2.Value, ComboBox3.Value)
        End If
    End If
    
    'オートシェイプ選択時
    If TypeName(Selection) = "TextBox" Or TypeName(Selection) = "Rectangle" Or TypeName(Selection) = "Oval" Then
        If (ComboBox2.Value = "\-") Then
            Call colorize_strike_shape(ComboBox1.list(idx, 1), Selection, CheckBox2.Value, ComboBox3.Value)
        Else
            Call add_font_color_re_shape(ComboBox2.Value, ComboBox1.list(idx, 1), Selection, CheckBox1.Value, CheckBox2.Value, ComboBox3.Value)
        End If
    End If
    
    'オートシェイプ複数選択時
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
'ユーザーフォーム初期化時
Private Sub UserForm_Initialize()
    Dim i As Integer
    
    'Captionを設定
    add_font_color_re_helper.Caption = "正規表現で色をつける " & regex_colorize.VERSION
    
    'カラーパレットイベントのクラスを初期化
    For i = 0 To 39 '色画像
        ColorPallet(i).NewClass Controls("Image" & i + 1), i, i + 1
    Next

    For i = 40 To 79 '決定画像
        ColorPallet(i).NewClass Controls("Image" & i + 1), i - 40, i + 1
    Next
    
    For i = 80 To 119 '候補画像
        ColorPallet(i).NewClass Controls("Image" & i + 1), i - 80, i + 1
    Next
    
    For i = 120 To 159 'ラベル
        ColorPallet(i).NewClass Controls("Label" & i - 120 + 4), i - 120, i - 120 + 4
    Next
    
    'カラーパレットのリスト設定
    Dim clr() As String
    clr = regex_colorize.colorlist
    
    With ComboBox1
        .list = clr
    End With
    
    'フォントサイズのリスト設定
    Dim fontsizelists() As Single
    fontsizelists = regex_colorize.fontsizelist
    
    With ComboBox3
        .list = fontsizelists
    End With
    
    cmd_execute = False
    
End Sub
''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''
'ユーザーフォーム終了直前
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim IntFlNo As Integer
    Dim SettingFile As String
    
    ' 設定ファイル書き込み
    SettingFile = Environ("APPDATA") & "\" & regex_colorize.MSADDIN_PATH_SETTING_FILE
    IntFlNo = FreeFile
    Open SettingFile For Output As #IntFlNo
    Write #IntFlNo, ComboBox2.Value     'Regex
    Write #IntFlNo, ComboBox1.ListIndex 'Color
    Write #IntFlNo, CheckBox1.Value     'IgnoreCase
    Write #IntFlNo, CheckBox2.Value     'Bold
    Write #IntFlNo, 0                   'Window
    Write #IntFlNo, add_font_color_re_helper.top  'Top
    Write #IntFlNo, add_font_color_re_helper.left 'Left
    Write #IntFlNo, ThisWorkbook.shortcutkey 'ShortCutKey
    Close #IntFlNo
    
    Call ThisWorkbook.reset_sht
End Sub
''''''''''''''''''''''''''''''''''''''



