VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} setting 
   Caption         =   "設定"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2310
   OleObjectBlob   =   "setting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub CommandButton1_Click()
    Dim IntFlNo As Integer
    Dim SettingFile As String
    Dim LoadInfo As SettingFileInfo
    
    'ショートカットキー変更
    If ThisWorkbook.shortcutkey <> "" Then
        Application.OnKey "^" & ThisWorkbook.shortcutkey 'クリア
    End If
    
    If ComboBox1.Value <> "" Then
        Application.OnKey "^" & ComboBox1.Value, "regex_colorize.regex_highlight" '再設定
    End If
    
    ThisWorkbook.shortcutkey = ComboBox1.Value
    
    '設定ファイル更新
    SettingFile = Environ("APPDATA") & "\" & regex_colorize.MSADDIN_PATH_SETTING_FILE
    LoadInfo = regex_colorize.LoadSettingFile(SettingFile)
    
    IntFlNo = FreeFile
    Open SettingFile For Output As #IntFlNo
    Write #IntFlNo, LoadInfo.regex 'Regex
    Write #IntFlNo, LoadInfo.color 'Color
    Write #IntFlNo, LoadInfo.icase 'IgnoreCase
    Write #IntFlNo, LoadInfo.bold 'Bold
    Write #IntFlNo, LoadInfo.window 'Window
    Write #IntFlNo, LoadInfo.top_pos 'Top
    Write #IntFlNo, LoadInfo.left_pos 'Left
    Write #IntFlNo, ThisWorkbook.shortcutkey 'ShortCutKey
    Close #IntFlNo
    
    Unload setting
    
End Sub


Private Sub UserForm_Initialize()

    ComboBox1.Value = ThisWorkbook.shortcutkey
    
End Sub
