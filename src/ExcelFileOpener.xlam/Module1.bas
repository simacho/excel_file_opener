Attribute VB_Name = "Module1"
Option Explicit
' ------------------------------------------------------------
'  変数定義
' ------------------------------------------------------------
Public activePath As String         ' アクティブパス管理用
Public crntPath As String           ' 現在のパス管理用
Public selectedName As String       ' 選択項目名受け渡し用
Public nodeCount As Long            ' 項目数
Public noMode As Integer            ' モード管理
Public filesBuffer() As String      ' リスト表示バッファ
Public maxCount As Long             ' 再帰時のファイル上限
Public amountFile As Long


' フラグ系
Public waitFlag As Boolean
Public escFlag As Boolean

' INIファイル用
Public iniWidth As Long
Public iniHeight As Long

' ------------------------------------------------------------
'  定数定義
' ------------------------------------------------------------
Enum mode
    ACTIVE_PATH = 1
    RECURSIVE_PATH = 2
    RECENT_FILE = 3
    SWITCH_BOOK = 4
End Enum

' ------------------------------------------------------------
'  INIファイル関連
' ------------------------------------------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long


' ------------------------------------------------------------
'  初期化
' ------------------------------------------------------------
Private Sub InitGlobal()
    Dim wScriptHost As Object
    Set wScriptHost = CreateObject("WScript.Shell")
    
    activePath = ActiveWorkbook.path
    crntPath = activePath
    selectedName = ""
    nodeCount = 0
    amountFile = 0
    noMode = mode.ACTIVE_PATH
    escFlag = False
    waitFlag = False
    
    ' 初期値
    iniWidth = 500
    iniHeight = 300
    maxCount = 10000
    
End Sub

'
' Iniファイル処理
'
Function GetINIValue(KEY As String, Section As String, ININame As String) As String
    Dim Value As String * 255
    Call GetPrivateProfileString(Section, KEY, "ERROR", Value, Len(Value), ININame)
    GetINIValue = Left$(Value, InStr(1, Value, vbNullChar) - 1)
End Function

Private Sub LoadIniFile()
    Dim wScriptHost As Object
    Dim mydoc_path As String
    Dim strWidth As String
    Dim strHeight As String
    Dim strMaxFile As String

    Set wScriptHost = CreateObject("WScript.Shell")
    mydoc_path = wScriptHost.SpecialFolders("MyDocuments")

    strWidth = GetINIValue("WIDTH", "Initial", mydoc_path & "\ExcelFileOpener.ini")
    strHeight = GetINIValue("HEIGHT", "Initial", mydoc_path & "\ExcelFileOpener.ini")

    If Not strWidth = "" Then
        iniWidth = Val(strWidth)
    End If

    If Not strHeight = "" Then
        iniHeight = Val(strHeight)
    End If

    'strMaxFile = GetINIValue("MAXFILE", "Initial", mydoc_path & "\ExcelFileOpener.ini")
    'If Not strMaxFile = "" Then
    '    maxCount = Val(strMaxFile)
    'End If
    
End Sub

'---------------------------------------------------------------------------------------------------

'
' パス上のファイルを選択する
'
Private Function SelectFile() As String

    Dim ret As Boolean
    Dim tgtfile As String
    Dim form As Boolean
    Dim dummy As Variant
    
    ret = False
    
    form = False
    
    Do
        ' 候補を取得する
        filesBuffer = GetFilesByMode(filesBuffer, noMode, crntPath)
        
        ' フォーム表示前更新
        Call UserForm2.TextBox2_Change
        UserForm2.TextBox2.Text = ""
        
        waitFlag = True
                
        ' フォームの表示
        If form = False Then
            form = True
            UserForm2.Show (vbModeless)
        End If

        ' フォーム表示後更新
        UserForm2.TextBox2.SetFocus
 
        ' 入力待ち
        Do While waitFlag
            dummy = DoEvents
        Loop
                        
        ' ------------------------------
        ' ESCキー特別処理
        If escFlag = True Then
            SelectFile = ""
            GoTo LastExit
        End If

        ' 指定物が見つからない場合は繰り返し
        If selectedName = "" Then
            GoTo Continue
        End If
                 
        ' 終了
        tgtfile = selectedName
        ret = True

Continue:
    Loop While ret = False


LastExit:
    ' フォームを消す
    Unload UserForm2

    SelectFile = tgtfile
End Function


'
' モード別のオープン処理
'
Private Sub OpenFileSub(tgtfile As String)

    Dim act_open As Boolean
    Dim idx As Integer
    Dim books() As String
    act_open = True
    
    ' Debug.Print "Open " & tgtfile
    
    'モード別に開く処理を変更する
    Select Case noMode
    Case mode.ACTIVE_PATH
        act_open = True
    Case mode.RECURSIVE_PATH
        act_open = True
    Case mode.RECENT_FILE
        act_open = True
    Case mode.SWITCH_BOOK
        act_open = False
    End Select
    
    If act_open = True Then
        On Error Resume Next  'エラーがあっても続行する
        Workbooks.Open tgtfile, ReadOnly:=False
        If Err.Number <> 0 Then
            MsgBox Error(Err.Number)
            Exit Sub
        End If
    Else
        On Error Resume Next
        
        books = GetWorkBookNames(books)
        idx = WorksheetFunction.Match(tgtfile, books, 0)
        
        'Debug.Print "Activate Index " & idx
        
        Workbooks(idx).Activate
        If Err.Number <> 0 Then
            MsgBox Error(Err.Number)
            Exit Sub
        End If
    End If
End Sub


Private Sub OpenFile0(mno As Integer)
    
    Dim tgtfile As String
              
    ' グローバル変数の初期化
    Call InitGlobal
    
    ' INIファイルの読み込み
    Call LoadIniFile

    ' モード指定
    noMode = mno
            
    'モード別処理
    Select Case noMode
    Case mode.ACTIVE_PATH
        UserForm2.OptionButton1 = True
    
    Case mode.RECURSIVE_PATH
        UserForm2.OptionButton2 = True
    
    Case mode.RECENT_FILE
        UserForm2.OptionButton3 = True
    
    Case mode.SWITCH_BOOK
        UserForm2.OptionButton4 = True
        
    End Select
        
    'Debug.Print "Initial UserPath " & userPath
    
    '選択処理
    tgtfile = SelectFile()
        
    If escFlag = True Then
        Exit Sub
    End If
    
    'ファイルオープン処理
    Call OpenFileSub(tgtfile)
    
End Sub

'======================================================================
'
' 公開部分
'
'======================================================================
'
' 公開関数（現在のファイルと同じフォルダを開く）
'
Public Sub OpenFile_ACTIVE_PATH()
    Call OpenFile0(mode.ACTIVE_PATH)
End Sub

'
' 公開関数（現在のフォルダから再帰的に開く）
'
Public Sub OpenFile_RECURSIVE_PATH()
    Call OpenFile0(mode.RECURSIVE_PATH)
End Sub

'
' 公開関数（履歴ファイルを選択して開く）
'
Public Sub OpenFile_RECENT_FILE()
    Call OpenFile0(mode.RECENT_FILE)
End Sub

'
' 公開関数（履歴ファイルを選択して開く）
'
Public Sub OpenFile_SWITCH_BOOK()
    Call OpenFile0(mode.SWITCH_BOOK)
End Sub

