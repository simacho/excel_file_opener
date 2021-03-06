Attribute VB_Name = "Module1"
Option Explicit
' ------------------------------------------------------------
'  変数定義
' ------------------------------------------------------------
Public activePath As String         ' アクティブパス管理用
Public crntPath As String           ' 現在のパス管理用
Public prevPath As String           ' 前回のパス管理用
Public initialString As String      ' 最初の文字列
Public selectedName As String       ' 選択項目名受け渡し用
Public nodeCount As Long            ' 項目数
Public noMode As Integer            ' モード管理
Public filesBuffer() As String      ' リスト表示バッファ
Public maxCount As Long             ' 再帰時のファイル上限
Public amountFile As Long
Public pathDic As New Dictionary         ' パス保存用
    
' フラグ系
Public waitFlag As Boolean
Public escFlag As Boolean
Public recursiveFlag As Boolean

' INIファイル用
Public iniWidth As Long
Public iniHeight As Long

' ------------------------------------------------------------
'  定数定義
' ------------------------------------------------------------
Enum mode
    ACTIVE_PATH = 1
    PREVIOUS_PATH = 2
    RECENT_FILE = 3
    SWITCH_BOOK = 4
End Enum

' ------------------------------------------------------------
'  INIファイル関連
' ------------------------------------------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' ------------------------------------------------------------
'  初期化
' ------------------------------------------------------------
Private Sub InitGlobal()
    Dim wScriptHost As Object
    Set wScriptHost = CreateObject("WScript.Shell")
    
    activePath = ""
    If Not ActiveWorkbook Is Nothing Then
        activePath = ActiveWorkbook.path
    End If
    crntPath = activePath
    prevPath = ""
    initialString = ""
    selectedName = ""
    nodeCount = 0
    amountFile = 0
    noMode = mode.ACTIVE_PATH
    
    ' ファイルバッファクリア
    ReDim filesBuffer(0)
    
    ' Path用の辞書
    Set pathDic = CreateObject("Scripting.Dictionary")
        
    ' フラグ関連クリア
    escFlag = False
    waitFlag = False
    recursiveFlag = False
    
    ' 初期値
    iniWidth = 500
    iniHeight = 300
    maxCount = 10000
    
End Sub


Private Sub ExitGlobal()
    Set pathDic = Nothing
End Sub


'
' Iniファイル読み込み
'
Function GetINIValue(KEY As String, Section As String, ININame As String) As String
    Dim Value As String * 8192
    Call GetPrivateProfileString(Section, KEY, "ERROR", Value, Len(Value), ININame)
    GetINIValue = Left$(Value, InStr(1, Value, vbNullChar) - 1)
End Function

Private Sub LoadIniFile()
    Dim wScriptHost As Object
    Dim mydoc_path As String
    Dim strWidth As String
    Dim strHeight As String
    Dim strMaxFile As String
    Dim strInitialString As String
    Dim strSaveRECURSIVE As String
    Dim strSavePATH As String
    Dim strSavePATHLIST As String

    Set wScriptHost = CreateObject("WScript.Shell")
    mydoc_path = wScriptHost.SpecialFolders("MyDocuments")

    '初期化関連読み出し
    strWidth = GetINIValue("WIDTH", "Initial", mydoc_path & "\ExcelFileOpener.ini")
    strHeight = GetINIValue("HEIGHT", "Initial", mydoc_path & "\ExcelFileOpener.ini")
    strMaxFile = GetINIValue("MAX_FILE", "Initial", mydoc_path & "\ExcelFileOpener.ini")
    strInitialString = GetINIValue("INITIAL_STRING", "Initial", mydoc_path & "\ExcelFileOpener.ini")

    If Not strWidth = "ERROR" Then
        iniWidth = Val(strWidth)
    End If

    If Not strHeight = "ERROR" Then
        iniHeight = Val(strHeight)
    End If

    If Not strMaxFile = "ERROR" Then
        maxCount = Val(strMaxFile)
    End If
    
    If Not strInitialString = "ERROR" Then
        initialString = strInitialString
    End If
    
    'セーブ関連復帰
    strSaveRECURSIVE = GetINIValue("RECURSIVE", "Save", mydoc_path & "\ExcelFileOpener.ini")
    strSavePATH = GetINIValue("PATH", "Save", mydoc_path & "\ExcelFileOpener.ini")
    
    '前回のカレントパスに復帰
    If Not strSaveRECURSIVE = "ERROR" Then
        prevPath = strSavePATH
    End If
        
    '再帰フラグ復帰 -> 重いので初回起動はオミット
    'If UCase(strSaveRECURSIVE) = "TRUE" Then
    '    recursiveFlag = True
    'End If
    
    'パス履歴復帰
    Dim pl As Variant
    Dim i As Long
    
    strSavePATHLIST = GetINIValue("PATHLIST", "Save", mydoc_path & "\ExcelFileOpener.ini")
    
    If Not strSavePATHLIST = "ERROR" Then
    
        pl = Split(strSavePATHLIST, ";")
        For i = 0 To UBound(pl)
            If Not pathDic.Exists(pl(i)) Then
                pathDic.Add KEY:=pl(i), Item:=1
            End If
        Next i
    
    End If
    
End Sub

'
' Iniファイル書き込み
'
Public Function SetINIValue(Value As String, KEY As String, Section As String, ININame As String) As Boolean

    Dim ret As Long

    ret = WritePrivateProfileString(Section, KEY, Value, ININame)
    SetINIValue = CBool(ret)

End Function

Private Sub SaveIniFile()

    Dim wScriptHost As Object
    Dim mydoc_path As String
    Dim ret As Boolean
    Dim i As Long
    Dim pathlist As String

    Set wScriptHost = CreateObject("WScript.Shell")
    mydoc_path = wScriptHost.SpecialFolders("MyDocuments")

    ' 前回パス保存
    ret = SetINIValue(crntPath, "PATH", "Save", mydoc_path & "\ExcelFileOpener.ini")

    ' 再帰フラグ保存
    If recursiveFlag = True Then
        ret = SetINIValue("True", "RECURSIVE", "Save", mydoc_path & "\ExcelFileOpener.ini")
    Else
        ret = SetINIValue("False", "RECURSIVE", "Save", mydoc_path & "\ExcelFileOpener.ini")
    End If
    
    ' 履歴パス保存
    pathlist = ""
    For i = 0 To UBound(pathDic.Keys)
        pathlist = pathlist & ";" & pathDic.Keys(i)
    Next i
    ret = SetINIValue(pathlist, "PATHLIST", "Save", mydoc_path & "\ExcelFileOpener.ini")

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
        Call GetFilesByMode
        
        ' フォーム表示前更新
        Call UserForm2.TextBox2_Change
        
        'モード別に開く処理を変更する
        Select Case noMode
        Case mode.ACTIVE_PATH
            UserForm2.TextBox2.Text = initialString
        Case mode.PREVIOUS_PATH
            UserForm2.TextBox2.Text = initialString
        Case mode.RECENT_FILE
            UserForm2.TextBox2.Text = ""
        Case mode.SWITCH_BOOK
            UserForm2.TextBox2.Text = ""
        End Select
        
        
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
    Case mode.PREVIOUS_PATH
        act_open = True
    Case mode.RECENT_FILE
        act_open = True
    Case mode.SWITCH_BOOK
        act_open = False
    End Select
    
    If act_open = True Then
        On Error Resume Next  'エラーがあっても続行する
        Workbooks.Open tgtfile, ReadOnly:=False, Notify:=False
        
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
        crntPath = activePath
        UserForm2.OptionButton1 = True
    
    Case mode.PREVIOUS_PATH
        crntPath = prevPath
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
        ' 入力パスを保存
        Call SaveIniFile
        Call ExitGlobal
        Exit Sub
    End If
    
    'ファイルオープン処理
    Call OpenFileSub(tgtfile)
    
    ' 入力パスを保存
    Call SaveIniFile
    Call ExitGlobal
    
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
' 公開関数（前回のフォルダを開く）
'
Public Sub OpenFile_PREVIOUS_PATH()
    Call OpenFile0(mode.PREVIOUS_PATH)
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

