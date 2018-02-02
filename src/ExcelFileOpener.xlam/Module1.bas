Attribute VB_Name = "Module1"
Option Explicit
' ------------------------------------------------------------
'  変数定義
' ------------------------------------------------------------
Public userPath As String           ' アクティブパス管理用
Public crntPath As String           ' 現在のパス管理用
Public selectedName As String       ' 選択項目名受け渡し用
Public nodeCount As Long            ' 項目数
Public noMode As Integer            ' モード管理
Public filesBuffer() As String      ' リスト表示バッファ

' フラグ系
Public waitFlag As Boolean
Public escFlag As Boolean
Public nodirFlag As Boolean

' ------------------------------------------------------------
'  定数定義
' ------------------------------------------------------------
Enum mode
    ACTIVE_PATH = 1
    RECENT_FILE = 2
    ACTIVE_BOOK = 3
End Enum

' ------------------------------------------------------------
'  初期化
' ------------------------------------------------------------
Private Sub InitGlobal()
    userPath = ""
    crntPath = ""
    selectedName = ""
    nodeCount = 0
    noMode = mode.ACTIVE_PATH
    escFlag = False
    nodirFlag = False
    waitFlag = False
End Sub

'---------------------------------------------------------------------------------------------------

'
' パス上のファイルを選択する
'
Private Function SelectFile(path As String) As String

    Dim ret As Boolean
    Dim tgtfile As String
    Dim form As Boolean
    Dim dummy As Variant
    
    ret = False
    
    form = False
    
    Do
        ' 受け渡し用パスにもセット
        crntPath = path
        
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
        
        ' フォルダ無し例外対応
        If nodirFlag = True Then
            path = ""
            nodirFlag = False
            GoTo Continue
        End If
                        
        ' 指定物が見つからない場合は繰り返し
        If selectedName = "" Then
            GoTo Continue
        End If
                 
        tgtfile = ""
                
        Select Case noMode
        Case mode.ACTIVE_PATH
            tgtfile = path & "\" & selectedName
                        
            If path = "" Then
                'トップディレクトリにいる場合
                path = selectedName
            Else
                'フォルダが存在している場合
                tgtfile = path & "\" & selectedName
                ' フォルダ遷移
                If selectedName = ".." Then
                ' 選択した物が..だった場合には親のフォルダに行く
                    path = GetParentFolder(path)
                ElseIf ArgumentTypeCheck(tgtfile) = 0 Then
                    path = tgtfile
                Else
                    'ファイル指定であった場合
                    ret = True
                End If
            End If
                                
        Case mode.RECENT_FILE
            tgtfile = selectedName
            ret = True
        Case mode.ACTIVE_BOOK
            tgtfile = selectedName
            ret = True
        End Select
                        

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
    Case mode.RECENT_FILE
        act_open = True
    Case mode.ACTIVE_BOOK
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
    
    ' モード指定
    noMode = mno
    
    If Not ActiveWorkbook Is Nothing Then
        userPath = ActiveWorkbook.path
    End If
        
    'モード別処理
    Select Case noMode
    Case mode.ACTIVE_PATH
        UserForm2.OptionButton1 = True
    Case mode.RECENT_FILE
        UserForm2.OptionButton3 = True
    Case mode.ACTIVE_BOOK
        UserForm2.OptionButton4 = True
    End Select
        
    'Debug.Print "Initial UserPath " & userPath
    
    '選択処理
    tgtfile = SelectFile(userPath)
        
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
' 公開関数（履歴ファイルを選択して開く）
'
Public Sub OpenFile_RECENT_FILE()
    Call OpenFile0(mode.RECENT_FILE)
End Sub

'
' 公開関数（履歴ファイルを選択して開く）
'
Public Sub OpenFile_ACTIVE_BOOK()
    Call OpenFile0(mode.ACTIVE_BOOK)
End Sub





