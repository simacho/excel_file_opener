Attribute VB_Name = "Module2"
Option Explicit

' ------------------------------------------------------------
'  共通関数
' ------------------------------------------------------------

' ------------------------------------------------------------
'  キー入力関連
' ------------------------------------------------------------


' ------------------------------------------------------------
'  ファイル関連
' ------------------------------------------------------------
'
' ファイルがディレクトリの判別
'
Public Function ArgumentTypeCheck(arg)

    Dim rtn
    Dim objFSO, strExName

    On Error Resume Next
    rtn = -1
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(arg) Then
        rtn = 0
    Else
        If objFSO.FileExists(arg) Then
            strExName = objFSO.GetExtensionName(arg)
            Select Case UCase(strExName)
                Case "LNK": rtn = 1
                Case "URL": rtn = 2
                Case Else: rtn = 3
            End Select
        End If
    End If
    Set objFSO = Nothing
    
    ' Debug.Print arg & rtn
        
    ArgumentTypeCheck = rtn
    
End Function


'
' 親のパスを取得する
'
Function GetParentFolder(crnt As String) As String
    Dim fso As New Scripting.FileSystemObject
    Dim parentFolder As String
    
    parentFolder = fso.GetParentFolderName(crnt)
    Set fso = Nothing
    
    '' Debug.Print "PARENT " & parentFoledr
    
    GetParentFolder = parentFolder
End Function

'
' ドライブの配列を返却する
'
Function GetDriveLetters(drives() As String) As String()
    Dim fso As Object
    Dim drv As Object
    Dim cnt As Integer
    
    Set fso = New FileSystemObject
  
    cnt = 0
    For Each drv In fso.drives
        ReDim Preserve drives(cnt)
        
        drives(cnt) = drv.DriveLetter & ":\"
        cnt = cnt + 1
    Next drv
     
    Set fso = Nothing
    GetDriveLetters = drives
End Function

'
' ファイルとフォルダの一覧を返却する
'
Function GetFolderFiles(path As String, files() As String) As String()
    Dim fso As FileSystemObject
    Dim fold As folder
    Dim file As Object
    Dim cnt As Integer
    
    
    On Error GoTo ErrorHandler
    
    ReDim files(0)
    cnt = 0
        
    ' ファイルシステム取得
    Set fso = New FileSystemObject
    
    ' 取得に失敗した場合はカレントをパスを初期化する
    If Not fso.FolderExists(path) Then
        Set fso = Nothing
        GetFolderFiles = files
        Exit Function
    End If
    
    ' 親フォルダも追加
    ReDim Preserve files(cnt)
    files(cnt) = ".."
    cnt = cnt + 1
            
    Set fold = fso.GetFolder(path)

    ' サブフォルダの名前を取得
    For Each file In fold.SubFolders
        ReDim Preserve files(cnt)
        files(cnt) = file.Name
        ' Debug.Print path & "\ SUBDIR -> " & files(cnt)
        
        cnt = cnt + 1
    Next file
     
    ' ファイルの名前を取得
    For Each file In fold.files
        ReDim Preserve files(cnt)
        files(cnt) = file.Name
        ' Debug.Print path & "\ FILE -> " & files(cnt)

        cnt = cnt + 1
    Next file
     
ErrorHandler:
          
    Set fso = Nothing
    GetFolderFiles = files
    
End Function

'
' ファイルの一覧を再帰的に取得する
'
Function GetFilesRecursive(path As String, files() As String, cnt As Long, rcsv As Boolean) As String()
    
    Dim fso As FileSystemObject
    Dim fold As folder
    Dim file As Object
    Dim dummy As Variant
    Dim temp As String
    
    On Error GoTo ErrorHandler
            
    ' ファイルシステム取得
    Set fso = New FileSystemObject
    Set fold = fso.GetFolder(path)
 
    ' ファイルの名前を取得
    For Each file In fold.files
        ReDim Preserve files(cnt)
        
        'files(cnt) = Combine(path, file.name)
    
        files(cnt) = Combine(Replace(path, crntPath, ""), file.Name)
        
        'Debug.Print "Recursive " & Combine(Replace(path, crntPath, ""), file.name)
        
        cnt = cnt + 1
    Next file
    
    '
    ' 時間待ち
    '
    UserForm2.LabelCounter.Caption = "Scanning " & cnt
    dummy = DoEvents

    If cnt > maxCount Then
        GoTo ErrorHandler
    End If
    
   ' サブフォルダで呼び出し
   If rcsv = True Then
        For Each file In fold.SubFolders
            files = GetFilesRecursive(file.path, files, cnt, rcsv)
        Next file
    End If


ErrorHandler:
    
    ' 個数を保存
    amountFile = cnt
          
    Set fso = Nothing
    GetFilesRecursive = files
End Function



'
' 開いているブックの一覧を返却する
'
Function GetWorkBookNames(files() As String) As String()
    Dim wbk As Workbook
    Dim file As Object
    Dim cnt As Integer
    
    On Error GoTo ErrorHandler
    
    ReDim files(0)
    cnt = 0

    ' ブック集合から取得
    For Each wbk In Workbooks
        ReDim Preserve files(cnt)
        files(cnt) = wbk.Name
        cnt = cnt + 1
    Next wbk

    ' 個数を保存
    amountFile = cnt


ErrorHandler:

    


    GetWorkBookNames = files
End Function


'
' 履歴一覧を返却する
'
Function GetRecentlyFiles(files() As String) As String()
    Dim FileCount As Long
    Dim i As Long
    Dim cnt As Integer
    
    ReDim files(0)
    cnt = 0
        
    FileCount = Application.RecentFiles.Count
    
    If FileCount > 1 Then
    
        For i = 1 To FileCount
            ReDim Preserve files(cnt)
          
            files(cnt) = Application.RecentFiles(i).path
            
            cnt = cnt + 1
        Next i
    End If

    ' 個数を保存
    amountFile = cnt
    

    GetRecentlyFiles = files
    
End Function

'
' ファイル名結合
'
Function Combine(ParamArray paths()) As String
    Dim i As Integer
    Dim path As String
    Dim Result As String
    For i = LBound(paths) To UBound(paths)
        path = CStr(paths(i))
        If i = LBound(paths) Then
            Result = path
        Else
            If Right(Result, 1) = "\" Then Result = Left(Result, Len(Result) - 1)
            If Left(path, 1) = "\" Then path = Mid(path, 2)
            Result = Result & "\" & path
        End If
    Next
     
    Combine = Result
End Function

'
' マッチ関数
'
Function MatchCheck(str As String, chkstr As String) As Boolean
    Dim spells As Variant
    Dim spell As String
    Dim i As Long
        
    spells = Split(chkstr, " ")
    
    For i = 0 To UBound(spells)
        spell = "*" & spells(i) & "*"

         If Not StrConv(UCase(str), vbNarrow) Like StrConv(UCase(spell), vbNarrow) Then
            MatchCheck = False
            Exit Function
         End If
    Next i
    
    MatchCheck = True
End Function


'
' マッチ関数 その2 (!が先頭行に来ていたら失敗にする)
'
Function MatchCheck2(str As String, chkstr As String) As Boolean
    Dim spells As Variant
    Dim spell As String
    Dim i As Long
    Dim ignore As Boolean
    
    spells = Split(chkstr, " ")
    
    For i = 0 To UBound(spells)
        spell = "*" & spells(i) & "*"

        ignore = False
        If InStr(spell, "!") > 0 Then
            ' !で排除フラグ判定
            ignore = True
            spell = Replace(spell, "!", "")
        
            If StrConv(UCase(str), vbNarrow) Like StrConv(UCase(spell), vbNarrow) Then
                MatchCheck2 = False
                Exit Function
            End If
        Else
            If Not StrConv(UCase(str), vbNarrow) Like StrConv(UCase(spell), vbNarrow) Then
                MatchCheck2 = False
                Exit Function
            End If
        End If
    Next i
    
    MatchCheck2 = True
End Function

'
' マッチ関数 正規表現版

Function MatchCheckRegExp(str As String, chkstr As String) As Boolean
    Dim reg             As New RegExp       '// 正規表現クラスオブジェクト
    Dim oMatches        As MatchCollection  '// RegExp.Execute結果
    Dim oMatch          As Match            '// 検索結果オブジェクト
    
    Dim spells As Variant
    Dim spell As String
    Dim i As Long
    Dim ignore As Boolean
    
    
    '// 検索条件設定
    reg.Global = True               '// 検索範囲（True：文字列の最後まで検索、False：最初の一致まで検索）
    reg.IgnoreCase = True           '// 大文字小文字の区別（True：区別しない、False：区別する）
    reg.Pattern = chkstr            '// 検索パターン（ここでは連続する数字を検索条件に設定）

    '// 検索実行
    Set oMatches = reg.Execute(str)
    
    If oMatches.Count >= 1 Then
        MatchCheckRegExp = True
    Else
        MatchCheckRegExp = False
    End If

End Function

'
' モード別に候補を取得する
'
Sub GetFilesByMode()

    ' 配列を初期化
    ReDim filesBuffer(0)

    'モード別に開く処理を変更する
    Select Case noMode
    Case mode.ACTIVE_PATH
        filesBuffer = GetFilesRecursive(crntPath, filesBuffer, 0, recursiveFlag)
    Case mode.PREVIOUS_PATH
        filesBuffer = GetFilesRecursive(crntPath, filesBuffer, 0, recursiveFlag)
    Case mode.RECENT_FILE
        filesBuffer = GetRecentlyFiles(filesBuffer)
    Case mode.SWITCH_BOOK    ' SWITCH_BOOK
        filesBuffer = GetWorkBookNames(filesBuffer)
    End Select

End Sub


    
