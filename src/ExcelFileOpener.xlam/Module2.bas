Attribute VB_Name = "Module2"
Option Explicit

' ------------------------------------------------------------
'  共通関数
' ------------------------------------------------------------

' ------------------------------------------------------------
'  INIファイル関連(制作中)
' ------------------------------------------------------------
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
'
' Iniファイル処理
'
'Function GetINIValue(KEY As String, Section As String, ININame As String) As String
'    Dim Value As String * 255
'    Call GetPrivateProfileString(Section, KEY, "ERROR", Value, Len(Value), ININame)
'    GetINIValue = Left$(Value, InStr(1, Value, vbNullChar) - 1)
'End Function
'Function LoadIniFile() As String
'    Dim wScriptHost As Object
'    Dim mydoc_path As String
'
'    Set wScriptHost = CreateObject("WScript.Shell")
'
'    mydoc_path = wScriptHost.SpecialFolders("MyDocuments")
'
'    LoadIniFile = GetINIValue("PATH", "Initial", mydoc_path & "\ExcelFileOpener.ini")
'
'End Function


' ------------------------------------------------------------
'  キー入力関連
' ------------------------------------------------------------
'
' コマンドキー判別


Public Function KeyCheck(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    ' 決定
    If KeyCode = vbKeyReturn Then
        selectedName = UserForm2.ListView1.SelectedItem.Text
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Function




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
    Dim fold As Folder
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
            
            files(cnt) = Application.RecentFiles(i).Name
            cnt = cnt + 1
        Next i
    End If

    GetRecentlyFiles = files
    
End Function

'
' ファイル名結合
'
Function Combine(ParamArray paths()) As String
    Dim i As Integer
    Dim path As String
    Dim result As String
    For i = LBound(paths) To UBound(paths)
        path = CStr(paths(i))
        If i = LBound(paths) Then
            result = path
        Else
            If Right(result, 1) = "\" Then result = Left(result, Len(result) - 1)
            If Left(path, 1) = "\" Then path = Mid(path, 2)
            result = result & "\" & path
        End If
    Next
     
    Combine = result
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
' モード別に候補を取得する
'
Function GetFilesByMode(files() As String, mno As Integer, path As String)

    'モード別に開く処理を変更する
    Select Case noMode
    Case mode.ACTIVE_PATH       ' ACTIVE_PATH
        If path = "" Then
            files = GetDriveLetters(files)
        Else
            files = GetFolderFiles(path, files)
            If UBound(files) = 0 Then
                path = ""
                nodirFlag = True
            End If
        End If
    Case mode.ACTIVE_BOOK    ' ACTIVE_BOOK
        files = GetWorkBookNames(files)
    
    Case mode.RECENT_FILE    ' ACTIVE_BOOK
        files = GetRecentlyFiles(files)
    End Select

    GetFilesByMode = files
End Function


    
