Attribute VB_Name = "Module2"
Option Explicit

' ------------------------------------------------------------
'  ���ʊ֐�
' ------------------------------------------------------------

' ------------------------------------------------------------
'  INI�t�@�C���֘A(���쒆)
' ------------------------------------------------------------
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
'
' Ini�t�@�C������
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
'  �L�[���͊֘A
' ------------------------------------------------------------
'
' �R�}���h�L�[����


Public Function KeyCheck(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    ' ����
    If KeyCode = vbKeyReturn Then
        selectedName = UserForm2.ListView1.SelectedItem.Text
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Function




' ------------------------------------------------------------
'  �t�@�C���֘A
' ------------------------------------------------------------
'
' �t�@�C�����f�B���N�g���̔���
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
' �e�̃p�X���擾����
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
' �h���C�u�̔z���ԋp����
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
' �t�@�C���ƃt�H���_�̈ꗗ��ԋp����
'
Function GetFolderFiles(path As String, files() As String) As String()
    Dim fso As FileSystemObject
    Dim fold As Folder
    Dim file As Object
    Dim cnt As Integer
    
    
    On Error GoTo ErrorHandler
    
    ReDim files(0)
    cnt = 0
        
    ' �t�@�C���V�X�e���擾
    Set fso = New FileSystemObject
    
    ' �擾�Ɏ��s�����ꍇ�̓J�����g���p�X������������
    If Not fso.FolderExists(path) Then
        Set fso = Nothing
        GetFolderFiles = files
        Exit Function
    End If
    
    ' �e�t�H���_���ǉ�
    ReDim Preserve files(cnt)
    files(cnt) = ".."
    cnt = cnt + 1
            
    Set fold = fso.GetFolder(path)

    ' �T�u�t�H���_�̖��O���擾
    For Each file In fold.SubFolders
        ReDim Preserve files(cnt)
        files(cnt) = file.Name
        ' Debug.Print path & "\ SUBDIR -> " & files(cnt)
        
        cnt = cnt + 1
    Next file
     
    ' �t�@�C���̖��O���擾
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
' �J���Ă���u�b�N�̈ꗗ��ԋp����
'
Function GetWorkBookNames(files() As String) As String()
    Dim wbk As Workbook
    Dim file As Object
    Dim cnt As Integer
    
    On Error GoTo ErrorHandler
    
    ReDim files(0)
    cnt = 0

    ' �u�b�N�W������擾
    For Each wbk In Workbooks
        ReDim Preserve files(cnt)
        files(cnt) = wbk.Name
        cnt = cnt + 1
    Next wbk

ErrorHandler:

    GetWorkBookNames = files
End Function


'
' �����ꗗ��ԋp����
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
' �t�@�C��������
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
' �}�b�`�֐�
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
' ���[�h�ʂɌ����擾����
'
Function GetFilesByMode(files() As String, mno As Integer, path As String)

    '���[�h�ʂɊJ��������ύX����
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


    
