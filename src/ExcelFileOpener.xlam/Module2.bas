Attribute VB_Name = "Module2"
Option Explicit

' ------------------------------------------------------------
'  ���ʊ֐�
' ------------------------------------------------------------

' ------------------------------------------------------------
'  �L�[���͊֘A
' ------------------------------------------------------------


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
    Dim fold As folder
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
' �t�@�C���̈ꗗ���ċA�I�Ɏ擾����
'
Function GetFilesRecursive(path As String, files() As String, cnt As Long, rcsv As Boolean) As String()
    
    Dim fso As FileSystemObject
    Dim fold As folder
    Dim file As Object
    Dim dummy As Variant
    Dim temp As String
    
    On Error GoTo ErrorHandler
            
    ' �t�@�C���V�X�e���擾
    Set fso = New FileSystemObject
    Set fold = fso.GetFolder(path)
 
    ' �t�@�C���̖��O���擾
    For Each file In fold.files
        ReDim Preserve files(cnt)
        
        'files(cnt) = Combine(path, file.name)
    
        files(cnt) = Combine(Replace(path, crntPath, ""), file.Name)
        
        'Debug.Print "Recursive " & Combine(Replace(path, crntPath, ""), file.name)
        
        cnt = cnt + 1
    Next file
    
    '
    ' ���ԑ҂�
    '
    UserForm2.LabelCounter.Caption = "Scanning " & cnt
    dummy = DoEvents

    If cnt > maxCount Then
        GoTo ErrorHandler
    End If
    
   ' �T�u�t�H���_�ŌĂяo��
   If rcsv = True Then
        For Each file In fold.SubFolders
            files = GetFilesRecursive(file.path, files, cnt, rcsv)
        Next file
    End If


ErrorHandler:
    
    ' ����ۑ�
    amountFile = cnt
          
    Set fso = Nothing
    GetFilesRecursive = files
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

    ' ����ۑ�
    amountFile = cnt


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
          
            files(cnt) = Application.RecentFiles(i).path
            
            cnt = cnt + 1
        Next i
    End If

    ' ����ۑ�
    amountFile = cnt
    

    GetRecentlyFiles = files
    
End Function

'
' �t�@�C��������
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
' �}�b�`�֐� ����2 (!���擪�s�ɗ��Ă����玸�s�ɂ���)
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
            ' !�Ŕr���t���O����
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
' �}�b�`�֐� ���K�\����

Function MatchCheckRegExp(str As String, chkstr As String) As Boolean
    Dim reg             As New RegExp       '// ���K�\���N���X�I�u�W�F�N�g
    Dim oMatches        As MatchCollection  '// RegExp.Execute����
    Dim oMatch          As Match            '// �������ʃI�u�W�F�N�g
    
    Dim spells As Variant
    Dim spell As String
    Dim i As Long
    Dim ignore As Boolean
    
    
    '// ���������ݒ�
    reg.Global = True               '// �����͈́iTrue�F������̍Ō�܂Ō����AFalse�F�ŏ��̈�v�܂Ō����j
    reg.IgnoreCase = True           '// �啶���������̋�ʁiTrue�F��ʂ��Ȃ��AFalse�F��ʂ���j
    reg.Pattern = chkstr            '// �����p�^�[���i�����ł͘A�����鐔�������������ɐݒ�j

    '// �������s
    Set oMatches = reg.Execute(str)
    
    If oMatches.Count >= 1 Then
        MatchCheckRegExp = True
    Else
        MatchCheckRegExp = False
    End If

End Function

'
' ���[�h�ʂɌ����擾����
'
Sub GetFilesByMode()

    ' �z���������
    ReDim filesBuffer(0)

    '���[�h�ʂɊJ��������ύX����
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


    
