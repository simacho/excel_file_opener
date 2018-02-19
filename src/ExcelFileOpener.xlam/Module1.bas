Attribute VB_Name = "Module1"
Option Explicit
' ------------------------------------------------------------
'  �ϐ���`
' ------------------------------------------------------------
Public activePath As String         ' �A�N�e�B�u�p�X�Ǘ��p
Public crntPath As String           ' ���݂̃p�X�Ǘ��p
Public prevPath As String           ' �O��̃p�X�Ǘ��p
Public initialString As String      ' �ŏ��̕�����
Public selectedName As String       ' �I�����ږ��󂯓n���p
Public nodeCount As Long            ' ���ڐ�
Public noMode As Integer            ' ���[�h�Ǘ�
Public filesBuffer() As String      ' ���X�g�\���o�b�t�@
Public maxCount As Long             ' �ċA���̃t�@�C�����
Public amountFile As Long
Public pathDic As New Dictionary         ' �p�X�ۑ��p
    
' �t���O�n
Public waitFlag As Boolean
Public escFlag As Boolean
Public recursiveFlag As Boolean

' INI�t�@�C���p
Public iniWidth As Long
Public iniHeight As Long

' ------------------------------------------------------------
'  �萔��`
' ------------------------------------------------------------
Enum mode
    ACTIVE_PATH = 1
    PREVIOUS_PATH = 2
    RECENT_FILE = 3
    SWITCH_BOOK = 4
End Enum

' ------------------------------------------------------------
'  INI�t�@�C���֘A
' ------------------------------------------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' ------------------------------------------------------------
'  ������
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
    
    ' �t�@�C���o�b�t�@�N���A
    ReDim filesBuffer(0)
    
    ' Path�p�̎���
    Set pathDic = CreateObject("Scripting.Dictionary")
        
    ' �t���O�֘A�N���A
    escFlag = False
    waitFlag = False
    recursiveFlag = False
    
    ' �����l
    iniWidth = 500
    iniHeight = 300
    maxCount = 10000
    
End Sub


Private Sub ExitGlobal()
    Set pathDic = Nothing
End Sub


'
' Ini�t�@�C���ǂݍ���
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

    '�������֘A�ǂݏo��
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
    
    '�Z�[�u�֘A���A
    strSaveRECURSIVE = GetINIValue("RECURSIVE", "Save", mydoc_path & "\ExcelFileOpener.ini")
    strSavePATH = GetINIValue("PATH", "Save", mydoc_path & "\ExcelFileOpener.ini")
    
    '�O��̃J�����g�p�X�ɕ��A
    If Not strSaveRECURSIVE = "ERROR" Then
        prevPath = strSavePATH
    End If
        
    '�ċA�t���O���A
    If UCase(strSaveRECURSIVE) = "TRUE" Then
        recursiveFlag = True
    End If
    
    '�p�X���𕜋A
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
' Ini�t�@�C����������
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

    ' �O��p�X�ۑ�
    ret = SetINIValue(crntPath, "PATH", "Save", mydoc_path & "\ExcelFileOpener.ini")

    ' �ċA�t���O�ۑ�
    If recursiveFlag = True Then
        ret = SetINIValue("True", "RECURSIVE", "Save", mydoc_path & "\ExcelFileOpener.ini")
    Else
        ret = SetINIValue("False", "RECURSIVE", "Save", mydoc_path & "\ExcelFileOpener.ini")
    End If
    
    ' �����p�X�ۑ�
    pathlist = ""
    For i = 0 To UBound(pathDic.Keys)
        pathlist = pathlist & ";" & pathDic.Keys(i)
    Next i
    ret = SetINIValue(pathlist, "PATHLIST", "Save", mydoc_path & "\ExcelFileOpener.ini")

End Sub


'---------------------------------------------------------------------------------------------------
'
' �p�X��̃t�@�C����I������
'
Private Function SelectFile() As String

    Dim ret As Boolean
    Dim tgtfile As String
    Dim form As Boolean
    Dim dummy As Variant
    
    ret = False
    
    form = False
    
    Do
        ' �����擾����
        Call GetFilesByMode
        
        ' �t�H�[���\���O�X�V
        Call UserForm2.TextBox2_Change
        
        '���[�h�ʂɊJ��������ύX����
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
                
        ' �t�H�[���̕\��
        If form = False Then
            form = True
            UserForm2.Show (vbModeless)
        End If

        ' �t�H�[���\����X�V
        UserForm2.TextBox2.SetFocus
 
        ' ���͑҂�
        Do While waitFlag
            dummy = DoEvents
        Loop
                        
        ' ------------------------------
        ' ESC�L�[���ʏ���
        If escFlag = True Then
            SelectFile = ""
            GoTo LastExit
        End If

        ' �w�蕨��������Ȃ��ꍇ�͌J��Ԃ�
        If selectedName = "" Then
            GoTo Continue
        End If
                 
        ' �I��
        tgtfile = selectedName
        ret = True

Continue:
    Loop While ret = False


LastExit:
    ' �t�H�[��������
    Unload UserForm2

    SelectFile = tgtfile
End Function


'
' ���[�h�ʂ̃I�[�v������
'
Private Sub OpenFileSub(tgtfile As String)

    Dim act_open As Boolean
    Dim idx As Integer
    Dim books() As String
    act_open = True
    
    ' Debug.Print "Open " & tgtfile
    
    '���[�h�ʂɊJ��������ύX����
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
        On Error Resume Next  '�G���[�������Ă����s����
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
              
    ' �O���[�o���ϐ��̏�����
    Call InitGlobal
    
    ' INI�t�@�C���̓ǂݍ���
    Call LoadIniFile

    ' ���[�h�w��
    noMode = mno
            
    '���[�h�ʏ���
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
    
    '�I������
    tgtfile = SelectFile()
        
    If escFlag = True Then
        ' ���̓p�X��ۑ�
        Call SaveIniFile
        Call ExitGlobal
        Exit Sub
    End If
    
    '�t�@�C���I�[�v������
    Call OpenFileSub(tgtfile)
    
    ' ���̓p�X��ۑ�
    Call SaveIniFile
    Call ExitGlobal
    
End Sub

'======================================================================
'
' ���J����
'
'======================================================================
'
' ���J�֐��i���݂̃t�@�C���Ɠ����t�H���_���J���j
'
Public Sub OpenFile_ACTIVE_PATH()
    Call OpenFile0(mode.ACTIVE_PATH)
End Sub

'
' ���J�֐��i�O��̃t�H���_���J���j
'
Public Sub OpenFile_PREVIOUS_PATH()
    Call OpenFile0(mode.PREVIOUS_PATH)
End Sub

'
' ���J�֐��i�����t�@�C����I�����ĊJ���j
'
Public Sub OpenFile_RECENT_FILE()
    Call OpenFile0(mode.RECENT_FILE)
End Sub

'
' ���J�֐��i�����t�@�C����I�����ĊJ���j
'
Public Sub OpenFile_SWITCH_BOOK()
    Call OpenFile0(mode.SWITCH_BOOK)
End Sub

