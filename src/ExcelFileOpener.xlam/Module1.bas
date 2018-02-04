Attribute VB_Name = "Module1"
Option Explicit
' ------------------------------------------------------------
'  �ϐ���`
' ------------------------------------------------------------
Public activePath As String         ' �A�N�e�B�u�p�X�Ǘ��p
Public crntPath As String           ' ���݂̃p�X�Ǘ��p
Public selectedName As String       ' �I�����ږ��󂯓n���p
Public nodeCount As Long            ' ���ڐ�
Public noMode As Integer            ' ���[�h�Ǘ�
Public filesBuffer() As String      ' ���X�g�\���o�b�t�@
Public maxCount As Long             ' �ċA���̃t�@�C�����
Public amountFile As Long


' �t���O�n
Public waitFlag As Boolean
Public escFlag As Boolean

' INI�t�@�C���p
Public iniWidth As Long
Public iniHeight As Long

' ------------------------------------------------------------
'  �萔��`
' ------------------------------------------------------------
Enum mode
    ACTIVE_PATH = 1
    RECURSIVE_PATH = 2
    RECENT_FILE = 3
    SWITCH_BOOK = 4
End Enum

' ------------------------------------------------------------
'  INI�t�@�C���֘A
' ------------------------------------------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long


' ------------------------------------------------------------
'  ������
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
    
    ' �����l
    iniWidth = 500
    iniHeight = 300
    maxCount = 10000
    
End Sub

'
' Ini�t�@�C������
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
        filesBuffer = GetFilesByMode(filesBuffer, noMode, crntPath)
        
        ' �t�H�[���\���O�X�V
        Call UserForm2.TextBox2_Change
        UserForm2.TextBox2.Text = ""
        
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
    Case mode.RECURSIVE_PATH
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
        UserForm2.OptionButton1 = True
    
    Case mode.RECURSIVE_PATH
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
        Exit Sub
    End If
    
    '�t�@�C���I�[�v������
    Call OpenFileSub(tgtfile)
    
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
' ���J�֐��i���݂̃t�H���_����ċA�I�ɊJ���j
'
Public Sub OpenFile_RECURSIVE_PATH()
    Call OpenFile0(mode.RECURSIVE_PATH)
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

