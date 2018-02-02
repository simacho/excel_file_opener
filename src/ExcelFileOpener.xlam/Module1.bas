Attribute VB_Name = "Module1"
Option Explicit
' ------------------------------------------------------------
'  �ϐ���`
' ------------------------------------------------------------
Public userPath As String           ' �A�N�e�B�u�p�X�Ǘ��p
Public crntPath As String           ' ���݂̃p�X�Ǘ��p
Public selectedName As String       ' �I�����ږ��󂯓n���p
Public nodeCount As Long            ' ���ڐ�
Public noMode As Integer            ' ���[�h�Ǘ�
Public filesBuffer() As String      ' ���X�g�\���o�b�t�@

' �t���O�n
Public waitFlag As Boolean
Public escFlag As Boolean
Public nodirFlag As Boolean

' ------------------------------------------------------------
'  �萔��`
' ------------------------------------------------------------
Enum mode
    ACTIVE_PATH = 1
    RECENT_FILE = 2
    ACTIVE_BOOK = 3
End Enum

' ------------------------------------------------------------
'  ������
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
' �p�X��̃t�@�C����I������
'
Private Function SelectFile(path As String) As String

    Dim ret As Boolean
    Dim tgtfile As String
    Dim form As Boolean
    Dim dummy As Variant
    
    ret = False
    
    form = False
    
    Do
        ' �󂯓n���p�p�X�ɂ��Z�b�g
        crntPath = path
        
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
        
        ' �t�H���_������O�Ή�
        If nodirFlag = True Then
            path = ""
            nodirFlag = False
            GoTo Continue
        End If
                        
        ' �w�蕨��������Ȃ��ꍇ�͌J��Ԃ�
        If selectedName = "" Then
            GoTo Continue
        End If
                 
        tgtfile = ""
                
        Select Case noMode
        Case mode.ACTIVE_PATH
            tgtfile = path & "\" & selectedName
                        
            If path = "" Then
                '�g�b�v�f�B���N�g���ɂ���ꍇ
                path = selectedName
            Else
                '�t�H���_�����݂��Ă���ꍇ
                tgtfile = path & "\" & selectedName
                ' �t�H���_�J��
                If selectedName = ".." Then
                ' �I����������..�������ꍇ�ɂ͐e�̃t�H���_�ɍs��
                    path = GetParentFolder(path)
                ElseIf ArgumentTypeCheck(tgtfile) = 0 Then
                    path = tgtfile
                Else
                    '�t�@�C���w��ł������ꍇ
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
    Case mode.RECENT_FILE
        act_open = True
    Case mode.ACTIVE_BOOK
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
    
    ' ���[�h�w��
    noMode = mno
    
    If Not ActiveWorkbook Is Nothing Then
        userPath = ActiveWorkbook.path
    End If
        
    '���[�h�ʏ���
    Select Case noMode
    Case mode.ACTIVE_PATH
        UserForm2.OptionButton1 = True
    Case mode.RECENT_FILE
        UserForm2.OptionButton3 = True
    Case mode.ACTIVE_BOOK
        UserForm2.OptionButton4 = True
    End Select
        
    'Debug.Print "Initial UserPath " & userPath
    
    '�I������
    tgtfile = SelectFile(userPath)
        
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
' ���J�֐��i�����t�@�C����I�����ĊJ���j
'
Public Sub OpenFile_RECENT_FILE()
    Call OpenFile0(mode.RECENT_FILE)
End Sub

'
' ���J�֐��i�����t�@�C����I�����ĊJ���j
'
Public Sub OpenFile_ACTIVE_BOOK()
    Call OpenFile0(mode.ACTIVE_BOOK)
End Sub





