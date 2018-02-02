VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "ExcelFileOpener"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UnitKeyDownReturn()
    If nodeCount = 0 Then
        selectedName = ""
    Else
        selectedName = ListView1.SelectedItem.Text
        If noMode = mode.RECENT_FILE Then
            selectedName = Combine(ListView1.SelectedItem.SubItems(1), ListView1.SelectedItem.Text)
        End If
    End If

End Sub


Private Sub ListView1_DblClick()
    Call ListView1_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call UnitKeyDownReturn
        waitFlag = False
    End If
    
    If KeyCode = vbKeyEscape Then
        escFlag = True
        waitFlag = False
    End If
        
End Sub

'
' ���[�h�؂�ւ�
'
Private Sub OptionButton1_Click()
    noMode = mode.ACTIVE_PATH
    filesBuffer = GetFilesByMode(filesBuffer, noMode, crntPath)
    
    TextBox2.Text = ""
    
    Call TextBox2_Change    ' ���e�X�V
End Sub

Private Sub OptionButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        waitFlag = False
        escFlag = True
    End If
End Sub

Private Sub OptionButton3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call OptionButton1_KeyDown(KeyCode, Shift)
End Sub

Private Sub OptionButton4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call OptionButton1_KeyDown(KeyCode, Shift)
End Sub


Private Sub OptionButton3_Click()
    noMode = mode.RECENT_FILE
    
    filesBuffer = GetFilesByMode(filesBuffer, noMode, crntPath)
    Call TextBox2_Change    ' ���e�X�V
End Sub

Private Sub OptionButton4_Click()
    noMode = mode.ACTIVE_BOOK
    
    filesBuffer = GetFilesByMode(filesBuffer, noMode, crntPath)
    Call TextBox2_Change    ' ���e�X�V
End Sub

'
' �i���֘A
'
Public Sub TextBox2_Change()
    
    Dim cnt As Integer
    Dim matchstr As String
    Dim buf As Variant
    Dim files() As String
    Dim searchstr As String
    Dim fso As New FileSystemObject
    
    ' ��⃊�X�g���擾
    files = filesBuffer
    
    ' ���X�g�̃N���A
    UserForm2.ListView1.ListItems.Clear
    
    Label3.Caption = crntPath

    matchstr = "*" & UserForm2.TextBox2.Value & "*"
                    
    cnt = 0
    
    ' �t�@�C���Q
    For Each buf In files()
            
        If MatchCheck(CStr(buf), UserForm2.TextBox2.Value) Then
            Dim fpath As String
            Dim itmWork As ListItem
    
            Set itmWork = ListView1.ListItems.Add   '�s�ǉ��A������ListItem�I�u�W�F�N�g�ϐ��ɑ��
    
            Select Case noMode
            Case mode.ACTIVE_PATH
                ' �A�N�e�B�u�p�X�͒P����o�^�A�t�H���_�̏ꍇ�́i�j�ɕύX
                itmWork.Text = buf
                fpath = Combine(crntPath, buf)
                If ArgumentTypeCheck(fpath) = 0 Then
                    itmWork.ForeColor = vbBlue
                Else
                    itmWork.ForeColor = vbBlack
                End If
        
            Case mode.RECENT_FILE
                ' ���������̓t�@�C�����ƃt�H���_���𕪂��ĕ\��
                itmWork.Text = fso.GetFileName(buf)
                itmWork.SubItems(1) = fso.GetParentFolderName(buf)
            Case mode.ACTIVE_BOOK
                ' �J���Ă���u�b�N�̌����͂��̂܂ܓo�^
                itmWork.Text = buf
            End Select
        
            ' ���\�[�X�J��
            Set itmWork = Nothing
        
            '�@���ڐ�
            cnt = cnt + 1
        
        End If
    
    Next buf

    '�ύX���͐擪�s��I����ԂɕύX
    nodeCount = cnt
    
    If cnt > 0 Then
        ListView1.ListItems(1).Selected = True
    End If

    ' ���\�[�X�J��
    Set fso = Nothing

End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        
    If KeyCode = vbKeyReturn Then
        Call UnitKeyDownReturn
        waitFlag = False
    End If
    
    If KeyCode = vbKeyEscape Then
        escFlag = True
        waitFlag = False
    End If
        
End Sub


Private Sub UserForm_Activate()
    
    Me.TextBox2.SetFocus

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'Form������Ƃ�
  
    If CloseMode = 0 Then       ' X�{�^���������͋����I�ɏI��
        End
    End If

End Sub


Private Sub UserForm_Initialize()
 With ListView1
    
    .AllowColumnReorder = True
    .BorderStyle = ccFixedSingle
    .OLEDragMode = ccOLEDragAutomatic
    .OLEDropMode = ccOLEDropManual
    .Gridlines = True
    .FullRowSelect = True
    .View = lvwReport
        
    '�r���[�̐擪��̕\��
    .ColumnHeaders.Add 1, "F", "File", 250
    .ColumnHeaders.Add 2, "D", "Appendix", 400
 
 End With

 
 
End Sub

