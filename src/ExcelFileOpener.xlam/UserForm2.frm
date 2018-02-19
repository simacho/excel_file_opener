VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "ExcelFileOpener"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11040
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Windows API�錾
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long



Private Sub UnitKeyDownReturn()
    If nodeCount = 0 Then
        selectedName = ""
    Else
        Select Case noMode
        Case mode.ACTIVE_PATH
            selectedName = Combine(crntPath, ListView1.SelectedItem.SubItems(1), ListView1.SelectedItem.Text)
        Case mode.PREVIOUS_PATH
            selectedName = Combine(crntPath, ListView1.SelectedItem.SubItems(1), ListView1.SelectedItem.Text)
        Case mode.RECENT_FILE
            selectedName = Combine(ListView1.SelectedItem.SubItems(1), ListView1.SelectedItem.Text)
        Case mode.SWITCH_BOOK
            selectedName = ListView1.SelectedItem.Text
        End Select
    
    End If

End Sub


Private Sub CheckBoxRecursive_Click()
    
    If Not (UserForm2.ActiveControl Is Nothing) Then
        If UserForm2.ActiveControl.Name = Me.ActiveControl.Name Then
            recursiveFlag = Not recursiveFlag
        
            Call GetFilesByMode
            Call TextBox2_Change    ' ���e�X�V

        End If
    End If
    
End Sub

Private Sub ComboBoxDirbox_Change()

End Sub

Private Sub LabelCounter_Click()

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
    
    TextBox2.Text = initialString
    crntPath = activePath
    selectedName = ""
    
    ComboBoxDirbox.ForeColor = &H80000012        ' �p�X�͔Z��
    Call GetFilesByMode
    
    Call TextBox2_Change    ' ���e�X�V
End Sub

Private Sub OptionButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        waitFlag = False
        escFlag = True
    End If
End Sub

Private Sub OptionButton2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call OptionButton1_KeyDown(KeyCode, Shift)
End Sub

Private Sub OptionButton3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call OptionButton1_KeyDown(KeyCode, Shift)
End Sub

Private Sub OptionButton4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call OptionButton1_KeyDown(KeyCode, Shift)
End Sub

Private Sub OptionButton2_Click()
    
    noMode = mode.PREVIOUS_PATH
    
    TextBox2.Text = initialString
    crntPath = prevPath
    selectedName = ""
    ComboBoxDirbox.ForeColor = &H80000012        ' �p�X�͔Z��
        
    Call GetFilesByMode
    Call TextBox2_Change    ' ���e�X�V
End Sub

Private Sub OptionButton3_Click()
    noMode = mode.RECENT_FILE
    
    TextBox2.Text = ""
    selectedName = ""
    ComboBoxDirbox.ForeColor = &H80000010        ' �p�X�͔���
    
    Call GetFilesByMode
    Call TextBox2_Change    ' ���e�X�V
End Sub

Private Sub OptionButton4_Click()
    noMode = mode.SWITCH_BOOK
    
    TextBox2.Text = ""
    selectedName = ""
    ComboBoxDirbox.ForeColor = &H80000010        ' �p�X�͔���
    
    Call GetFilesByMode
    Call TextBox2_Change    ' ���e�X�V
End Sub

Private Sub OptionButton5_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

'
' �i���֘A
'
Public Sub TextBox2_Change()
    
    Dim cnt As Integer
    Dim matchstr As String
    Dim buf As Variant
    Dim searchstr As String
    Dim fso As New FileSystemObject
    Dim tempstr As String
            
    ' �v�f������
    'If UBound(filesBuffer) - LBound(filesBuffer) = 0 Then
    '    Exit Sub
    'End If
    
    ' ���X�g�̃N���A
    UserForm2.ListView1.ListItems.Clear
    
    ComboBoxDirbox.Text = crntPath

    matchstr = "*" & UserForm2.TextBox2.Value & "*"
                    
    cnt = 0
    
   For Each buf In filesBuffer()
        If MatchCheck2(CStr(buf), UserForm2.TextBox2.Value) Then
       ' If MatchCheckRegExp(CStr(buf), UserForm2.TextBox2.Value) Then
            
            Dim fpath As String
            Dim itmWork As ListItem
    
            Set itmWork = ListView1.ListItems.Add   '�s�ǉ��A������ListItem�I�u�W�F�N�g�ϐ��ɑ��
                
            If noMode = mode.SWITCH_BOOK Then
                ' �J���Ă���u�b�N�̌����͂��̂܂ܓo�^
                itmWork.Text = buf
            Else
                ' ��������(�t�@�C��)�̓t�@�C�����ƃt�H���_���𕪂��ĕ\��
                itmWork.Text = fso.GetFileName(buf)
                itmWork.SubItems(1) = fso.GetParentFolderName(buf)
            End If
        
            ' ���\�[�X�J��
            Set itmWork = Nothing
        
            '�@���ڐ�
            cnt = cnt + 1
        
        End If
    
    Next buf

    '����\�����Ă���
    If amountFile > maxCount Then
        UserForm2.LabelCounter.Caption = CStr(cnt) & "/" & CStr(amountFile) & " more are omitted!!!"
    Else
        UserForm2.LabelCounter.Caption = CStr(cnt) & "/" & CStr(amountFile) & " matched"
    End If

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

'
' �_�u���N���b�N�Ńt�H���_�I��
'
Private Sub ComboBoxDirbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            ComboBoxDirbox.Text = .SelectedItems(1)
        End If
    
            crntPath = ComboBoxDirbox.Text       ' �e�L�X�g�{�b�N�X�ŃJ�����g�p�X���㏑��
            selectedName = ""
            waitFlag = False
    
            ' �R���{�{�b�N�X�ɒǉ�
            If Not pathDic.Exists(ComboBoxDirbox.Text) Then
                pathDic.Add KEY:=ComboBoxDirbox.Text, Item:=1
            End If
    
    End With
End Sub

'
' �p�X�̕ύX�m�F
'
Private Sub ComboBoxDirbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyReturn Then
        crntPath = ComboBoxDirbox.Text       ' �e�L�X�g�{�b�N�X�ŃJ�����g�p�X���㏑��
        selectedName = ""
        waitFlag = False
    
            ' �R���{�{�b�N�X�ɒǉ�
            If Not pathDic.Exists(ComboBoxDirbox.Text) Then
                pathDic.Add KEY:=ComboBoxDirbox.Text, Item:=1
            End If
    End If
    
    If KeyCode = vbKeyEscape Then
        escFlag = True
        waitFlag = False
    End If

End Sub

Private Sub UserForm_Activate()
    
    Me.TextBox2.SetFocus
    Call FormSetting

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'Form������Ƃ�
  
    If CloseMode = 0 Then       ' X�{�^���������͋����I�ɏI��
        End
    End If

End Sub



' �t�H�[�������T�C�Y�\�ɂ��邽�߂̐ݒ�
Public Sub FormSetting()
    Dim Result As Long
    Dim hwnd As Long
    Dim Wnd_STYLE As Long
 
    hwnd = GetActiveWindow()
    Wnd_STYLE = GetWindowLong(hwnd, GWL_STYLE)
    Wnd_STYLE = Wnd_STYLE Or WS_THICKFRAME Or &H30000
 
    Result = SetWindowLong(hwnd, GWL_STYLE, Wnd_STYLE)
    Result = DrawMenuBar(hwnd)
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    
 With ListView1
    .AllowColumnReorder = True
    .BorderStyle = ccFixedSingle
    .OLEDragMode = ccOLEDragAutomatic
    .OLEDropMode = ccOLEDropManual
    .Gridlines = True
    .FullRowSelect = True
    .View = lvwReport
        
    '�r���[�̐擪��̕\��
    .ColumnHeaders.Add 1, "F", "File", iniWidth / 2
    .ColumnHeaders.Add 2, "D", "Appendix", iniWidth / 2
 End With

 ' ���[�U�[�T�C�Y�ύX
 Me.Width = iniWidth
 Me.Height = iniHeight
 
 ' �ۑ��t���O�ύX
 Me.CheckBoxRecursive.Value = recursiveFlag
 
 ' �p�X����ۑ�
 For i = 0 To pathDic.Count - 1
    Me.ComboBoxDirbox.AddItem (pathDic.Keys(i))
 Next i

 ' �T�C�Y�ύX
 Call UserForm_Resize

End Sub


'
' �E�B���h�E�T�C�Y����
'
Private Sub UserForm_Resize()
    Dim mgn As Long
    Dim xx As Variant
    Dim yy As Variant
    
    xx = Array(8, 38, 16)
    yy = Array(12, 36, 60, 30, 12)
    
    'wdWidth = Me.Width
    'wdHeight = Me.Height
    
    ' x0 �����킹
    Label2.Left = xx(0)
    Label4.Left = xx(0)
    Label6.Left = xx(0)
    
    ' x1 �����킹
    ComboBoxDirbox.Left = xx(1)
    TextBox2.Left = xx(1)
    ListView1.Left = xx(1)
        
    OptionButton1.Left = xx(1)
    OptionButton2.Left = xx(1) + 100
    OptionButton3.Left = xx(1) + 200
    OptionButton4.Left = xx(1) + 300
        
    ' x2 �E���킹
    If Me.InsideWidth > 200 Then
           
        CheckBoxRecursive.Width = 60
        CheckBoxRecursive.Left = Me.InsideWidth - xx(2) - CheckBoxRecursive.Width
        
        ComboBoxDirbox.Width = Me.InsideWidth - xx(1) - xx(2) - CheckBoxRecursive.Width
        TextBox2.Width = Me.InsideWidth - xx(1) - xx(2)
        ListView1.Width = Me.InsideWidth - xx(1) - xx(2)
        ListView1.ColumnHeaders.Item(1).Width = ListView1.Width / 2
        ListView1.ColumnHeaders.Item(2).Width = ListView1.Width / 2
        
        LabelCounter.Width = 200
        LabelCounter.Left = Me.InsideWidth - xx(2) - LabelCounter.Width
        
    End If
        
    ' y0 �㍇�킹
    Label4.Top = yy(0)
    ComboBoxDirbox.Top = yy(0)
    CheckBoxRecursive.Top = yy(0)
    ' y1 �㍇�킹
    Label2.Top = yy(1)
    TextBox2.Top = yy(1)
    ' y2 �㍇�킹
    Label6.Top = yy(2)
    ListView1.Top = yy(2)
    
    ' y3 �����킹
    If Me.InsideHeight > 200 Then
        ListView1.Height = Me.InsideHeight - yy(2) - yy(3)
    
        OptionButton1.Top = Me.InsideHeight - yy(3)
        OptionButton2.Top = Me.InsideHeight - yy(3)
        OptionButton3.Top = Me.InsideHeight - yy(3)
        OptionButton4.Top = Me.InsideHeight - yy(3)
        LabelCounter.Top = Me.InsideHeight - yy(3)
    End If


End Sub
