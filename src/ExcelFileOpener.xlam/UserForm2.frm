VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "ExcelFileOpener"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11040
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Windows API宣言
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
            Call TextBox2_Change    ' 内容更新

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
' モード切り替え
'
Private Sub OptionButton1_Click()
    noMode = mode.ACTIVE_PATH
    
    TextBox2.Text = initialString
    crntPath = activePath
    selectedName = ""
    
    ComboBoxDirbox.ForeColor = &H80000012        ' パスは濃く
    Call GetFilesByMode
    
    Call TextBox2_Change    ' 内容更新
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
    ComboBoxDirbox.ForeColor = &H80000012        ' パスは濃く
        
    Call GetFilesByMode
    Call TextBox2_Change    ' 内容更新
End Sub

Private Sub OptionButton3_Click()
    noMode = mode.RECENT_FILE
    
    TextBox2.Text = ""
    selectedName = ""
    ComboBoxDirbox.ForeColor = &H80000010        ' パスは薄く
    
    Call GetFilesByMode
    Call TextBox2_Change    ' 内容更新
End Sub

Private Sub OptionButton4_Click()
    noMode = mode.SWITCH_BOOK
    
    TextBox2.Text = ""
    selectedName = ""
    ComboBoxDirbox.ForeColor = &H80000010        ' パスは薄く
    
    Call GetFilesByMode
    Call TextBox2_Change    ' 内容更新
End Sub

Private Sub OptionButton5_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

'
' 絞込関連
'
Public Sub TextBox2_Change()
    
    Dim cnt As Integer
    Dim matchstr As String
    Dim buf As Variant
    Dim searchstr As String
    Dim fso As New FileSystemObject
    Dim tempstr As String
            
    ' 要素無し時
    'If UBound(filesBuffer) - LBound(filesBuffer) = 0 Then
    '    Exit Sub
    'End If
    
    ' リストのクリア
    UserForm2.ListView1.ListItems.Clear
    
    ComboBoxDirbox.Text = crntPath

    matchstr = "*" & UserForm2.TextBox2.Value & "*"
                    
    cnt = 0
    
   For Each buf In filesBuffer()
        If MatchCheck2(CStr(buf), UserForm2.TextBox2.Value) Then
       ' If MatchCheckRegExp(CStr(buf), UserForm2.TextBox2.Value) Then
            
            Dim fpath As String
            Dim itmWork As ListItem
    
            Set itmWork = ListView1.ListItems.Add   '行追加、同時にListItemオブジェクト変数に代入
                
            If noMode = mode.SWITCH_BOOK Then
                ' 開いているブックの検索はそのまま登録
                itmWork.Text = buf
            Else
                ' 履歴検索(ファイル)はファイル名とフォルダ名を分けて表示
                itmWork.Text = fso.GetFileName(buf)
                itmWork.SubItems(1) = fso.GetParentFolderName(buf)
            End If
        
            ' リソース開放
            Set itmWork = Nothing
        
            '　項目数
            cnt = cnt + 1
        
        End If
    
    Next buf

    '個数を表示しておく
    If amountFile > maxCount Then
        UserForm2.LabelCounter.Caption = CStr(cnt) & "/" & CStr(amountFile) & " more are omitted!!!"
    Else
        UserForm2.LabelCounter.Caption = CStr(cnt) & "/" & CStr(amountFile) & " matched"
    End If

    '変更時は先頭行を選択状態に変更
    nodeCount = cnt
    
    If cnt > 0 Then
        ListView1.ListItems(1).Selected = True
    End If

    ' リソース開放
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
' ダブルクリックでフォルダ選択
'
Private Sub ComboBoxDirbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            ComboBoxDirbox.Text = .SelectedItems(1)
        End If
    
            crntPath = ComboBoxDirbox.Text       ' テキストボックスでカレントパスを上書き
            selectedName = ""
            waitFlag = False
    
            ' コンボボックスに追加
            If Not pathDic.Exists(ComboBoxDirbox.Text) Then
                pathDic.Add KEY:=ComboBoxDirbox.Text, Item:=1
            End If
    
    End With
End Sub

'
' パスの変更確認
'
Private Sub ComboBoxDirbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyReturn Then
        crntPath = ComboBoxDirbox.Text       ' テキストボックスでカレントパスを上書き
        selectedName = ""
        waitFlag = False
    
            ' コンボボックスに追加
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


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'Formが閉じるとき
  
    If CloseMode = 0 Then       ' Xボタン押下時は強制的に終了
        End
    End If

End Sub



' フォームをリサイズ可能にするための設定
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
        
    'ビューの先頭列の表示
    .ColumnHeaders.Add 1, "F", "File", iniWidth / 2
    .ColumnHeaders.Add 2, "D", "Appendix", iniWidth / 2
 End With

 ' ユーザーサイズ変更
 Me.Width = iniWidth
 Me.Height = iniHeight
 
 ' 保存フラグ変更
 Me.CheckBoxRecursive.Value = recursiveFlag
 
 ' パス履歴保存
 For i = 0 To pathDic.Count - 1
    Me.ComboBoxDirbox.AddItem (pathDic.Keys(i))
 Next i

 ' サイズ変更
 Call UserForm_Resize

End Sub


'
' ウィンドウサイズ整理
'
Private Sub UserForm_Resize()
    Dim mgn As Long
    Dim xx As Variant
    Dim yy As Variant
    
    xx = Array(8, 38, 16)
    yy = Array(12, 36, 60, 30, 12)
    
    'wdWidth = Me.Width
    'wdHeight = Me.Height
    
    ' x0 左合わせ
    Label2.Left = xx(0)
    Label4.Left = xx(0)
    Label6.Left = xx(0)
    
    ' x1 左合わせ
    ComboBoxDirbox.Left = xx(1)
    TextBox2.Left = xx(1)
    ListView1.Left = xx(1)
        
    OptionButton1.Left = xx(1)
    OptionButton2.Left = xx(1) + 100
    OptionButton3.Left = xx(1) + 200
    OptionButton4.Left = xx(1) + 300
        
    ' x2 右合わせ
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
        
    ' y0 上合わせ
    Label4.Top = yy(0)
    ComboBoxDirbox.Top = yy(0)
    CheckBoxRecursive.Top = yy(0)
    ' y1 上合わせ
    Label2.Top = yy(1)
    TextBox2.Top = yy(1)
    ' y2 上合わせ
    Label6.Top = yy(2)
    ListView1.Top = yy(2)
    
    ' y3 下合わせ
    If Me.InsideHeight > 200 Then
        ListView1.Height = Me.InsideHeight - yy(2) - yy(3)
    
        OptionButton1.Top = Me.InsideHeight - yy(3)
        OptionButton2.Top = Me.InsideHeight - yy(3)
        OptionButton3.Top = Me.InsideHeight - yy(3)
        OptionButton4.Top = Me.InsideHeight - yy(3)
        LabelCounter.Top = Me.InsideHeight - yy(3)
    End If


End Sub
