VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "DM INI Editor"
   ClientHeight    =   4905
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView Lst1 
      Height          =   1335
      Left            =   2415
      TabIndex        =   3
      Top             =   420
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4410
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":09F6
            Key             =   "Alpha"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0B08
            Key             =   "Digit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0C1A
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0F6C
            Key             =   "Top"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":12BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tv1 
      Height          =   1305
      Left            =   0
      TabIndex        =   2
      Top             =   435
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   2302
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4530
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14340
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAVE"
            Object.ToolTipText     =   "ExportXML"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SELECTION"
            Object.ToolTipText     =   "Selection"
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "SelAdd"
                  Text            =   "Add"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "SelDel"
                  Text            =   "Delete"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "SelRename"
                  Text            =   "Rename"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "SelExport"
                  Text            =   "Export"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ITEM"
            Object.ToolTipText     =   "Item"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "iNew"
                  Text            =   "New"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "iEdit"
                  Text            =   "Edit"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "iDel"
                  Text            =   "Delete"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3855
      Top             =   465
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   585
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   585
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Export XML"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuItem 
         Caption         =   "Item"
         Begin VB.Menu mnuNewi 
            Caption         =   "&New"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuEdit1 
            Caption         =   "Edit"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDel1 
            Caption         =   "Delete"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuSel 
         Caption         =   "Selection"
         Begin VB.Menu mnuAdd1 
            Caption         =   "Add"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDel2 
            Caption         =   "Delete"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRename 
            Caption         =   "Rename"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuExport 
            Caption         =   "Export"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIni As dINIFile
Private Const VBQuote As String = """"
Private Const Filter1 As String = "INI Files(*.ini)|*.ini|"
Private Const Filter2 As String = "XML Files(*.xml)|*.xml|"

Private Sub ExportXML(ByVal OutName As String)
Dim oCol As Collection
Dim iCol As Collection
Dim sLine As String
Dim sItem
Dim iItem
Dim fp As Long

    'Converts an INI File to XML
    Set oCol = mIni.GetSelections()
    fp = FreeFile
    
    Open OutName For Append As #fp
        
        Print #fp, "<?xml version=" & VBQuote & "1.0" & VBQuote & " encoding=" & VBQuote & "UTF-8" & VBQuote & "?>"
        Print #fp, "<sections>"
        
        For Each sItem In oCol
            Set iCol = mIni.GetValues(sItem)
            '
            Print #fp, Space(4) & "<section name=" & VBQuote & sItem & VBQuote & ">"
            
            For Each iItem In iCol
                sLine = mIni.ReadValue(sItem, iItem)
                Print #fp, Space(6) & "<item key=" & VBQuote & iItem & VBQuote & " value=" & VBQuote & sLine & VBQuote & " />    "
            Next iItem
            
            Print #fp, Space(4) & "</section>"
            
        Next sItem
    
        Print #fp, "</sections>"
    Close #fp
    
    Set oCol = Nothing
    Set iCol = Nothing
    sLine = vbNullString
    
End Sub

Private Sub ExportSelection(ByVal OutName As String, Selection As String)
Dim oCol As Collection
Dim tmp() As String
Dim Cnt As Integer
Dim fp As Long
    
    'Writes a selected selection to a file.
    If mIni.SelectionExists(Selection) Then
        Set oCol = mIni.GetValues(Selection)
        
        ReDim Preserve tmp(0 To oCol.Count - 1)
        
        For Cnt = 1 To oCol.Count
            tmp(Cnt - 1) = oCol(Cnt) & "=" & mIni.ReadValue(Selection, oCol(Cnt))
        Next Cnt
        
        fp = FreeFile
        
        Open OutName For Output As #fp
            Print #fp, "[" & Selection & "]"
            Print #fp, Join(tmp, vbCrLf)
        Close #fp
        
    End If
    
    Erase tmp
    Set oCol = Nothing
    
    
End Sub

Private Sub LoadSelections()
Dim oCol As Collection
Dim sSel
    
    'Load INI selections into the Treeview control
    With Tv1
        Set oCol = mIni.GetSelections
        .Nodes.Clear
        If (oCol.Count > 0) Then
            'Add top node
            .Nodes.Add , tvwFirst, "TOP", GetFilename(mIni.Filename), "Top"
            'Add child nodes INI Selection names.
            For Each sSel In oCol
                .Nodes.Add 1, tvwChild, "c" & .Nodes.Count, sSel, "Folder"
            Next sSel
            'Select the first node
            .Nodes(2).Selected = True
            .Refresh
        End If
        
        Call Tv1_Click
    End With
    
    Set oCol = Nothing
    Toolbar1.Buttons(5).ButtonMenus(1).Enabled = True
    mnuAdd1.Enabled = True
End Sub

Private Sub LoadValues(Selection As String, Optional SelectIdx As Integer = 1)
On Error Resume Next

Dim sSel As String
Dim oItems As Collection
Dim sItem
Dim SelName As String
Dim sIcon As String
Dim sVal As Variant
Dim lItem As ListItem

    'LOads INI Value names and data into Listview control.
    Set oItems = mIni.GetValues(Selection)
    
    With Lst1
        'Clear control.
        .ListItems.Clear
        'Add each item to the listview.
        For Each sItem In oItems
            'Add vlaue name
            sVal = mIni.ReadValue(Selection, sItem)
            
            If IsNumeric(sVal) Then
                'Add String bitmap
                sIcon = "Digit"
            Else
                'Add Alpha bitmap
                sIcon = "Alpha"
            End If
            
            .ListItems.Add , "c," & Selection & "," & .ListItems.Count, sItem, , sIcon
            .ListItems(.ListItems.Count).SubItems(1) = sVal
        Next sItem
        
        'Get list item
        Set lItem = .ListItems(SelectIdx)
        lItem.Selected = True
        Call Lst1_ItemClick(lItem)
        
        Set lItem = Nothing
        Set oItems = Nothing
    End With
    
    
End Sub

Private Sub Command1_Click()
ExportXML "C:\work\ben.txt"

End Sub

Private Sub Form_Load()
    LastDirLoc = FixPath(App.Path)
    Set mIni = New dINIFile
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Line1(0).X2 = frmmain.ScaleWidth
    Line1(1).X2 = Line1(0).X2
    
    Tv1.Height = (frmmain.ScaleHeight - sBar1.Height) - Tv1.Top
    Lst1.Height = Tv1.Height
    Lst1.Width = (frmmain.ScaleWidth - Lst1.Left)
    
End Sub

Private Function GetNameFromDLG(ShowOpen As Boolean, dFilter As String, Optional Title As String = "Open") As String
On Error GoTo OpenErr:
    'Returns a filename from the Dialog control
    With CD1
        .CancelError = True
        .DialogTitle = Title
        .Filter = dFilter
        .InitDir = LastDirLoc
        .Filename = vbNullString
        'What dialog to show, Open or save
        If (ShowOpen) Then
            .ShowOpen
        Else
            .ShowSave
        End If
        'Return Filename.
        GetNameFromDLG = .Filename
        'Preserve last known path
        LastDirLoc = GetAbsPath(.Filename)
    End With
    
    Exit Function
    'Cancel Error flag
OpenErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mIni = Nothing
    Set frmAdd = Nothing
    Set frmmain = Nothing
End Sub

Private Sub Lst1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mValName = Lst1.SelectedItem.Text 'Item Name
    mValData = Lst1.SelectedItem.SubItems(1) 'Item Data
    
    Toolbar1.Buttons(6).ButtonMenus(2).Enabled = True
    Toolbar1.Buttons(6).ButtonMenus(3).Enabled = True
    '
    mnuEdit1.Enabled = True
    mnuDel1.Enabled = True
End Sub

Private Sub mnuAbout_Click()
    MsgBox frmmain.Caption & "Ver 1.1" & vbCrLf & vbTab & "By Ben Jones" _
    & vbCrLf & vbTab & vbTab & "Please vote if you like this code.", vbInformation, "About"
End Sub

Private Sub mnuAdd1_Click()
    'Add new Selection
    mEditMode = INI_NEW_SELECTION
    frmAdd.Show vbModal, frmmain
    'Check if OK button was pressed
    If (ButtonPress = vbOK) Then
        Call mIni.AddSelection(mValName)
    End If
    'Update Treeview with new changes
    Call LoadSelections
End Sub

Private Sub mnuDel1_Click()
    'Delete Item
    If MsgBox("Are you sure you want to delete this item?", vbYesNo Or vbQuestion, "Delete Item") = vbYes Then
        Call mIni.DeleteValue(mCurSelection, mValName)
        Call LoadValues(mCurSelection)
        Lst1.SetFocus
        
        Toolbar1.Buttons(6).ButtonMenus(2).Enabled = Lst1.ListItems.Count
        Toolbar1.Buttons(6).ButtonMenus(3).Enabled = Lst1.ListItems.Count
    End If
End Sub

Private Sub mnuDel2_Click()
    'Delete a selection and all it's keys
    If MsgBox("Are you sure you want to delete this selection?", vbYesNo Or vbQuestion, "Delete Selection") = vbYes Then
        Call mIni.DeleteSelection(mCurSelection)
    End If
    
    Lst1.ListItems.Clear
    'Update Treeview with new changes
    Call LoadSelections
End Sub

Private Sub mnuEdit1_Click()
    'Edit Item
    mEditMode = INI_EDIT_VALUE
    'Show add form
    frmAdd.Show vbModal, frmmain
    'Check if OK button was pressed
    If (ButtonPress = vbOK) Then
        mIni.SetValue mCurSelection, mValName, mValData
        Call LoadValues(mCurSelection, Lst1.SelectedItem.Index)
        Lst1.SetFocus
    End If
End Sub

Private Sub mnuexit_Click()
    Unload frmmain
End Sub

Private Sub mnuExport_Click()
Dim TmpFile As String

    'Export Selection
    TmpFile = GetNameFromDLG(False, Filter2, "Export")
    If Len(TmpFile) > 0 Then
        Call ExportSelection(TmpFile, mCurSelection)
    End If
End Sub

Private Sub mnuNew_Click()
    With mIni
        .Filename = GetNameFromDLG(False, Filter1, "New")
    
        If Len(mIni.Filename) > 0 Then
            Call .CreateIni
            'Add default selection
            Call .SetValue("Selection", "Default", "Test_Value")
            'Open the ini file.
            Call LoadSelections
        End If
    End With
    
End Sub

Private Sub mnuNewi_Click()
    mEditMode = INI_NEW_VALUE
    frmAdd.Show vbModal, frmmain
    'Check if OK button was pressed
    If (ButtonPress = vbOK) Then
        mIni.SetValue mCurSelection, mValName, mValData
        Call LoadValues(mCurSelection)
        Lst1.SetFocus
    End If
End Sub

Private Sub mnuOpen_Click()
    mIni.Filename = GetNameFromDLG(True, Filter1)
    If Len(mIni.Filename) > 0 Then Call LoadSelections
End Sub

Private Sub mnuRename_Click()
    'Rename Selection
    mEditMode = INI_RENAME_SELECTION
    frmAdd.Show vbModal, frmmain
    'Check if OK button was pressed
    If (ButtonPress = vbOK) Then
        Call mIni.RenameSelection(mCurSelection, mValName)
    End If
    'Update Treeview with new changes
    Call LoadSelections
End Sub

Private Sub mnuSave_Click()
Dim TmpFile As String
    TmpFile = GetNameFromDLG(False, Filter2, "ExportXML")
    
    If Len(TmpFile) > 0 Then
        Call ExportXML(TmpFile)
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "OPEN"
            Call mnuOpen_Click
        Case "SAVE"
            Call mnuSave_Click
        Case "NEW"
            Call mnuNew_Click
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "SelAdd"
            Call mnuAdd1_Click
        Case "SelDel"
            Call mnuDel2_Click
        Case "SelRename"
            Call mnuRename_Click
        Case "SelExport"
            Call mnuExport_Click
        Case "iNew"
            Call mnuNewi_Click
        Case "iEdit"
            Call mnuEdit1_Click
        Case "iDel"
            Call mnuDel1_Click
    End Select
    
    'Update Button status
    ButtonPress = vbCancel
End Sub

Private Sub Tv1_Click()
On Error Resume Next
Dim mKey As String
    
    If (Tv1.Nodes.Count = 0) Then
        Exit Sub
    End If
    
    'Treeview Node key
    mKey = Tv1.SelectedItem.Key
    'Treeview selected Node Text
    mCurSelection = Tv1.SelectedItem.Text
    
    Toolbar1.Buttons(5).ButtonMenus(2).Enabled = (mKey <> "TOP")
    Toolbar1.Buttons(5).ButtonMenus(3).Enabled = (mKey <> "TOP")
    Toolbar1.Buttons(5).ButtonMenus(4).Enabled = (mKey <> "TOP")
    Toolbar1.Buttons(6).ButtonMenus(1).Enabled = (mKey <> "TOP")
    mnuNewi.Enabled = (mKey <> "TOP")
    mnuDel2.Enabled = mnuNewi.Enabled
    mnuRename.Enabled = mnuNewi.Enabled
    mnuExport.Enabled = mnuNewi.Enabled
    
    Toolbar1.Buttons(6).ButtonMenus(2).Enabled = False
    Toolbar1.Buttons(6).ButtonMenus(3).Enabled = False
    mnuEdit1.Enabled = False
    mnuDel1.Enabled = False
    
    If (Tv1.SelectedItem.Key = "TOP") Then
        Exit Sub
    Else
        Call LoadValues(Tv1.SelectedItem.Text, 0)
    End If
    
End Sub

