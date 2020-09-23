VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBut 
      Caption         =   "C&ancel"
      Height          =   375
      Index           =   1
      Left            =   4830
      TabIndex        =   5
      Top             =   840
      Width           =   1055
   End
   Begin VB.CommandButton cmdBut 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4830
      TabIndex        =   4
      Top             =   405
      Width           =   1055
   End
   Begin VB.TextBox txtData 
      Height          =   350
      Left            =   225
      TabIndex        =   3
      Top             =   1155
      Width           =   4470
   End
   Begin VB.TextBox txtValue 
      Height          =   350
      Left            =   225
      TabIndex        =   1
      Top             =   435
      Width           =   4470
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Data:"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   900
      Width           =   735
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   90
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBut_Click(Index As Integer)
    Select Case Index
        Case 0
            'Set variable data
            mValName = txtValue.Text
            mValData = txtData.Text
            
            ButtonPress = vbOK
            Unload frmAdd
        Case 1
            ButtonPress = vbCancel
            Unload frmAdd
        End Select
End Sub

Private Sub Form_Load()
    Set frmAdd.Icon = Nothing
    
    lblName.Caption = "Item Name:"
    
    'Add new item
    If (mEditMode = INI_NEW_VALUE) Then
        frmAdd.Caption = "New Item"
    End If
    
    'Edit Item
    If (mEditMode = INI_EDIT_VALUE) Then
        frmAdd.Caption = "Edit Item"
        txtValue.Text = mValName
        txtData.Text = mValData
        txtValue.Enabled = False
    End If
    
    'Add New Selection
    If (mEditMode = INI_NEW_SELECTION) Then
        frmAdd.Caption = "Add Selection"
        lblName.Caption = "Selection Name:"
        txtData.Visible = False
        lblData.Visible = False
    End If
    
    'Rename Selection
    If (mEditMode = INI_RENAME_SELECTION) Then
        frmAdd.Caption = "Rename Selection"
        lblName.Caption = "Selection Name:"
        'Update textbox with selection name.
        txtValue.Text = mCurSelection
        txtData.Visible = False
        lblData.Visible = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAdd = Nothing
End Sub

Private Sub txtData_Change()
    Call txtValue_Change
End Sub

Private Sub txtValue_Change()
    cmdBut(0).Enabled = Len(Trim$(txtValue.Text))
End Sub
