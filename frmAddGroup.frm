VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddRemoveGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Group"
   ClientHeight    =   1830
   ClientLeft      =   3870
   ClientTop       =   3855
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Group"
      Height          =   735
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   4395
      Begin VB.ComboBox lstGroups 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   4155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group name"
      Height          =   735
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4395
      Begin VB.TextBox txtGroup 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   4155
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   1290
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   953
      ButtonWidth     =   1032
      ButtonHeight    =   953
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAddRemoveGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DEL As Boolean

Private Sub Form_Load()
frmMain.Enabled = False

If DEL = True Then
    'deleteing group
    Me.Caption = "Remove Group"
    Toolbar1.Buttons(3).Caption = "Remove"
    Toolbar1.Buttons(5).Visible = True
    Frame1.Visible = False
    Frame2.Visible = True
    ListGroups
Else
    Frame1.Visible = True
    Frame2.Visible = False
    Me.Caption = "Add Group"
    Toolbar1.Buttons(3).Caption = "Add"
    Toolbar1.Buttons(5).Visible = False
    
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMain.Enabled = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1  'cancel
    Unload Me
Case 3  'add/remove
    Select Case LCase(Trim(Button.Caption))
        Case "add"
            If AddGroup(txtGroup.Text) = 1 Then
                Unload Me
            Else
                txtGroup.SetFocus
                txtGroup.SelStart = 0
                txtGroup.SelLength = Len(txtGroup.Text)
            End If
        Case "remove"
            If RemoveGroup(lstGroups.Text) = 1 Then
                ListGroups
            End If
    End Select
Case 5 'exit (remove only)
    Unload Me
End Select
End Sub

Private Sub ListGroups()
lstGroups.Clear
With GROUPS
    .MoveFirst
    Do While Not .EOF
        lstGroups.AddItem .Fields("GROUPNAME")
        .MoveNext
    DoEvents
    Loop
End With
lstGroups.Text = lstGroups.List(0)
End Sub
