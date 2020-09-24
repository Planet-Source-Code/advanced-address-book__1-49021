VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   5460
   ClientLeft      =   3285
   ClientTop       =   1620
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.Frame f_general 
      Height          =   1995
      Left            =   180
      TabIndex        =   3
      Top             =   480
      Width           =   5655
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Text            =   "data.mbd"
         Top             =   180
         Width           =   3675
      End
      Begin VB.Label Label4 
         Caption         =   "Note: Database file MUST be in the application directory! Database file must be a Valid database formatted for this application!"
         ForeColor       =   &H80000011&
         Height          =   915
         Left            =   120
         TabIndex        =   11
         Top             =   660
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Database File:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1635
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   953
      ButtonWidth     =   1032
      ButtonHeight    =   953
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Apply"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Previous"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Next"
         EndProperty
      EndProperty
   End
   Begin VB.Frame f_web 
      Height          =   1995
      Left            =   180
      TabIndex        =   4
      Top             =   480
      Width           =   5655
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1500
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1500
         TabIndex        =   6
         Text            =   "81"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Web Server Port:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1755
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Web Access"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   4350
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Text3.Text = GetSetting(App.EXEName, "Set", "Path", "data.mdb")
With CONFIG
    .MoveFirst
    
    Text1.Text = .Fields("WEB_PORT")
    Text2.Text = .Fields("WEB_PASS")
End With
End Sub

Private Sub TabStrip1_Click()
f_general.Visible = False
f_web.Visible = False
Select Case TabStrip1.SelectedItem.Index
    Case 1  'general
        f_general.Visible = True
        
    Case 2  'web
        f_web.Visible = True
        
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'cancel
        Unload Me
    Case 3 'apply
        SaveSetting App.EXEName, "Set", "Path", Text3.Text
        
        With CONFIG
            .MoveFirst
            .Edit
                !WEB_PORT = Val(Text1.Text)
                !WEB_PASS = Text2.Text
                
            .Update
        End With
        Unload Me
End Select
End Sub
