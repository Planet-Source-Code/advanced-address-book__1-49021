VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Advanced Address Book"
   ClientHeight    =   5175
   ClientLeft      =   915
   ClientTop       =   2565
   ClientWidth     =   8310
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3300
      Top             =   1680
   End
   Begin MSWinsockLib.Winsock sckServer 
      Left            =   4800
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList dul 
      Left            =   2940
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":222C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList dis 
      Left            =   1680
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":293E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList nrm 
      Left            =   1380
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   25
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   1217
      ButtonWidth     =   847
      ButtonHeight    =   1164
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "dul"
      DisabledImageList=   "dis"
      HotImageList    =   "nrm"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Object.ToolTipText     =   "New Contact"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Group"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Contact"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Object.ToolTipText     =   "Find Contacts"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Address Book"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find..."
      End
      Begin VB.Menu mnuMKAOL 
         Caption         =   "&Web Access"
         Begin VB.Menu mnuEnableWeb 
            Caption         =   "&Enable"
         End
         Begin VB.Menu mnuDisableWeb 
            Caption         =   "&Disable"
         End
         Begin VB.Menu mb78 
            Caption         =   "-"
         End
         Begin VB.Menu WebStatus 
            Caption         =   "Status"
         End
      End
      Begin VB.Menu mb67 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuGroups 
      Caption         =   "&Groups"
      Begin VB.Menu mnuAddGroup 
         Caption         =   "&Add..."
      End
      Begin VB.Menu mnuRemGroup 
         Caption         =   "&Remove..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuVCLG 
         Caption         =   "&Contact List Grid"
         Checked         =   -1  'True
      End
      Begin VB.Menu mb553 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCFG 
         Caption         =   "C&onfiguration..."
      End
   End
   Begin VB.Menu mnuSMgroups 
      Caption         =   "&Gp"
      Visible         =   0   'False
      Begin VB.Menu mnuGM 
         Caption         =   "Group Menu"
      End
      Begin VB.Menu mb1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAG 
         Caption         =   "&Add Group..."
      End
      Begin VB.Menu mnuRG 
         Caption         =   "&Remove Group"
      End
      Begin VB.Menu mb5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuACITG 
         Caption         =   "&Add contact in this group..."
      End
      Begin VB.Menu mb785 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindMn 
         Caption         =   "&Find in group..."
      End
   End
   Begin VB.Menu mnuSmContacts 
      Caption         =   "CT"
      Visible         =   0   'False
      Begin VB.Menu mnusmAC 
         Caption         =   "&Add Contact..."
      End
      Begin VB.Menu mnuETC 
         Caption         =   "&Edit this Contact..."
      End
      Begin VB.Menu mbub77 
         Caption         =   "-"
      End
      Begin VB.Menu mnusRC 
         Caption         =   "&Remove Contact"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
LoadDatabase
frmContacts.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
UnloadDatabase
End Sub

Private Sub mnuACITG_Click()
frmAddEditContact.EDT = False

frmAddEditContact.Show
frmAddEditContact.lstGroups.Text = frmContacts.SelNde.Text
End Sub

'Private Sub mnugroups_Click()
'frmMain.mnuGP.Visible = False
'frmMain.mnuB1.Visible = False
'End Sub

Private Sub mnuAddGroup_Click()
frmAddRemoveGroup.DEL = False
frmAddRemoveGroup.Show
End Sub

Private Sub mnuAG_Click()
frmAddRemoveGroup.DEL = False
            frmAddRemoveGroup.Show
End Sub

Private Sub mnuCFG_Click()
frmConfig.Show
End Sub

Private Sub mnuDisableWeb_Click()
StopWeb
WEbon = False
End Sub

Private Sub mnuEnableWeb_Click()
StartWeb
WEbon = True
End Sub

Private Sub mnuETC_Click()
Dim G() As String
G = Split(frmContacts.SelNde.Tag, "|")
frmAddEditContact.EDT = True
frmAddEditContact.EDTID = Trim(G(1))

frmAddEditContact.Show

End Sub

Private Sub mnuExit_Click()
UnloadDatabase
Unload Me
End
End Sub

Private Sub mnuFind_Click()
frmFind.Show
End Sub

Private Sub mnuFindMn_Click()
frmFind.Show
frmFind.lstGroups.Text = frmContacts.SelNde.Text
End Sub

Private Sub mnuRemGroup_Click()
frmAddRemoveGroup.DEL = True
frmAddRemoveGroup.Show
End Sub

Private Sub mnuRG_Click()
RemoveGroup frmContacts.SelNde.Text
End Sub

Private Sub mnusmAC_Click()
frmAddEditContact.EDT = False
frmAddEditContact.Show
End Sub

Private Sub mnusRC_Click()
Dim G() As String
G = Split(frmContacts.SelNde.Tag, "|")
RemoveContact G(1)
End Sub

Private Sub mnuVCLG_Click()
mnuVCLG.Checked = Not (mnuVCLG.Checked)
frmContacts.lstContacts.GridLines = mnuVCLG.Checked

End Sub

Private Sub sckServer_ConnectionRequest(ByVal requestID As Long)
If sckServer.State <> sckClosed Then sckServer.Close
sckServer.Accept requestID

End Sub

Private Sub sckServer_DataArrival(ByVal bytesTotal As Long)

On Error Resume Next
Dim DATA As String
Dim HDR As String
Dim BDY As String
Dim DSPL() As String
Dim sw As String

Dim N_F As String
Dim N_L As String
Dim N_M As String
Dim P_H As String
Dim P_W As String
Dim P_C As String
Dim P_P As String
Dim O_E As String
Dim O_A As String
Dim O_N As String
Dim O_G As String
Dim FORM_SEP() As String
Dim FORM_SEP2() As String

Dim FIELD() As String
Dim VALUE() As String
Dim od As String

Dim uid() As String

sckServer.GetData DATA

DSPL = Split(DATA, vbCrLf & vbCrLf)

If Client.sIp <> sckServer.RemoteHostIP Then
    Client.sIp = sckServer.RemoteHostIP
    Client.sAuthenticated = False
End If


If UBound(DSPL) <= 0 Then GoTo SNDDEF:
HDR = DSPL(0)
BDY = DSPL(1)

FORM_SEP = Split(BDY, "&")
If UBound(FORM_SEP) <= 0 Then GoTo SNDDEF:

ReDim FIELD(UBound(FORM_SEP))
ReDim VALUE(UBound(FORM_SEP))
'get parameters
For i = 0 To UBound(FORM_SEP)
    FORM_SEP2 = Split(FORM_SEP(i), "=")
    If UBound(FORM_SEP2) > 0 Then
        FIELD(i) = LCase(Trim(FORM_SEP2(0)))
        VALUE(i) = FORM_SEP2(1)
            
    End If
Next

'determine the current form
    Select Case FIELD(0)
        Case "authenticate"
            Select Case FIELD(1)
                Case "auth"
                    If CONFIG.Fields("WEB_PASS") = VALUE(1) Then
                        'password ok
                        Client.sAuthenticated = True
                        sw = "addressbook"
                        'Exit Sub
                    Else
                        'password error
                        sw = "invalidauth"
                        
                        'Exit Sub
                    End If
                Case Else
                    sw = "Invalid Parameter!"
            End Select
        Case "contactlist"
            If UBound(FIELD) >= 2 Then
            
            Select Case FIELD(2)
                Case "contact_view"
                    sw = "viewcontact"
                    od = VALUE(1)
                    'Exit Sub
                Case "contact_remove"
                    sw = "removecontact"
                    od = VALUE(1)
                Case "contact_add"
                    sw = "addcontact"
                Case "group_add"
                    sw = "addgroup"
                Case "group_remove"
                    sw = "removegroup"
                    od = VALUE(1)
                Case "view_clist"
                    sw = "listc"
            End Select
            Else
                Select Case FIELD(1)
                Case "contact_view"
                    sw = "viewcontact"
                    od = VALUE(1)
                    'Exit Sub
                Case "contact_remove"
                    sw = "removecontact"
                    od = VALUE(1)
                Case "contact_add"
                    sw = "addcontact"
                Case "group_add"
                    sw = "addgroup"
                Case "group_remove"
                    sw = "removegroup"
                    od = VALUE(1)
                Case "view_clist"
                    sw = "listc"
            End Select
            End If
        Case "viewcontact"
            Select Case FIELD(1)
                Case "back"
                    sw = "addressbook"
                Case "remove"
                    sw = "removecontact"
            End Select
        Case "removecontact"
            
            Select Case FIELD(1)
                Case "remove"
                    If Client.sAuthenticated = False Then GoTo LGI:
                    If MoveContacts(VALUE(0)) = 1 Then
                    With CONTACTS
                        .Delete
                    End With
                    sw = "addressbook"
                    Else
                    sw = "Error while removing contact!"
                    End If
                Case "cancel"
                    sw = "addressbook"
            End Select
        Case "addcontact"
            If Client.sAuthenticated = False Then GoTo LGI:
            For i = 1 To UBound(FIELD)
                Select Case FIELD(i)
                    Case "group"
                        O_G = VALUE(i)
                    Case "name_first"
                        N_F = VALUE(i)
                    Case "name_last"
                        N_L = VALUE(i)
                    Case "name_middle"
                        N_M = VALUE(i)
                    Case "phone_home"
                        P_H = VALUE(i)
                    Case "phone_work"
                        P_W = VALUE(i)
                    Case "phone_cell"
                        P_C = VALUE(i)
                    Case "phone_pager"
                        P_P = VALUE(i)
                    Case "email"
                        O_E = VALUE(i)
                    Case "address"
                        O_A = VALUE(i)
                        O_A = ReplaceString(O_A, "+", " ")
                        O_A = ReplaceString(O_A, "%0", vbCrLf)
                    Case "notes"
                        O_N = VALUE(i)
                        O_N = ReplaceString(O_N, "+", " ")
                        O_N = ReplaceString(O_N, "%0", vbCrLf)
                        
                    Case "add"
                        If LCase(Trim(O_G)) = "<none>" Then O_G = ""
                        AddContact N_F, N_M, N_L, P_H, P_W, P_C, P_P, O_E, O_A, O_N, O_G
                    Case "cancel"
                        sw = "addressbook"
                End Select
            Next
        Case "addgroup"
            Select Case FIELD(2)
                Case "add"
                    AddGroup VALUE(1)
                    sw = "addressbook"
                Case "cancel"
                    sw = "addressbook"
            End Select
        Case "removegroup"
            Select Case FIELD(1)
                Case "remove"
                    With GROUPS
                        RemoveGroup .Fields("GROUPNAME")
                    End With
                    sw = "addressbook"
                Case "cancel"
                    sw = "addressbook"
            End Select
        Case Else
            sw = "addressbook"
    End Select
            
SNDDEF:
LGI:
If Client.sAuthenticated = False Then
    If sw = "" Then
        Client.sIp = sckServer.RemoteHostIP
        SendWeb "authenticate"
    Else
        Client.sIp = sckServer.RemoteHostIP
        SendWeb "invalidauth"
    End If
Else
    If sw = "" Then sw = "addressbook"
    SendWeb sw, od
End If
End Sub

Private Sub Timer1_Timer()
Select Case sckServer.State
    Case sckConnected, sckConnecting, sckConnectionPending, sckResolvingHost, sckHostResolved
        If sckServer.State = sckConnected Then
            WebStatus.Caption = "Client Connected"
        Else
            WebStatus.Caption = "Client Connecting..."
        End If
        'mnuEnableWeb.Enabled = False
        mnuEnableWeb.Checked = True
        'mnuDisableWeb.Enabled = True
        mnuDisableWeb.Checked = False
    Case sckListening
        WebStatus.Caption = "Waiting for Client..."
        'mnuEnableWeb.Enabled = False
        mnuEnableWeb.Checked = True
        'mnuDisableWeb.Enabled = True
        mnuDisableWeb.Checked = False
        
    Case Else
        If WEbon Then
            StartWeb
            Exit Sub
        End If
        If sckServer.State <> sckClosed Then sckServer.Close
        WebStatus.Caption = "Idle"
        'mnuEnableWeb.Enabled = True
        mnuEnableWeb.Checked = False
        'mnuDisableWeb.Enabled = False
        mnuDisableWeb.Checked = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1  'add contact
    frmAddEditContact.EDT = False
    frmAddEditContact.Show
Case 3 'find
    frmFind.Show
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Parent.Index
Case 1  'add contact/group
    Select Case ButtonMenu.Index
        Case 1  'add group
            frmAddRemoveGroup.DEL = False
            frmAddRemoveGroup.Show
        Case 3  'add contact
            frmAddEditContact.EDT = False
            frmAddEditContact.Show
    End Select
End Select
End Sub

Private Sub StartWeb()
If sckServer.State <> sckClosed Then sckServer.Close
'With Client
'    .sAuthenticated = False
'    .sIp = ""
'End With
With CONFIG
sckServer.LocalPort = .Fields("WEB_PORT")
sckServer.Listen
End With

End Sub

Private Sub StopWeb()
If sckServer.State <> sckClosed Then sckServer.Close

End Sub
