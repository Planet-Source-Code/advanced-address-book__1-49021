VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddEditContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Contact"
   ClientHeight    =   6150
   ClientLeft      =   5280
   ClientTop       =   2550
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Contact Information"
      Height          =   4455
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   5655
      Begin VB.TextBox tFields 
         Height          =   765
         Index           =   9
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   3540
         Width           =   5355
      End
      Begin VB.TextBox tFields 
         Height          =   765
         Index           =   8
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2460
         Width           =   5355
      End
      Begin VB.TextBox tFields 
         Height          =   285
         Index           =   7
         Left            =   780
         TabIndex        =   22
         Top             =   1680
         Width           =   4755
      End
      Begin VB.TextBox tFields 
         Height          =   285
         Index           =   6
         Left            =   4380
         TabIndex        =   19
         Top             =   1200
         Width           =   1155
      End
      Begin VB.TextBox tFields 
         Height          =   285
         Index           =   5
         Left            =   3180
         TabIndex        =   18
         Top             =   1200
         Width           =   1155
      End
      Begin VB.TextBox tFields 
         Height          =   285
         Index           =   4
         Left            =   1980
         TabIndex        =   17
         Top             =   1200
         Width           =   1155
      End
      Begin VB.TextBox tFields 
         Height          =   285
         Index           =   3
         Left            =   780
         TabIndex        =   12
         Top             =   1200
         Width           =   1155
      End
      Begin VB.TextBox tFields 
         Height          =   285
         Index           =   2
         Left            =   3300
         TabIndex        =   10
         Top             =   540
         Width           =   2235
      End
      Begin VB.TextBox tFields 
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   9
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox tFields 
         Height          =   285
         Index           =   0
         Left            =   780
         TabIndex        =   6
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Other/Notes:"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   3300
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Address:"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   2220
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Email:"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1740
         Width           =   3315
      End
      Begin VB.Label Label9 
         Caption         =   "Pager:"
         Height          =   255
         Left            =   4380
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Cell:"
         Height          =   255
         Left            =   3180
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Work:"
         Height          =   255
         Left            =   1980
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Phone:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1260
         Width           =   3315
      End
      Begin VB.Label Label5 
         Caption         =   "Home:"
         Height          =   255
         Left            =   780
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Last:"
         Height          =   255
         Left            =   3300
         TabIndex        =   11
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Middle:"
         Height          =   195
         Left            =   2400
         TabIndex        =   8
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "First:"
         Height          =   255
         Left            =   780
         TabIndex        =   7
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   600
         Width           =   3315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group"
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5655
      Begin VB.ComboBox lstGroups 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   5355
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Width           =   5805
      _ExtentX        =   10239
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
      TabIndex        =   3
      Top             =   5610
      Width           =   5805
      _ExtentX        =   10239
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
            Caption         =   "Add"
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
End
Attribute VB_Name = "frmAddEditContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EDT As Boolean
Public EDTID As String

Private Sub Form_Load()
ListGroups
If EDT Then
    
    MoveContacts EDTID
    GCTIFO
    
    Toolbar1.Buttons(3).Caption = "Apply"
    Toolbar1.Buttons(1).Caption = "Exit"
    
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Me.Caption = "Edit contact"
Else

    Toolbar1.Buttons(3).Caption = "Add"
    Toolbar1.Buttons(1).Caption = "Cancel"
    
    Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
    
    Me.Caption = "Add contact"
End If

frmMain.Enabled = False


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMain.Enabled = True
End Sub

Private Sub tFields_LostFocus(Index As Integer)
Dim TMPSTR As String
Dim NF() As String
TMPSTR = tFields(Index).Text
Select Case Index
    Case 3, 4, 5, 6
        NF = Split(TMPSTR, "-")
        If UBound(NF) <= 0 Then
        Select Case Len(tFields(Index).Text)
            Case 11 'CountryCode + areacode + number
                tFields(Index).Text = Mid(TMPSTR, 1, 1) & "-" & "(" & Mid(TMPSTR, 2, 3) & ")" & "-" & Mid(TMPSTR, 5, 3) & "-" & Mid(TMPSTR, 8, 4)
            Case 10 'area code + number
                                      'area code                'first three              'last 4
                tFields(Index).Text = "(" & Mid(TMPSTR, 1, 3) & ")" & "-" & Mid(TMPSTR, 4, 3) & "-" & Mid(TMPSTR, 7, 4)
            Case 7  'number
                tFields(Index).Text = Mid(TMPSTR, 1, 3) & "-" & Mid(TMPSTR, 4, 4)
            Case Else 'leave as is
        End Select
        End If
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Dim GP As String
Select Case Button.Index
    Case 1 'cancel
        Unload Me
    Case 3 'add/edit
        Select Case Trim(LCase(Button.Caption))
            Case "add"
                If LCase(Trim(lstGroups.Text)) = "<none>" Then
                    GP = ""
                Else
                    GP = lstGroups.Text
                End If
                AddContact tFields(0).Text, tFields(1).Text, tFields(2).Text, tFields(3).Text, tFields(4).Text, tFields(5).Text, tFields(6).Text, tFields(7).Text, tFields(8).Text, tFields(9).Text, GP
                Unload Me
            Case "apply"
                If LCase(Trim(lstGroups.Text)) = "<none>" Then
                    GP = ""
                    GoTo ADDCT:
                Else
                    GP = lstGroups.Text
                End If
                
                If Not (GroupExists(GP)) Then
                    If MsgBox("The group '" & GP & "' does not exist, Would you like to create it now?", vbYesNo + vbExclamation, "Could not find group") = vbYes Then
                        AddGroup GP
                        DoEvents
                        GoTo ADDCT:
                    Else
                        GP = ""
                        GoTo ADDCT:
                    End If
                End If
ADDCT:
                With CONTACTS
                    .Edit
                        !NAME_FIRST = tFields(0).Text
                        !NAME_MIDDLE = tFields(1).Text
                        !NAME_LAST = tFields(2).Text
                        
                        !PHONE_HOME = tFields(3).Text
                        !PHONE_WORK = tFields(4).Text
                        !PHONE_CELL = tFields(5).Text
                        !PHONE_PAGER = tFields(6).Text
                        
                        !EMAIL = tFields(7).Text
                        
                        !ADDRESS = tFields(8).Text
                        !NOTES = tFields(9).Text
                        
                        !groupname = GP
                    .Update
                End With
                frmContacts.RefContactList
        End Select
    Case 5 'prev contact
        CONTACTS.MovePrevious
        GCTIFO
    Case 6 'next contact
        CONTACTS.MoveNext
        GCTIFO
End Select
End Sub

Private Sub ListGroups()
On Error GoTo ERR1:
lstGroups.Clear
With GROUPS
    .MoveFirst
    Do While Not .EOF
        lstGroups.AddItem .Fields("GROUPNAME")
        .MoveNext
    DoEvents
    Loop
End With
lstGroups.AddItem "<None>"
lstGroups.Text = lstGroups.List(0)
Exit Sub
ERR1:
Select Case Err.Number
    Case 3021 'ignore, no entries in list
        lstGroups.Text = "<None>"
        Resume Next
    Case Else
        MsgBox "Unknown database error: " & vbCrLf & vbCrLf & Err.Description, vbCritical, "Database Error!"
        
End Select
End Sub

Public Sub GCTIFO()
On Error Resume Next
    
    With CONTACTS
        lstGroups.Text = .Fields("GROUPNAME")
        
        tFields(0).Text = .Fields("NAME_FIRST")
        tFields(1).Text = .Fields("NAME_MIDDLE")
        tFields(2).Text = .Fields("NAME_LAST")
        
        tFields(3).Text = .Fields("PHONE_HOME")
        tFields(4).Text = .Fields("PHONE_WORK")
        tFields(5).Text = .Fields("PHONE_CELL")
        tFields(6).Text = .Fields("PHONE_PAGER")
        
        tFields(7).Text = .Fields("EMAIL")
        
        tFields(8).Text = .Fields("ADDRESS")
        tFields(9).Text = .Fields("NOTES")
        
        EDTID = .Fields("CONTACTID")
    End With
End Sub
