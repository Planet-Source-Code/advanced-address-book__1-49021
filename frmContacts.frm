VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContacts 
   Caption         =   "Contact Viewer"
   ClientHeight    =   4290
   ClientLeft      =   3165
   ClientTop       =   2820
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContacts.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5775
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":1ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":1E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":24B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":2802
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":2D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":32A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstContacts 
      Height          =   3675
      Left            =   2280
      TabIndex        =   1
      Top             =   60
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   6482
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Home Phone"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Buisness Phone"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView lstGroups 
      Height          =   3675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   6482
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   139
      LabelEdit       =   1
      Style           =   3
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private BTN As Integer
Public SelNde As Node


Private Sub Form_Load()
RefContactList
End Sub

Private Sub Form_Resize()
lstGroups.Move 50, 50, lstGroups.Width, Me.ScaleHeight - 100
lstContacts.Move lstGroups.Left + lstGroups.Width + 50, 50, Me.ScaleWidth - (150 + lstGroups.Width), Me.ScaleHeight - 100

End Sub

Public Sub RefContactList()

On Error GoTo ERR1:
Dim SC As Node
Dim RT As Node
Dim GP As Node
Dim CT As Node
Dim GID As Integer

lstGroups.Nodes.Clear

Set RT = lstGroups.Nodes.Add(, , "ROOT" & GENID, "Contact Groups", 7, 7)
'RT.Expanded = True
lstGroups.Indentation = 79
Set SC = lstGroups.Nodes.Add(, , "STRY" & GENID, "Stray Contacts", 7, 7)
SC.Bold = True
RT.Bold = True

RT.Expanded = True
SC.Expanded = True

With GROUPS
    .MoveFirst
    Do While Not .EOF
        If Trim(.Fields("GROUPNAME")) <> "" Then
            Set GP = lstGroups.Nodes.Add(RT.Key, tvwChild, "G" & .Fields("GROUPID"), .Fields("GROUPNAME"), 2, 1)
                GP.Tag = "G|" & .Fields("GROUPID")
        End If
    .MoveNext
    DoEvents
    Loop
End With

'On Error Resume Next
With CONTACTS
    .MoveFirst
    Do While Not .EOF
        If Trim(.Fields("GROUPNAME")) <> "" Then
            On Error GoTo addrt:
            GID = FindGroupID(.Fields("GROUPNAME"))
            If GID <> -1 Then
                Set CT = lstGroups.Nodes.Add("G" & GID, tvwChild, "C" & .Fields("CONTACTID"), .Fields("NAME_LAST") & ", " & .Fields("NAME_FIRST") & " " & .Fields("NAME_MIDDLE"), 4, 3)
                CT.Tag = "C|" & .Fields("CONTACTID")
                CT.ForeColor = RGB(100, 100, 130)
            End If
        Else
addrt:
            Set CT = lstGroups.Nodes.Add(SC.Key, tvwChild, "C" & .Fields("CONTACTID"), .Fields("NAME_LAST") & ", " & .Fields("NAME_FIRST") & " " & .Fields("NAME_MIDDLE"), 6, 5)
                CT.Tag = "C|" & .Fields("CONTACTID")
                CT.ForeColor = RGB(100, 100, 130)
        End If
    .MoveNext
    DoEvents
    Loop
End With

Exit Sub
ERR1:
Select Case Err.Number
    Case 3021 'ignore, no entries in list
        Resume Next
    Case Else
        MsgBox "Unknown database error: " & vbCrLf & vbCrLf & Err.Description, vbCritical, "Database Error!"
        
End Select
End Sub

Private Sub lstContacts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
MsgBox ColumnHeader.Index

End Sub

Private Sub lstGroups_DblClick()
On Error GoTo nd:
Dim G() As String
G = Split(frmContacts.SelNde.Tag, "|")
If LCase(G(0)) = "c" Then 'only contacts
    frmAddEditContact.EDT = True
    frmAddEditContact.EDTID = Trim(G(1))
    frmAddEditContact.Show
End If
nd:
End Sub

Private Sub lstGroups_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
BTN = Button
End Sub

Private Sub lstGroups_NodeClick(ByVal Node As MSComctlLib.Node)
'On Error GoTo ERR1:

Dim D() As String
Dim A As ListItem

D = Split(Node.Tag, "|")


If UBound(D) <= 0 Then Exit Sub
Set SelNde = Node

If BTN = 2 Then

Select Case LCase(D(0))
    Case "c" 'contact
        PopupMenu frmMain.mnuSmContacts
    Case "g" 'group
        
        PopupMenu frmMain.mnuSMgroups
End Select

    Exit Sub
End If

lstContacts.ListItems.Clear



With CONTACTS


Select Case LCase(D(0))
    Case "c" 'contact
        
            .MoveFirst
            Do While Not .EOF
                If Val(.Fields("CONTACTID")) = Val(D(1)) Then
                    GoTo Foundit:
                End If
                .MoveNext
            Loop
            'no record
            Exit Sub
Foundit:
        lstContacts.ColumnHeaders.Clear
            lstContacts.ColumnHeaders.Add , , "Field"
            lstContacts.ColumnHeaders.Add , , "Value"
            Set A = lstContacts.ListItems.Add(, , "Name:")
            A.SubItems(1) = .Fields("NAME_FIRST") & " " & .Fields("NAME_MIDDLE") & " " & .Fields("NAME_LAST")
            Set A = lstContacts.ListItems.Add(, , "Telephone:")
            A.SubItems(1) = .Fields("PHONE_HOME") & " [Home]"
        
            If .Fields("PHONE_WORK") <> "" Then
                Set A = lstContacts.ListItems.Add(, , " ")
                A.SubItems(1) = .Fields("PHONE_WORK") & " [Work]"
            End If
            
            If .Fields("PHONE_CELL") <> "" Then
                Set A = lstContacts.ListItems.Add(, , " ")
                A.SubItems(1) = .Fields("PHONE_CELL") & " [Cell]"
            End If
            
            If .Fields("PHONE_PAGER") <> "" Then
                Set A = lstContacts.ListItems.Add(, , " ")
                A.SubItems(1) = .Fields("PHONE_PAGER") & " [Pgr.]"
            End If
            
            Set A = lstContacts.ListItems.Add(, , "Email Address:")
                A.SubItems(1) = .Fields("EMAIL")
            
            Dim tt() As String
            tt = Split(.Fields("ADDRESS"), vbCrLf)
            For i = 0 To UBound(tt)
                If i = 0 Then
                    Set A = lstContacts.ListItems.Add(, , "Street Address:")
                    A.SubItems(1) = tt(i)
                Else
                    Set A = lstContacts.ListItems.Add(, , " ")
                    A.SubItems(1) = tt(i)
                End If
            Next
            
            
            tt = Split(.Fields("NOTES"), vbCrLf)
            For i = 0 To UBound(tt)
                If i = 0 Then
                    Set A = lstContacts.ListItems.Add(, , "Notes:")
                    A.SubItems(1) = tt(i)
                Else
                    Set A = lstContacts.ListItems.Add(, , " ")
                    A.SubItems(1) = tt(i)
                End If
            Next
            
            Set A = lstContacts.ListItems.Add(, , "Group:")
                A.SubItems(1) = .Fields("GROUPNAME")
    Case "g" 'group
    lstContacts.ColumnHeaders.Clear
    lstContacts.ColumnHeaders.Add , , "Name"
    lstContacts.ColumnHeaders.Add , , "Home Phone"
    lstContacts.ColumnHeaders.Add , , "Work Phone"
    lstContacts.ColumnHeaders.Add , , "Email Address"
     On Error Resume Next
    .MoveFirst
    Do While Not .EOF
        If .Fields("GROUPNAME") = Node.Text Then
            Set A = lstContacts.ListItems.Add(, , .Fields("NAME_FIRST") & " " & .Fields("NAME_MIDDLE") & " " & .Fields("NAME_LAST"))
                A.SubItems(1) = .Fields("PHONE_HOME")
                A.SubItems(2) = .Fields("PHONE_WORK")
                A.SubItems(3) = .Fields("EMAIL")
        End If
    .MoveNext
    DoEvents
    Loop
    
End Select

End With

'On Error Resume Next
Dim mx As String
For i = 1 To lstContacts.ColumnHeaders.Count - 1
    'If i <> 1 Then
    mx = ""
    For ii = 1 To lstContacts.ListItems.Count
        If Len(lstContacts.ListItems(ii).ListSubItems(i).Text) > Len(mx) Then
            mx = lstContacts.ListItems(ii).ListSubItems(i).Text
        End If
    Next
    lstContacts.ColumnHeaders(i + 1).Width = frmContacts.TextWidth(mx) + 300
    'End If
Next

Exit Sub
ERR1:
Select Case Err.Number
    Case 3021 'ignore, no entries in list
        Resume Next
    Case 94 'ignore, null entry string
        Resume Next
    Case Else
        MsgBox "Unknown database error: " & vbCrLf & vbCrLf & Err.Description, vbCritical, "Database Error!"
        
        
End Select


End Sub

