VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Contacts"
   ClientHeight    =   3930
   ClientLeft      =   2100
   ClientTop       =   3270
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList dul 
      Left            =   1560
      Top             =   960
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
            Picture         =   "frmFind.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":077E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList dis 
      Left            =   300
      Top             =   1260
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
            Picture         =   "frmFind.frx":0E90
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":1602
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList nrm 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmFind.frx":1D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":2486
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmFind.frx":2B98
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":2EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":323C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":358E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":38E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":3E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":4384
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame f_search 
      BorderStyle     =   0  'None
      Caption         =   "f_Search"
      Height          =   3255
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   5475
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   660
         Left            =   4740
         TabIndex        =   11
         Top             =   2580
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1164
         ButtonWidth     =   1058
         ButtonHeight    =   1164
         Style           =   1
         ImageList       =   "dul"
         DisabledImageList=   "dis"
         HotImageList    =   "nrm"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Search"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   5235
         Begin VB.CheckBox Check2 
            Caption         =   "Match Case"
            Height          =   195
            Left            =   360
            TabIndex        =   10
            Top             =   1440
            Width           =   3255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Match Whole Word"
            Height          =   195
            Left            =   360
            TabIndex        =   9
            Top             =   1200
            Width           =   4455
         End
         Begin VB.TextBox txtSearch 
            Height          =   285
            Left            =   1140
            TabIndex        =   8
            Top             =   660
            Width           =   3975
         End
         Begin VB.ComboBox lstFields 
            Height          =   315
            Left            =   1140
            TabIndex        =   6
            Text            =   "<All>"
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label2 
            Caption         =   "Search Text:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Search Field:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   300
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Find In Group"
         Height          =   795
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   5235
         Begin VB.ComboBox lstGroups 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Text            =   "<All>"
            Top             =   300
            Width           =   4995
         End
      End
   End
   Begin MSComctlLib.ListView lstResults 
      Height          =   3255
      Left            =   180
      TabIndex        =   12
      Top             =   480
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Phone"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Notes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Group"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbSearch 
      Height          =   3795
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   6694
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Results"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ListGroups

lstGroups.AddItem "<All>"
lstGroups.AddItem "<None>"
lstGroups.Text = lstGroups.List(0)

lstFields.AddItem "<All>"

lstFields.AddItem "Full Name"
lstFields.AddItem "First Name"
lstFields.AddItem "Last Name"
lstFields.AddItem "Middle Name"

lstFields.AddItem "All Phone Numbers"
lstFields.AddItem "Home Phone Number"
lstFields.AddItem "Work Phone Number"
lstFields.AddItem "CellPhone Number"
lstFields.AddItem "Pager Number"

lstFields.AddItem "Email Address"
lstFields.AddItem "Street Address"
lstFields.AddItem "Notes"

lstFields.AddItem "Group Name"

End Sub

Private Sub tbSearch_Click()
lstResults.Visible = False

f_search.Visible = False
Select Case tbSearch.SelectedItem.Index
Case 1  'search
    f_search.Visible = True
    
Case 2  'results
    lstResults.Visible = True
    
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    Search lstGroups.Text, lstFields.Text, txtSearch.Text, Check1.VALUE, Check2.VALUE
    
End Select
End Sub

Public Sub Search(sGroupName As String, sFields As String, sFindText As String, Optional MtchWord As Byte = 0, Optional MtchCase As Byte = 0)
Dim CT As ListItem
lstResults.ListItems.Clear

If LCase(Trim(sGroupName)) = "<all>" Then
    sGroupName = "*"
End If
If LCase(Trim(sGroupName)) = "<none>" Then
    sGroupName = ""
End If

If MtchWord = 1 And MtchCase = 1 Then
    mde = 3
Else
    If MtchWord = 1 And MtchCase = 0 Then
        mde = 2
    Else
        If MtchCase = 1 And MtchWord = 0 Then
            mde = 1
        Else
            If MtchCase = 0 And MtchWord = 0 Then
                mde = 0
            End If
        End If
    End If
End If

With CONTACTS
    .MoveFirst
    Do While Not .EOF
        
        If Trim(sGroupName) <> "*" Then
            'search all
            If .Fields("GROUPNAME") <> sGroupName Then GoTo nxtCtct:
            GoTo SRCCH:
                
        Else
SRCCH:
        Select Case Trim(LCase(lstFields.Text))
                Case "<all>"
                    'all fields
                    If isThere(.Fields("NAME_FIRST"), sFindText, mde) = True Then GoTo Foundit:
                    If isThere(.Fields("NAME_MIDDLE"), sFindText, mde) = True Then GoTo Foundit:
                    If isThere(.Fields("NAME_LAST"), sFindText, mde) = True Then GoTo Foundit:
                    If isThere(.Fields("PHONE_HOME"), sFindText, mde) = True Then GoTo Foundit:
                    If isThere(.Fields("PHONE_WORK"), sFindText, mde) = True Then GoTo Foundit:
                    If isThere(.Fields("PHONE_CELL"), sFindText, mde) = True Then GoTo Foundit:
                    If isThere(.Fields("PHONE_PAGER"), sFindText, mde) = True Then GoTo Foundit:
                    If isThere(.Fields("EMAIL"), sFindText, mde) = True Then GoTo Foundit:
                    If isThere(.Fields("ADDRESS"), sFindText, mde) = True Then GoTo Foundit:
                    If isThere(.Fields("NOTES"), sFindText, mde) = True Then GoTo Foundit:
                Case "full name"
                    If isThere(.Fields("NAME_FIRST") & " " & .Fields("NAME_MIDDLE") & " " & .Fields("NAME_LAST"), sFindText, mde) = True Then GoTo Foundit:
                Case "last name"
                    If isThere(.Fields("NAME_LAST"), sFindText, mde) = True Then GoTo Foundit:
                Case "first name"
                    If isThere(.Fields("NAME_FIRST"), sFindText, mde) = True Then GoTo Foundit:
                Case "middle name"
                    If isThere(.Fields("NAME_MIDDLE"), sFindText, mde) = True Then GoTo Foundit:
                Case "all phone numbers"
                    If isThere(.Fields("PHONE_HOME") & " " & .Fields("PHONE_WORK") & " " & .Fields("PHONE_CELL") & " " & .Fields("PHONE_PAGER"), sFindText, mde) = True Then GoTo Foundit:
                Case "home phone number"
                    If isThere(.Fields("PHONE_HOME"), sFindText, mde) = True Then GoTo Foundit:
                Case "cellphone number"
                    If isThere(.Fields("PHONE_CELL"), sFindText, mde) = True Then GoTo Foundit:
                Case "work phone number"
                    If isThere(.Fields("PHONE_WORK"), sFindText, mde) = True Then GoTo Foundit:
                Case "pager number"
                    If isThere(.Fields("PHONE_PAGER"), sFindText, mde) = True Then GoTo Foundit:
                Case "email address"
                    If isThere(.Fields("EMAIL"), sFindText, mde) = True Then GoTo Foundit:
                Case "street address"
                    If isThere(.Fields("ADDRESS"), sFindText, mde) = True Then GoTo Foundit:
                Case "notes"
                    If isThere(.Fields("NOTES"), sFindText, mde) = True Then GoTo Foundit:
                Case "groupname"
                    If isThere(.Fields("GROUPNAME"), sFindText, mde) = True Then GoTo Foundit:
            End Select
            
        
        End If
GoTo nxtCtct:

Foundit:
    
    'find here
    Set CT = lstResults.ListItems.Add(, , .Fields("NAME_FIRST") & " " & .Fields("NAME_MIDDLE") & " " & .Fields("NAME_LAST"), 3, 3)
        CT.SubItems(1) = IIf(.Fields("PHONE_HOME") <> "", "Home: " & .Fields("PHONE_HOME"), "") & IIf(.Fields("PHONE_WORK") <> "", " Work: " & .Fields("PHONE_WORK"), "") & IIf(.Fields("PHONE_CELL") <> "", " Cell: " & .Fields("PHONE_CELL"), "") & IIf(.Fields("PHONE_PAGER") <> "", " Pager: " & .Fields("PHONE_PAGER"), "")
        CT.SubItems(2) = .Fields("EMAIL")
        CT.SubItems(3) = .Fields("ADDRESS")
        CT.SubItems(4) = .Fields("NOTES")
        CT.SubItems(5) = .Fields("GROUPNAME")
nxtCtct:
    .MoveNext
    DoEvents
    Loop
End With
End Sub

Private Function isThere(sTxt As String, sComp As String, ByVal sMode As Long) As Boolean
Dim TST() As String
isThere = False

Select Case sMode
    Case 1
        'match case
        TST = Split(sTxt, sComp)
        If UBound(TST) > 0 Then isThere = True: Exit Function
        Exit Function
    Case 2
        'match word
        TST = Split(LCase(sTxt), LCase(sComp))
        If UBound(TST) <= 0 Then isThere = False: Exit Function
        If Trim(Right(TST(0), 1)) = "" And Trim(Right(TST(1), 1)) = "" Then isThere = True: Exit Function
        
    Case 3
        'match word+case
         TST = Split(sTxt, sComp)
        If UBound(TST) <= 0 Then isThere = False: Exit Function
        If Trim(Right(TST(0), 1)) = "" And Trim(Right(TST(1), 1)) = "" Then isThere = True: Exit Function
        
    Case 0
        'match any
        TST = Split(LCase(sTxt), LCase(sComp))
        If UBound(TST) > 0 Then isThere = True: Exit Function
        Exit Function
End Select
End Function

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
'lstGroups.AddItem "<None>"
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
