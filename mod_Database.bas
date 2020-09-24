Attribute VB_Name = "mod_Database"
Public ADDRESSBOOK As Database
Public GROUPS As Recordset
Public CONTACTS As Recordset
Public CONFIG As Recordset

Public WEbon As Boolean
Public Type WebClient
    sAuthenticated As Boolean
    sIp As String
End Type

Public Client As WebClient
Public Const ChunkSize As Long = 4096




Public Function AddGroup(sGroupName As String) As Long
On Error GoTo ERR1:
With GROUPS
    .AddNew
        !groupname = sGroupName
    .Update
End With
AddGroup = 1
frmContacts.RefContactList

Exit Function
ERR1:
AddGroup = -1
If Not (WEbon) Then
NewError Err.Number, 2, "", sGroupName
Else
    SendWeb "The group " & sGroupName & " Exists and will not be added!"
End If
End Function

Public Sub LoadDatabase(Optional sFile As String = "data.mdb")
On Error GoTo ERR1:

sFile = GetSetting(App.EXEName, "Set", "Path", "data.mdb")

Set ADDRESSBOOK = OpenDatabase(IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & sFile)
Set GROUPS = ADDRESSBOOK.OpenRecordset("GROUPS", dbOpenDynaset)
Set CONTACTS = ADDRESSBOOK.OpenRecordset("CONTACTS", dbOpenDynaset)
Set CONFIG = ADDRESSBOOK.OpenRecordset("CONFIG", dbOpenDynaset)

Exit Sub
ERR1:
MsgBox "Error while initializing database: " & sFile & vbCrLf & vbCrLf & Err.Description, vbCritical, "Database Error!"
End Sub

Public Sub UnloadDatabase()
GROUPS.Close
CONTACTS.Close
CONFIG.Close

ADDRESSBOOK.Close
End Sub

Public Function RemoveGroup(sGroupName As String)
On Error GoTo ERR1:

Dim UC As Long

If Trim(sGroupName) = "" Then Exit Function

With CONTACTS
    .MoveFirst
    Do While Not .EOF
        If .Fields("GROUPNAME") = sGroupName Then
            UC = UC + 1
        End If
    .MoveNext
    DoEvents
    Loop
End With

If UC > 0 Then
    'users are in this group
    If Not (WEbon) Then
    If MsgBox("There are " & UC & " contact(s) in this group, this action will also remove these contact entries, are you sure you want to delete the group '" & sGroupName & "' ?", vbExclamation + vbYesNo, "Contacts in group") = vbNo Then Exit Function
    
        ClsGroup sGroupName
    Else
        ClsGroup sGroupName
    End If
    
End If

With GROUPS
    .MoveFirst
    Do While Not .EOF
        If .Fields("GROUPNAME") = sGroupName Then
            .Delete
            '.Update
            RemoveGroup = 1
            frmContacts.RefContactList
            Exit Function
        End If
        .MoveNext
    DoEvents
    Loop
End With

RemoveGroup = -1
If Not (WEbon) Then
MsgBox "Could not Find the group '" & sGroupName & "'", vbExclamation + vbApplicationModal, "Error!"
End If
Exit Function
ERR1:
If Not (WEbon) Then
MsgBox "Database error: " & vbCrLf & vbCrLf & Err.Description, vbCritical, "Database Error!"
End If
End Function


Public Function GENID() As Long
Randomize
GENID = Int(Rnd * 1000000000)
End Function

Public Function AddContact(snameFirst As String, snameMiddle As String, snameLast As String, sPhoneHome As String, sPhoneWork As String, sPhoneCell As String, sPhonePager As String, sEmail As String, sAddress As String, sNotes As String, Optional sContactGroup As String = "")
'On Error GoTo ERR1:

AddContact = -1
If WEbon = False Then
If Trim(snameFirst) = "" Then MsgBox "A First name is required!", vbExclamation, "Error": Exit Function
End If
If sContactGroup <> "" Then
    If GroupExists(sContactGroup) = False Then
        If WEbon = False Then
        If MsgBox("The group '" & sContactGroup & "' does not exist, Would you like to create it now?", vbYesNo + vbExclamation, "Could not find group") = vbNo Then
            AddContact = -1
            Exit Function
        Else
            AddGroup sContactGroup
        End If
        Else
            AddGroup sContactGroup
        End If
    End If
End If

With CONTACTS
    .AddNew
        !NAME_FIRST = snameFirst
        !NAME_MIDDLE = snameMiddle
        !NAME_LAST = snameLast
        !PHONE_HOME = sPhoneHome
        !PHONE_WORK = sPhoneWork
        !PHONE_CELL = sPhoneCell
        !PHONE_PAGER = sPhonePager
        !EMAIL = sEmail
        !ADDRESS = sAddress
        !NOTES = sNotes
        !groupname = sContactGroup
    .Update
End With
frmContacts.RefContactList

AddContact = 1
Exit Function
ERR1:
AddContact = -1
NewError Err.Number, 1, sContactName
End Function

Public Function GroupExists(sGroupName As String) As Boolean
GroupExists = False
With GROUPS
    .MoveFirst
    Do While Not .EOF
        If .Fields("GROUPNAME") = sGroupName Then
            GroupExists = True
            Exit Function
        End If
    DoEvents
    .MoveNext
    Loop
End With
GroupExists = False
End Function


Public Sub ClsGroup(sGroupName As String)
With CONTACTS
    .MoveFirst
    Do While Not .EOF
        If .Fields("GROUPNAME") = sGroupName Then
            .Delete
            
        End If
    .MoveNext
    DoEvents
    Loop
    End With
End Sub

Public Function FindGroupID(sGroupName As String) As Integer
FindGroupID = -1
With GROUPS
    .MoveFirst
    Do While Not .EOF
        If .Fields("GROUPNAME") = sGroupName Then
            FindGroupID = .Fields("GROUPID")
            Exit Function
        End If
    DoEvents
    .MoveNext
    Loop
End With

End Function

Public Function MoveContacts(sID As String) As Long
On Error Resume Next
With CONTACTS
    .MoveFirst
    Do While Not .EOF
        If .Fields("CONTACTID") = sID Then
        MoveContacts = 1
        Exit Function
            
        End If
    .MoveNext
    DoEvents
    Loop
    End With
    MoveContacts = -1
End Function

Public Function MoveGroups(ByVal sID As String) As Long
On Error Resume Next
With GROUPS
    .MoveFirst
    Do While Not .EOF
        If .Fields("GROUPID") = sID Then
        MoveGroups = 1
        Exit Function
            
        End If
    .MoveNext
    DoEvents
    Loop
    End With
    MoveGroups = -1
End Function

Public Sub RemoveContact(sContactID As String)
MoveContacts sContactID
With CONTACTS
    If MsgBox("Are you sure you want to remove this contact?" & vbCrLf & vbCrLf & .Fields("NAME_FIRST") & " " & .Fields("NAME_MIDDLE") & " " & .Fields("NAME_LAST"), vbExclamation + vbYesNo, "Confirm Remove contact") = vbNo Then Exit Sub
    .Delete
End With
frmContacts.RefContactList
End Sub

Public Sub NewError(Optional ByVal errNum As Long, Optional ByVal OtherDat As Long, Optional ByVal sContactName As String, Optional ByVal sGroupName As String)
Select Case Err.Number
    Case 3022 'duplicate
        Select Case OtherDat
            Case 1
                MsgBox "The Contact '" & sContactName & "' Exists!" & vbCrLf & vbCrLf & "Please choose a unique contact name", vbExclamation + vbApplicationModal, "Duplicate Contact"
            Case 2
                MsgBox "The group '" & sGroupName & "' Exists!" & vbCrLf & vbCrLf & "Please choose a unique group name", vbExclamation + vbApplicationModal, "Duplicate Group"
            Case Else
        End Select
    Case Else
        MsgBox "Unknown database error: " & vbCrLf & vbCrLf & Err.Description, vbCritical, "Database Error!"
        
End Select
End Sub

Public Sub SendWeb(sMsg As String, Optional sData As String = -1)

On Error Resume Next

Dim WEB_HEADer As String
Dim WEB_PATH As String
Dim DTS As String
Dim DATA_FILE As String
Dim N_L As String
Dim p_L As String
Dim e_L As String
Dim a_L As String
Dim No_L As String
Dim g_L As String
Dim CR As String

        
Dim mde As Long


With CONFIG
    If .Fields("WEB_PATH") = "" Then
        WEB_PATH = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "web\"
    Else
        WEB_PATH = .Fields("WEB_PATH")
    End If
End With

DTS = ""

If frmMain.sckServer.State <> sckConnected Then Exit Sub
    Select Case Trim(LCase(sMsg))
        Case "authenticate"
            DATA_FILE = "authenticate.html"
                mde = 0
                GoTo LOADFILE:
        Case "invalidauth"
            DATA_FILE = "passrejected.html"
                mde = 1
                GoTo LOADFILE:
        Case "addressbook"
            DATA_FILE = "addressbookm.html"
                mde = 2
                GoTo LOADFILE:
        Case "viewcontact"
            DATA_FILE = "viewcontact.html"
            mde = 3
            GoTo LOADFILE:
        Case "removecontact"
            DATA_FILE = "confirmdelete.html"
            mde = 4
            GoTo LOADFILE:
        Case "addcontact"
            DATA_FILE = "contactadd.html"
            mde = 5
            GoTo LOADFILE:
        Case "addgroup"
            DATA_FILE = "addgroup.html"
            mde = 6
            GoTo LOADFILE:
        Case "removegroup"
            DATA_FILE = "confirmdeleteg.html"
            mde = 7
            GoTo LOADFILE:
        Case "listc"
            DATA_FILE = "clist.html"
            mde = 8
            GoTo LOADFILE:
        Case Else
            mde = -1
            'Exit Sub
    End Select

If False Then
LOADFILE:
DTS = Space$(FileLen(WEB_PATH & DATA_FILE))
Open WEB_PATH & DATA_FILE For Binary Access Read As #1
    Get #1, , DTS
Close #1
End If



Select Case mde
    Case 2
        Dim FORMEDLIST As String
        Dim HH() As String
        
        FORMEDLIST = FORMEDLIST & "<option value=" & """" & -999 & """" & ">" & "Contact List" & "</option>" & vbCrLf
        FORMEDLIST = FORMEDLIST & "<option value=" & """" & -999 & """" & ">" & "--------------------------------------------------" & "</option>" & vbCrLf
        
        With GROUPS
            .MoveFirst
            Do While Not .EOF
                'FORMEDLIST = FORMEDLIST & "<option value=" & """" & -999 & """" & ">" & "     ---------------------------" & "</option>" & vbCrLf
                FORMEDLIST = FORMEDLIST & "<option value=" & """" & .Fields("GROUPID") & """" & ">" & .Fields("GROUPNAME") & "</option>" & vbCrLf
                'FORMEDLIST = FORMEDLIST & "<option value=" & """" & -999 & """" & ">" & "     ---------------------------" & "</option>" & vbCrLf
                With CONTACTS
                    .MoveFirst
                    Do While Not .EOF
                        If .Fields("GROUPNAME") = GROUPS.Fields("GROUPNAME") Then
                            FORMEDLIST = FORMEDLIST & "<option value=" & """" & .Fields("CONTACTID") & """" & ">" & "+ " & .Fields("NAME_LAST") & "," & .Fields("NAME_FIRST") & " " & .Fields("NAME_MIDDLE") & "</option>" & vbCrLf
                        End If
                        
                    .MoveNext
                    
                    DoEvents
                    Loop
                End With
            DoEvents
            .MoveNext
            
            Loop
            
            With CONTACTS
                    .MoveFirst
                    Do While Not .EOF
                        If Trim(.Fields("GROUPNAME")) = "" Then
                            FORMEDLIST = FORMEDLIST & "<option value=" & """" & .Fields("CONTACTID") & """" & ">" & "- " & .Fields("NAME_LAST") & "," & .Fields("NAME_FIRST") & " " & .Fields("NAME_MIDDLE") & "</option>" & vbCrLf
                        End If
                        
                    .MoveNext
                    
                    DoEvents
                    Loop
                End With
            '<option value="sel">sel</option>
        End With
        DTS = ReplaceString(DTS, "!CONTACTLIST!", FORMEDLIST)
        DTS = ReplaceString(DTS, "!IP!", frmMain.sckServer.RemoteHostIP)
    Case 3
        If sData <> -1 Then
        If MoveContacts(sData) = 1 Then GoTo founduser:
        DTS = "Invalid User ID!"
        GoTo SND:
        End If
founduser:
        DTS = GetVars(DTS)
    Case 4
        If sData = -1 Then GoTo founduser2:
        If MoveContacts(sData) = 1 Then GoTo founduser2:
        DTS = "Invalid User ID!"
        GoTo SND:
founduser2:
        DTS = GetVars(DTS)
    Case 5
        FORMEDLIST = ""
        With GROUPS
            .MoveFirst
            Do While Not .EOF
                
                FORMEDLIST = FORMEDLIST & "<option value=" & """" & .Fields("GROUPNAME") & """" & ">" & .Fields("GROUPNAME") & "</option>" & vbCrLf
                .MoveNext
            DoEvents
            Loop
            FORMEDLIST = FORMEDLIST & "<option value=" & """" & -999 & """" & ">" & "<NONE>" & "</option>" & vbCrLf
        End With
        DTS = ReplaceString(DTS, "!GROUPLIST!", FORMEDLIST)
    Case 7
        If MoveGroups(sData) <> -1 Then
            With GROUPS
            DTS = ReplaceString(DTS, "!GROUPNAME!", .Fields("GROUPNAME"))
            End With
        Else
            DTS = "Invalid Group ID!"
            GoTo SND:
        End If
    Case 8
        N_L = ""
        p_L = ""
        e_L = ""
        a_L = ""
        No_L = ""
        g_L = ""
        
        With CONTACTS
            .MoveFirst
            Do While Not .EOF
            
            N_L = .Fields("NAME_FIRST") & " " & .Fields("NAME_MIDDLE") & .Fields("NAME_LAST")
            p_L = "Home: " & .Fields("PHONE_HOME") & " Work: " & .Fields("PHONE_WORK") & " Cell: " & .Fields("PHONE_CELL") & " Pgr: " & .Fields("PHONE_PAGER")
            e_L = .Fields("EMAIL")
            a_L = .Fields("ADDRESS")
            No_L = .Fields("NOTES")
            g_L = .Fields("GROUPNAME")
            
            CR = CR & "<tr><td>" & N_L & "</td>" & "<td>" & g_L & "</td>" & "<td>" & p_L & "</td>" & "<td>" & e_L & "<td>" & "</td>" & "<td>" & a_L & "</td>" & "<td>" & No_L & "</td>" & "</tr>"
            .MoveNext
            
            DoEvents
            Loop
        End With
        
        'DTS = CR
        DTS = ReplaceString(DTS, "!NAMELIST!", CR)
        'DTS = ReplaceString(DTS, "!PHONELIST!", p_L)
        'DTS = ReplaceString(DTS, "!EMAILLIST!", e_L)
        'DTS = ReplaceString(DTS, "!ADDRESSLIST!", a_L)
        'DTS = ReplaceString(DTS, "!NOTESLIST!", No_L)
        'DTS = ReplaceString(DTS, "!GROUPLIST!", g_L)
    Case -1
        DTS = sMsg
    Case Else
End Select

SND:
If Len(DTS) = 0 Then
fh:
    WEB_HEADer = "HTTP/1.1 404 Addressbook Error" & vbCrLf & _
        "Connection: Close" & vbCrLf & _
        vbCrLf & vbCrLf
Else
    WEB_HEADer = "HTTP/1.1 200 Data follows..." & vbCrLf & _
        "Connection: Close" & vbCrLf & _
        vbCrLf & vbCrLf
End If

DTS = WEB_HEADer & DTS

If Len(DTS) > ChunkSize Then
    'send in chunks
    For i = 1 To Len(DTS) Step ChunkSize
        frmMain.sckServer.SendData Mid(DTS, i, ChunkSize)
        DoEvents
        
    Next
Else
    'send in one chunk
    frmMain.sckServer.SendData DTS
End If

DoEvents
If frmMain.sckServer.State <> sckClosed Then frmMain.sckServer.Close


End Sub

Public Sub SendWebSimple(sMsg As String)
Dim WEB_HEADer As String

If frmMain.sckServer.State <> sckConnected Then Exit Sub
WEB_HEADer = "HTTP/1.1 200 Data follows..." & vbCrLf & _
        "Connection: Keep-Alive" & vbCrLf & _
        vbCrLf & vbCrLf
frmMain.sckServer.SendData WEB_HEADer & sMsg
End Sub

Public Function GetVars(sText As String) As String
Dim NSTR As String

With CONTACTS

NSTR = ReplaceString(sText, "!NAMEFIRST!", .Fields("NAME_FIRST"))
NSTR = ReplaceString(NSTR, "!NAMEMIDDLE!", .Fields("NAME_MIDDLE"))
NSTR = ReplaceString(NSTR, "!NAMELAST!", .Fields("NAME_LAST"))

NSTR = ReplaceString(NSTR, "!EMAIL!", .Fields("EMAIL"))
NSTR = ReplaceString(NSTR, "!ADDRESS!", .Fields("ADDRESS"))
NSTR = ReplaceString(NSTR, "!NOTES!", .Fields("NOTES"))

NSTR = ReplaceString(NSTR, "!PHONEHOME!", .Fields("PHONE_HOME"))
NSTR = ReplaceString(NSTR, "!PHONEWORK!", .Fields("PHONE_WORK"))
NSTR = ReplaceString(NSTR, "!PHONECELL!", .Fields("PHONE_CELL"))
NSTR = ReplaceString(NSTR, "!PHONEPAGER!", .Fields("PHONE_PAGER"))

NSTR = ReplaceString(NSTR, "!GROUPNAME!", .Fields("GROUPNAME"))
NSTR = ReplaceString(NSTR, "!CONTACTID!", .Fields("CONTACTID"))

End With
GetVars = NSTR
End Function

Public Function ReplaceString(sText As String, srchText As String, replText As String) As String
Dim NEWSTRING As String

Dim spl() As String
spl = Split(sText, UCase(srchText))
If UBound(spl) > 0 Then
    'in string
    For i = 0 To UBound(spl)
        If i <> UBound(spl) Then
            NEWSTRING = NEWSTRING & spl(i) & replText
        Else
            NEWSTRING = NEWSTRING & spl(i)
        End If
    Next
Else
    NEWSTRING = sText
End If

ReplaceString = NEWSTRING
End Function
