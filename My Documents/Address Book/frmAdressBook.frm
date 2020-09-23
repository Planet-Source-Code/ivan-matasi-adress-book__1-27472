VERSION 5.00
Begin VB.Form frmBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address book"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAdressBook.frx":0000
   ScaleHeight     =   5730
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   120
      Width           =   2055
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   5535
         Left            =   0
         Picture         =   "frmAdressBook.frx":4586
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      Caption         =   "Surname and name"
      ForeColor       =   &H80000018&
      Height          =   3255
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   2775
      Begin VB.ListBox List1 
         Height          =   2790
         ItemData        =   "frmAdressBook.frx":7DFE
         Left            =   120
         List            =   "frmAdressBook.frx":7E00
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000010&
      Caption         =   "Options"
      ForeColor       =   &H80000018&
      Height          =   3255
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "New data input"
         DownPicture     =   "frmAdressBook.frx":7E02
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Erase selected data"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Erase all data"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000010&
      Caption         =   "Selected data information"
      ForeColor       =   &H80000018&
      Height          =   2175
      Left            =   2160
      TabIndex        =   0
      Top             =   3480
      Width           =   5175
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of birth:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblBirth 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lbladdress 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label lblsurname 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   2760
         TabIndex        =   15
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblPhone_nr 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Surname:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Image Image2 
      Height          =   5745
      Left            =   0
      Picture         =   "frmAdressBook.frx":B67A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7440
   End
End
Attribute VB_Name = "frmBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAdressBook As Recordset
Dim intAnswer As Integer

Private Sub Command1_Click()
frmInput.Show
End Sub

Private Sub Command2_Click()
intAnswer = MsgBox("Are you sure you want" & vbCrLf & "to delete information about" & vbCrLf & List1.List(List1.ListIndex) & "?", _
vbYesNo + vbQuestion, "Erase seleced record" & List1.List(List1.ListIndex))
If intAnswer = vbYes Then
rsAdressBook.MoveFirst
While Not rsAdressBook!surname & " " & rsAdressBook!Name = List1.List(List1.ListIndex)
  rsAdressBook.MoveNext
Wend
List1.RemoveItem List1.ListIndex
  rsAdressBook.Delete adAffectCurrent
  rsAdressBook.Update
  Unload Me
  Me.Show
  
End If

End Sub

Private Sub Command3_Click()
intAnswer = MsgBox("Are you sure you want" & vbCrLf & "to erase entire Addressbook?", _
vbYesNo + vbQuestion, "Erase entire Addressbook")
rsAdressBook.MoveFirst

If intAnswer = vbYes Then
While Not rsAdressBook.EOF
rsAdressBook.Delete adAffectCurrent
rsAdressBook.MoveNext
Wend

List1.Clear
Unload Me
  Me.Show
End If
End Sub

Private Sub Command4_Click()
intAnswer = MsgBox("Are you sure you want" & vbCrLf & "to exit address book?", vbYesNo + vbQuestion, "Exit?")
If intAnswer = vbYes Then Unload Me

End Sub

Private Sub Form_Load()
Set cnConnect1 = New Connection
  cnConnect1.Provider = "microsoft.jet.oledb.4.0"
  cnConnect1.ConnectionString = "user ID=admin; password=;" & _
  "data source=" & App.Path & "\address book.mdb"
  cnConnect1.Open
  
  Set rsAdressBook = New Recordset
  
  With rsAdressBook
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockOptimistic
  .Open "select * from Addressbook order by surname,name", cnConnect1
  End With
  
  While Not rsAdressBook.EOF
  List1.AddItem rsAdressBook!surname & " " & rsAdressBook!Name
  rsAdressBook.MoveNext
  Wend

If List1.ListCount > 0 Then
  List1.ListIndex = 0
  Else
  Command3.Enabled = False
  Command2.Enabled = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsAdressBook.ActiveConnection = Nothing
cnConnect1.Close
End Sub

Private Sub List1_Click()
Command3.Enabled = True
  Command2.Enabled = True
rsAdressBook.MoveFirst
While Not rsAdressBook!surname & " " & rsAdressBook!Name = List1.List(List1.ListIndex)
  rsAdressBook.MoveNext
Wend
lblName = rsAdressBook!Name
lblsurname = rsAdressBook!surname
lblPhone_nr = rsAdressBook!Phone_nr
lbladdress = rsAdressBook!address
lblBirth = rsAdressBook!Date_of_birth
End Sub
