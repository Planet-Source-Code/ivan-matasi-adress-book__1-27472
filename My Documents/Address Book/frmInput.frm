VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Input"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInput.frx":0000
   ScaleHeight     =   1935
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd. MM. yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1050
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of birth:"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone number:"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Surname:"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAdressBook As Recordset

Private Sub Command1_Click()
rsAdressBook.AddNew

rsAdressBook!Name = Text1.Text
rsAdressBook!surname = Text2.Text
rsAdressBook!Phone_nr = Text3.Text
rsAdressBook!address = Text4.Text
rsAdressBook!Date_of_birth = Text5.Text
rsAdressBook.Update

Unload frmBook
Load frmBook
frmBook.Show
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set rsAdressBook = New Recordset
  
  With rsAdressBook
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockOptimistic
  .Open "select * from Addressbook", cnConnect1
  End With
  
  If Text1.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If

If Text2.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub

Private Sub Text1_Change()
 If Text1.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If

If Text2.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub

Private Sub Text2_Change()

If Text1.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If

If Text2.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub
