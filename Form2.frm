VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Password"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form2.frx":030A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim fnum, pWord
fnum = FreeFile
Open "c:\windows\system\pword.txt" For Input As fnum
Input #fnum, pWord
Close fnum
If Text1.Text = pWord Then
frmMain.Show
Form2.Visible = False
Else
MsgBox "Wrong password !", vbCritical, "Crypt It 2.0"
End If
End Sub

Private Sub Form_Load()
Dim fnum
fnum = FreeFile
If Dir("c:\windows\system\pword.txt") = "" Then
Open "c:\windows\system\pword.txt" For Output As fnum
Print #fnum, "Emran"
Close fnum
MsgBox "You're running Crypt It 2.0 For the first time. Enter 'Emran' as the password. But change it from the Change password option for more security.", vbInformation, "Crypt It 2.0"
End If
End Sub
