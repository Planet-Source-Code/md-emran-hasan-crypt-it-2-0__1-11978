VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2835
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   2835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Encryption Method"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton optBin 
         Caption         =   "Binary Encryption"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton optRun 
         Caption         =   "Rudimentary Encryption"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optXOR 
         Caption         =   "Simple XOR Encryption"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim fNum
Dim EncMethod
fNum = FreeFile
If optRun.Value = True Then
    EncMethod = "Run"
ElseIf optXOR.Value = True Then
    EncMethod = "XOR"
Else
    EncMethod = "Bin"
End If

Open "Info.ini" For Output As fNum
Print #fNum, EncMethod
Close fNum

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim fNum
Dim m
fNum = FreeFile

Open "Info.ini" For Input As fNum
Input #fNum, m
Close fNum

If m = "Run" Then
    optRun.Value = True
    optXOR.Value = False
    optBin.Value = False
ElseIf m = "XOR" Then
    optXOR.Value = True
    optRun.Value = False
    optBin.Value = False
Else
    optBin.Value = True
    optXOR.Value = False
    optRun.Value = False
End If
    
End Sub
