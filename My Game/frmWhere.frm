VERSION 5.00
Begin VB.Form frmWhere 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DragonQuest"
   ClientHeight    =   3090
   ClientLeft      =   4095
   ClientTop       =   3120
   ClientWidth     =   7230
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7230
   Begin VB.CommandButton Command2 
      Caption         =   "&GO TRAIN"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "&HOME"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdForest 
      Caption         =   "&FOREST"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdTown 
      Caption         =   "&TOWN"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "WHRE DO YOU WANT TO GO?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************
'This is pretty self explainatory
'it just give you the option of choosing where to go
'******************************************

Option Explicit

Private Sub cmdForest_Click()
    Unload Me
    frmForest.Show
End Sub

Private Sub cmdHome_Click()
    Unload Me
    Form1.Show
End Sub

Private Sub cmdTown_Click()
    Unload Me
    frmOut.Show
End Sub

Private Sub Command2_Click()
    Unload Me
    frmTraining.Show
End Sub

Private Sub Form_Load()
    frmWhere.Left = (Screen.Width / 2 - frmWhere.Width / 2)
    frmWhere.Top = (Screen.Height / 2 - frmWhere.Height / 2)
End Sub
