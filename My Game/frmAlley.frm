VERSION 5.00
Begin VB.Form frmAlley 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alley"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "UNDER CONSTRUCTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   6255
   End
End
Attribute VB_Name = "frmAlley"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'This is a side project to add weird events or things
'havent worked on it
'add whatever you wish here


Private Sub cmdOK_Click()
Unload Me
End Sub
