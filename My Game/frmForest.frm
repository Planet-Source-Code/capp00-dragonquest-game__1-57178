VERSION 5.00
Begin VB.Form frmForest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DragonQuest"
   ClientHeight    =   7170
   ClientLeft      =   2370
   ClientTop       =   1905
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmForest.frx":0000
   ScaleHeight     =   478
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStats 
      BackColor       =   &H00808080&
      Cancel          =   -1  'True
      Caption         =   "Stats"
      Height          =   495
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Look for a fight"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Find the Dragon"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00808080&
      Caption         =   "&Go Back"
      Height          =   495
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO THE FOREST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmForest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    Unload Me
    frmWhere.Show
End Sub

Private Sub cmdStats_Click()
    frmStats.Show   'shows current stats
End Sub

'Search for the dragon

Private Sub Command1_Click()
    Dim intLook As Integer
    Randomize
    intLook = Int(Rnd * 30)
    If intLook > 25 Then
        CheckLevel  'checks to see if you are in a high enough level to fight dragon
    Else
        MsgBox "The dragon is no where to be found" & vbCrLf & vbCrLf & "Please try again later"
        Exit Sub
    End If
End Sub

'calls battle

Private Sub Command2_Click()
    frmBattle.Show
End Sub

Private Sub Form_Load()
    frmForest.Left = (Screen.Width - frmForest.Left) / 2        'center form
    frmForest.Top = (Screen.Height - frmForest.Height) / 2      'center form
End Sub


Sub CheckLevel()
   If Stats.intLevel < 8 Then
        MsgBox "You have found the dragon, but" & vbCrLf & "you are not strong enough to challenge it yet" & vbCrLf & vbCrLf & "Please try again later", vbInformation, "You Found the Dragon!"
        Exit Sub
    Else
        MsgBox "You have found the dragon." & vbCrLf & vbCrLf & "Prepare for battle..."
        frmDragon.Show
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmWhere.Show
End Sub
