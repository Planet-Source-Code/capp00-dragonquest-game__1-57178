VERSION 5.00
Begin VB.Form frmTraining 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DragonQuest"
   ClientHeight    =   7185
   ClientLeft      =   2655
   ClientTop       =   1905
   ClientWidth     =   9570
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTraining.frx":0000
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   638
   Begin VB.CommandButton Command2 
      Caption         =   "Stats"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdStats 
      Caption         =   "Trainer &Stats"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Data datTrain 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "DragonQuest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Trainers"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdTrain 
      Caption         =   "&Train"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Go Back"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "frmTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStats_Click()
Dim strSQL As String, intLevel As String
intLevel = Stats.intLevel
    strSQL = "SELECT * FROM Trainers WHERE Level = '" & intLevel & "'"
    datTrain.RecordSource = strSQL
    datTrain.Refresh
            MsgBox "TRAINER STATS" & vbCrLf & vbCrLf & "Name:  " & datTrain.Recordset("Name") & vbCrLf & "Required Exp:  " & datTrain.Recordset("Exp")
            
End Sub

'If you are strong enough to challenge the trainer
'this is the battle pretty much
'it is scripted so there is no real battle

Private Sub cmdTrain_Click()
    Dim strSQL As String, intLevel As String
    Dim IntExp As Integer, intX As Integer
    Dim intTrainerGold As Integer
    intLevel = Stats.intLevel
    IntExp = Stats.IntExp
    strSQL = "SELECT * FROM Trainers WHERE Level = '" & intLevel & "'"
    datTrain.RecordSource = strSQL
    datTrain.Refresh
    intX = datTrain.Recordset("Exp")
    intTrainerGold = datTrain.Recordset("Gold")
    If IntExp >= intX Then
        MsgBox "You begin dualling your trainer"
        MsgBox "THE BATTLE" & vbCrLf & vbCrLf & datTrain.Recordset("Name") & " " & datTrain.Recordset("Attack") & vbCrLf & vbCrLf & _
            "You dodge the attack" & vbCrLf & "You catch your trainer off guard with a clean hit"
        MsgBox "You have defeated your trainer"     'You Win
        Stats.intLevel = Stats.intLevel + 1         'Level goes up
        Stats.intGold = Stats.intGold + intTrainerGold  'gold goes up
        Stats.intMaxHP = (Stats.intMaxHP / 2) + Stats.intMaxHP  'stats go up
        Stats.intCurrentHP = Stats.intMaxHP
        Stats.intStr = Stats.intStr + 10
        Save                                         'save
        frmStats.Show
        Exit Sub
    Else
        MsgBox "You are not experience enough to challenge a trainer"
        Exit Sub
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
    frmWhere.Show
End Sub

Private Sub Command2_Click()
    frmStats.Show
End Sub

Private Sub Form_Load()
    frmTraining.Left = (Screen.Width - frmTraining.Width) / 2       'center on screen
    frmTraining.Top = (Screen.Height - frmTraining.Height) / 2      'center on screen
End Sub


