VERSION 5.00
Begin VB.Form frmCharacters 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DragonQuest"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Data datChr 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "DragonQuest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Players"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Info"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&New"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   1095
      Left            =   1800
      TabIndex        =   2
      Top             =   2880
      Width           =   2415
   End
   Begin VB.ListBox lstName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      ItemData        =   "frmCharacters.frx":0000
      Left            =   840
      List            =   "frmCharacters.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Characters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This sub adds a new character to the database with default settings

Private Sub cmdAdd_Click()
    Dim strSQL As String, strName As String, intresults As Integer
    intresults = MsgBox("Do you want to add a new player?", vbYesNoCancel, "Add Player")
    If intresults = vbYes Then

        strName = InputBox("Enter new player's name", "Name")
        strSQL = "SELECT * FROM Players"
        datChr.RecordSource = strSQL
        datChr.Refresh
        datChr.Recordset.AddNew
            datChr.Recordset("Name") = strName
            datChr.Recordset("Level") = 1
            datChr.Recordset("BaseHP") = 30
            datChr.Recordset("HP") = 30
            datChr.Recordset("Str") = 30
            datChr.Recordset("BaseStr") = 30
            datChr.Recordset("Weapon") = "HANDS"
            datChr.Recordset("Armor") = "Naked"
            datChr.Recordset("Exp") = 0
            datChr.Recordset("Gold") = 0
            datChr.Recordset("Kills") = 0
        datChr.Recordset.Update
        lstName.AddItem strName
    Else
        Exit Sub
    End If

End Sub

'Delete the selected character

Private Sub cmdDelete_Click()
    Dim intresults As Integer, strName As String
    strName = lstName.Text
    intresults = MsgBox("Are you sure you want to delete " & strName, vbYesNo, "Delete")
    If intresults = vbYes Then
        datChr.Recordset.Delete
        lstName.RemoveItem lstName.ListIndex
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me 'Unload form from memory
    End
End Sub

'Get selected player's stats

Private Sub cmdInfo_Click()
    Dim strSQL As String
    strSQL = "SELECT * FROM Players WHERE Name = '" & lstName.Text & "'"
    datChr.RecordSource = strSQL
    datChr.Refresh

    frmStats.Show
End Sub

'Begin game with selected player

Private Sub cmdStart_Click()
    Unload Me
    Form1.Show
    Battling = False
End Sub

Private Sub Form_Load()
    frmCharacters.Left = (Screen.Width - frmCharacters.Width) / 2   'Centers form on screen
    frmCharacters.Top = (Screen.Height - frmCharacters.Height) / 2  'Centers form on screen
    Dim strSQL As String
        strSQL = "SELECT * FROM Players"
        datChr.RecordSource = strSQL
        datChr.Refresh
        If datChr.Recordset.EOF Then Exit Sub
            Do Until datChr.Recordset.EOF
                lstName.AddItem datChr.Recordset("Name")
                datChr.Recordset.MoveNext
            Loop
    cmdStart.Enabled = False
    cmdInfo.Enabled = False
    cmdDelete.Enabled = False
End Sub


'Loads character stats into variables to begin play

Private Sub lstName_Click()
     Dim strSQL As String, strSearch As String
    Dim intDamage As Integer, intStr As String
    strSQL = "SELECT * FROM Players WHERE Name = '" & lstName.Text & "'"
    datChr.RecordSource = strSQL
    datChr.Refresh
    intStr = datChr.Recordset("Str")
        Stats.strName = datChr.Recordset("Name")
        Stats.intLevel = datChr.Recordset("Level")
        Stats.intGold = datChr.Recordset("Gold")
        Stats.intMaxHP = datChr.Recordset("HP")
        Stats.intCurrentHP = datChr.Recordset("HP")
        Stats.intStr = intStr + intDamage
        Stats.strWeapon = datChr.Recordset("Weapon")
        Stats.strArmor = datChr.Recordset("Armor")
        Stats.IntExp = datChr.Recordset("Exp")
        Stats.intKills = datChr.Recordset("Kills")
        Stats.intBaseStr = datChr.Recordset("BaseStr")
        Modifier
    cmdStart.Enabled = True
    cmdInfo.Enabled = True
    cmdDelete.Enabled = True
End Sub
