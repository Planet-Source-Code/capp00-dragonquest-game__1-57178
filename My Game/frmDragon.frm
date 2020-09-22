VERSION 5.00
Begin VB.Form frmDragon 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   Caption         =   "DragonQuest"
   ClientHeight    =   7950
   ClientLeft      =   1800
   ClientTop       =   1320
   ClientWidth     =   7770
   LinkTopic       =   "Form2"
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   518
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3360
      Top             =   7320
   End
   Begin VB.CommandButton cmdAttack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Attack"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Attempt to &Run"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      X1              =   88
      X2              =   168
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   344
      X2              =   424
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   344
      X2              =   424
      Y1              =   328
      Y2              =   328
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      X1              =   88
      X2              =   168
      Y1              =   328
      Y2              =   328
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   8
      Left            =   6000
      Picture         =   "frmDragon.frx":0000
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   7
      Left            =   5400
      Picture         =   "frmDragon.frx":0C42
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   6
      Left            =   4800
      Picture         =   "frmDragon.frx":1884
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   5
      Left            =   4200
      Picture         =   "frmDragon.frx":24C6
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   4
      Left            =   3600
      Picture         =   "frmDragon.frx":3108
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   3
      Left            =   3000
      Picture         =   "frmDragon.frx":3D4A
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   2
      Left            =   1800
      Picture         =   "frmDragon.frx":498C
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   1
      Left            =   2400
      Picture         =   "frmDragon.frx":55CE
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   0
      Left            =   1200
      Picture         =   "frmDragon.frx":6210
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Your Roll"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "Enemy Roll"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "Your HP"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Caption         =   "Dragon HP"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lblEnemyHP 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   6480
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   464
      X2              =   464
      Y1              =   304
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   40
      X2              =   464
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   40
      X2              =   40
      Y1              =   0
      Y2              =   304
   End
   Begin VB.Image Image1 
      Height          =   4665
      Left            =   600
      Picture         =   "frmDragon.frx":6E52
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   6330
   End
End
Attribute VB_Name = "frmDragon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************
'This is the final battle for the game
'******************************


Option Explicit
    Dim Damage As Integer
    Dim MonsterDamage As Integer
    Dim MonsterHitPoints As Integer

'Sub for attacking the dragon - starts dice roll

Private Sub cmdAttack_Click()
            If MonsterHitPoints < 1 Then        'if you kill the dragon
                Battling = False                'Cancels battle
                WIN                             'calls Win sub
                Dragon = False                  'sets dragon fight to false
                lblEnemyHP.Caption = 0
                Save                            'saves stats
                cmdAttack.Enabled = False
                cmdUse.Enabled = False
                cmdRun.Enabled = False
                cmdOK.Enabled = True
                Unload frmBattle
                Exit Sub
            End If

    Dim PlayerRoll As Integer
    Dim MonsterRoll As Integer
    Dim Alive As Boolean
    Dim X As Integer
    X = Stats.intBaseStr
    cmdAttack.Enabled = False
    cmdUse.Enabled = False
    cmdOK.Enabled = False
'Same as for the battle
'uses randome dice rolls
'then determines damage

    PlayerRoll = RollDice(120, 336) 'paint dice
        Delay (0.2)
        MonsterRoll = RollDice(360, 336)    'paint dice
        Damage = (X / 10) + Stats.intLevel + Int(Rnd * 30)
        MonsterDamage = Int(Rnd * 30) + 10
            If MonsterDamage < 1 Then MonsterDamage = 1
        If PlayerRoll >= MonsterRoll Then
            Delay (0.1)
            MonsterHitPoints = MonsterHitPoints - Damage
            lblEnemyHP.Caption = MonsterHitPoints
            cmdAttack.Enabled = True
            cmdAttack.Enabled = True
            cmdUse.Enabled = True
            cmdRun.Enabled = True

            If MonsterHitPoints < 1 Then
                Battling = False
                Dragon = False
                WIN
                lblEnemyHP.Caption = "0"
                Save
                cmdAttack.Enabled = False
                cmdUse.Enabled = False
                cmdRun.Enabled = False
                cmdOK.Enabled = True
                Unload frmDragon
            End If
        ElseIf PlayerRoll < MonsterRoll Then
            Stats.intCurrentHP = Stats.intCurrentHP - MonsterDamage
            lblHP.Caption = Stats.intCurrentHP
            Delay (0.2)
            cmdAttack.Enabled = True
            cmdUse.Enabled = True
            cmdRun.Enabled = True
            If Stats.intCurrentHP <= 0 Then
                lblHP.Caption = "0"
                Battling = False
                Dead
            End If
        End If
    

End Sub

'Call inventory screen

Private Sub cmdItem_Click()
    frmInventory.Show
End Sub

'Attempt to run from battle

Private Sub cmdRun_Click()
Dim intChance As Integer    'creates random chance
intChance = Int(Rnd * 10)
    If intChance >= 7 Then
        MsgBox "You barely escape", vbInformation, "Run"
        Unload Me
    Else
        MsgBox "You attempt to run" & vbCrLf & vbCrLf & "...But cannot escape", vbInformation, "Run"
        Exit Sub
    End If
End Sub

Private Sub Command1_Click()
    Unload frmDragon
End Sub

Private Sub cmdUse_Click()
    Load frmInventory
    frmInventory.Show
End Sub

Private Sub Form_Load()
    frmDragon.Left = Me.Left
    frmDragon.Top = Me.Top
    Dim strHP As String
    Dragon = True
    Battling = True
    Timer1.Enabled = False
    MonsterHitPoints = 750
        strHP = Stats.intCurrentHP
        lblHP.Caption = strHP
        lblEnemyHP.Caption = MonsterHitPoints

End Sub

'Dice roll function
'same as battle

Function RollDice(X As Integer, Y As Integer)
    Dim num As Integer
    Dim roll As Integer
    For roll = 1 To 10
    Randomize
        num = Int(Rnd * 9)
        frmDragon.PaintPicture imgDice(num), X, Y
        Delay (0.1)
    Next roll
    RollDice = num
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Battling = True Then
        Exit Sub
    End If
End Sub
