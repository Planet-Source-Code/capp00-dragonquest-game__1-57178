VERSION 5.00
Begin VB.Form frmBattle 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DragonQuest"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5730
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   382
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Data datBattle 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "DragonQuest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Monsters"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAttack 
      Caption         =   "&Attack"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdUse 
      Caption         =   "&Use Item"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Image imgWater 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   2040
      Picture         =   "frmBattle.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image imgMonster 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   1200
      Picture         =   "frmBattle.frx":17C24A
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image imgReaper 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   360
      Picture         =   "frmBattle.frx":187A3C
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image imgWizard 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   2280
      Picture         =   "frmBattle.frx":192ED6
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image imgSpider 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   1680
      Picture         =   "frmBattle.frx":19870C
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image imgOrch 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   720
      Picture         =   "frmBattle.frx":1AEA9E
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image imgLizMan 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   120
      Picture         =   "frmBattle.frx":1B9078
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image imgTree 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   2760
      Picture         =   "frmBattle.frx":1C8F2F
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image imgDwarf 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   2280
      Picture         =   "frmBattle.frx":1ECC6C
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblEnemyHP 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   975
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   272
      X2              =   352
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Enemy HP"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   4320
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   16
      X2              =   96
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Your HP"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   4320
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   272
      X2              =   352
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Enemy Roll"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   16
      X2              =   96
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Your Roll"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmBattle.frx":20E1EC
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   1
      Left            =   1320
      Picture         =   "frmBattle.frx":20EE2E
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   2
      Left            =   720
      Picture         =   "frmBattle.frx":20FA70
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   3
      Left            =   1920
      Picture         =   "frmBattle.frx":2106B2
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   4
      Left            =   2520
      Picture         =   "frmBattle.frx":2112F4
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   5
      Left            =   3120
      Picture         =   "frmBattle.frx":211F36
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   6
      Left            =   3720
      Picture         =   "frmBattle.frx":212B78
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   7
      Left            =   4320
      Picture         =   "frmBattle.frx":2137BA
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   8
      Left            =   4920
      Picture         =   "frmBattle.frx":2143FC
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSallos 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   1800
      Picture         =   "frmBattle.frx":21503E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2385
   End
   Begin VB.Image imgSnake 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   1800
      Picture         =   "frmBattle.frx":2165DD
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2085
   End
   Begin VB.Image imgIpos 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   840
      Picture         =   "frmBattle.frx":216CCB
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2385
   End
   Begin VB.Image imgCarc 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   480
      Picture         =   "frmBattle.frx":217FF5
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2385
   End
   Begin VB.Image imgBarbados 
      BorderStyle     =   1  'Fixed Single
      Height          =   2715
      Left            =   240
      Picture         =   "frmBattle.frx":218F2E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2385
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'*****************************************
'This is where all the action takes place
'The battle zone
'*****************************************

Option Explicit
    Dim Damage As Integer
    Dim IntExp As Integer
    Dim intGold As Integer
    Dim MonsterDamage As Integer
    Dim MonsterLife As Integer


'Starts dice roll to determine who won the attack and hands out damage

Private Sub cmdAttack_Click()
            If MonsterHitPoints < 1 Then    'checks to see if monster is dead yet
                Battling = False
                MsgBox "You have defeated the enemy"
                MsgBox "You have gained " & IntExp & " Experience" & vbCrLf & vbCrLf & _
                    "You have gained " & intGold & " Gold"
                Stats.IntExp = Stats.IntExp + IntExp
                Stats.intGold = Stats.intGold + intGold
                Stats.intKills = Stats.intKills + 1
                Battling = False            'ends battle
                lblEnemyHP.Caption = 0
                Save                        'calls save
                cmdAttack.Enabled = False
                cmdUse.Enabled = False
                cmdRun.Enabled = False
                cmdOK.Enabled = True
                Unload frmBattle
                Exit Sub
            End If

    Dim PlayerRoll As Integer           'stores player dice roll
    Dim MonsterRoll As Integer          'stores monster dice roll
    Dim Alive As Boolean                'checks if you are alive still
    Dim X As Integer
    X = Stats.intBaseStr
    cmdAttack.Enabled = False
    cmdUse.Enabled = False
    cmdOK.Enabled = False
    PlayerRoll = RollDice(32, 216)      'Roll dice for player and places coords for dice pics
        Delay (0.2)                     'calls Delay Timer
        MonsterRoll = RollDice(296, 216)    'roll dice for monster and places coords for dice pics
        Damage = (X / 5) + Stats.intLevel + Int(Rnd * 5)    'gets random damage amount
        MonsterDamage = Int(Rnd * 10)       'monsters damage amount
            If MonsterDamage < 1 Then MonsterDamage = 1
        If PlayerRoll >= MonsterRoll Then   'compares dice rolls
            Delay (0.1)                     'Calls Delay timer
            MonsterHitPoints = MonsterHitPoints - Damage    'calculates damage and remaining hitpoints
            lblEnemyHP.Caption = MonsterHitPoints
            cmdAttack.Enabled = True
            cmdAttack.Enabled = True
            cmdUse.Enabled = True
            cmdRun.Enabled = True

            If MonsterHitPoints < 1 Then     'if the monster is dead
                Battling = False
                MsgBox "You have defeated the enemy"
                MsgBox "You have gained " & IntExp & " Experience" & vbCrLf & vbCrLf & _
                    "You have gained " & intGold & " Gold"
                Stats.IntExp = Stats.IntExp + IntExp
                Stats.intGold = Stats.intGold + intGold
                Stats.intKills = Stats.intKills + 1
                Battling = False
                lblEnemyHP.Caption = 0
                Save
                cmdAttack.Enabled = False
                cmdUse.Enabled = False
                cmdRun.Enabled = False
                cmdOK.Enabled = True
                Unload frmBattle
            End If
        ElseIf PlayerRoll < MonsterRoll Then        'if the monster rolled higher than you
            Stats.intCurrentHP = Stats.intCurrentHP - MonsterDamage
            lblHP.Caption = Stats.intCurrentHP
            Delay (0.2)                             'calls delay timer
            cmdAttack.Enabled = True
            cmdAttack.Enabled = True
            cmdUse.Enabled = True
            cmdRun.Enabled = True

            If Stats.intCurrentHP <= 0 Then
                Battling = False                    'ends battle
                Dead                                'calls Dead sub
            End If
        End If
    

End Sub

Private Sub cmdOK_Click()
    Battling = False
    Unload frmBattle
    frmForest.Show
End Sub

'Run from the battle

Private Sub cmdRun_Click()
    Battling = False
    MsgBox "You run away scared"
    'Enter negative effects of running here
    Unload frmBattle
    frmForest.Show
End Sub

'calls the inventory

Private Sub cmdUse_Click()
    frmInventory.Show
End Sub


Private Sub Form_Load()
Dim strSQL As String, strLevel As String, strSQL1 As String, strHP As String, strName As String
    frmBattle.Left = (frmForest.Left)
    frmBattle.Top = frmForest.Top
    imgSallos.Visible = False
    imgSnake.Visible = False
    imgBarbados.Visible = False
    imgCarc.Visible = False
    imgIpos.Visible = False
    
'This determines the randomization of which Monsters you fight
'and then puts their picture visible with their stats

Dim Random As String
Dim G As Integer
Randomize
    Random = Int(Rnd * 16)  'Generates random number between 0 and 15
        If Random = 0 Then
            Random = 1      'if random number is 0 then make it 1
        End If
        If Random = 15 Then 'if random number is 15 then you get a random amount of gold
            G = Stats.intLevel * Int(Rnd * 25)
            MsgBox "You find " & G & "Gold"
            Stats.intGold = Stats.intBaseHP + G
            Save
            Exit Sub
        End If
        strLevel = Stats.intLevel
        strSQL1 = "SELECT * FROM Monsters WHERE Number = '" & Random & "'"
        datBattle.RecordSource = strSQL1
        datBattle.Refresh
            strName = datBattle.Recordset("Name")
            IntExp = (datBattle.Recordset("Exp") + (strLevel * Int(Rnd * 10)))
            intGold = (datBattle.Recordset("Gold") + (strLevel))
            MsgBox "You search around for a fight"
            MsgBox "You run into a " & strName
Dim intGroup
'loads and paints monster from database
'the numbers after "paintPicture" are coords and size

        intGroup = datBattle.Recordset("Number")
            If intGroup = 1 Then
                frmBattle.PaintPicture imgBarbados, 118, 8
            ElseIf intGroup = 2 Then
                frmBattle.PaintPicture imgCarc, 118, 8
            ElseIf intGroup = 3 Then
                frmBattle.PaintPicture imgIpos, 118, 8
            ElseIf intGroup = 4 Then
                frmBattle.PaintPicture imgSnake, 118, 8, 159, 181
            ElseIf intGroup = 5 Then
                frmBattle.PaintPicture imgSallos, 118, 8
            ElseIf intGroup = 6 Then
                frmBattle.PaintPicture imgTree, 118, 8, 159, 181
            ElseIf intGroup = 7 Then
                frmBattle.PaintPicture imgDwarf, 118, 8, 159, 181
            ElseIf intGroup = 8 Then
                frmBattle.PaintPicture imgOrch, 118, 8, 159, 181
            ElseIf intGroup = 9 Then
                frmBattle.PaintPicture imgLizMan, 118, 8, 159, 181
            ElseIf intGroup = 10 Then
                frmBattle.PaintPicture imgSpider, 118, 8, 159, 181
            ElseIf intGroup = 11 Then
                frmBattle.PaintPicture imgWizard, 118, 8, 159, 181
            ElseIf intGroup = 12 Then
                frmBattle.PaintPicture imgReaper, 118, 8, 159, 181
            ElseIf intGroup = 13 Then
                frmBattle.PaintPicture imgMonster, 118, 8, 159, 181
            ElseIf intGroup = 14 Then
                frmBattle.PaintPicture imgWater, 118, 8, 159, 181
            End If
        MonsterLife = datBattle.Recordset("HP")             'sets monsters stats
        MonsterHitPoints = (MonsterLife * strLevel) / 2
        strHP = Stats.intCurrentHP
        lblHP.Caption = strHP
        lblEnemyHP.Caption = MonsterHitPoints
        lblName.Caption = datBattle.Recordset("Name")
        Battling = True
        cmdOK.Enabled = False
End Sub

'Function for rolling dice
'if the player ties or has a higher roll, he wins

Function RollDice(X As Integer, Y As Integer)
    Dim num As Integer
    Dim roll As Integer
    Randomize                   'sets to random
    For roll = 1 To 10
        num = Int(Rnd * 9)      'creates random number
        frmBattle.PaintPicture imgDice(num), X, Y   'paints correcte dice
        Delay (0.1)             'calls delay function
    Next roll                   'moves to next roll
    RollDice = num
End Function


