VERSION 5.00
Begin VB.Form frmInventory 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DragonQuest"
   ClientHeight    =   4935
   ClientLeft      =   2940
   ClientTop       =   2505
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.Data datUse 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "DragonQuest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Data datInv 
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
      RecordSource    =   "Inventory"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdUse 
      Caption         =   "&Use Item"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ListBox lstInventory 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "frmInventory.frx":0000
      Left            =   360
      List            =   "frmInventory.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "INVENTORY"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************
'This is your inventory
'You have items you can use in battle or to heal your self
'**************************************

Option Explicit
    Dim strType As String
    Dim strSQL As String
    Dim strItem As String
    Dim strSQL1 As String

'Determines which window to go back to for battle

Private Sub cmdOK_Click()
    If Battling = True And Dragon <> True Then
        frmBattle.Show
        Unload Me
    ElseIf Battling = True And Dragon = True Then
        frmDragon.Show
        Unload Me
    Else
        Unload Me
    End If
End Sub

'Sub for using items

Private Sub cmdUse_Click()
    Dim intAdded As Integer
    If Battling = False And strType = "weapon" Then             'checks to see if the item you are using is meant for battle
        MsgBox "You can't use that item when not in battle"
        Exit Sub
    End If
    strItem = lstInventory.Text
    If strItem = "POTION" Then                                  'If Item is "Potion"
        If Stats.intCurrentHP = Stats.intBaseHP Then
            MsgBox "You don't need to use this"
            Exit Sub
            Unload Me
        End If
            intAdded = (Stats.intMaxHP) / 3                     'adds 1/3 of life back
            Stats.intCurrentHP = Stats.intCurrentHP + intAdded
                If Stats.intCurrentHP > Stats.intBaseHP Then
                    Stats.intCurrentHP = Stats.intBaseHP
                End If
                MsgBox "You have been healed for " & intAdded & " hit points"
                    If Battling = True Then
                        frmBattle.lblHP = Stats.intCurrentHP
                    End If
    Use
    ElseIf strItem = "SLEEPING BAG" Then                        'If item is "Sleeping Bag"
        If Stats.intCurrentHP = Stats.intBaseHP Then
            MsgBox "You don't need to use this"
            Exit Sub
            Unload Me
        End If
            Stats.intCurrentHP = Stats.intBaseHP
            MsgBox "You have been completely healed"
            Use
    ElseIf strItem = "THROWING KNIFE" Then                      'If item is throwing knife
        MsgBox "you have hit the enemy for 10 damage"
        MonsterHitPoints = MonsterHitPoints - 10
        Use
    ElseIf strItem = "GRENADE" Then                             'if item is "grenade"
        MsgBox "you throw a grenade and damage the enemy for 10 damage"
        MonsterHitPoints = MonsterHitPoints - 10
        Use
    ElseIf strItem = "BIG GRENADE" Then                         'if item is "big grenade"
        MsgBox "you throw a big grenade and damage the enemy for 25 damage"
        MonsterHitPoints = MonsterHitPoints - 25
        Use
        Unload frmInventory
    ElseIf strItem = "BIG BANG" Then                            'if item is "big Bang"
            If Dragon = True Then
                MsgBox "You cannot use this against the dragon"
                Exit Sub
            End If
        MsgBox "you use the " & "Big Bang" & " and have killed the enemy"
        MonsterHitPoints = 0
        Battling = False
        frmBattle.Show
        frmBattle.cmdAttack.Caption = "OK"
        frmBattle.cmdRun.Enabled = False
        frmBattle.cmdUse.Enabled = False
        Use
    ElseIf strItem = "DAGGER" Then                              'if item is "dagger"
        MsgBox "you throw a dagger and hit the enemy for 15 damage"
        MonsterHitPoints = MonsterHitPoints - 15
        Use
    End If
        Unload frmInventory
        frmBattle.lblEnemyHP.Caption = MonsterHitPoints
    If Battling = True And Dragon <> True Then              'if in a battle then go back to battle
        Unload Me
        frmBattle.Show
    ElseIf Battling = True And Dragon = True Then           'if battling dragon
        Unload Me
        frmDragon.Show

    End If
    cmdUse.Enabled = False
End Sub

'gets players inventory
Private Sub Form_Load()
    frmInventory.Left = Me.Left
    frmInventory.Top = Me.Top
    Dim strName As String
    strName = Stats.strName
    strSQL = "SELECT * FROM Inventory Where Player = '" & strName & "'"
    datUse.RecordSource = strSQL
    datUse.Refresh
        Do Until datUse.Recordset.EOF
            lstInventory.AddItem UCase(datUse.Recordset("Item"))
            datUse.Recordset.MoveNext
        Loop
    cmdUse.Enabled = False
End Sub

'when you click the item in database list
Private Sub lstInventory_Click()
    txtDesc.Text = ""
    strItem = lstInventory.Text
    strSQL1 = "SELECT * FROM Items WHERE Name = '" & strItem & "'"
    datInv.RecordSource = strSQL1
    datInv.Refresh
    strType = datInv.Recordset("Type")
    txtDesc.Text = datInv.Recordset("Description")
    cmdUse.Enabled = True
End Sub

'remove item from database and inventory list when used
Public Sub Use()
    strSQL = "SELECT * FROM Inventory WHERE Item = '" & strItem & "'"
    datUse.RecordSource = strSQL
    datUse.Refresh
    datUse.Recordset.Delete

End Sub
