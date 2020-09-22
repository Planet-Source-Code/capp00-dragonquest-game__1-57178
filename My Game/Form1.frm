VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DragonQuest"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9585
   HelpContextID   =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   StartUpPosition =   3  'Windows Default
   Begin VB.Data datArmor 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "DragonQuest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Armor"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datWeapon 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "DragonQuest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Weapons"
      Top             =   600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datHome 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "DragonQuest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLog 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "LOG OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      MouseIcon       =   "Form1.frx":E134C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Click to save and return to character selection screen"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "INVENTORY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      MouseIcon       =   "Form1.frx":E1656
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Click to view current inventory"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Shape shpStorage 
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MouseIcon       =   "Form1.frx":E1960
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Click to Save and Exit"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6720
      MouseIcon       =   "Form1.frx":E1C6A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":E1F74
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image imgRoom 
      Height          =   4575
      Left            =   8280
      MouseIcon       =   "Form1.frx":E227E
      MousePointer    =   99  'Custom
      ToolTipText     =   "Go to bedroom to rest"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image imgOut 
      Height          =   4695
      Left            =   360
      MouseIcon       =   "Form1.frx":E2588
      MousePointer    =   99  'Custom
      ToolTipText     =   "Leave the house"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Welcome Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "STATS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MouseIcon       =   "Form1.frx":E2892
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Click to view stats"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape shpStats 
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Load Cheesy interface
'This is your home
'Where you go to rest at the end of the day


'This sub handles cheats that you can access at home by pressing <Enter>

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim strCheat As String
    If KeyAscii = vbKeyReturn Then
    If intcheatcount > 5 Then                       'allows only 5 cheats
        MsgBox "No More cheats for today"
        Exit Sub
    End If
    strCheat = InputBox("Enter Cheat", "cheat")
    If strCheat = "blingbling" Then                 'Cheat for more gold
        Stats.intGold = Stats.intGold + 10
        intcheatcount = intcheatcount + 1
    ElseIf strCheat = "muscleman" Then              'Cheat for increased Strength
        Stats.intStr = Stats.intStr + 1
        intcheatcount = intcheatcount + 1
    ElseIf strCheat = "power" Then                  'Cheat for increased hitpoints
        Stats.intMaxHP = Stats.intMaxHP + 1
        intcheatcount = intcheatcount + 1
    ElseIf strCheat = "gimmemore" Then              'Resets cheatcounter to 0
        intcheatcount = 0
    End If
End If

End Sub

Private Sub Form_Load()
    Form1.Left = (Screen.Width / 2 - Form1.Width / 2)   'centers form on screen
    Form1.Top = (Screen.Height / 2 - Form1.Height / 2)  'centers form on screen
End Sub

'This keeps the form at the appropriate size for the picture

Private Sub Form_Resize()
On Error Resume Next
    If Form1.Width < 9690 Then
        Form1.Width = 9690
    End If
    If Form1.Height < 7560 Then
        Form1.Height = 7560
    End If
    On Error GoTo 0
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.ToolTipText = Time       'Shows time on clock's tooltip
End Sub


'Click the left door to leave the house

Private Sub imgOut_Click()
    Dim intresults As Integer
    intresults = MsgBox("Do you want to leave the house", vbYesNoCancel, "Leave?")
    If intresults = vbYes Then
        Unload Me
        frmWhere.Show
    Else
        Exit Sub
    End If
End Sub

'Click the right door to enter your bedroom and rest (refill's HP)

Private Sub imgRoom_Click()
Dim intresults As Integer
    intresults = MsgBox("Do you want to go to bed?", vbYesNoCancel, "Bed?")
    If intresults = vbYes Then
        If Stats.intCurrentHP < Stats.intBaseHP Then
            MsgBox "Your HP is refilled"
            Stats.intCurrentHP = Stats.intBaseHP
        ElseIf Stats.intCurrentHP = Stats.intBaseHP Then
            MsgBox "You don't need to rest"
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpStats.BorderColor = &H80000005   'imitates button click
    Label2.ForeColor = &H80000005
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpStats.BorderColor = &H80000008   'imitates button click
    Label2.ForeColor = &H80000008
    frmStats.Show

End Sub


Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = &H80000005
    Shape2.BorderColor = &H80000005
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = &H80000008
    Shape2.BorderColor = &H80000008
    Save    'calls save command
End Sub


Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = &H80000005
    Shape1.BorderColor = &H80000005

End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = &H80000008
    Shape1.BorderColor = &H80000008
    Save
    End
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpStorage.BorderColor = &H80000005
    Label5.ForeColor = &H80000005

End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpStorage.BorderColor = &H80000008
    Label5.ForeColor = &H80000008
    frmInventory.Show

End Sub

Private Sub lblLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLog.ForeColor = &H80000005
    Shape3.BorderColor = &H80000005

End Sub

Private Sub lblLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLog.ForeColor = &H80000008
    Shape3.BorderColor = &H80000008
    Save
    Unload Me
    frmCharacters.Show

End Sub
