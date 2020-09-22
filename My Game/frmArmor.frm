VERSION 5.00
Begin VB.Form frmArmor 
   Caption         =   "DragonQuest"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6210
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstItems 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "frmArmor.frx":0000
      Left            =   120
      List            =   "frmArmor.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   5895
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   5895
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "&BUY"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "&Sell"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Leave"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Data datArmor 
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
      RecordSource    =   "Armor"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   1920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line2 
      X1              =   3960
      X2              =   5640
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   2520
      Picture         =   "frmArmor.frx":0004
      Top             =   3960
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ARMOR SHOP"
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
      TabIndex        =   13
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "PRICE"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "PROTECTION"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblAction 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Label Label4 
      Caption         =   "RESALE VALUE"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Current Gold:"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblGold 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Current Armor"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmArmor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***********************************************
'Welcome to the Armor Shop
'Here you can buy and sell your armor
'to increase your strength and
'protection in the game
'***********************************************


Option Explicit
    Dim strPrice As String * 15
    Dim strArmor As String

Private Sub cmdBuy_Click()
    Dim canAfford As Boolean
    Dim amount As Integer
    Dim intR As Integer
    intR = MsgBox("Do you want to buy this item?", vbYesNoCancel, "Buy?")
    If intR = vbYes Then
        If Stats.strArmor <> "Naked" Then   'checks to see if you have any armor yet
            MsgBox "You must sell your current armor first", vbOKOnly, "Sell?"
            Exit Sub
        End If
        canAfford = CheckGold(datArmor.Recordset("Price"))  'determines if you can afford item
        If canAfford Then   'if yes then buy the armor
            Stats.intGold = Stats.intGold - strPrice
            Stats.strArmor = strArmor
            Save
            lblGold.Caption = Stats.intGold
            lblWeapon.Caption = Stats.strArmor
            lblAction.Caption = "Thank You"
            ArmorMod        'Calls Armor modification
            Save            'calls Save
        Else                'if not, then you can't
            MsgBox "You don't have enough gold"
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    cmdBuy.Enabled = False
End Sub

Private Sub cmdOK_Click()
    Save
    Unload Me
End Sub

'Sub for selling your current armor

Private Sub cmdSell_Click()
    Dim intR As Integer
    Dim strSQL As String
    Dim strName As String
    strName = Stats.strArmor    'stores your current armor
    If strName = "Naked" Then   'checks to see if you have any armor to sell
        MsgBox "You do not have any armor to sell"
        Exit Sub
    Else
        intR = MsgBox("Are you sure you want to sell your " & Stats.strArmor, vbYesNoCancel, "Sell")
            If intR = vbYes Then    'Get information and save new info
                strSQL = "SELECT * FROM Armor WHERE Name = '" & strName & "'"
                datArmor.RecordSource = strSQL
                datArmor.Refresh
                Stats.intGold = Stats.intGold + datArmor.Recordset("Sell")
                Stats.strArmor = "Naked"
                lblGold.Caption = Stats.intGold
                lblWeapon.Caption = Stats.strArmor
                lblAction.Caption = "Thank You"
                Save
            Else
                Exit Sub
            End If
    End If
End Sub

'loads form and all information from database onto form

Private Sub Form_Load()
    Dim strSQL As String
    strSQL = "SELECT * FROM Armor"
    datArmor.RecordSource = strSQL
    datArmor.Refresh
        Do Until datArmor.Recordset.EOF
            lstItems.AddItem UCase(datArmor.Recordset("Name"))
            datArmor.Recordset.MoveNext
        Loop
    frmArmor.Top = frmOut.Top
    frmArmor.Left = frmOut.Left
    lblAction.Caption = "How may I help you?"
    lblGold.Caption = Stats.intGold
    lblWeapon.Caption = Stats.strArmor
    cmdBuy.Enabled = False
End Sub

'Gets item information from database when item name is clicked

Private Sub lstItems_Click()
    Dim strSQL As String
    Dim strProtect As String * 15
    Dim strResale As String * 15
    strArmor = lstItems.Text
    txtDesc.Text = ""
    strSQL = "SELECT * FROM Armor WHERE Name = '" & strArmor & "'"
    datArmor.RecordSource = strSQL
    datArmor.Refresh
    strPrice = datArmor.Recordset("Price")
    strProtect = datArmor.Recordset("Protection")
    strResale = datArmor.Recordset("Sell")
    txtDesc.Text = strPrice & strProtect & strResale
    lblAction.Caption = "How may I help you?"
    cmdBuy.Enabled = True
End Sub


