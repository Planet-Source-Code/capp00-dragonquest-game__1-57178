VERSION 5.00
Begin VB.Form frmWeapons 
   Caption         =   "DragonQuest"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.Data datWeapons 
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
      RecordSource    =   "Weapons"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Leave"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "&Sell"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "&BUY"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
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
      TabIndex        =   1
      Top             =   2520
      Width           =   5895
   End
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
      ItemData        =   "frmWeapons.frx":0000
      Left            =   120
      List            =   "frmWeapons.frx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
   Begin VB.Line Line2 
      X1              =   3960
      X2              =   5640
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label7 
      Caption         =   "Current Weapon"
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblGold 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   1920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label5 
      Caption         =   "Current Gold:"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "RESALE VALUE"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
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
      TabIndex        =   8
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "DAMAGE"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "PRICE"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "WEAPON SHOP"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   2400
      Picture         =   "frmWeapons.frx":0004
      Top             =   3960
      Width           =   1140
   End
End
Attribute VB_Name = "frmWeapons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************
'This works the exact same as the armor shop
'*****************************************




Option Explicit
Dim strPrice As String * 15
Dim strWeapon As String

Private Sub cmdBuy_Click()
    Dim canAfford As Boolean
    Dim amount As Integer
    Dim intR As Integer
    intR = MsgBox("Do you want to buy this item?", vbYesNoCancel, "Buy?")
    If intR = vbYes Then
        If Stats.strWeapon <> "HANDS" Then
            MsgBox "You must sell your weapon first", vbOKOnly, "Sell?"
            Exit Sub
        End If
        canAfford = CheckGold(datWeapons.Recordset("Price"))
        If canAfford Then
            Stats.intGold = Stats.intGold - strPrice
            Stats.strWeapon = strWeapon
            lblGold.Caption = Stats.intGold
            lblWeapon.Caption = Stats.strWeapon
            lblAction.Caption = "Thank You"
            Modifier
            Save
            Unload Form1

        Else
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
    Unload Form1

    Unload Me
End Sub

Private Sub cmdSell_Click()
    Dim intR As Integer
    Dim strSQL As String
    Dim strName As String
    strName = Stats.strWeapon
    If strName = "Hands" Then
        MsgBox "You do not have any weapons to sell"
        Exit Sub
    Else
        intR = MsgBox("Are you sure you want to sell your " & Stats.strWeapon, vbYesNoCancel, "Sell")
        If intR = vbYes Then
            strSQL = "SELECT * FROM Weapons WHERE Name = '" & strName & "'"
            datWeapons.RecordSource = strSQL
            datWeapons.Refresh
            Stats.intGold = Stats.intGold + datWeapons.Recordset("Sell")
            Stats.strWeapon = "HANDS"
            lblGold.Caption = Stats.intGold
            lblWeapon.Caption = Stats.strWeapon
            lblAction.Caption = "Thank You"
            Modifier
            Save
            Unload Form1

        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    strSQL = "SELECT * FROM Weapons"
    datWeapons.RecordSource = strSQL
    datWeapons.Refresh
    Do Until datWeapons.Recordset.EOF
        lstItems.AddItem UCase(datWeapons.Recordset("Name"))
        datWeapons.Recordset.MoveNext
    Loop
    frmWeapons.Top = frmOut.Top
    frmWeapons.Left = frmOut.Left
    lblAction.Caption = "How may I help you?"
    lblGold.Caption = Stats.intGold
    lblWeapon.Caption = Stats.strWeapon
    cmdBuy.Enabled = False
End Sub

Private Sub Label6_Click()

End Sub

Private Sub lstItems_Click()
    Dim strSQL As String
    Dim strDamage As String * 15
    Dim strResale As String * 15
    strWeapon = lstItems.Text
    txtDesc.Text = ""
    strSQL = "SELECT * FROM Weapons WHERE Name = '" & strWeapon & "'"
    datWeapons.RecordSource = strSQL
    datWeapons.Refresh
    strPrice = datWeapons.Recordset("Price")
    strDamage = datWeapons.Recordset("Damage")
    strResale = datWeapons.Recordset("Sell")
    txtDesc.Text = strPrice & strDamage & strResale
    lblAction.Caption = "How may I help you?"
    cmdBuy.Enabled = True
End Sub
