VERSION 5.00
Begin VB.Form frmItems 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DragonQuest"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstInventory 
      BackColor       =   &H80000000&
      Height          =   2010
      ItemData        =   "frmItems.frx":0000
      Left            =   3720
      List            =   "frmItems.frx":0002
      TabIndex        =   11
      Top             =   4080
      Width           =   2415
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
      Height          =   615
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2520
      Width           =   4455
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
      ItemData        =   "frmItems.frx":0004
      Left            =   120
      List            =   "frmItems.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   5895
   End
   Begin VB.TextBox txtPrice 
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
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "&BUY"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Leave"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Data datItems 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "DragonQuest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Items"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Current Items:"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   1920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   2400
      Picture         =   "frmItems.frx":0008
      Top             =   3840
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ITEM SHOP"
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
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "PRICE"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   855
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
      Left            =   960
      TabIndex        =   7
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "DESCRIPTION"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Current Gold:"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblGold 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************
'This works the exact same as the armor shop
'*****************************************



Option Explicit
Dim strPrice As String * 10
Dim strItems As String

Private Sub cmdBuy_Click()
    Dim canAfford As Boolean
    Dim amount As Integer, strItem As String, strPlayer As String
    Dim intR As Integer
    Dim strSQL As String
    strPlayer = Stats.strName
    strItem = strItems
    intR = MsgBox("Do you want to buy this item?", vbYesNoCancel, "Buy?")
    If intR = vbYes Then
        canAfford = CheckGold(datItems.Recordset("Price"))
        If canAfford Then
            strSQL = "SELECT * FROM Inventory WHERE Player = '" & Stats.strName & "'"
            datItems.RecordSource = strSQL
            datItems.Refresh
                 With datItems
                     .Recordset.AddNew
                     .Recordset("Player") = strPlayer
                     .Recordset("Item") = strItem
                     .Recordset.Update
                 End With

            Stats.intGold = Stats.intGold - strPrice
            Save
            lblGold.Caption = Stats.intGold
            lblAction.Caption = "Thank You"
            Inv
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
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String, strSQL2 As String
    strSQL = "SELECT * FROM Items"
    datItems.RecordSource = strSQL
    datItems.Refresh
    Do Until datItems.Recordset.EOF
        lstItems.AddItem UCase(datItems.Recordset("Name"))
        datItems.Recordset.MoveNext
    Loop
    Inv
    frmItems.Top = frmOut.Top
    frmItems.Left = frmOut.Left
    lblAction.Caption = "How may I help you?"
    lblGold.Caption = Stats.intGold
    cmdBuy.Enabled = False
End Sub

Private Sub lstItems_Click()
    Dim strSQL As String
    Dim strDescription As String
    strItems = lstItems.Text
    txtPrice.Text = ""
    txtDesc.Text = ""
    strSQL = "SELECT * FROM Items WHERE Name = '" & strItems & "'"
    datItems.RecordSource = strSQL
    datItems.Refresh
    strPrice = datItems.Recordset("Price")
    strDescription = datItems.Recordset("Description")
    txtPrice.Text = strPrice
    txtDesc.Text = strDescription
    lblAction.Caption = "How may I help you?"
    cmdBuy.Enabled = True

End Sub

Public Sub Inv()
    Dim strSQL2 As String
    lstInventory.Clear
    strSQL2 = "SELECT * From Inventory WHERE Player = '" & Stats.strName & "'"
    datItems.RecordSource = strSQL2
    datItems.Refresh
    Do Until datItems.Recordset.EOF
        lstInventory.AddItem UCase(datItems.Recordset("Item"))
        datItems.Recordset.MoveNext
    Loop
End Sub


