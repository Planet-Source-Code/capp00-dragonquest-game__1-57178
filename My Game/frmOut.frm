VERSION 5.00
Begin VB.Form frmOut 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DragonQuest"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOut.frx":0000
   ScaleHeight     =   7170
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Go Back"
      Height          =   495
      Left            =   4920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click to show current inventory"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Inventory"
      Height          =   495
      Left            =   3240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click to show current inventory"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdStats 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Stats"
      Height          =   495
      Left            =   1800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click to show player stats"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Go Home"
      Height          =   495
      Left            =   6480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click to go back home"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image imgAlley 
      Height          =   615
      Left            =   1200
      ToolTipText     =   "Alley"
      Top             =   5520
      Width           =   735
   End
   Begin VB.Image imgItems 
      Height          =   4455
      Left            =   7680
      MouseIcon       =   "frmOut.frx":E1042
      MousePointer    =   99  'Custom
      ToolTipText     =   "Buy Items"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Image imgArmor 
      Height          =   1935
      Left            =   5640
      MouseIcon       =   "frmOut.frx":E134C
      MousePointer    =   99  'Custom
      ToolTipText     =   "Buy/Sell Armor"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Image imgWeapons 
      Height          =   2175
      Left            =   2040
      MouseIcon       =   "frmOut.frx":E1656
      MousePointer    =   99  'Custom
      ToolTipText     =   "Buy/Sell Weapons"
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "frmOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    Unload Me
    Unload Form1

    frmWhere.Show
End Sub

Private Sub cmdHome_Click()
    Save            'calls save
    Unload Form1    'unload from memory
    Unload Me
    Form1.Show      'go home
End Sub

Private Sub cmdStats_Click()
    frmStats.Show       'view stats page
End Sub

Private Sub Command1_Click()
    frmInventory.Show   'view inventory
End Sub

Private Sub Form_Load()
    frmOut.Left = (Screen.Width / 2 - frmOut.Width / 2)     'centers form
    frmOut.Top = (Screen.Height / 2 - frmOut.Height / 2)    'centers form
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form1.Show
End Sub

Private Sub imgAlley_Click()
    frmAlley.Show       'show alley
End Sub

Private Sub imgArmor_Click()
    frmArmor.Show       'shows armor shop
End Sub

Private Sub imgItems_Click()
    frmItems.Show       'shows item shop
End Sub

Private Sub imgWeapons_Click()
    frmWeapons.Show     'shows weapon shop
End Sub
