Attribute VB_Name = "Module1"
'=================================================================
'DragonQuest 2003
'Programmed and Designed by Capp
'=================================================================


'****************************
'This module is used to store variables
'and common functions/subs used by other parts
'of the game
'****************************



Option Explicit




'Player Stats and information
Type StatsType
    strName As String           'Players name
    intLevel As Integer         'Players Current level
    intGold As Integer          'Players Gold (money) count
    intBaseHP As Integer        'Base HitPoints
    intMaxHP As Integer         'Max HitPoints
    intCurrentHP As Integer     'Current Hitpoints
    intBaseStr As Integer       'Base Strength
    intStr As Integer           'Current Strength
    strWeapon As String         'Players current weapon
    strArmor As String          'Players current armor
    intKills As Integer         'Players kill count
    IntExp As Integer           'Players Experience count
End Type

'In-Game variables
Public Stats As StatsType
Public intcheatcount As Integer     'Holds cheat code
Public Battling As Boolean          'Determines if a battle is in progress
Public MonsterHitPoints As Integer  'Self Explanitory
Public Dragon As Boolean            'Determines if you are battling the dragon


'This function checks to see if you can afford weapons/armor/items

Function CheckGold(Cost As Integer) As Boolean
    If Stats.intGold < Cost Then
        MsgBox "You don't have enough gold"
        CheckGold = False
        Exit Function
    Else
        CheckGold = True
    End If
End Function

'This Sub saves the player information into the database

Public Sub Save()
    Dim strName As String, intStrength As Integer, strSearch As String, intDamage As Integer
    Dim strSQL As String, intBaseStr As Integer, intBaseHP As Integer
    Dim intGold As Integer, strWeapon As String, strArmor As String, IntExp As Integer
    Dim intHP As Integer, intStr As Integer, intLevel As Integer, intKills As Integer
    Load Form1 'uses DAO control from Form1
    intGold = Stats.intGold
    strWeapon = Stats.strWeapon
    strArmor = Stats.strArmor
    intHP = Stats.intMaxHP
    intStr = Stats.intStr
    intLevel = Stats.intLevel
    IntExp = Stats.IntExp
    intKills = Stats.intKills
    intBaseStr = Stats.intBaseStr
    intBaseHP = Stats.intBaseHP
    strName = Stats.strName
    strSQL = "SELECT * FROM Players WHERE Name = '" & strName & "'"
    Form1.datHome.RecordSource = strSQL
    Form1.datHome.Refresh
        With Form1.datHome
            .Recordset.Edit
            .Recordset("Gold") = intGold
            .Recordset("Weapon") = strWeapon
            .Recordset("Armor") = strArmor
            .Recordset("HP") = intHP
            .Recordset("BaseHP") = intBaseHP
            .Recordset("Str") = intStr
            .Recordset("BaseStr") = intBaseStr
            .Recordset("Level") = intLevel
            .Recordset("Exp") = IntExp
            .Recordset("Kills") = intKills
            .Recordset.Update
        End With
End Sub

'This sub determines if you have died or not during battle

Public Sub Dead()
' Notifies you when you are low on life
    If Stats.intCurrentHP > 0 And Stats.intCurrentHP < 8 Then
        MsgBox "You are almost dead"
' Notifies you when you are dead
    ElseIf Stats.intCurrentHP <= 0 Then
        MsgBox "You have died" & vbCrLf & vbCrLf & "Your gold has been halved"
        Stats.intGold = Stats.intGold / 2
        Save 'Calls Save sub
        End
    End If
End Sub

'this is a timer sub used during the dice roll to simulate a roll

Public Sub Delay(amount As Single)
    Dim StartTime As Single
    Dim CurrentTime As Single
    StartTime = Timer
    Do
        CurrentTime = Timer
        DoEvents
    Loop While CurrentTime < StartTime + amount
End Sub

'Modifies how much damage you can due based on level and strength

Public Sub Modifier()
    Dim strSQL As String, strName As String, intStr As Integer
    Dim Damage As Integer, Strength As Integer
    intStr = Stats.intStr
    Load Form1
    strName = Stats.strWeapon
    If strName = "HANDS" Then
        Damage = 0
        Stats.intBaseStr = Stats.intStr + Damage
        Save
        Exit Sub
    ElseIf strName <> "HANDS" Then
    strSQL = "SELECT * FROM Weapons WHERE Name = '" & strName & "'"
    Form1.datWeapon.RecordSource = strSQL
    Form1.datWeapon.Refresh
        Damage = Form1.datWeapon.Recordset("Damage")
        Stats.intBaseStr = Stats.intStr + Damage
    Save
    End If
End Sub

'modifies how much strength you have based on armor and life you have

Public Sub ArmorMod()
    Dim strSQL As String, strName As String, intHP As String
    Dim Protection As Integer
    intHP = Stats.intMaxHP
    Load Form1
    strName = Stats.strArmor
        If strName = "Naked" Then
            Protection = "0"
            Stats.intBaseHP = intHP
            Save
            Exit Sub
        ElseIf strName <> "Naked" Then
    strSQL = "SELECT * FROM Armor WHERE Name = '" & strName & "'"
    Form1.datArmor.RecordSource = strSQL
    Form1.datArmor.Refresh
        Protection = Form1.datArmor.Recordset("Protection")
        Stats.intBaseHP = Stats.intMaxHP + Protection
            If Stats.intCurrentHP > Stats.intBaseHP Then
                Stats.intCurrentHP = Stats.intBaseHP
            End If
    Save
    End If
End Sub


'What happens when you beat the dragon

Public Sub WIN()
    MsgBox ("Congratulations!!" & vbCrLf & vbCrLf & "You have defeated the Dragon")
    'Enter code here for hall-of-fame
    'Do what you want with this one
    
End Sub
