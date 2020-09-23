Attribute VB_Name = "modClasses"
Option Explicit

Dim strAppend As String

Public Enum enumClass_template
    Scout
    Ranger
    HeavyWeapon
    Spy
    Soldier
    Support
End Enum

Public Function fstrReturnDescription(ClassType As enumClass_template) As String

    strAppend = ""
    
    Select Case ClassType
    
        Case Scout
            fstrAppend "The Scout"
            fstrAppend ""
            fstrAppend "The Scout is a soldier specialised in pre-recognisition of enemies. It prefers light-weight mechs with long range sensors instead of weaponary"
        
        Case Ranger
            fstrAppend "The Ranger"
            fstrAppend ""
            fstrAppend "The Ranger is a sniper soldier. It's specialised in small mechs with long range laser weapons."
        
        Case HeavyWeapon
            fstrAppend "The Heavy Weapon specialist"
            fstrAppend ""
            fstrAppend "The Heavy weapons soldier is specialised in demolitions and making sure it does heavy damage. It prefers heavy armored mechs with medium and short attack lasers and rocket barrages."
            
        Case Spy
            fstrAppend "The Spy"
            fstrAppend ""
            fstrAppend "The spy is accustomed to having mechs with cloaking ability. It will try to hide instead of fight and attack them from the back when they aint expecting it."
            
        Case Soldier
            fstrAppend "The Soldier"
            fstrAppend ""
            fstrAppend "The soldier is a standard mech pilot/warrior that does not have any special skills. It can pilot any mech but needs to learn more then the rest."
            
        Case Support
            fstrAppend "The Support units"
            fstrAppend ""
            fstrAppend "The support units dont attack most of the time; They can.. but they are mostly busy repairing or helping their fellow teammates."
    End Select
    
    fstrReturnDescription = strAppend
End Function

Private Function fstrAppend(strString As String)
    strAppend = strAppend & strString & vbCrLf
End Function
