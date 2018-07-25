Attribute VB_Name = "modConstructor"
Option Explicit


Public Function MCI() As MCI
    Set MCI = Static_MCI
End Function
Public Function MCIDevices() As MCIDevices
    Set MCIDevices = Static_MCIDevices
End Function

Public Function MCIPlayer() As MCIPlayer
    Set MCIPlayer = Static_MCIPlayer
End Function
