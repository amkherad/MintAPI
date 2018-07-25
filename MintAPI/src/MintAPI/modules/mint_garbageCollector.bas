Attribute VB_Name = "mint_garbageCollector"
Option Explicit

Public Type GarbageCollectorArguments
    Break As Boolean
    Force As Boolean
    IsEndOfTheApplication As Boolean
    IsIdleTime As Boolean
    TicksPerModule As Long
End Type

Public Sub mint_garbagec_BeginCollector(Args As GarbageCollectorArguments)
    If Args.Break Then Exit Sub
    
End Sub
