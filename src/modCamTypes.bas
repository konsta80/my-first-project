Attribute VB_Name = "modCamTypes"
Option Explicit

' Общие типы и глобальное состояние каркаса.

Public Type Vec2
    X As Double
    Y As Double
End Type

' Замкнутая или открытая ломаная. X/Y — параллельные массивы, Cap — ёмкость.
Public Type Poly
    X() As Double
    Y() As Double
    N As Long
    Cap As Long
    Closed As Boolean
End Type

Public Type PolySet
    Items() As Poly
    Count As Long
End Type

Public Type Hole
    X As Double
    Y As Double
    Dia As Double
End Type

Public Type HoleSet
    Items() As Hole
    Count As Long
End Type

Public Type CamSettings
    ToolName As String
    ToolDiameter As Double
    ToolNumber As Long
    TipAngle As Double
    FeedXY As Double
    FeedZ As Double
    SpindleRPM As Long
    SafeZ As Double
    StartZ As Double
    StepDown As Double
    TotalDepth As Double
    StepoverPercent As Double
    Mode As String
    Side As String
    Climb As Boolean
    ProgramNumber As Long
    UseATC As Boolean
    Coolant As Boolean
    OutputPath As String
    DrawPreview As Boolean
    FlattenTol As Double
End Type

' Элементарное перемещение. Kind — одна из констант mm*.
Public Type MillMove
    Kind As Long
    X As Double
    Y As Double
    Z As Double
    F As Double
    R As Double
End Type

Public Const mmRapid As Long = 1
Public Const mmPlunge As Long = 2
Public Const mmCut As Long = 3
Public Const mmDrill As Long = 4

Public gCfg As CamSettings
Public gMoves() As MillMove
Public gMoveCount As Long
