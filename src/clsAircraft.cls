VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAircraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTailNumber As String
Private pModel As String
Private pEngineType As String
Private pMaxFuelCapacityLt As Double
Private pMaxHopperCapacityGal As Double

' --- TAIL NUMBER ---
Public Property Let TailNumber(ByVal Value As String)
    pTailNumber = UCase(Trim(Value))
End Property

Public Property Get TailNumber() As String
    TailNumber = pTailNumber
End Property

' --- MODEL ---
Public Property Let Model(ByVal Value As String)
    pModel = UCase(Trim(Value))
End Property

Public Property Get Model() As String
    Model = pModel
End Property

' --- ENGINE TYPE ---
Public Property Let EngineType(ByVal Value As String)
    pEngineType = UCase(Trim(Value))
End Property

Public Property Get EngineType() As String
    EngineType = pEngineType
End Property

' --- MAX FUEL CAPACITY (Liters) ---
Public Property Let MaxFuelCapacityLt(ByVal Value As Double)
    pMaxFuelCapacityLt = Value
End Property

Public Property Get MaxFuelCapacityLt() As Double
    MaxFuelCapacityLt = pMaxFuelCapacityLt
End Property

' --- MAX HOPPER CAPACITY (Gallons) ---
Public Property Let MaxHopperCapacityGal(ByVal Value As Double)
    pMaxHopperCapacityGal = Value
End Property

Public Property Get MaxHopperCapacityGal() As Double
    MaxHopperCapacityGal = pMaxHopperCapacityGal
End Property
