VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuelLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pAircraft As clsAircraft
Private pRemainingFuelLt As Double
Private pDispensedFuelLt As Double

' --- AIRCRAFT OBJECT ---
Public Property Set Aircraft(ByVal obj As clsAircraft)
    Set pAircraft = obj
End Property

Public Property Get Aircraft() As clsAircraft
    Set Aircraft = pAircraft
End Property

' --- REMAINING FUEL (Liters) ---
Public Property Let RemainingFuelLt(ByVal Value As Double)
    If Value < 0 Then
        MsgBox "Remaining fuel cannot be negative.", vbCritical, "Error"
        pRemainingFuelLt = 0
    Else
        pRemainingFuelLt = Value
    End If
End Property

Public Property Get RemainingFuelLt() As Double
    RemainingFuelLt = pRemainingFuelLt
End Property

' --- DISPENSED FUEL (Liters) ---
Public Property Let DispensedFuelLt(ByVal Value As Double)
    ' 1. Validar números negativos
    If Value < 0 Then
        MsgBox "Dispensed fuel cannot be negative.", vbCritical, "Error"
        Exit Property
    End If
    
    ' 2. Validar que exista un avión asignado antes de hacer cálculos
    If pAircraft Is Nothing Then
        MsgBox "Aircraft not assigned to this fuel log.", vbCritical, "Error"
        Exit Property
    End If
    
    ' 3. Lógica central: Remanente + Dispensado vs Capacidad Máxima
    If (Value + pRemainingFuelLt) > pAircraft.MaxFuelCapacityLt Then
        MsgBox "Capacity Exceeded! Aircraft " & pAircraft.TailNumber & _
               " can only hold " & pAircraft.MaxFuelCapacityLt & " liters total." & vbCrLf & _
               "Current remaining: " & pRemainingFuelLt & vbCrLf & _
               "Attempting to add: " & Value & vbCrLf & _
               "Total: " & (Value + pRemainingFuelLt), vbCritical, "Fuel Overflow Alert"
    Else
        pDispensedFuelLt = Value
    End If
End Property

Public Property Get DispensedFuelLt() As Double
    DispensedFuelLt = pDispensedFuelLt
End Property

' --- TOTAL FUEL ON BOARD (Read Only) ---
Public Property Get TotalFuelOnBoardLt() As Double
    TotalFuelOnBoardLt = pRemainingFuelLt + pDispensedFuelLt
End Property

