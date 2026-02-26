Attribute VB_Name = "modTests"
Option Explicit


Public Sub TestFuelLogLogic()
    Dim myAircraft As clsAircraft
    Dim myLog As clsFuelLog
    
    ' 1. Creamos el avión
    Set myAircraft = New clsAircraft
    myAircraft.TailNumber = "EC-ALPHA"
    myAircraft.MaxFuelCapacityLt = 1438
    
    ' 2. Creamos el registro de repostaje
    Set myLog = New clsFuelLog
    Set myLog.Aircraft = myAircraft ' Vinculamos el avión al registro
    
    ' 3. Simulamos la operación
    myLog.RemainingFuelLt = 200 ' Quedaban 200 litros del vuelo anterior
    
    ' Intentamos meter 1300 litros (200 + 1300 = 1500 -> Supera los 1438)
    ' Esto debe hacer saltar el MsgBox de alerta y bloquear la asignación.
    myLog.DispensedFuelLt = 1600
    
    ' Si metemos una cantidad válida, la acepta:
    ' myLog.DispensedFuelLt = 1000
    
    Debug.Print "Total on board: " & myLog.TotalFuelOnBoardLt
End Sub

    Set fRadar = Nothing
    
    On Error GoTo 0
End Sub

Sub AbrirFuelLogEditor()
    Load frmEditLog
    frmEditLog.Show
End Sub
