VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFuelLog 
   Caption         =   "FUEL LOG"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4005
   OleObjectBlob   =   "frmFuelLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFuelLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub frmFuelLog_Initialize()
    cmbTailNumber.List = Array("N-123DA", "EC-ALPHA", "EC-BRAVO", "EC-LNG")
    
    cmbTailNumber.Enabled = False: cmbTailNumber.BackColor = RGB(245, 245, 245)
    txtBase.Enabled = False: txtBase.BackColor = RGB(245, 245, 245)
    txtDispensedFuel.Enabled = False: txtDispensedFuel.BackColor = RGB(245, 245, 245)
    txtRemainingFuel.Enabled = False: txtRemainingFuel.BackColor = RGB(245, 245, 245)
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub txtPIN_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    ' El KeyCode 13 es la tecla ENTER (la que envían los escáneres al final)
    If KeyCode = 13 Then
        
        ' Anulamos el sonido de "Ding" de Windows al pulsar Enter en un TextBox
        KeyCode = 0
        
        Dim operatorInfo As String
        
        ' 1. Mostramos estado de carga
        lblOperatorName.Caption = "VERIFYING..."
        lblOperatorName.ForeColor = vbBlack
        DoEvents ' Fuerza a Excel a repintar la pantalla al instante
        
        ' 2. Disparamos la consulta a la API
        operatorInfo = modApi.GetOperatorByPIN(txtPIN.Text)
        
        ' 3. Evaluamos la respuesta del servidor
        If operatorInfo = "NOT FOUND" Then
            lblOperatorName.Caption = "NOT FOUND"
            lblOperatorName.ForeColor = vbRed
            txtPIN.Text = "" ' Borramos para que vuelva a intentarlo
            
        ElseIf operatorInfo = "ERROR" Then
            lblOperatorName.Caption = "SERVER ERROR"
            lblOperatorName.ForeColor = vbRed
            
        Else
            ' ÉXITO: Mostramos el nombre en verde
            lblOperatorName.Caption = operatorInfo
            lblOperatorName.ForeColor = RGB(0, 153, 0)
            
            ' TÁCTICO: The server verified the PIN. NOW we unlock the form.
            cmbTailNumber.Enabled = True: cmbTailNumber.BackColor = &H80000005
            txtBase.Enabled = True: txtBase.BackColor = &H80000005
            txtDispensedFuel.Enabled = True: txtDispensedFuel.BackColor = &H80000005
            txtRemainingFuel.Enabled = True: txtRemainingFuel.BackColor = &H80000005
            
            ' Movemos el cursor automáticamente a la casilla de avión
            cmbTailNumber.SetFocus
        End If
        
    End If
End Sub

Private Sub btnSubmit_Click()
    Dim objAircraft As clsAircraft
    Dim objLog As clsFuelLog
    Dim remFuel As Double
    Dim dispFuel As Double
    Dim PersonName As String
    Dim Location As String
    Dim TailNum As String
    Dim Result As Boolean
    
    ' 1. FRONTEND VALIDATION (Solo si está vacío o no es número)
    
    If txtPIN.Value = "" Then
        MsgBox "Please type your PIN."
        Exit Sub
    End If
    
    If txtBase.Value = "" Then
        MsgBox "Please insert a Location/Base.", vbExclamation
        Exit Sub
    End If
    
    If txtRemainingFuel.Value = "" Or Not IsNumeric(txtRemainingFuel.Value) Then
        MsgBox "Please insert a valid numeric quantity for Remaining Fuel.", vbExclamation
        txtRemainingFuel.Value = ""
        Exit Sub
    End If
    
    If txtDispensedFuel.Value = "" Or Not IsNumeric(txtDispensedFuel.Value) Then
        MsgBox "Please insert a valid numeric quantity for Dispensed Fuel.", vbExclamation
        txtDispensedFuel.Value = ""
        Exit Sub
    End If
    
    If cmbTailNumber.Value = "" Then
        MsgBox "Please select an aircraft.", vbExclamation
        Exit Sub
    End If
    
    If lblOperatorName.Caption = "" Or InStr(1, lblOperatorName.Caption, "DENEGADO") > 0 Or InStr(1, lblOperatorName.Caption, "ERROR") > 0 Then
        MsgBox "SECURITY LOCK: Please scan a valid Operator PIN before submitting.", vbCritical
        txtPIN.SetFocus
        Exit Sub
    End If
    
    ' Ya es seguro convertir a número
    remFuel = CDbl(txtRemainingFuel.Value)
    dispFuel = CDbl(txtDispensedFuel.Value)
    
    ' 2. FETCH AIRCRAFT DATA
    Set objAircraft = GetAircraftData(cmbTailNumber.Value)
    
If objAircraft Is Nothing Then
        MsgBox "Aircraft not found in Database.", vbCritical
        Exit Sub
    End If
    
    ' 3. INSTANTIATE FUEL LOG & ASSIGN AIRCRAFT
    Set objLog = New clsFuelLog
    Set objLog.Aircraft = objAircraft
    
    ' 4. INPUT FUEL DATA
    objLog.RemainingFuelLt = CDbl(txtRemainingFuel.Value)
    objLog.DispensedFuelLt = CDbl(txtDispensedFuel.Value)
    TailNum = objAircraft.TailNumber

    ' 5. CHECK VALIDATION RESULT
    If objLog.DispensedFuelLt = 0 And CDbl(txtDispensedFuel.Value) > 0 Then
        Exit Sub
    End If
    
    ' 6. PREPARAR DATOS PARA ENVÍO
    ' Usamos On Error Resume Next por si el Label no tiene el formato esperado "("
    On Error Resume Next
    PersonName = Trim(Mid(lblOperatorName.Caption, 1, (InStr(1, lblOperatorName.Caption, "(") - 1)))
    If Err.Number <> 0 Then PersonName = lblOperatorName.Caption
    On Error GoTo 0
    
    Location = Trim(UCase(txtBase.Value))
    
    ' ---------------------------------------------------------
    ' 7. LÓGICA DE EDICIÓN (PATCH)
    ' ---------------------------------------------------------
    If Trim(UCase(Me.btnSubmit.Caption)) = "EDIT LOG" Then
        Result = modApi.UpdateFuelLog(IDFuelLog, TailNum, objLog.RemainingFuelLt, objLog.DispensedFuelLt, PersonName, Location)
        
        If Result = True Then
            MsgBox "ENTRY EDITED", vbInformation
            
            ' REPARACIÓN: IDFuelLog es Long, para "limpiarlo" se pone a 0
            IDFuelLog = 0
            
            ' CERRAMOS Y CORTAMOS: Esto evita el Error 91
            Unload Me
            Exit Sub
        End If
        Exit Sub ' Si falló el Update, no queremos que intente hacer un Post abajo
    End If
        
    ' ---------------------------------------------------------
    ' 8. LÓGICA DE NUEVO REGISTRO (POST)
    ' ---------------------------------------------------------
    Result = modApi.PostFuelLog(TailNum, objLog.RemainingFuelLt, objLog.DispensedFuelLt, PersonName, Location)
    
    If Result = True Then
        MsgBox "REFUEL LOGGED", vbInformation
        ' Limpiamos antes de cerrar para evitar que eventos zombis lean datos
        txtRemainingFuel.Value = ""
        txtBase.Value = ""
        txtDispensedFuel.Value = ""
        Unload Me
    Else
        MsgBox "Not logged. Error in database communication.", vbCritical
    End If

End Sub

