Attribute VB_Name = "modApi"
Option Explicit

Private Const SUPABASE_URL As String = "URL"
Private Const SUPABASE_KEY As String = "KEY"

' ==========================================================================
' PURPOSE:      Fetch AC DATA
' ==========================================================================

Public Function GetAircraftData(ByVal TailNum As String) As clsAircraft
    
    ' --------------------------------------------------
    ' 1. CONNECTION, AUTHENTICATION AND REQUEST
    ' --------------------------------------------------
    
    Dim http As Object
    Dim url As String
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    'Clean plate number
    TailNum = UCase(Trim(TailNum))
    
    '[TU_URL_PROYECTO]/rest/v1/[NOMBRE_TABLA]?[NOMBRE_COLUMNA]=eq.[VALOR_BUSCADO]
    Let url = SUPABASE_URL & "/rest/v1/aircraft?tail_number=eq." & TailNum
    http.Open "GET", url, False
    
    'Set headers, el servidor espera la identificación
    http.setRequestHeader "apikey", SUPABASE_KEY
    http.setRequestHeader "Authorization", "Bearer " & SUPABASE_KEY
    
    'send request
    http.send
    
    'Check server response
    If http.Status <> 200 Then
        Debug.Print "SERVER ERROR: " & http.Status & " - " & http.responseText
        Set GetAircraftData = Nothing
        Exit Function
    End If

    ' --------------------------------------------------
    ' 2. JSON RESPONSE, ASSIGNATION
    ' --------------------------------------------------
    
    Dim jsonResponse As Object
    Set jsonResponse = JsonConverter.ParseJson(http.responseText)
    If jsonResponse.Count = 0 Then
        Debug.Print "Aircraft not found."
        Set GetAircraftData = Nothing
        Exit Function
    End If
    
    Dim Aircraft As clsAircraft
    Set Aircraft = New clsAircraft
    
    With Aircraft
        .TailNumber = jsonResponse(1)("tail_number")
        .Model = jsonResponse(1)("model")
        .EngineType = jsonResponse(1)("engine_type")
        .MaxFuelCapacityLt = jsonResponse(1)("max_fuel_capacity_liters")
        .MaxHopperCapacityGal = jsonResponse(1)("max_hopper_load_capacity_gal")
    End With
    
    Set GetAircraftData = Aircraft
End Function

' ==========================================================================
' PURPOSE:      Fetch fuel logs
' ==========================================================================
Public Function GetFuelLogs() As String
    Dim http As Object
    Dim url As String
    
    ' Usamos la URL normal
    url = SUPABASE_URL & "/rest/v1/fuel_logs?select=*&order=id.asc"
    
    ' CAMBIO CLAVE: Usamos ServerXMLHTTP para evitar el caché de Windows
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    With http
        .Open "GET", url, False
        
        ' Cabeceras de seguridad
        .setRequestHeader "apikey", SUPABASE_KEY
        .setRequestHeader "Authorization", "Bearer " & SUPABASE_KEY
        
        ' Cabeceras para forzar datos nuevos
        .setRequestHeader "Cache-Control", "no-cache"
        .setRequestHeader "Pragma", "no-cache"
        .setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
        
        .send

        If .Status = 200 Then
            GetFuelLogs = .responseText
        Else
            GetFuelLogs = ""
            Debug.Print "Error de Radar: " & .Status
        End If
    End With
    
    Set http = Nothing
End Function

' ==========================================================================
' PURPOSE:      Post fuel-log values into database from userform
' ==========================================================================

Public Function PostFuelLog(ByVal TailNum As String, ByVal RemainingFuelAmount As Double, ByVal ReFuelAmount As Double, ByVal PersonName As String, Location As String) As Boolean
    
    Dim url As String
    Dim http As Object
    Dim dictData As Scripting.Dictionary
    Dim Payload As String
    
    ' --------------------------------------------------
    ' 2.  BUILD THE PAYLOAD
    ' --------------------------------------------------
    Set dictData = New Scripting.Dictionary
    dictData.Add "log_timestamp", Format(VBA.Now, ("YYYY-MM-DD" & " " & "HH:NN:SS"))
    dictData.Add "tail_number", UCase(Trim(TailNum))
    dictData.Add "remaining_fuel_amount", RemainingFuelAmount
    dictData.Add "refuel_amount", ReFuelAmount
    dictData.Add "operator_name", PersonName
    dictData.Add "base", Location
    
    Payload = JsonConverter.ConvertToJson(dictData)
    
    ' --------------------------------------------------
    ' 2. CONNECTION, AUTHENTICATION AND REQUEST
    ' --------------------------------------------------
    
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    'Llamada
    Let url = SUPABASE_URL & "/rest/v1/fuel_logs" ' al insertar un avion nuevo no se filtra nada
    http.Open "POST", url, False
    
    'Autenticación
    http.setRequestHeader "apikey", SUPABASE_KEY
    http.setRequestHeader "Authorization", "Bearer " & SUPABASE_KEY
    http.setRequestHeader "Content-Type", "application/json"
    
    'send request
    http.send Payload
    
    'check server response
    If http.Status = 201 Then
        PostFuelLog = True
    Else
        MsgBox "ERROR CRITICO " & http.Status & ": " & http.responseText
        PostFuelLog = False
    End If

End Function

' ==========================================================================
' PURPOSE:      Edition of fuel-log values into database from userform
' ==========================================================================

Public Function UpdateFuelLog(ByVal IDFuelLog As Long, ByVal TailNum As String, ByVal RemainingFuelAmount As Double, ByVal ReFuelAmount As Double, ByVal PersonName As String, Location As String) As Boolean
    
    Dim url As String
    Dim http As Object
    Dim dictData As Scripting.Dictionary
    Dim EditedPayload As String
    
    ' --------------------------------------------------
    ' 2.  BUILD THE PAYLOAD
    ' --------------------------------------------------
Set dictData = New Scripting.Dictionary
        dictData.Add "tail_number", UCase(Trim(TailNum))
        dictData.Add "remaining_fuel_amount", Val(RemainingFuelAmount)
        dictData.Add "refuel_amount", Val(ReFuelAmount)
        dictData.Add "base", Location
        dictData.Add "edited_by", PersonName
    EditedPayload = JsonConverter.ConvertToJson(dictData)
    Debug.Print "REFUELED AMOUNT: " & EditedPayload
    
    ' --------------------------------------------------
    ' 2. CONNECTION, AUTHENTICATION AND REQUEST
    ' --------------------------------------------------
    
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    'Llamada
    Let url = SUPABASE_URL & "/rest/v1/fuel_logs?id=eq." & IDFuelLog ' we use ID number to link the information
    http.Open "PATCH", url, False
    
    'Autenticación
    http.setRequestHeader "apikey", SUPABASE_KEY
    http.setRequestHeader "Authorization", "Bearer " & SUPABASE_KEY
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Prefer", "return=representation"
    
Debug.Print "URL DE ENVÍO: " & url
Debug.Print "PAYLOAD: " & EditedPayload
    'send request
    http.send EditedPayload
    
    'check server response
    If http.Status >= 200 And http.Status < 300 Then
        UpdateFuelLog = True
    Else
        MsgBox "EDIT ERROR " & http.Status & ": " & http.responseText
        UpdateFuelLog = False
    End If

End Function

' ==========================================================================
' PURPOSE:      Fetch Operator Name by PIN (Barcode/RFID simulation)
' ==========================================================================
Public Function GetOperatorByPIN(ByVal strPIN As String) As String
    
    Dim http As MSXML2.XMLHTTP60
    Dim url As String
    Dim jsonResponse As Object
    
    ' --------------------------------------------------
    ' 1.  INITIALIZE THE CONNECTION & GET DATA
    ' --------------------------------------------------
    Set http = New MSXML2.XMLHTTP60
    Let url = SUPABASE_URL & "/rest/v1/personnel?pin_code=eq." & strPIN
    
    http.Open "GET", url, False
    http.setRequestHeader "apikey", SUPABASE_KEY
    http.setRequestHeader "Authorization", "Bearer " & SUPABASE_KEY
    
    http.send
    
    If http.Status <> 200 Then
       GetOperatorByPIN = "ERROR"
        Exit Function
    End If
    
    ' --------------------------------------------------
    ' 2.  OBTAIN THE NAME
    ' --------------------------------------------------
    
    Set jsonResponse = JsonConverter.ParseJson(http.responseText) 'JsonResponse is an object
    If jsonResponse.Count > 0 Then
        GetOperatorByPIN = jsonResponse(1)("full_name") & " (" & jsonResponse(1)("role") & ")"
    Else
        GetOperatorByPIN = "NOT FOUND"
    End If
    
End Function

Sub TestConnectionGetAircraftData()
    Dim testPlane As clsAircraft
    Set testPlane = GetAircraftData("EC-LNG")
    
    If Not testPlane Is Nothing Then
        Debug.Print "--------------------------"
        Debug.Print "HYDRATION SUCCESSFUL"
        Debug.Print "Tail Number : " & testPlane.TailNumber
        Debug.Print "Max Fuel    : " & testPlane.MaxFuelCapacityLt & " L"
        Debug.Print "--------------------------"
    Else
        Debug.Print "HYDRATION FAILED."
    End If
End Sub
