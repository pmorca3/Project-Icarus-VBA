VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditLog 
   Caption         =   "UserForm1"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8490
   OleObjectBlob   =   "frmEditLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With lstLogs
        .ColumnCount = 7 ' Necesitamos 6 para: ID, Tail, Base, Refuel, Remaining, EditedBy
        .ColumnWidths = "30 pt; 60 pt; 80 pt; 50 pt; 50 pt; 50 pt; 80 pt"
    End With
    Call btnRefresh_Click
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnEdit_Click()
    If lstLogs.ListIndex = -1 Then
        MsgBox "Please select a record."
        Exit Sub
    End If
    
    ' 1. Guardamos el ID global
    IDFuelLog = CLng(lstLogs.List(lstLogs.ListIndex, 0))
    
    ' 2. Cargamos el form (dispara Initialize, campos quedan bloqueados)
    Load frmFuelLog
    
    ' 3. Seteamos el caption del botón
    frmFuelLog.btnSubmit.Caption = "EDIT LOG"
    
    ' 4. Habilitamos temporalmente para poder asignar valores
    '    (mismos colores que cuando el PIN desbloquea)
    With frmFuelLog
        .cmbTailNumber.Enabled = True:  .cmbTailNumber.BackColor = &H80000005
        .txtBase.Enabled = True:        .txtBase.BackColor = &H80000005
        .txtDispensedFuel.Enabled = True: .txtDispensedFuel.BackColor = &H80000005
        .txtRemainingFuel.Enabled = True: .txtRemainingFuel.BackColor = &H80000005
        
        ' 5. Asignamos los valores (ahora sí los acepta)
        .cmbTailNumber.Value = lstLogs.List(lstLogs.ListIndex, 1)
        .txtBase.Value = lstLogs.List(lstLogs.ListIndex, 2)
        .txtDispensedFuel.Value = lstLogs.List(lstLogs.ListIndex, 3)
        .txtRemainingFuel.Value = lstLogs.List(lstLogs.ListIndex, 4)
        
        ' 6. Volvemos a bloquear — el PIN los desbloqueará de nuevo
        .cmbTailNumber.Enabled = False:  .cmbTailNumber.BackColor = RGB(245, 245, 245)
        .txtBase.Enabled = False:        .txtBase.BackColor = RGB(245, 245, 245)
        .txtDispensedFuel.Enabled = False: .txtDispensedFuel.BackColor = RGB(245, 245, 245)
        .txtRemainingFuel.Enabled = False: .txtRemainingFuel.BackColor = RGB(245, 245, 245)
    End With
    
    ' 7. Show FUERA del With — esto mata el Error 91
    frmFuelLog.Show
    
    ' 8. Al volver aquí el form ya cerró — refrescamos
    Call btnRefresh_Click
    
End Sub

Private Sub btnRefresh_Click()
    Dim nuevoJson As String
    
    ' 1. Vaciamos la lista actual
    Me.lstLogs.Clear
    
    ' 2. Llamamos a la función que acabamos de mejorar
    nuevoJson = modApi.GetFuelLogs()
    
    ' 3. Volvemos a llenar la lista
    If nuevoJson <> "" Then
        Call FillListBox(nuevoJson)
    End If
End Sub

Sub FillListBox(jsonText As String)
    Dim Json As Object
    Dim Item As Object
    
    ' 1. Verificación de seguridad: ¿Viene vacío?
    If jsonText = "" Or jsonText = "[]" Then
        lstLogs.Clear
        Exit Sub
    End If

    ' 2. Parseo con manejo de errores
    On Error Resume Next
    Set Json = JsonConverter.ParseJson(jsonText)
    On Error GoTo 0
    
    ' Si el objeto no se creó, salimos antes de que explote
    If Json Is Nothing Then Exit Sub
    
    lstLogs.Clear
    
    ' 3. Llenado con seguridad de objeto
    For Each Item In Json
        If Not Item Is Nothing Then ' Verificamos que el item exista
            With lstLogs
                .AddItem Item("id") & ""
                .List(.ListCount - 1, 1) = Item("tail_number") & ""
                .List(.ListCount - 1, 2) = Item("base") & ""
                .List(.ListCount - 1, 3) = Item("refuel_amount") & ""
                .List(.ListCount - 1, 4) = Item("remaining_fuel_amount") & ""
                .List(.ListCount - 1, 5) = Item("edited_by") & ""
                Dim rawTS As String
                Dim posT As Integer
                rawTS = Item("log_timestamp") & ""
                posT = InStr(1, rawTS, "T")
                If posT > 0 Then
                    rawTS = Left(rawTS, posT - 1) & " " & Mid(rawTS, posT + 1, 5)
                End If
            .List(.ListCount - 1, 6) = rawTS
            End With
        End If
    Next Item
    
    ' 4. Limpieza total de referencias
    Set Item = Nothing
    Set Json = Nothing
End Sub
