Attribute VB_Name = "Modulo_Panel"

' Módulo de gestión de la base de datos adaptado al nuevo modelo relacional
Option Explicit

'-------------------------------------------------------------------------------
' Función: GetNextID
' Obtiene el próximo identificador disponible para el campo ID_Registro.
' Recorre la columna ID_Registro de la tabla Base_Planing y devuelve el máximo
' valor más uno. Si la tabla está vacía devuelve 1.
Private Function GetNextID() As Long
    Dim tbl As ListObject
    Set tbl = Worksheets("Base_Planing").ListObjects("Base_Planing")
    On Error Resume Next
    GetNextID = Application.WorksheetFunction.Max(tbl.ListColumns("ID_Registro").DataBodyRange) + 1
    If GetNextID = 0 Then GetNextID = 1
End Function

'-------------------------------------------------------------------------------
' Procedimiento: NuevoRegistro
' Prepara el formulario estableciendo el siguiente ID disponible y limpiando
' los campos de entrada. La celda de ID se deja bloqueada para evitar
' modificaciones manuales.
Public Sub NuevoRegistro()
    Dim ws As Worksheet
    Set ws = Worksheets("Panel_de_Control")

    ' Asignar el siguiente ID
    ws.Range("B5").Value = GetNextID()

    ' Limpiar campos de entrada (excepto ID)
    ws.Range("B6:B14").ClearContents
    ws.Range("B15:D17").ClearContents ' Comentarios (celdas combinadas)
End Sub

'-------------------------------------------------------------------------------
' Procedimiento: GuardarRegistro
' Valida los campos obligatorios y guarda o actualiza el registro en la
' tabla Base_Planing. También establece automáticamente el valor del campo
' Origen como "Planing".
Public Sub GuardarRegistro()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim id As Variant
    Dim rowIndex As Long
    Dim celda As Range
    Dim mensaje As String

    Set ws = Worksheets("Panel_de_Control")
    Set tbl = Worksheets("Base_Planing").ListObjects("Base_Planing")

    id = ws.Range("B5").Value
    If id = "" Then
        MsgBox "El ID es obligatorio.", vbExclamation
        Exit Sub
    End If

    ' Validación de campos obligatorios
    If ws.Range("B6") = "" Or _
       ws.Range("B7") = "" Or _
       ws.Range("B8") = "" Or _
       ws.Range("B9") = "" Or _
       ws.Range("B13") = "" Then
        MsgBox "Complete los campos obligatorios: Categoría, ID_Jefatura, Encargado, Proyecto y Fecha.", vbExclamation
        Exit Sub
    End If

    ' Validar fecha
    If Not IsDate(ws.Range("B13").Value) Then
        MsgBox "La fecha no es válida. Utilice el formato DD-MM-AAAA.", vbExclamation
        Exit Sub
    End If

    ' Validar horas numéricas
    If ws.Range("B14").Value <> "" Then
        If Not IsNumeric(ws.Range("B14").Value) Then
            MsgBox "El campo Horas debe ser numérico.", vbExclamation
            Exit Sub
        End If
    End If

    ' Verificar si existe el ID en la tabla
    Set celda = tbl.ListColumns("ID_Registro").DataBodyRange.Find(What:=CLng(id), LookIn:=xlValues, LookAt:=xlWhole)
    If celda Is Nothing Then
        ' Nuevo registro
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add
        rowIndex = newRow.Index
        mensaje = "Registro agregado correctamente."
    Else
        ' Actualizar registro existente
        rowIndex = celda.Row - tbl.HeaderRowRange.Row
        mensaje = "Registro actualizado correctamente."
    End If

    ' Asignar valores a la fila correspondiente
    With tbl.DataBodyRange
        .Cells(rowIndex, tbl.ListColumns("ID_Registro").Index).Value = id
        .Cells(rowIndex, tbl.ListColumns("Categoria").Index).Value = ws.Range("B6").Value
        .Cells(rowIndex, tbl.ListColumns("ID_Jefatura").Index).Value = ws.Range("B7").Value
        .Cells(rowIndex, tbl.ListColumns("Encargado").Index).Value = ws.Range("B8").Value
        .Cells(rowIndex, tbl.ListColumns("Proyecto").Index).Value = ws.Range("B9").Value
        .Cells(rowIndex, tbl.ListColumns("Origen").Index).Value = "Planning"
        .Cells(rowIndex, tbl.ListColumns("Tarea_asignada").Index).Value = ws.Range("B12").Value
        .Cells(rowIndex, tbl.ListColumns("Fecha").Index).Value = ws.Range("B13").Value
        .Cells(rowIndex, tbl.ListColumns("Horas").Index).Value = ws.Range("B14").Value
        .Cells(rowIndex, tbl.ListColumns("Comentarios").Index).Value = ws.Range("B15").Value
    End With

    MsgBox mensaje, vbInformation
    ' Preparar formulario para el siguiente ingreso
    NuevoRegistro
End Sub

'-------------------------------------------------------------------------------
' Procedimiento: ConsultarRegistro
' Carga en el formulario los datos de un ID específico para su revisión o
' actualización. Si el ID no existe se muestra un mensaje.
Public Sub ConsultarRegistro()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim id As Variant
    Dim celda As Range
    Dim rowIndex As Long

    Set ws = Worksheets("Panel_de_Control")
    Set tbl = Worksheets("Base_Planing").ListObjects("Base_Planing")

    id = ws.Range("B5").Value
    If id = "" Then Exit Sub

    Set celda = tbl.ListColumns("ID_Registro").DataBodyRange.Find(What:=CLng(id), LookIn:=xlValues, LookAt:=xlWhole)
    If celda Is Nothing Then
        MsgBox "No existe un registro con el ID indicado.", vbInformation
        Exit Sub
    End If
    rowIndex = celda.Row - tbl.HeaderRowRange.Row
    With tbl.DataBodyRange
        ws.Range("B6").Value = .Cells(rowIndex, tbl.ListColumns("Categoria").Index).Value
        ws.Range("B7").Value = .Cells(rowIndex, tbl.ListColumns("ID_Jefatura").Index).Value
        ws.Range("B8").Value = .Cells(rowIndex, tbl.ListColumns("Encargado").Index).Value
        ws.Range("B9").Value = .Cells(rowIndex, tbl.ListColumns("Proyecto").Index).Value
        ws.Range("B12").Value = .Cells(rowIndex, tbl.ListColumns("Tarea_asignada").Index).Value
        ws.Range("B13").Value = .Cells(rowIndex, tbl.ListColumns("Fecha").Index).Value
        ws.Range("B14").Value = .Cells(rowIndex, tbl.ListColumns("Horas").Index).Value
        ws.Range("B15").Value = .Cells(rowIndex, tbl.ListColumns("Comentarios").Index).Value
    End With
    MsgBox "Registro cargado. Puede modificar los datos y presionar 'Guardar' para actualizar.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' Procedimiento: EliminarRegistro
' Elimina un registro de la tabla según el ID indicado. Solicita confirmación
' antes de eliminar definitivamente.
Public Sub EliminarRegistro()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim id As Variant
    Dim celda As Range
    Dim respuesta As VbMsgBoxResult

    Set ws = Worksheets("Panel_de_Control")
    Set tbl = Worksheets("Base_Planing").ListObjects("Base_Planing")

    id = ws.Range("B5").Value
    If id = "" Then Exit Sub

    Set celda = tbl.ListColumns("ID_Registro").DataBodyRange.Find(What:=CLng(id), LookIn:=xlValues, LookAt:=xlWhole)
    If celda Is Nothing Then
        MsgBox "No existe un registro con el ID indicado.", vbInformation
        Exit Sub
    End If
    respuesta = MsgBox("¿Desea eliminar el registro?", vbYesNo + vbQuestion, "Confirmación")
    If respuesta = vbYes Then
        tbl.ListRows(celda.Row - tbl.HeaderRowRange.Row).Delete
        MsgBox "Registro eliminado.", vbInformation
        NuevoRegistro
    End If
End Sub

'-------------------------------------------------------------------------------
' Procedimiento: LimpiarFormulario
' Limpia los campos de entrada manteniendo el ID y la fórmula de búsqueda del
' titular. Se utiliza cuando el usuario desea desechar cambios sin borrar el
' registro existente.
Public Sub LimpiarFormulario()
    Dim ws As Worksheet
    Set ws = Worksheets("Panel_de_Control")
    ' Mantener ID y fórmulas, limpiar datos de entrada
    ws.Range("B6:B14").ClearContents
    ws.Range("B15:D17").ClearContents
End Sub
