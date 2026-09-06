Option Explicit

' ===========================================================================
' ThisWorkbook - Eventos de libro del Motor de Calculo ASME PCC Rev. 3
' ===========================================================================
' Fuente versionada en texto. make_vba_seed.py vuelca este codigo dentro del
' modulo de documento ThisWorkbook del maestro (no lo importa como clase
' nueva: ThisWorkbook no se puede sustituir, solo rellenar).
'
' Toda la logica vive en modNav. Aqui solo estan los enganches de evento.
' ===========================================================================

Private Sub Workbook_Open()
    AplicarVisibilidad
    MarcarMacrosActivas

    On Error Resume Next
    ThisWorkbook.Worksheets(HOJA_INICIO).Activate
    ThisWorkbook.Worksheets(HOJA_INICIO).Range("A1").Select
    On Error GoTo 0

    ' Las unicas modificaciones hechas hasta aqui son las nuestras: no hay
    ' trabajo del usuario que perder. Marcar el libro como guardado evita
    ' que Excel pida guardar al cerrar un libro que solo se abrio y miro.
    ThisWorkbook.Saved = True
End Sub


' Intercepta el clic sobre las celdas-boton del Dashboard y sobre el enlace
' "VOLVER AL DASHBOARD" de cada motor.
'
' La clave de destino NO viaja en el hipervinculo, sino en una celda oculta
' que modNav.CeldaClave localiza a partir de la posicion del boton. Dos
' razones: openpyxl no sabe crear formas ni controles con OnAction, y un
' hipervinculo que apunte a una hoja oculta es invalido dentro del archivo.
' El hipervinculo apunta siempre a Dashboard!A1 -destino inocuo- y solo sirve
' para disparar este evento.
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    Dim clave As String

    On Error GoTo Salir

    clave = Trim$(CStr(CeldaClave(Sh, Target.Range.Row, Target.Range.Column).Value))
    If Len(clave) = 0 Then Exit Sub

    If StrComp(clave, CLAVE_VOLVER, vbTextCompare) = 0 Then
        VolverAlDashboard
    ElseIf StrComp(Sh.Name, HOJA_INICIO, vbTextCompare) = 0 Then
        AbrirHoja clave
    End If

Salir:
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Se limpia ANTES de escribir a disco, de modo que el archivo guardado
    ' nunca conserve una hoja de trabajo destapada.
    RestaurarEstadoLimpio
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Solo se limpia si no hay cambios pendientes del usuario. Si los hay, se
    ' deja intacto: tocar el libro aqui volveria a ensuciarlo y Excel pediria
    ' guardar de todos modos, y forzar Saved = True descartaria en silencio
    ' los datos de entrada que el ingeniero acabase de teclear. Si el usuario
    ' opta por guardar, Workbook_BeforeSave hace la limpieza; si no, la hara
    ' Workbook_Open la proxima vez.
    If ThisWorkbook.Saved Then
        RestaurarEstadoLimpio
        ThisWorkbook.Saved = True
    End If
End Sub
