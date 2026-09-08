Option Explicit

' ===========================================================================
' modNav - Navegacion del Motor de Calculo ASME PCC Rev. 3
' ===========================================================================
' Fuente versionada en texto. Se inyecta en templates/maestro_con_macros.xlsm
' mediante scripts/make_vba_seed.py. No editar dentro de Excel: el editor VBA
' no es la fuente de verdad y cualquier cambio ahi se pierde al resembrar.
'
' Las hojas se referencian SIEMPRE por .Name y NUNCA por CodeName. Motivo:
' openpyxl no asigna codeName a las hojas que crea (todas salvo las tres que
' vienen del maestro), y Excel se los inventa al abrir. Un CodeName escrito
' aqui apuntaria a una hoja distinta -o a ninguna- en el proximo build.
' ===========================================================================

Public Const HOJA_INICIO As String = "Dashboard"
Public Const CLAVE_VOLVER As String = "VOLVER"

' Celda del aviso de macros en el Dashboard, fusionada A..L por el builder.
' El builder la deja en rojo ("deshabilitadas"); si las macros corren,
' Workbook_Open la reescribe en verde. Es el unico indicador de que la
' navegacion esta viva, asi que su direccion vive aqui y en el builder: si
' se mueve, hay que moverla en los dos sitios.
Public Const CELDA_AVISO As String = "A4"

' La clave de destino de cada boton se guarda en celdas ocultas, fuera de
' todo layout: fila del boton, columna COL_CLAVE_BASE + columna del boton.
'
' Depende de la columna, y no solo de la fila, porque el Dashboard pone tres
' tarjetas por banda: con una sola columna de claves las tres escribirian en
' la misma celda y solo sobreviviria la ultima.
'
' La base es 66 (BN) porque BM es la columna mas alta que usa cualquier hoja
' navegable. No se usa N, que seria lo natural, porque Parche_PCC2_Art212 ya
' ocupa K..P con sus listas de cascada.
Public Const COL_CLAVE_BASE As Long = 66

Private Const TXT_ACTIVAS As String = "MACROS ACTIVAS - navegacion habilitada"
Private Const TXT_INACTIVAS As String = _
    "MACROS DESHABILITADAS - habilitelas para navegar entre los motores"

' ---- fin de las declaraciones de modulo ----------------------------------
' VBA exige que TODA declaracion de nivel de modulo (Const, Dim, Type, Declare)
' preceda a la primera rutina. Una constante colocada mas abajo no da error en
' su linea: da "Variable not defined" en cada sitio que la usa, y el libro no
' llega a guardarse porque Excel compila al guardar. No insertar codigo aqui
' arriba.


' Celda que guarda la clave de destino del boton anclado en (fila, columna).
Public Function CeldaClave(ByVal hoja As Object, ByVal fila As Long, _
                           ByVal columna As Long) As Object
    Set CeldaClave = hoja.Cells(fila, COL_CLAVE_BASE + columna)
End Function


' Unica fuente de verdad de que hojas puede llegar a ver el usuario.
' Todo lo que no este en esta lista y no sea el Dashboard queda
' xlSheetVeryHidden: ni siquiera aparece en el menu Mostrar de Excel.
' Las hojas restantes (DB_*, MAP_*, Notas_Codigo, DB_Listas, Datos_Ref,
' _meta, _Curvas) son insumo auditado, no interfaz.
Public Function HojasNavegables() As Variant
    HojasNavegables = Array( _
        "Parche_PCC2_Art212", _
        "Buscar_B31_3", _
        "Buscar_BPVC_IID", _
        "Buscar_BPVC_IID_B", _
        "Buscar_Su", _
        "Buscar_Sy", _
        "Buscar_Prop_IID", _
        "Buscar_Prop_B31_3", _
        "Buscar_NoMetalicos", _
        "Buscar_Ec_A2", _
        "Buscar_Ej_A3", _
        "Instrucciones")
End Function


Private Function EsNavegable(ByVal nombre As String) As Boolean
    Dim v As Variant
    For Each v In HojasNavegables()
        If StrComp(CStr(v), nombre, vbTextCompare) = 0 Then
            EsNavegable = True
            Exit Function
        End If
    Next v
End Function


' Deja el libro en su estado de reposo: solo el Dashboard a la vista.
Public Sub AplicarVisibilidad()
    Dim ws As Object

    Application.ScreenUpdating = False

    ' El Dashboard se hace visible ANTES de ocultar nada: Excel se niega a
    ' ocultar la ultima hoja visible de un libro y abortaria el bucle.
    On Error Resume Next
    ThisWorkbook.Worksheets(HOJA_INICIO).Visible = xlSheetVisible
    On Error GoTo 0

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, HOJA_INICIO, vbTextCompare) <> 0 Then
            If EsNavegable(ws.Name) Then
                ws.Visible = xlSheetHidden
            Else
                ws.Visible = xlSheetVeryHidden
            End If
        End If
    Next ws

    Application.ScreenUpdating = True
End Sub


' Abre una hoja navegable y esconde el Dashboard, de forma que el usuario
' siga viendo una sola pestana.
Public Sub AbrirHoja(ByVal nombre As String)
    Dim ws As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "La hoja '" & nombre & "' no existe en este libro." & vbCrLf & _
               "Reconstruya el libro con build_db_materiales.py.", _
               vbExclamation, "Motor de Calculo ASME PCC"
        Exit Sub
    End If

    ' Cortafuegos: aunque las celdas de clave del Dashboard se corrompiesen,
    ' esta rutina nunca puede destapar una base de datos.
    If Not EsNavegable(nombre) Then
        MsgBox "La hoja '" & nombre & "' no es un motor navegable." & vbCrLf & _
               "Las bases de datos son insumo auditado y no se abren desde aqui.", _
               vbExclamation, "Motor de Calculo ASME PCC"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    ws.Visible = xlSheetVisible
    ws.Activate
    ' Se oculta el Dashboard DESPUES de activar el destino, por el mismo
    ' motivo que en AplicarVisibilidad.
    ThisWorkbook.Worksheets(HOJA_INICIO).Visible = xlSheetHidden
    Application.ScreenUpdating = True
End Sub


Public Sub VolverAlDashboard()
    Dim previa As Object

    Application.ScreenUpdating = False
    Set previa = ActiveSheet

    ThisWorkbook.Worksheets(HOJA_INICIO).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(HOJA_INICIO).Activate

    If StrComp(previa.Name, HOJA_INICIO, vbTextCompare) <> 0 Then
        previa.Visible = xlSheetHidden
    End If

    Application.ScreenUpdating = True
End Sub


Public Sub MarcarMacrosActivas()
    On Error Resume Next
    With ThisWorkbook.Worksheets(HOJA_INICIO).Range(CELDA_AVISO)
        .Value = TXT_ACTIVAS
        .Font.Color = RGB(0, 97, 0)          ' 006100
        .Interior.Color = RGB(198, 239, 206) ' C6EFCE
    End With
End Sub


Public Sub MarcarMacrosInactivas()
    On Error Resume Next
    With ThisWorkbook.Worksheets(HOJA_INICIO).Range(CELDA_AVISO)
        .Value = TXT_INACTIVAS
        .Font.Color = RGB(156, 0, 6)         ' 9C0006
        .Interior.Color = RGB(255, 199, 206) ' FFC7CE
    End With
End Sub


' Estado con el que el archivo debe quedar GRABADO en disco: solo el
' Dashboard visible y el aviso en rojo. Asi, quien lo abra con las macros
' bloqueadas no ve ninguna base de datos y sabe por que no navega.
Public Sub RestaurarEstadoLimpio()
    AplicarVisibilidad
    MarcarMacrosInactivas
    On Error Resume Next
    ThisWorkbook.Worksheets(HOJA_INICIO).Activate
End Sub
