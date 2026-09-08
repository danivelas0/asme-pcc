Option Explicit

' ===========================================================================
' modNav - Navegacion del Motor de Calculo ASME PCC Rev. 5
' ===========================================================================
' Fuente versionada en texto. Se inyecta en templates/maestro_con_macros.xlsm
' mediante scripts/make_vba_seed.py. No editar dentro de Excel: el editor VBA
' no es la fuente de verdad y cualquier cambio ahi se pierde al resembrar.
'
' Las hojas se referencian SIEMPRE por .Name y NUNCA por CodeName. Motivo:
' openpyxl no asigna codeName a las hojas que crea (todas salvo las tres que
' vienen del maestro), y Excel se los inventa al abrir. Un CodeName escrito
' aqui apuntaria a una hoja distinta -o a ninguna- en el proximo build.
'
' Rev. 5: la navegacion es un ARBOL de cinco niveles (publicante > disciplina
' > codigo > standard > hoja), no un indice plano. Ya no hay dos clases de
' clave: la celda oculta guarda SIEMPRE el nombre de la hoja destino, se este
' bajando o subiendo. Subir es ir a la hoja PADRE, que el builder graba en la
' clave del boton VOLVER. Por eso desaparecieron CLAVE_VOLVER y
' VolverAlDashboard, y AbrirHoja se convirtio en IrAHoja: una sola rama que no
' crece cuando crece el arbol.
' ===========================================================================

Public Const HOJA_INICIO As String = "Dashboard"

' Celda del aviso de macros en el Dashboard, fusionada A..L por el builder.
' El builder la deja en rojo ("deshabilitadas"); si las macros corren,
' Workbook_Open la reescribe en verde. Es el unico indicador de que la
' navegacion esta viva, asi que su direccion vive aqui y en el builder: si
' se mueve, hay que moverla en los dos sitios.
Public Const CELDA_AVISO As String = "A4"

' La clave de destino de cada boton se guarda en celdas ocultas, fuera de
' todo layout: fila del boton, columna COL_CLAVE_BASE + columna del boton.
'
' Depende de la columna, y no solo de la fila, porque cada nivel pone tres
' tarjetas por banda: con una sola columna de claves las tres escribirian en
' la misma celda y solo sobreviviria la ultima.
'
' Cada tarjeta es clicable ENTERA: sus cuatro filas llevan hipervinculo y una
' clave propia, en filas distintas de la misma columna.
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
'
' Las nueve de la Seccion II (CAT_/IDX_/DB_SecII_*) SI estan aqui, y es
' deliberado: son hojas de datos, sin buscador ni formulas, pero el ingeniero
' tiene que poder abrirlas. Sin esto, «solo hojas de datos» habria significado
' «inalcanzables».
'
' Las quince NAV_* son los nodos interiores del arbol de navegacion. Van en
' PREORDEN, intercaladas con los destinos que cuelgan de cada una: es el orden
' en que el builder las deriva del arbol, y este es el UNICO punto del VBA que
' crece cuando crece el arbol.
'
' Los prefijos CAL / BUS / DAT son las tres bandas del Dashboard (motores de
' calculo, motores de busqueda, bases de datos). Cada banda tiene su PROPIA
' cascada, y por eso hay tres NAV_*_ASME y dos NAV_*_SEC_II: la rama que lleva
' a los buscadores de la Parte D no es la que lleva a las hojas de datos de
' las Partes A, B y C.
'
' DEBE coincidir, en contenido Y EN ORDEN, con NAVEGABLES de
' build_db_materiales.py. Lo comprueba test_dashboard.py::TestSincroniaPythonVba.
'
' La lista se arma concatenando, y NO con un Array( ... ) de una sola linea
' logica partida con guiones bajos. Motivo, ya pagado: VBA admite como maximo
' 25 continuaciones de linea, y con 31 hojas hacian falta 30. AddFromString
' falla con "Demasiadas continuaciones de linea", el modulo se queda VACIO, y
' el sintoma que se ve no es ese sino un "No se ha definido Sub o Function" al
' guardar -porque ThisWorkbook llama a rutinas que ya no existen- en un
' dialogo modal que cuelga la automatizacion. Concatenar no tiene tope: el
' arbol puede crecer sin volver a chocar con el limite, y cada hoja se queda
' en su propia linea, que es lo que permite diffear esta lista contra la de
' Python.
Public Function HojasNavegables() As Variant
    Dim s As String
    s = "NAV_CAL_ASME"
    s = s & "|NAV_CAL_REPARACION"
    s = s & "|NAV_CAL_PCC"
    s = s & "|NAV_CAL_PCC2"
    s = s & "|Parche_PCC2_Art212"
    s = s & "|NAV_BUS_ASME"
    s = s & "|NAV_BUS_PIPING"
    s = s & "|NAV_BUS_B31"
    s = s & "|NAV_BUS_B31_3"
    s = s & "|Buscar_B31_3"
    s = s & "|Buscar_Prop_B31_3"
    s = s & "|Buscar_B31_B1"
    s = s & "|Buscar_Ec_A2"
    s = s & "|Buscar_Ej_A3"
    s = s & "|NAV_BUS_PVESSELS"
    s = s & "|NAV_BUS_BPVC"
    s = s & "|NAV_BUS_SEC_II"
    s = s & "|Buscar_BPVC_IID"
    s = s & "|Buscar_BPVC_IID_B"
    s = s & "|Buscar_Su"
    s = s & "|Buscar_Sy"
    s = s & "|Buscar_Prop_IID"
    s = s & "|NAV_DAT_ASME"
    s = s & "|NAV_DAT_PVESSELS"
    s = s & "|NAV_DAT_BPVC"
    s = s & "|NAV_DAT_SEC_II"
    s = s & "|DB_SecII_A1"
    s = s & "|DB_SecII_A2"
    s = s & "|DB_SecII_B"
    s = s & "|DB_SecII_C"
    s = s & "|CAT_SecII"
    s = s & "|IDX_SecII_Tablas"
    s = s & "|DB_SecII_Notas"
    s = s & "|DB_SecII_Quimica"
    s = s & "|DB_SecII_Traccion"
    s = s & "|Instrucciones"
    HojasNavegables = Split(s, "|")
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


' Va de `origen` a `destino` dejando una sola pestana a la vista. Es la UNICA
' rutina de navegacion: sirve igual para bajar un nivel del arbol, para subir
' al padre y para saltar por la miga de pan, porque en los tres casos la clave
' del boton es el nombre de la hoja destino.
Public Sub IrAHoja(ByVal destino As String, ByVal origen As Object)
    Dim ws As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(destino)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "La hoja '" & destino & "' no existe en este libro." & vbCrLf & _
               "Reconstruya el libro con build_db_materiales.py.", _
               vbExclamation, "Motor de Calculo ASME PCC"
        Exit Sub
    End If

    ' Cortafuegos: aunque una celda de clave se corrompiese, esta rutina nunca
    ' puede destapar una base de datos. El Dashboard se admite ademas de las
    ' navegables porque es la raiz del arbol -no esta en la lista, es la unica
    ' hoja visible- y hay botones que suben hasta el.
    '
    ' Ya no se exige que el origen sea el Dashboard, y no es un guardarrail que
    ' se pierda: era redundante -EsNavegable ya impide abrir una DB_* o una
    ' MAP_*- y con el arbol el origen legitimo dejo de ser una sola hoja.
    If Not EsNavegable(destino) And _
       StrComp(destino, HOJA_INICIO, vbTextCompare) <> 0 Then
        MsgBox "La hoja '" & destino & "' no es un motor navegable." & vbCrLf & _
               "Las bases de datos son insumo auditado y no se abren desde aqui.", _
               vbExclamation, "Motor de Calculo ASME PCC"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    ws.Visible = xlSheetVisible
    ws.Activate
    ' El origen se oculta DESPUES de activar el destino, por el mismo motivo
    ' que en AplicarVisibilidad: Excel se niega a ocultar la ultima hoja
    ' visible del libro.
    If StrComp(origen.Name, destino, vbTextCompare) <> 0 Then
        origen.Visible = xlSheetHidden
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
