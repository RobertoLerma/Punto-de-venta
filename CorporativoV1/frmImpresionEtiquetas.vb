Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmImpresionEtiquetas
    Inherits System.Windows.Forms.Form

    Dim isLoad As Boolean = False

    Private components As System.ComponentModel.IContainer
    'Programa: Ventas Salida de Mercancia
    'Autor: Rosaura Torres López
    'Fecha de Creación: 27/Mayo/2003
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents optCodActual As System.Windows.Forms.RadioButton
    Public WithEvents optCodAnterior As System.Windows.Forms.RadioButton
    Public WithEvents fraOrdenamiento As System.Windows.Forms.GroupBox
    Public WithEvents tlbCode As AxTALBarCode.AxTALBarCd
    Public WithEvents txtOrdenCompra As System.Windows.Forms.TextBox
    Public WithEvents optOrdenCompra As System.Windows.Forms.RadioButton
    Public WithEvents optArticulos As System.Windows.Forms.RadioButton
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents txtDetalle As System.Windows.Forms.TextBox
    Public WithEvents msgEtiquetas As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public CommonOpen As System.Windows.Forms.OpenFileDialog
    Public CommonSave As System.Windows.Forms.SaveFileDialog
    Public CommonFont As System.Windows.Forms.FontDialog
    Public CommonColor As System.Windows.Forms.ColorDialog
    Public CommonPrint As System.Windows.Forms.PrintDialog
    Public WithEvents txtDesArticulo As System.Windows.Forms.Label
    Public WithEvents _Marco_141 As System.Windows.Forms.GroupBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Marco As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray

    Dim SignoPesos As String
    Dim ColInicial As Integer
    Dim RenInicial As Integer
    Dim Espacio As Integer
    Dim FueraChange As Boolean
    Dim mblnSalir As Boolean 'Se utiliza para saber cuando el usuraio a presionado Escape estando en el primer control del form.
    Dim mblnNuevo As Boolean ''''''
    Dim I As Integer 'Para manejar el For
    Dim rsLocal As ADODB.Recordset ''''''
    Dim intCodArticulo As Integer
    Dim Cambios As Integer
    Dim mintTotalEtiquetas As Integer

    ' Para Manejar el FlexGrid  de Detalle
    Const C_COLCODIGO As Integer = 0
    Const C_ColDESCRIPCION As Integer = 1
    Const C_ColCANTIDAD As Integer = 3
    Const C_ColPRECIOPUBLICO As Integer = 4
    Const C_COLCOSTO As Integer = 5
    Const C_ColORIGEN As Integer = 6
    Const C_ColPESOSFIJOS As Integer = 7
    Const C_ColCODANTERIOR As Integer = 2
    Public WithEvents btnNuevo As Button
    Public WithEvents btnGuardar As Button
    Friend WithEvents btnBuscar As Button
    Const C_COLCODARTPROVEEDOR As Integer = 8

    Sub Buscar()
        BuscarArticulos(False, CStr(0))
    End Sub

    Sub LimpiaGrid()

        If (isLoad = False) Then
            Exit Sub
        End If

        msgEtiquetas.Clear()
        Encabezado()
    End Sub

    Sub Nuevo()
        'Este procedimiento genera un nuevo registro para una venta
        'Se deben Limpiar todos los controles del formulario con excepcion del Control de la Llavve principal
        On Error GoTo Merr
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        txtDesArticulo.Text = ""
        txtOrdenCompra.Text = ""
        txtOrdenCompra.Tag = ""
        msgEtiquetas.Clear()
        Encabezado()
        mintTotalEtiquetas = 0
        mblnNuevo = True
        mblnSalir = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub

Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function ValidaDatos() As Object
        ''Esta Función valida que todos los datos e hayan introducido en el Form de Ventas , para poder procesar la venta
        '    On Local Error GoTo MErr:
        '
        '    If Trim(txtMotivo) = "" Then
        '        MsgBox "Proporcione el Motivo del Obsequio.", vbExclamation, gstrCorpoNombreEmpresa
        '        Me.txtMotivo.SetFocus
        '        Exit Function
        '    End If
        '    If txtTotal = 0 Then
        '        MsgBox "Proporcione el Detalle del Obsequio.", vbExclamation, gstrCorpoNombreEmpresa
        '        Me.msgEtiquetas.SetFocus
        '        Exit Function
        '    End If
        '    With msgEtiquetas
        '        For I = 1 To .Rows - 1
        '            If Numerico(.TextMatrix(I, C_ColCODIGO)) <> 0 And Numerico(.TextMatrix(I, C_ColCANTIDAD)) = 0 Then
        '                MsgBox "La cantidad de Artículos debe ser mayor de Cero." + vbNewLine + "Verifique Por Favor..", vbExclamation, gstrCorpoNombreEmpresa
        '                .Row = I
        '                .Col = C_ColCANTIDAD
        '                .SetFocus
        '                Exit Function
        '            End If
        '        Next
        '    End With
        '    ValidaDatos = True
        '    Exit Function
        'MErr:
        '    If Err.Number <> 0 Then ModEstandar.MostrarError
    End Function

    Function PuedeAbdandonarCaptIniciada() As Boolean
        'Esta Función Valida si se requiere autorizacion para abandonar una captura iniciada.
        'De ser asi, se pide el nombre y password de un usuario que pueda autorizar la salida.
        'Regresa Falso, si no puede Abandonar la captura sin guardar, de lo contrario regresa true.
        On Error GoTo Merr
        PuedeAbdandonarCaptIniciada = False
        If gblnAutAbandCapturaIniciada = True Then
            'Pedir el usuario y password para modificar el descto
            'Para esto se usará la forma: frmAutorizacionConfig.
            frmAutorizacionConfig.Text = "Autorizacion para Abandonar Captura Iniciada"
            frmAutorizacionConfig.ShowDialog()

            If gblnAutorizacionAceptada = False Then
                'Si la Peticion no fue aceptada, es decir que el usuario que se proporciono no tiene derecho para autorizar o para modificar
                'entonces no podrá ser modificado el descuento
                If gblnSalioSinValidar = False Then 'Si valido el Usuari y Password y no tuvo derecho, mostrar el aviso de ke no puede hacerlo
                    MsgBox(C_msgSINAUTORIZACION & "Abandonar la Captura sin Guardar la Información.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AVISO")
                End If
                Exit Function
            End If
        End If
        gblnAutorizacionAceptada = False 'Se pone Falso, para que cuando se requiera un nueva autorizacion, el valor inicial de esta sea falso. y unicamente si el usuario tiene autorizacion se modifique a true
        PuedeAbdandonarCaptIniciada = True

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Limpiar()
        'Esta función Limpia todos los controles del formulario.
        'No se valida si hubo cambios, ya que no es posible modificar una venta
        On Error GoTo Merr

        Nuevo()
        optArticulos.Focus()
        Exit Sub
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub


    Sub Encabezado()
        'Genera el encabezao del Grid, asigna el tamaño y número de columas y centra el texto dentro de ellas
        Dim LnContador As Integer

        With msgEtiquetas
            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusHeavy 'flexFocusLight 'flexFocusNone
            .WordWrap = True
            .FixedRows = 1
            .FixedCols = 0
            .set_RowHeight(0, 500)
            .set_ColWidth(C_COLCODIGO, 0, 1200)
            .set_ColWidth(C_ColDESCRIPCION, 0, 4100)
            .set_ColWidth(C_ColCANTIDAD, 0, 1100)
            .set_ColWidth(C_ColPRECIOPUBLICO, 0, 0)
            .set_ColWidth(C_COLCOSTO, 0, 0)
            .set_ColWidth(C_ColORIGEN, 0, 0)
            .set_ColWidth(C_ColPESOSFIJOS, 0, 0)
            .set_ColWidth(C_ColCODANTERIOR, 0, 1200)
            .set_ColWidth(C_COLCODARTPROVEEDOR, 0, 0)

            .set_ColAlignment(C_COLCODIGO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(C_ColDESCRIPCION, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(C_ColCODANTERIOR, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(C_ColCANTIDAD, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)

            .set_TextMatrix(0, C_COLCODIGO, "CODIGO")
            .set_TextMatrix(0, C_ColDESCRIPCION, "DESCRIPCION")
            .set_TextMatrix(0, C_ColCANTIDAD, "CANTIDAD")
            .set_TextMatrix(0, C_ColPRECIOPUBLICO, "P. PUBLICO")
            .set_TextMatrix(0, C_COLCOSTO, "costo")
            .set_TextMatrix(0, C_ColORIGEN, "ORIGEN")
            .set_TextMatrix(0, C_ColCODANTERIOR, "ANTERIOR")

            .Row = 0
            For LnContador = 0 To C_ColORIGEN
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next LnContador
            .Row = 1
            .TopRow = 1
            .Col = C_COLCODIGO
            .WordWrap = False 'Hacer esto , para que no se puedan escribir dos o mal lineas de texto en una  sola fila, solo se usa para el encabezado
        End With
    End Sub

    Sub InicializaVariables()
        Cambios = 0
        mblnNuevo = True
        mblnSalir = False
        FueraChange = False
    End Sub

    Private Sub frmImpresionEtiquetas_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmImpresionEtiquetas_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmImpresionEtiquetas_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        isLoad = True
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
        InicializaVariables()
        CargarFilayColumnaInicial()
    End Sub

    Private Sub frmImpresionEtiquetas_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        ' En este evento del formulario se valida la tecla presionada.
        ' Si es Enter se simula un tab(Avanza al siguiente control)
        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ' Si el control en que se presiono enter, es el Grid de Detalle de la venta que no se ejecute el avanzar tab

                If ActiveControl.Name <> "msgEtiquetas" Then
                    ModEstandar.AvanzarTab(Me)
                End If
            Case System.Windows.Forms.Keys.Escape
                If ActiveControl.Name <> "txtDetalle" Then
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmImpresionEtiquetas_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmImpresionEtiquetas_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSalir Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        'Else 'Se quiere salir con escape
        '    mblnSalir = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0 'Sale de la Captura, Con 1: Sigue en la captura
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmImpresionEtiquetas_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        CerrarAccess()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub msgEtiquetas_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgEtiquetas.DblClick
        msgEtiquetas_KeyPressEvent(msgEtiquetas, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub msgEtiquetas_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgEtiquetas.EnterCell
        'Aqui poner la descripcion del articulo cuando se este moviendo entre las filas del grid.
        'Poner la descripcion del articulo seleccionado, o dle que tenga la fila seleccionada.
        ' en el Textbox de Descripcion completa de abajo.
        With msgEtiquetas
            txtDesArticulo.Text = Trim(.get_TextMatrix(.Row, C_ColDESCRIPCION))
        End With
    End Sub

    Private Sub msgEtiquetas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgEtiquetas.Enter
        msgEtiquetas.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
        'Verificar si la columna en que esta posicionado dentro del grid es mayor de 7 (Son Ocho columnas las visibles.)
        'Entonces que se posicione en la columna del codigo
        If msgEtiquetas.Row = 0 Then
            msgEtiquetas.Row = 1
        End If
        If msgEtiquetas.Col > 5 Then
            msgEtiquetas.Col = C_COLCODIGO
        End If
        Pon_Tool()
        msgEtiquetas_EnterCell(msgEtiquetas, New System.EventArgs())
    End Sub

    Private Sub msgEtiquetas_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles msgEtiquetas.KeyDownEvent
        'Aqui debe cvalidarse el movimiento de teclas, si es ke se va a tomar en cuenta
        If optArticulos.Checked = False Then Exit Sub
        With msgEtiquetas
            Select Case eventArgs.keyCode
                Case System.Windows.Forms.Keys.Delete
                    'Si el cursor está en en renglón 0, que es el nombre de columna, entonces no se toma en cuenta la tecla Supr
                    If .Row = 0 Then Exit Sub
                    If .get_TextMatrix(.Row, C_COLCODIGO) <> "" And mblnNuevo = True Then
                        .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                        Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "Mensaje")
                            Case MsgBoxResult.No
                                'Poner el setfocus en el grid, para que siga dentro del mismo
                                msgEtiquetas.Focus()
                                Exit Sub
                            Case MsgBoxResult.Cancel
                                'Poner el setfocus en el grid, para que siga dentro del mismo
                                msgEtiquetas.Focus()
                                Exit Sub
                        End Select
                        BorraGrid(.Row) 'Cuando se Borra, se obtienen los nuevos totales, la cntidad de Articulos (Dento del proc.)                    .Col = C_ColCODIGO
                        .Focus()
                    End If
            End Select
        End With
    End Sub

    Private Sub msgEtiquetas_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles msgEtiquetas.KeyPressEvent
        '    If optArticulos.Value = False Then Exit Sub
        Dim ColSiguiente As Integer
        Dim rowsiguiente As Integer
        With msgEtiquetas
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then 'Para que cuando sea escape, no entre a editar el codigo,simplemente que se regrese al control anterior
                Select Case .Col
                    Case C_COLCODIGO ''-------------- SE EDITA EL codigo ---------------------'''''
                        If optArticulos.Checked = False Then Exit Sub
                        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        eventArgs.keyAscii = ModEstandar.MskCantidad(txtDetalle.Text, eventArgs.keyAscii, 8, 0, (txtDetalle.SelectionStart))
                        txtDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                        '''en esta parte se validará si es el rengón, columna que le corresponde editarse
                        If (.Row > 1) Then
                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
                            If CDbl(Numerico(.get_TextMatrix(.Row - 1, C_COLCODIGO))) = 0 Or CDbl(Numerico(.get_TextMatrix(.Row, C_COLCODIGO))) <> 0 Then
                                .Focus()
                                Exit Sub
                            End If
                        End If
                        '                    If Numerico((.TextMatrix(.Row, C_ColCODIGO))) <> 0 Then
                        '                        .SetFocus
                        '                        Exit Sub
                        '                    End If
                        ModEstandar.MSHFlexGridEdit(msgEtiquetas, txtDetalle, eventArgs.keyAscii)
                        txtDetalle.SelectionStart = Len(txtDetalle.Text)
                        If Len(Trim(txtDetalle.Text)) <> 1 Then
                            ModEstandar.SelTextoTxt(txtDetalle)
                        End If

                    Case C_ColDESCRIPCION ''-------------- SE EDITA LA DESCRIPCION ---------------------'''''
                        If optArticulos.Checked = False Then Exit Sub
                        txtDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                        '''en esta parte se validará si es el rengón, columna que le corresponde editarse
                        If (.Row > 1) Then
                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
                            If CDbl(Numerico(.get_TextMatrix(.Row - 1, C_COLCODIGO))) = 0 Then
                                .Focus()
                                Exit Sub
                            End If
                        End If
                        ModEstandar.MSHFlexGridEdit(msgEtiquetas, txtDetalle, eventArgs.keyAscii)
                        If Len(Trim(txtDetalle.Text)) <> 1 Then
                            ModEstandar.SelTextoTxt(txtDetalle)
                        End If

                    Case C_ColCANTIDAD '-------------- SE EDITA LA CANTIDAD ---------------------'''''
                        eventArgs.keyAscii = ModEstandar.MskCantidad(txtDetalle.Text, eventArgs.keyAscii, 3, 0, (txtDetalle.SelectionStart))
                        txtDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                        'Validar que en el codigo y en la descripcion exista valor para editar este campo
                        If CDbl(Numerico(.get_TextMatrix(.Row, C_COLCODIGO))) = 0 And .get_TextMatrix(.Row, C_ColDESCRIPCION) = "" Then
                            .Focus()
                            Exit Sub
                        End If
                        ModEstandar.MSHFlexGridEdit(msgEtiquetas, txtDetalle, eventArgs.keyAscii)
                        txtDetalle.SelectionStart = Len(txtDetalle.Text)
                        If Len(Trim(txtDetalle.Text)) <> 1 Then
                            ModEstandar.SelTextoTxt(txtDetalle)
                        End If
                End Select
            End If
        End With

    End Sub

    Private Sub msgEtiquetas_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgEtiquetas.Leave
        msgEtiquetas.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub

    Sub LimpiaDatosArticulo(ByRef Control As String)
        'Control, Determina el nombre del control que ejecuta este procedimiento. Esto es, para saber si borrar el código o la Descripción
        '"D"=Descripcion... "C"=Codigo
        On Error GoTo Merr
        'Este Procedimiento Limpialos Campos Correspondientes a un Artículo, cuando se cambie de Articulo, que se limpien los datos
        With msgEtiquetas
            If Control = "C" Then
                .set_TextMatrix(.Row, C_ColDESCRIPCION, "")
            Else
                .set_TextMatrix(.Row, C_COLCODIGO, "")
            End If
            .set_TextMatrix(.Row, C_ColCANTIDAD, "")
            txtDesArticulo.Text = ""
            If CDbl(Numerico(.get_TextMatrix(.Row, C_COLCODIGO))) = 0 Then
                .set_TextMatrix(.Row, C_COLCODIGO, "")
            End If
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenarDatosArticulo(ByRef Articulo As Integer, ByRef Campo As String)
        'Articulo: Este parametro contiene el artiuclo que se buscará, que puede ser el codigo del articulo o la descripcion desl mis.
        'Para saber si es codigo o descripcion se toma en cuenta el parametro campo, si este es "C", sera un codigo, de lo contraio es descripcion.
        'Campo: Campo desde donde se esta ejecutando la busqueda de Articulo C=Codigo, D=DEscripcion
        On Error GoTo Merr
        Dim CodArticulo As Integer 'Esta Varibale contiene el codigo del articulo sobre el que se este buscando informacion. ya que se puede buscar por codigo o por descripcion
        Dim ImpDescuento As Decimal 'Contiene el Importe de Decuento
        Dim ImpPromocion As Decimal 'Contiene el Importe de Promoción
        Dim PrecioPubDolar As Decimal 'Contiene el Precio al público en Dólares
        Dim importe As Decimal
        Dim Cantidad As Integer
        Dim ImporteNeto As Object 'Contiene el Valor del Importe Neto (Importe - (Descuento o Promocion * CAntidad))
        Dim ColSiguiente As Integer 'Contiene el Número de COlumna en la que se quedará al Salir de Este proceimiento de ACuerdo a los datos del Artículo
        Dim rowsiguiente As Integer 'Contiene el Número de Fila en la que se quedará al Salir de Este proceimiento de ACuerdo a los datos del Artículo
        If Articulo = 0 Or Trim(CStr(Articulo)) = "" Then
            LimpiaDatosArticulo(Campo)
            Exit Sub
        End If

        'Este Procedimiento muestra los datos de un Articulo dado
        If Campo = "C" Then 'Se buscará por codigo de articulo
            gStrSql = "Select CodArticulo,ltrim(rtrim(DescArticulo)) as DescArticulo,DescUnidad as Unidad,CodGrupo,MonedaCompra,PrecioPubDolar,CostoReal," & "Isnull(CodFamilia,0 ) as CodFamilia,Isnull(CodFamilia ,0) as CodFamilia,isnull(CodLinea,0) as CodLinea,isnull(CodSublinea,0) as CodSublinea," & "isnull(CodMarca,0) as CodMarca,isnull(CodModelo,0) CodModelo,isnull(CodTipoMaterial,0 ) as CodTipoMaterial,isnull(Genero,0) as Genero, isnull(Movimiento ,0) as Movimiento,PesosFijos,codAlmacenOrigen," & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-'+ RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as CODANT,CodigoArticuloProv " & "From CatArticulos,CatUnidades " & "Where CatArticulos.CodUnidad = CatUnidades.CodUnidad And CodArticulo =  " & Articulo
        ElseIf Campo = "D" Then  ' Se buscará por descripcion
            gStrSql = "Select CodArticulo,ltrim(rtrim(DescArticulo)) as DescArticulo,DescUnidad as Unidad,CodGrupo,MonedaCompra,PrecioPubDolar,CostoReal," & "Isnull(CodFamilia,0 ) as CodFamilia,Isnull(CodFamilia ,0) as CodFamilia,isnull(CodLinea,0) as CodLinea,isnull(CodSublinea,0) as CodSublinea," & "isnull(CodMarca,0) as CodMarca,isnull(CodModelo,0) CodModelo,isnull(CodTipoMaterial,0 ) as CodTipoMaterial,isnull(Genero,0) as Genero,isnull(Movimiento ,0) as Movimiento,PesosFijos,CodAlmacenOrigen," & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-'+ RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as CODANT,CodigoArticuloProv " & "From CatArticulos,CatUnidades " & "Where CatArticulos.CodUnidad = CatUnidades.CodUnidad And DescArticulo Like  '" & Trim(CStr(Articulo)) & "'"
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        With msgEtiquetas
            If RsGral.RecordCount > 0 Then
                CodArticulo = CDbl(RsGral.Fields("CodArticulo").Value)
                PrecioPubDolar = RsGral.Fields("PrecioPubDolar").Value

                .set_TextMatrix(.Row, C_COLCODIGO, RsGral.Fields("CodArticulo").Value)
                .set_TextMatrix(.Row, C_ColDESCRIPCION, Trim(RsGral.Fields("DescArticulo").Value))
                .set_TextMatrix(.Row, C_ColCODANTERIOR, RsGral.Fields("CodAnt").Value)
                .set_TextMatrix(.Row, C_ColORIGEN, RsGral.Fields("CodAlmacenOrigen").Value)
                .set_TextMatrix(.Row, C_ColPESOSFIJOS, RsGral.Fields("PesosFijos").Value)
                .set_TextMatrix(.Row, C_COLCODARTPROVEEDOR, RsGral.Fields("CodigoArticuloProv").Value)
                .set_TextMatrix(.Row, C_ColPRECIOPUBLICO, System.Math.Round(RsGral.Fields("PrecioPubDolar").Value, 0))
                .set_TextMatrix(.Row, C_COLCOSTO, RsGral.Fields("CostoReal").Value)
                .set_TextMatrix(.Row, C_ColCANTIDAD, 1)
                rowsiguiente = .Row + 1
                AgregarFilaFinal()
                TotaldeEtiquetas()
                .Col = C_COLCODIGO
                .Row = rowsiguiente
            Else
                MsjNoExiste("El Artículo", gstrCorpoNOMBREEMPRESA)
                If Campo = "C" Then
                    msgEtiquetas.set_TextMatrix(msgEtiquetas.Row, C_COLCODIGO, "")
                ElseIf Campo = "D" Then
                    msgEtiquetas.set_TextMatrix(msgEtiquetas.Row, C_ColDESCRIPCION, "")
                End If
                LimpiaDatosArticulo(Campo)
                .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                .Focus()
                .Col = C_COLCODIGO
                .Row = .Row
            End If
        End With
        Exit Sub
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function ObtenerCantidadArticulos() As Integer
        'Este Procedmiento obtiene el total de Articulos de uan venta elaborada
        On Error GoTo Merr
        Dim FilaAnt As Integer
        Dim Cantidad As Integer
        With msgEtiquetas
            Cantidad = 0
            FilaAnt = .Row
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, C_ColCANTIDAD) = "" Then Exit For
                Cantidad = Cantidad + CDbl(Numerico(.get_TextMatrix(I, C_ColCANTIDAD)))
            Next
            .Row = FilaAnt
        End With
        ObtenerCantidadArticulos = Cantidad
        '    txtTotalArticulos = Cantidad
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub BuscarArticulos(ByRef BusquedaEspecial As Boolean, ByRef CodArticulo As String)
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim strControlActual As String 'Nombre del control actual
        Dim Columna As Integer

        strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "MSGETIQUETAS", "TXTDETALLE"
                With msgEtiquetas
                    'Obtener la columna de donde se está ejecutando la consulta
                    Columna = .Col
                    If Columna = C_COLCODIGO Then 'Se Busca por código
                        strCaptionForm = "Consulta de Articulos"
                        If BusquedaEspecial Then
                            '                        strSQL = "SELECT     CodArticulo AS CODIGO, RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION, CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-'+ RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]  " & _
                            ''                            "From CatArticulos WHERE (CodArticulo = " & CLng(CodArticulo) & ") " & _
                            '"OR   (OrigenAnt = " & CLng(Left(CodArticulo, 1)) & ") AND (CodigoAnt = " & CLng(Right(CodArticulo, 5)) & ")"
                            strSQL = "SELECT     CodArticulo AS CODIGO, RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION, " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR], " & "dbo.FormatCantidad(A.PrecioPubDolar)  AS [PRECIO PÚBLICO] , " & "case PesosFijos WHEN 0 THEN 'DÓLARES' WHEN 1 THEN 'PESOS' END AS [MONEDA] " & "From CatArticulos A cross Join Configuraciongeneral c WHERE (CodArticulo = " & CInt(CodArticulo) & ") " & "OR   (OrigenAnt = " & CInt(VB.Left(CodArticulo, 1)) & ") AND (CodigoAnt = " & CInt(VB.Right(CodArticulo, 5)) & ")"

                        Else
                            strSQL = "sELECT CodArticulo as CODIGO, lTRIM(RTRIM(DescArticulo)) as DESCRIPCION,  CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]      From CatArticulos"
                        End If
                    ElseIf Columna = C_ColDESCRIPCION Then  'Se busca por descripción
                        strCaptionForm = "Consulta de Articulos"
                        strSQL = "sELECT  Rtrim(Ltrim(DescArticulo))  as DESCRIPCION ,  CodArticulo as CODIGO,  CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]       From CatArticulos where DescArticulo Like '" & Trim(txtDetalle.Text) & "%'"
                    Else
                        'Sale del Sub si no es ninguna de estas columnas de donde se ejecuto la consulta, y no hace nada
                        Exit Sub
                    End If
                End With
            Case "TXTORDENCOMPRA"
                strCaptionForm = "Consulta de Ordenes de Compra"
                If Trim(txtOrdenCompra.Text) = "" Then
                    strSQL = "SELECT FolioOrdenCompra AS 'FOLIO DE ORDEN',LTRIM(RTRIM(DBO.FormatFecha(FechaOrdenCompra,10))) AS 'FECHA',CASE WHEN Estatus = 'G' THEN 'GENERADA' WHEN Estatus = 'R' THEN 'REGISTRADA' END AS 'ESTATUS',DBO.FormatCantidad(Total) AS 'IMPORTE' " & "FROM OrdenesCompra WHERE Estatus In ('G','R') ORDER BY FechaOrdenCompra DESC,FolioOrdenCompra DESC"
                Else
                    strSQL = "SELECT FolioOrdenCompra AS 'FOLIO DE ORDEN',LTRIM(RTRIM(DBO.FormatFecha(FechaOrdenCompra,10))) AS 'FECHA',CASE WHEN Estatus = 'G' THEN 'GENERADA' WHEN Estatus = 'R' THEN 'REGISTRADA' END AS 'ESTATUS',DBO.FormatCantidad(Total) AS 'IMPORTE' " & "FROM OrdenesCompra WHERE FolioOrdenCompra LIKE '" & txtOrdenCompra.Text & "%' AND Estatus In ('G','R') ORDER BY FechaOrdenCompra DESC,FolioOrdenCompra DESC"
                End If
            Case Else
                'Sale de este sub para ke no ejecute ninguna opcion
                Exit Sub
        End Select

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, strSQL))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsgBox(C_msgSINDATOS & vbNewLine & "Verifique Por Favor....", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta
        'Load(FrmConsultas)
        If strControlActual = "TXTORDENCOMPRA" Then
            Call ConfiguraConsultas(FrmConsultas, 6300, RsGral, strTag, strCaptionForm)
            With FrmConsultas.Flexdet
                .set_ColWidth(0, 0, 2000)
                .set_ColWidth(1, 0, 1300)
                .set_ColWidth(2, 0, 1200)
                .set_ColWidth(3, 0, 1500)
                .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
                .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            End With
        Else
            If BusquedaEspecial = True Then
                Call ConfiguraConsultas(FrmConsultas, 10300, RsGral, strTag, strCaptionForm)
            Else
                Call ConfiguraConsultas(FrmConsultas, 7600, RsGral, strTag, strCaptionForm)
            End If
            With FrmConsultas.Flexdet
                Select Case strControlActual
                    Case "MSGETIQUETAS", "TXTDETALLE"
                        With msgEtiquetas
                            'Obtener la columna de donde se está ejecutando la consulta
                            Columna = .Col
                        End With
                        If BusquedaEspecial = True Then
                            If Columna = C_COLCODIGO Then 'Se Busca por código
                                .set_ColWidth(0, 0, 900)
                                .set_ColWidth(1, 0, 4800)
                                .set_ColWidth(2, 0, 1700)
                                .set_ColWidth(3, 0, 1700)
                                .set_ColWidth(4, 0, 1200)
                                .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                                .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                                .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                                .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                                .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
                            ElseIf Columna = C_ColDESCRIPCION Then  'Se busca por descripción
                                .set_ColWidth(0, 0, 4800)
                                .set_ColWidth(1, 0, 900)
                                .set_ColWidth(2, 0, 1700)
                                .set_ColWidth(3, 0, 1700)
                                .set_ColWidth(4, 0, 1200)
                                .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                                .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                                .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                                .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                                .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
                            Else
                                'Sale del Sub si no es ninguna de estas columnas de donde se ejecuto la consulta, y no hace nada
                                Exit Sub
                            End If
                        Else

                            If Columna = C_COLCODIGO Then 'Se Busca por código
                                .set_ColWidth(0, 0, 900)
                                .set_ColWidth(1, 0, 4800)
                                .set_ColWidth(2, 0, 1900)
                                .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                                .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                                .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                            ElseIf Columna = C_ColDESCRIPCION Then  'Se busca por descripción
                                .set_ColWidth(0, 0, 4800)
                                .set_ColWidth(1, 0, 900)
                                .set_ColWidth(2, 0, 1900)
                                .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                                .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                                .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                            Else
                                'Sale del Sub si no es ninguna de estas columnas de donde se ejecuto la consulta, y no hace nada
                                Exit Sub
                            End If
                        End If
                End Select
            End With
        End If
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub AgregarFilaFinal()
        'Agrea una Fila al Final del Grid, Sólo cuando sea necesario hacerlo
        With msgEtiquetas
            If .Row = .Rows - 1 Then
                ' Si se Presiono enter y estamos en la ultima fila, entonces se agregrara una nueva fila
                .AddItem("")
                ScrollGrid()
            End If
        End With
    End Sub

    Sub AgregarFilaFinalGrid(ByRef lRen As Integer)
        'Agrea una Fila al Final del Grid, Sólo cuando sea necesario hacerlo
        With msgEtiquetas
            If lRen = .Rows - 1 Then
                ' Si se Presiono enter y estamos en la ultima fila, entonces se agregrara una nueva fila
                .AddItem("")
                ScrollGrid()
            End If
        End With
    End Sub

    Sub BorraGrid(ByRef Row As Integer)
        'Este Procediento borra un renglon del Grid
        'Si el Número de Filas que kedan en el grid, es menor de 8, se insertará una nueva fila al final del grid
        With msgEtiquetas
            .RemoveItem(Row)
            'Si el número de filas es menor de 10 o esta posicionado en la utlima fila, entonces, agrega una fila
            If .Rows < 11 Or .Row = .Rows - 1 Then
                .AddItem("")
                .Row = .Row
            End If
        End With
        'Al borrar se deben obtener los nuevo totales y Actualizar la cantidad de articulos
    End Sub

    Public Sub ScrollGrid()
        'Procedimiento que pone el enfoque en el primer renglón vacío del Grid
        Dim I As Integer
        Dim nCont As Integer 'Cuenta los renglones que están ocupados (que no están vacíos)
        Dim nRen As Integer
        'Aparecen 7 renglones disponibles en el Grid
        'Si son menos de siete registros ocupados, no se utiliza el .TopRow
        'Pero, si son 7 ó más registros, el .TopRow manda el enfoque al primer renglón vacío
        'después de los renglones ocupados
        nRen = 7 'El máximo de renglones que aparece en el grid (Además del encabezado)
        nCont = 0
        With msgEtiquetas
            For I = 1 To .Rows
                If Trim(.get_TextMatrix(I, C_COLCODIGO)) <> "" Then
                    nCont = nCont + 1
                Else
                    Exit For
                End If
            Next I
            If nCont < 7 Then
                'Hay menos de 7 registros
                '            .TopRow = 7
                .Row = nCont + 1
                .Col = C_COLCODIGO

            Else
                'Hay 7 ó más registros, hay que recorrer el grid
                .TopRow = (nCont - nRen) + 2
                .Row = nCont + 1
                .Col = C_COLCODIGO
            End If
        End With
    End Sub

    Private Sub msgEtiquetas_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgEtiquetas.Scroll
        txtDetalle.Visible = False
    End Sub

    'UPGRADE_WARNING: Event optArticulos.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optArticulos_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optArticulos.CheckedChanged
        If eventSender.Checked Then
            If optArticulos.Checked = True Then
                mintTotalEtiquetas = 0
                txtOrdenCompra.Enabled = False
                txtOrdenCompra.Text = ""
                LimpiaGrid()
            End If
        End If
    End Sub

    Private Sub optArticulos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles optArticulos.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If ModEstandar.Salir Then Me.Close()
        End If
    End Sub

    Private Sub optOrdenCompra_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optOrdenCompra.CheckedChanged
        If eventSender.Checked Then
            If optOrdenCompra.Checked = True Then
                mintTotalEtiquetas = 0
                txtOrdenCompra.Enabled = True
                LimpiaGrid()
            End If
        End If
    End Sub

    Private Sub optOrdenCompra_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles optOrdenCompra.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If ModEstandar.Salir Then Me.Close()
        End If
    End Sub
    Private Sub txtCodArticulo_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
        If KeyCode = System.Windows.Forms.Keys.Escape Then optArticulos.Focus()
    End Sub

    Private Sub txtCodArticulo_KeyPress(ByRef KeyAscii As Integer)
        gp_CampoNumerico(KeyAscii)
    End Sub

    'Private Sub txtCodArticulo_LostFocus()
    '    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    '    Dim ResBusquedaArt  As Long
    '    Dim CodAux As Long
    '    If Trim(txtCodArticulo) <> "" Then
    '        ResBusquedaArt = BuscarCodigoArticulo(CLng((Val(txtCodArticulo))))
    '        If ResBusquedaArt > 0 Or ResBusquedaArt = -1 Then
    '            MostrarArticulo CDbl(ResBusquedaArt), "", "A"
    '        ElseIf ResBusquedaArt = -2 Then
    '            CodAux = CLng(txtCodArticulo)
    '            txtCodArticulo = ""
    '            BuscarArticulos True, Right(String(6, "0") + Trim(CodAux), 6)
    '        End If
    '    End If
    'End Sub

    'Private Sub txtDescArticulo_Change()
    '    If FueraChange = True Then Exit Sub
    '    FueraChange = True
    '    txtCodArticulo = ""
    '    FueraChange = False
    'End Sub



    Private Sub txtDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetalle.Enter
        txtDetalle.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtDetalle.Width) + 10)
        txtDetalle.Text = Trim(txtDetalle.Text)
        Pon_Tool()
    End Sub

    Private Sub txtDetalle_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDetalle.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Aqui se muestran los datos del control editable, en el Grid
        'Se deberá formatear el Valor de Acuerdo al Tipo de Dato en uso
        Dim rowsiguiente As Integer
        Dim ColSiguiente As Integer
        Dim FormatoCantidad As String
        Dim ResBusquedaArt As Integer
        With msgEtiquetas
            Select Case KeyCode

                Case System.Windows.Forms.Keys.Escape
                    txtDetalle.Visible = False
                    txtDetalle.Text = ""
                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                    .Focus()
                Case System.Windows.Forms.Keys.Return
                    rowsiguiente = .Row + 1
                    ColSiguiente = C_ColCANTIDAD
                    'Si la Columna en que se está escribiendo es Codigo o Cantidad, Formatear el Valor par que quede numérico
                    If .Col = C_COLCODIGO Then
                        .set_TextMatrix(.Row, .Col, Trim(txtDetalle.Text))
                    ElseIf .Col = C_ColCANTIDAD Then
                        If gbytPosicionesDecimal = 0 Then
                            FormatoCantidad = CStr(0)
                        Else
                            FormatoCantidad = "0." & New String("0", gbytPosicionesDecimal)
                        End If
                        .set_TextMatrix(.Row, .Col, VB6.Format(Numerico(txtDetalle.Text), FormatoCantidad))
                        TotaldeEtiquetas()
                    Else
                        .set_TextMatrix(.Row, .Col, Trim(txtDetalle.Text))
                    End If
                    FueraChange = True
                    txtDetalle.Text = ""
                    txtDetalle.Visible = False
                    .Focus()
                    If .Col = C_COLCODIGO Then
                        ResBusquedaArt = BuscarCodigoArticulo(Trim(.get_TextMatrix(.Row, C_COLCODIGO)))
                        If ResBusquedaArt > 0 Or ResBusquedaArt = -1 Then
                            LlenarDatosArticulo(CDbl(ResBusquedaArt), "C")
                            Exit Sub
                        ElseIf ResBusquedaArt = -2 And CDbl(Numerico(.get_TextMatrix(.Row, C_COLCODIGO))) <> 0 Then
                            'BuscarArticulos True, Right(String(6, "0") + Trim(.TextMatrix(.Row, C_ColCODIGO)), 6)
                            ResBusquedaArt = CInt(Trim(.get_TextMatrix(.Row, C_COLCODIGO)))
                            .set_TextMatrix(.Row, C_COLCODIGO, "")
                            BuscarArticulos(True, VB.Right(New String("0", 6) & CStr(ResBusquedaArt), 6))
                            ColSiguiente = C_COLCODIGO
                        Else
                            rowsiguiente = .Row
                        End If
                        AgregarFilaFinal()
                    ElseIf .Col = C_ColDESCRIPCION Then
                        'Si estamos en la columna descripcion, se deberá también mostrar los datos del articulo.
                        '                    LlenarDatosArticulo Trim(.TextMatrix(.Row, C_ColDESCRIPCION)), "D"
                    ElseIf .Col = C_ColCANTIDAD Then
                        If CDbl(Numerico(.get_TextMatrix(.Row, C_ColCANTIDAD))) <= 0 Then
                            MsgBox("La Cantidad mínima debe ser 1." & vbNewLine & "Verfique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AVISO")
                            .Focus()
                            .set_TextMatrix(.Row, C_COLCODIGO, 1)
                            Exit Sub
                        End If
                    End If
                    .Row = rowsiguiente
                    .Col = ColSiguiente
                Case System.Windows.Forms.Keys.Delete
                    'Si la Tecla Presiona fue SUPR y estamos en la COL Código, se Limpian los controles correspondientes al articulo
                    If .Col = C_COLCODIGO Then
                        LimpiaDatosArticulo("C")
                    ElseIf .Col = C_ColDESCRIPCION Then
                        LimpiaDatosArticulo("D")
                    End If
            End Select
        End With
    End Sub

    Private Sub txtDetalle_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDetalle.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'En este Evento se validan los datos que se introduzcan al control txtDetalle,dependiendo de la columan en que se esté editando
        If KeyAscii = 0 Or KeyAscii = 13 Then GoTo EventExitSub
        With msgEtiquetas
            If .Col = C_COLCODIGO Then
                KeyAscii = ModEstandar.MskCantidad(txtDetalle.Text, KeyAscii, 8, 0, (txtDetalle.SelectionStart))
            End If
            If .Col = C_ColCANTIDAD Then
                KeyAscii = ModEstandar.MskCantidad(txtDetalle.Text, KeyAscii, 3, CInt(CStr(gbytPosicionesDecimal)), (txtDetalle.SelectionStart))
            End If
        End With
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    'Sub Guardar()
    '    On Error GoTo MErr:
    '    Dim CodArticulo As Long
    '    Dim Cantidad As Integer
    '    Dim FechaInventario
    '
    '    FechaInventario = dtpFechaInventario
    '    'Indicar en que procedimiento de Guardar nos encontramos.
    '    gstrProcesoqueGeneraError = "frmInvElectronico (Guardar) "
    '    If mblnNuevo = False Then Exit Sub   'Si no es un Registro Nuevo, Se sale del Proc
    '    If Me.ValidaDatos = False Then
    '        Exit Sub
    '    End If
    '
    '    'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
    '    Screen.MousePointer = vbHourglass
    '    Cnn.BeginTrans
    '    'Obtener el Folio con el cual se almacenará la Venta:
    '    'La estructura es: Prefijo - Sucursal - CodCaja - Fecha(Año-Mes-Dia) - Consecutivo
    '    gStrSql = "Select prefijo, consecutivo as consecutivo from CatFolios where DescFolio= 'OBSEQUIOS' " & " And CodAlmacen = " & gintCodAlmacen
    '
    '    ModEstandar.BorraCmd
    '    Cmd.CommandText = "dbo.Up_Select_Datos"
    '    Cmd.CommandType = adCmdStoredProc
    '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
    '    Set RsGral = Cmd.Execute
    '
    '    'Guaradar la Información del Obsequio, incluyendo el Detalle
    '    With msgEtiquetas
    '        For I = 1 To .Rows - 1
    '            If Numerico(.TextMatrix(I, C_ColCODIGO)) = 0 Then Exit For
    '            NumPartida = I
    '            CodArticulo = Numerico(.TextMatrix(I, C_ColCODIGO))
    '            Cantidad = Numerico(.TextMatrix(I, C_ColCANTIDAD))
    '            PrecioPublico = .TextMatrix(I, C_ColPRECIOPUBLICO)
    '            CostoVenta = ObtenerCostoArticulo(CodArticulo)
    '            ModStoredProcedures.PR_IEMObsequios FolioObsequio, Format(FechaObsequio, C_FORMATFECHAGUARDAR), CStr(gintCodAlmacen), CStr(gintCodCaja), CStr(intCodCliente), Trim(dbcCliente), Trim(txtRFCCliente), Trim(ModEstandar.QuitaEnter(txtMotivo)), CStr(intCodVendedor), CStr(NumPartida), CStr(CodArticulo), CStr(Cantidad), CStr(PrecioPublico), CStr(CostoVenta), Estatus, "01/01/1900", CStr(TipoCambioDolar), C_INSERCION, 0
    '            Cmd.Execute
    '        Next
    '    End With
    '    'Realizar el Movimiento de Alamcen Correspondiente, En este Caso se Registra una salida de Alamacen
    '    If RealizarMovimientosDeAlmacen(FolioObsequio, FechaObsequio, CodMovtoAlm, "S") = False Then
    '        Err.Raise 0, , "Error al realizar el movimiento de almacen(salida) en obsequios"
    '    End If
    '    Screen.MousePointer = vbDefault
    '    Cnn.CommitTrans
    '
    '    'Por cuestiones de estética el cambio al puntero del mouse se hace antes de iniciar la transacción y al finalizar la misma.
    '
    '    If mblnNuevo Then
    '        MsgBox "El obsequio ha sido grabado correctamente con el código: " & _
    ''            FolioObsequio, vbInformation + vbOKOnly, "Mensaje"
    '    Else
    '        MsgBox C_msgACTUALIZADO, vbInformation + vbOKOnly, ModVariables.gstrCorpoNombreEmpresa
    '    End If
    '    'Dejar el Procedimiento Nuevo, sirve para que al usar limpiar,. no pregunte si se desea guardar cambios en el codigo
    '    Nuevo
    '    Limpiar
    '    Exit Sub
    'MErr:
    '    If Err.Number <> 0 Then
    '        Screen.MousePointer = vbDefault
    '        Cnn.RollbackTrans
    '        ModEstandar.MostrarError "Ocurrió un Error en el Formulario y Proceso: " + gstrProcesoqueGeneraError
    '    End If
    ''Resume
    'End Sub

    Sub MostrarArticulo(ByRef CodArticulo As Integer, ByRef OrdenCompra As String, ByRef TipoBusqueda As String)
        If CDbl(Numerico(Trim(CStr(CodArticulo)))) = 0 And OrdenCompra = "" Then Exit Sub
        If TipoBusqueda = "A" Then
            gStrSql = "SELECT Ltrim(Rtrim(DescArticulo)) as DEscArticulo,PrecioPubDolar,CostoReal,CodAlmacenOrigen,PesosFijos, " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-'+ RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as CODANT,CodigoArticuloProv " & "From CatArticulos A where CodArticulo = " & CodArticulo & ""
        Else
            gStrSql = "SELECT A.CodArticulo, A.DescArticulo, A.PrecioPubDolar, A.CostoReal, A.CodAlmacenOrigen, A.PesosFijos,  " & "CASE A.CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), A.OrigenAnt) + '-'+ RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5), A.CodigoAnt))) ,5) End as CODANT, A.CodigoArticuloProv, O.CantidadRecepcion " & "FROM dbo.OrdenesCompraPreCat O INNER JOIN " & "dbo.CatArticulos A ON O.CodArticulo = A.CodArticulo AND O.FolioOrdenCompra = '" & OrdenCompra & "'"
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        With msgEtiquetas
            msgEtiquetas.Clear()
            Encabezado()
            If RsGral.RecordCount > 0 Then
                If TipoBusqueda = "A" Then
                    FueraChange = True
                    FueraChange = False
                    .set_TextMatrix(.Row, C_COLCODIGO, CodArticulo)
                    .set_TextMatrix(.Row, C_ColDESCRIPCION, RsGral.Fields("DescArticulo").Value)
                    .set_TextMatrix(.Row, C_ColCANTIDAD, 1)
                    .set_TextMatrix(.Row, C_ColPRECIOPUBLICO, System.Math.Round(RsGral.Fields("PrecioPubDolar").Value, 0))
                    .set_TextMatrix(.Row, C_COLCOSTO, RsGral.Fields("CostoReal").Value)
                    .set_TextMatrix(.Row, C_ColORIGEN, RsGral.Fields("CodAlmacenOrigen").Value)
                    .set_TextMatrix(.Row, C_ColPESOSFIJOS, RsGral.Fields("PesosFijos").Value)
                    .set_TextMatrix(.Row, C_ColCODANTERIOR, RsGral.Fields("CodAnt").Value)
                    .set_TextMatrix(.Row, C_COLCODARTPROVEEDOR, RsGral.Fields("CodigoArticuloProv").Value)
                Else
                    For I = 1 To RsGral.RecordCount
                        .set_TextMatrix(I, C_COLCODIGO, RsGral.Fields("CodArticulo").Value)
                        .set_TextMatrix(I, C_ColDESCRIPCION, RsGral.Fields("DescArticulo").Value)
                        .set_TextMatrix(I, C_ColCANTIDAD, RsGral.Fields("CantidadRecepcion").Value)
                        .set_TextMatrix(I, C_ColPRECIOPUBLICO, System.Math.Round(RsGral.Fields("PrecioPubDolar").Value, 0))
                        .set_TextMatrix(I, C_ColORIGEN, RsGral.Fields("CodAlmacenOrigen").Value)
                        .set_TextMatrix(I, C_COLCOSTO, RsGral.Fields("CostoReal").Value)
                        .set_TextMatrix(I, C_ColPESOSFIJOS, RsGral.Fields("PesosFijos").Value)
                        .set_TextMatrix(I, C_ColCODANTERIOR, RsGral.Fields("CodAnt").Value)
                        .set_TextMatrix(I, C_COLCODARTPROVEEDOR, RsGral.Fields("CodigoArticuloProv").Value)
                        AgregarFilaFinalGrid(I)
                        RsGral.MoveNext()
                    Next
                End If
            Else
                MsgBox("La orden de compra " & Trim(txtOrdenCompra.Text) & " no existe" & vbNewLine & "Verifique por favor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                If optArticulos.Checked = True Then optArticulos.Focus() Else optOrdenCompra.Focus()
                Exit Sub
            End If
            TotaldeEtiquetas()
            .Row = 1
            .Col = C_COLCODIGO
        End With
    End Sub

    Private Sub txtDetalle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetalle.Leave
        txtDetalle.Visible = False
    End Sub

    Private Sub txtOrdenCompra_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrdenCompra.TextChanged
        If Len(txtOrdenCompra.Text) = Len(txtOrdenCompra.Tag) - 1 Or Trim(txtOrdenCompra.Text) = "" Then
            msgEtiquetas.Clear()
            Encabezado()
            TotaldeEtiquetas()
        End If
    End Sub

    Private Sub txtOrdenCompra_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrdenCompra.Enter
        ModEstandar.SelTextoTxt(txtOrdenCompra)
    End Sub

    Private Sub txtOrdenCompra_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrdenCompra.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        MostrarArticulo(0, Trim(txtOrdenCompra.Text), "O")
        txtOrdenCompra.Tag = Trim(txtOrdenCompra.Text)
    End Sub

    Sub Guardar()
        On Error GoTo Errores
        Dim miArchivo As String

        miArchivo = Dir(gstrCorpoDriveLocal & "\Sistema\InvElect\" & C_BDACCESS & ".MDB", FileAttribute.Archive)
        If Trim(miArchivo) = "" Then
            MsgBox("No existe el archivo de impresion de etiquetas, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            If mintTotalEtiquetas = 0 Then
                MsgBox("Debe indicarse el no. de etiquetas a imprimir", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                Exit Sub
            End If
            If AbrirAccess() Then
                Atributos_Recordset_Access(RsGralAccess)
                ImprimeEtiqueta()
                MsgBox("La información fue generada", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                Limpiar()
            Else
                MsgBox("Hay problemas con el archivo de impresión de etiquetas, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            End If
        End If

Errores:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    'Sub Imprime()
    '   On Error GoTo Errores
    '   If mintTotalEtiquetas = 0 Then
    '      MsgBox "Debe indicarse el no. de etiquetas a imprimir", vbExclamation, gstrCorpoNOMBREEMPRESA
    '      Exit Sub
    '   End If
    '   Common.ShowPrinter
    '   ImprimeEtiqueta
    'Errores:
    '   If (Err.Number <> cdlCancel And Err.Number <> 0) Then
    '      ModEstandar.MostrarError
    '   End If
    'End Sub

    Sub ImprimeEtiqueta()
        On Error GoTo Err_Renamed
        Dim I As Integer
        Dim Y As Integer
        Dim Ren, Col As Integer
        'Dim Espacio As Integer
        Dim lMsg As String
        Dim lDESC As String
        Dim lComment As String
        Dim lArt As String
        Dim lArt2 As String
        Dim PrecioPubDolar As String
        Dim CostoArticulo As String
        Dim Origen As String
        Dim blnTransaccion As Boolean
        gStrSql = "Select SimboloMonedaNac From ConfiguracionGralPv Where COdAlmacen = " & gintCodAlmacenGral
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        SignoPesos = IIf((RsGral.RecordCount > 0), RsGral.Fields("SimboloMonedaNac").Value, "")
        lMsg = ""
        lDESC = ""
        lComment = ""
        lArt = ""
        PrecioPubDolar = CStr(0)
        CostoArticulo = CStr(0)
        Dim Fuente As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
        With tlbCode

            .BarHeight = 185
            .BarWidthReduction = 40
            .PDFAspectRatio = 1
            .NarrowBarWidth = 10
            .NarrowToWideRatio = 2
            '.CommentAlignment = TALBarCode.TALCOMMENTALIGNMENTENUM.bcCenterAlign
            .CommentOnTop = True
            .CtlAutoSize = True
            .CodaBarOptionalCheckDigit = False
            '.Symbology = TALBarCode.TALSYMBOLOGYENUM.bcInterleaved_2of5
            .I2of5OptionalCheckDigit = False
            .ShowHRText = False
            .TextOnTop = False
            ''CODIGO EAN
            Fuente = VB6.FontChangeBold(Fuente, True)
            Fuente = VB6.FontChangeSize(Fuente, 5) '5.6
            Fuente = VB6.FontChangeName(Fuente, "Arial")
            .Font = Fuente
        End With
        Ren = RenInicial
        Col = ColInicial
        'tlbCode.Rotation = TALBarCode.TALROTATEVALSENUM.bcClockwise_180

        CnnAccess.BeginTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        With msgEtiquetas
            gStrSql = "DELETE FROM Etiquetas"
            CmdAccess.CommandText = gStrSql
            CmdAccess.Execute()
            For Y = 1 To .Rows - 1
                If Trim(.get_TextMatrix(Y, C_COLCODIGO)) = "" Then Exit For

                If CDbl(Numerico(.get_TextMatrix(Y, C_COLCODIGO))) = 0 Then Exit For
                lMsg = VB.Right("0000000" & Trim(.get_TextMatrix(Y, C_COLCODIGO)), 7)
                '       asi se hacia antes
                '       lComment = Mid(Trim(.TextMatrix(Y, C_ColDESCRIPCION)) & Space(15), 16, 15)
                '       lArt = Left(Trim(.TextMatrix(Y, C_ColDESCRIPCION)), 15)
                lComment = Trim(.get_TextMatrix(Y, C_COLCODARTPROVEEDOR))
                lArt = VB.Left(Trim(.get_TextMatrix(Y, C_ColDESCRIPCION)), 15)
                lArt2 = Mid(Trim(.get_TextMatrix(Y, C_ColDESCRIPCION)), 16, 15)
                If CBool(.get_TextMatrix(Y, C_ColPESOSFIJOS)) = True Then
                    PrecioPubDolar = Trim(SignoPesos) & .get_TextMatrix(Y, C_ColPRECIOPUBLICO)
                Else
                    PrecioPubDolar = "D" & .get_TextMatrix(Y, C_ColPRECIOPUBLICO)
                End If
                CostoArticulo = Cifrar(CStr(Fix(CInt(Numerico(.get_TextMatrix(Y, C_COLCOSTO))))))
                Origen = Trim(.get_TextMatrix(Y, C_ColORIGEN))
                If optCodActual.Checked = True Then
                    lDESC = Trim(.get_TextMatrix(Y, C_COLCODIGO)) & "-" & Origen
                Else
                    lDESC = Trim(.get_TextMatrix(Y, C_ColCODANTERIOR))
                End If
                gStrSql = "INSERT INTO Etiquetas (CodArticulo,CodigoAnterior,Descripcion,DescripcionA,DescripcionB,Costo,PrecioPublico,Cantidad,blnCodigoAnt,Origen,Usuario,Renglon,Columna,Espacio,CodArtProveedor) " & "VALUES (" & Numerico(.get_TextMatrix(Y, C_COLCODIGO)) & ",'" & IIf(Trim(.get_TextMatrix(Y, C_ColCODANTERIOR)) = "", " ", Trim(.get_TextMatrix(Y, C_ColCODANTERIOR))) & "','" & Trim(.get_TextMatrix(Y, C_ColDESCRIPCION)) & "','" & IIf(Trim(lArt) = "", " ", lArt) & "','" & IIf(Trim(lArt2) = "", " ", lArt2) & "','" & IIf(Trim(CostoArticulo) = "", " ", CostoArticulo) & "','" & IIf(Trim(PrecioPubDolar) = "", " ", PrecioPubDolar) & "'," & .get_TextMatrix(Y, C_ColCANTIDAD) & "," & IIf(optCodActual.Checked = True, 1, 0) & "," & Origen & ",'" & gStrNomUsuario & "'," & RenInicial & "," & ColInicial & "," & Espacio & ",'" & IIf(Trim(lComment) = "", " ", lComment) & "')"
                'CmdAccess.ActiveConnection = CnnAccess
                'BorraCmdAccess
                CmdAccess.CommandText = gStrSql
                CmdAccess.Execute()

                '          With tlbCode
                '             For I = 1 To CCur(Numerico(Trim(msgEtiquetas.TextMatrix(Y, C_ColCANTIDAD))))
                '                .Refresh
                '                'Codigo Ean
                '                .Message = lMsg
                '                .Comment = lComment
                '                .CommentAlignment = bcLeftAlign
                '                Printer.PaintPicture .Picture, Col, Ren
                '                Ren = Ren + 500    '+600
                '                'Descripcion1
                '                Printer.CurrentX = Col
                '                Printer.CurrentY = Ren
                '                Printer.FontName = "Arial"
                '                Printer.FontSize = 5   '5.6
                '                Printer.FontBold = False
                '                Printer.Print lArt
                '
                '                Ren = Ren + 100
                '                'Descripcion2
                '                Printer.CurrentX = Col
                '                Printer.CurrentY = Ren
                '                Printer.FontName = "Arial"
                '                Printer.FontSize = 5   '5.6
                '                Printer.FontBold = False
                '                Printer.Print lArt2
                '
                '                Ren = Ren + 150   '150
                '                'Precio Publico
                '                Printer.CurrentX = Col
                '                Printer.CurrentY = Ren
                '                Printer.FontName = "Arial"
                '                Printer.FontSize = 5.2   '5.8
                '                Printer.FontBold = True
                '                Printer.Print PrecioPubDolar
                '
                '                Ren = Ren + 150     '150
                '                'Precio CLAVE
                '                Printer.CurrentX = Col
                '                Printer.CurrentY = Ren
                '                Printer.FontName = "Arial"
                '                Printer.FontSize = 5.2   '5.8
                '                Printer.FontBold = True
                '                Printer.Print CostoArticulo
                '
                '                Col = Col + 475
                '                Printer.CurrentX = Col
                '                Printer.CurrentY = Ren
                '                Printer.FontName = "Arial"
                '                Printer.FontSize = 5   '5.6
                '                Printer.FontBold = False
                '                Printer.Print Right(Space(10) & lDESC, 10)
                '                Ren = RenInicial
                '                Col = ColInicial
                '                Printer.NewPage
                '             Next I
                '          End With
            Next
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            CnnAccess.CommitTrans()
            blnTransaccion = False
            '    Printer.EndDoc
        End With
Err_Renamed:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If blnTransaccion = True Then CnnAccess.RollbackTrans()
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub CargarFilayColumnaInicial()
        On Error GoTo Merr
        gStrSql = "Select * from CoordenadasEtiqueta"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            RenInicial = RsGral.Fields("Renglon").Value
            ColInicial = RsGral.Fields("Columna").Value
            Espacio = RsGral.Fields("Espacio").Value
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub TotaldeEtiquetas()
        mintTotalEtiquetas = 0
        With msgEtiquetas
            For I = 1 To .Rows - 1
                mintTotalEtiquetas = mintTotalEtiquetas + CDec(Numerico(.get_TextMatrix(I, C_ColCANTIDAD)))
            Next
        End With
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImpresionEtiquetas))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtOrdenCompra = New System.Windows.Forms.TextBox()
        Me.txtDesArticulo = New System.Windows.Forms.Label()
        Me._Marco_141 = New System.Windows.Forms.GroupBox()
        Me.fraOrdenamiento = New System.Windows.Forms.GroupBox()
        Me.optCodActual = New System.Windows.Forms.RadioButton()
        Me.optCodAnterior = New System.Windows.Forms.RadioButton()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.optOrdenCompra = New System.Windows.Forms.RadioButton()
        Me.optArticulos = New System.Windows.Forms.RadioButton()
        Me.txtDetalle = New System.Windows.Forms.TextBox()
        Me.msgEtiquetas = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.CommonOpen = New System.Windows.Forms.OpenFileDialog()
        Me.CommonSave = New System.Windows.Forms.SaveFileDialog()
        Me.CommonFont = New System.Windows.Forms.FontDialog()
        Me.CommonColor = New System.Windows.Forms.ColorDialog()
        Me.CommonPrint = New System.Windows.Forms.PrintDialog()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Marco = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me._Marco_141.SuspendLayout()
        Me.fraOrdenamiento.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.msgEtiquetas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Marco, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtOrdenCompra
        '
        Me.txtOrdenCompra.AcceptsReturn = True
        Me.txtOrdenCompra.BackColor = System.Drawing.SystemColors.Window
        Me.txtOrdenCompra.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOrdenCompra.Enabled = False
        Me.txtOrdenCompra.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOrdenCompra.Location = New System.Drawing.Point(136, 36)
        Me.txtOrdenCompra.MaxLength = 19
        Me.txtOrdenCompra.Name = "txtOrdenCompra"
        Me.txtOrdenCompra.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOrdenCompra.Size = New System.Drawing.Size(126, 20)
        Me.txtOrdenCompra.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtOrdenCompra, "Folio de la Orden de Compra")
        '
        'txtDesArticulo
        '
        Me.txtDesArticulo.BackColor = System.Drawing.SystemColors.Info
        Me.txtDesArticulo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDesArticulo.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDesArticulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtDesArticulo.Location = New System.Drawing.Point(8, 280)
        Me.txtDesArticulo.Name = "txtDesArticulo"
        Me.txtDesArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDesArticulo.Size = New System.Drawing.Size(512, 21)
        Me.txtDesArticulo.TabIndex = 10
        Me.txtDesArticulo.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolTip1.SetToolTip(Me.txtDesArticulo, "Descripción de Artículos")
        '
        '_Marco_141
        '
        Me._Marco_141.BackColor = System.Drawing.SystemColors.Control
        Me._Marco_141.Controls.Add(Me.fraOrdenamiento)
        Me._Marco_141.Controls.Add(Me.Frame1)
        Me._Marco_141.Controls.Add(Me.txtDetalle)
        Me._Marco_141.Controls.Add(Me.msgEtiquetas)
        Me._Marco_141.Controls.Add(Me.txtDesArticulo)
        Me._Marco_141.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Marco_141.Location = New System.Drawing.Point(8, 8)
        Me._Marco_141.Name = "_Marco_141"
        Me._Marco_141.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Marco_141.Size = New System.Drawing.Size(547, 313)
        Me._Marco_141.TabIndex = 0
        Me._Marco_141.TabStop = False
        '
        'fraOrdenamiento
        '
        Me.fraOrdenamiento.BackColor = System.Drawing.SystemColors.Control
        Me.fraOrdenamiento.Controls.Add(Me.optCodActual)
        Me.fraOrdenamiento.Controls.Add(Me.optCodAnterior)
        Me.fraOrdenamiento.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraOrdenamiento.Location = New System.Drawing.Point(335, 13)
        Me.fraOrdenamiento.Name = "fraOrdenamiento"
        Me.fraOrdenamiento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraOrdenamiento.Size = New System.Drawing.Size(201, 67)
        Me.fraOrdenamiento.TabIndex = 5
        Me.fraOrdenamiento.TabStop = False
        Me.fraOrdenamiento.Text = " Código impreso  "
        '
        'optCodActual
        '
        Me.optCodActual.BackColor = System.Drawing.SystemColors.Control
        Me.optCodActual.Checked = True
        Me.optCodActual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCodActual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCodActual.Location = New System.Drawing.Point(72, 22)
        Me.optCodActual.Name = "optCodActual"
        Me.optCodActual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCodActual.Size = New System.Drawing.Size(73, 15)
        Me.optCodActual.TabIndex = 6
        Me.optCodActual.TabStop = True
        Me.optCodActual.Text = "Actual"
        Me.optCodActual.UseVisualStyleBackColor = False
        '
        'optCodAnterior
        '
        Me.optCodAnterior.BackColor = System.Drawing.SystemColors.Control
        Me.optCodAnterior.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCodAnterior.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCodAnterior.Location = New System.Drawing.Point(72, 43)
        Me.optCodAnterior.Name = "optCodAnterior"
        Me.optCodAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCodAnterior.Size = New System.Drawing.Size(73, 18)
        Me.optCodAnterior.TabIndex = 7
        Me.optCodAnterior.TabStop = True
        Me.optCodAnterior.Text = "Anterior"
        Me.optCodAnterior.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtOrdenCompra)
        Me.Frame1.Controls.Add(Me.optOrdenCompra)
        Me.Frame1.Controls.Add(Me.optArticulos)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(9, 13)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(283, 67)
        Me.Frame1.TabIndex = 1
        Me.Frame1.TabStop = False
        '
        'optOrdenCompra
        '
        Me.optOrdenCompra.BackColor = System.Drawing.SystemColors.Control
        Me.optOrdenCompra.Cursor = System.Windows.Forms.Cursors.Default
        Me.optOrdenCompra.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optOrdenCompra.Location = New System.Drawing.Point(24, 39)
        Me.optOrdenCompra.Name = "optOrdenCompra"
        Me.optOrdenCompra.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optOrdenCompra.Size = New System.Drawing.Size(112, 19)
        Me.optOrdenCompra.TabIndex = 3
        Me.optOrdenCompra.TabStop = True
        Me.optOrdenCompra.Text = "Orden de Compra"
        Me.optOrdenCompra.UseVisualStyleBackColor = False
        '
        'optArticulos
        '
        Me.optArticulos.BackColor = System.Drawing.SystemColors.Control
        Me.optArticulos.Checked = True
        Me.optArticulos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optArticulos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optArticulos.Location = New System.Drawing.Point(24, 17)
        Me.optArticulos.Name = "optArticulos"
        Me.optArticulos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optArticulos.Size = New System.Drawing.Size(73, 17)
        Me.optArticulos.TabIndex = 2
        Me.optArticulos.TabStop = True
        Me.optArticulos.Text = "Artículos"
        Me.optArticulos.UseVisualStyleBackColor = False
        '
        'txtDetalle
        '
        Me.txtDetalle.AcceptsReturn = True
        Me.txtDetalle.BackColor = System.Drawing.SystemColors.Window
        Me.txtDetalle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDetalle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDetalle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDetalle.Location = New System.Drawing.Point(9, 111)
        Me.txtDetalle.MaxLength = 0
        Me.txtDetalle.Name = "txtDetalle"
        Me.txtDetalle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDetalle.Size = New System.Drawing.Size(63, 20)
        Me.txtDetalle.TabIndex = 9
        Me.txtDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDetalle.Visible = False
        '
        'msgEtiquetas
        '
        Me.msgEtiquetas.DataSource = Nothing
        Me.msgEtiquetas.Location = New System.Drawing.Point(8, 88)
        Me.msgEtiquetas.Name = "msgEtiquetas"
        Me.msgEtiquetas.OcxState = CType(resources.GetObject("msgEtiquetas.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msgEtiquetas.Size = New System.Drawing.Size(529, 184)
        Me.msgEtiquetas.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(368, 248)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(81, 33)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Label2"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(127, 334)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 94
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(12, 334)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 93
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(242, 335)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 92
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmImpresionEtiquetas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(567, 382)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me._Marco_141)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Label2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(139, 202)
        Me.MaximizeBox = False
        Me.Name = "frmImpresionEtiquetas"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Impresión de Etiquetas"
        Me._Marco_141.ResumeLayout(False)
        Me._Marco_141.PerformLayout()
        Me.fraOrdenamiento.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.msgEtiquetas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Marco, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

End Class