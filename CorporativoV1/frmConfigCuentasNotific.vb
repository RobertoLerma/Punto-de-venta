Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmConfigCuentasNotific
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdAceptar As System.Windows.Forms.Button
    Public WithEvents cmbEstatus As System.Windows.Forms.ComboBox
    Public WithEvents txtDetalle As System.Windows.Forms.TextBox
    Public WithEvents Flexdetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents FlexBorrados As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    'Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents lblMsg As System.Windows.Forms.Label

    Const C_COLCUENTA As Integer = 0
    Const C_COLESTADO As Integer = 1
    Const C_COLCODIGO As Integer = 2
    Const C_COLCUENTATAG As Integer = 3
    Const C_COLESTATUS As Integer = 4
    Const C_MaxPartidas As Integer = 10

    Const C_COLCODIGOB As Integer = 1
    Const C_COLESTATUSB As Integer = 2

    Dim mblnSalir As Boolean
    Public mblnCambios As Boolean
    Public mintBorrados As Integer
    Dim ColumnaEnUso As Integer
    Dim FueraChange As Boolean
    Friend WithEvents Panel1 As Panel
    Public WithEvents btnGuardar As Button
    Dim mintRenglonB As Integer

    Private Sub cmbEstatus_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbEstatus.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbEstatus_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmbEstatus.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Aqui se muestran los datos del control editable, en el Grid
        'Se deberá formatear el Valor de Acuerdo al Tipo de Dato en uso
        Dim rowsiguiente As Integer
        Dim ColSiguiente As Integer

        With Flexdetalle
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    'txtDetalle.Visible = False
                    'txtDetalle.Text = ""
                    cmbEstatus.Visible = False
                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone

                Case System.Windows.Forms.Keys.Return
                    'Si la Columna en que se está escribiendo es Codigo o Cantidad, Formatear el Valor par que quede numérico
                    If .Col = C_COLESTADO Then
                        .set_TextMatrix(.Row, .Col, Trim(cmbEstatus.Text))

                        frmConfigGralCorporativo.mintCuentas = frmConfigGralCorporativo.mintCuentas + 1
                        FueraChange = True
                        cmbEstatus.Visible = False
                        Flexdetalle.Col = .Col
                        Flexdetalle.Row = .Row
                        .Focus()

                        rowsiguiente = .Row + 1
                        ColSiguiente = C_COLCUENTA
                        .Row = rowsiguiente
                        .Col = ColSiguiente
                    End If
            End Select
        End With
    End Sub

    Private Sub cmbEstatus_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmbEstatus.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub cmbEstatus_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbEstatus.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        'Aqui se Valida que los datos que se introdujeron en el Control Aux (txtDetalle), sea válido, dependiendo de la columna en que se está posicionado, y que se esté editando dicha columna
        'Para el Caso de la Columna de Codigo, al perder el foco el control, se mostrará los datos del Articulo
        cmbEstatus_KeyDown(cmbEstatus, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
        cmbEstatus.Visible = False
    End Sub

    Private Sub cmdAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAceptar.Click
        '                                 Nuevo     Guardar       Cancelar       Eliminar         Buscar       Imprimir      Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        If ValidaDatos() Then
            Cambios()
            Me.Hide()
            'frmConfigCuentasNotific_FormClosed(New Object, New EventArgs)
        End If

    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Flexdetalle.DblClick
        FlexDetalle_KeyPressEvent(Flexdetalle, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Flexdetalle.Enter

    End Sub

    Private Sub FlexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles Flexdetalle.KeyDownEvent
        Dim RowActual As Integer
        'Aqui debe cvalidarse el movimiento de teclas, si es ke se va a tomar en cuenta
        With Flexdetalle
            Select Case eventArgs.keyCode
                Case System.Windows.Forms.Keys.Delete
                    RowActual = .Row
                    'Si el cursor está en en renglón 0, que es el nombre de columna, entonces no se toma en cuenta la tecla Supr
                    If .Row = 0 Then Exit Sub
                    'Validar si s Posible Eliminar una Línea Capturada
                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                    Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "Mensaje")
                        Case MsgBoxResult.No
                            'Poner el setfocus en el grid, para que siga dentro del mismo
                            Flexdetalle.Focus()
                            Exit Sub
                        Case MsgBoxResult.Cancel
                            'Poner el setfocus en el grid, para que siga dentro del mismo
                            Flexdetalle.Focus()
                            Exit Sub
                    End Select
                    PasaElementoBorrado(.Row)
                    BorraGrid(.Row) 'Cuando se Borra, se obtienen los nuevos totales, la cntidad de Articulos (Dento del proc.)
                    'Poner el Foco en la Primera linea del Grid
                    .Col = C_COLCUENTA
                    .Row = RowActual
                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                    .Focus()
                Case System.Windows.Forms.Keys.Insert
                    .AddItem("", .Row)
            End Select
        End With
    End Sub

    Private Sub FlexDetalle_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles Flexdetalle.KeyPressEvent
        'En este Evento, Se muestra el Control para editar el Grid, con los datos que ya tiene el grid
        'Las Columnas Editables son: Codigo, DEscripcion, Cantidad,Descuento
        'Si no es nuevo, no podrá editarse nada en el Grid
        Dim ColSiguiente As Integer
        Dim rowsiguiente As Integer

        With Flexdetalle
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then 'Para que cuando sea escape, no entre a editar el codigo,simplemente que se regrese al control anterior
                Select Case .Col
                    Case C_COLCUENTA ''-------------- SE EDITA LA CUENTA ---------------------'''''
                        txtDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                        '''en esta parte se validará si es el rengón, columna que le corresponde editarse
                        If (.Row > 1) Then
                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
                            If Trim(.get_TextMatrix(.Row - 1, C_COLCUENTA)) = "" Then
                                .Focus()
                                Exit Sub
                            End If
                        End If
                        ModEstandar.MSHFlexGridEdit(Flexdetalle, txtDetalle, eventArgs.keyAscii)
                        If Len(Trim(txtDetalle.Text)) <> 1 Then ModEstandar.SelTextoTxt(txtDetalle)
                        ColumnaEnUso = C_COLCUENTA

                    Case C_COLESTADO ''-------------- SE EDITA LA CUENTA ---------------------'''''
                        '''En esta parte se validará si es el rengón, columna que le corresponde editarse
                        If (.Row > 1) Then
                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
                            If Trim(.get_TextMatrix(.Row - 1, C_COLCUENTA)) = "" Then
                                .Focus()
                                Exit Sub
                            End If
                        End If
                        ModEstandar.MSHFlexGridEdit(Flexdetalle, cmbEstatus, eventArgs.keyAscii)
                        cmbEstatus.Text = ""
                        ColumnaEnUso = C_COLESTADO

                End Select
            End If
        End With
    End Sub

    Private Sub FlexDetalle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Flexdetalle.Leave
        Flexdetalle.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub

    Private Sub FlexDetalle_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Flexdetalle.Scroll
        txtDetalle.Visible = False
        cmbEstatus.Visible = False
    End Sub

    Private Sub txtDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetalle.Enter
        Pon_Tool()
        Flexdetalle.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
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
        With Flexdetalle
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    'txtDetalle.Visible = False
                    'txtDetalle.Text = ""
                    cmbEstatus.Visible = False
                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone

                Case System.Windows.Forms.Keys.Return
                    'Si la Columna en que se está escribiendo es Codigo o Cantidad, Formatear el Valor par que quede numérico
                    If .Col = C_COLCUENTA Then .set_TextMatrix(.Row, .Col, Trim(txtDetalle.Text))
                    FueraChange = True
                    txtDetalle.Text = ""
                    txtDetalle.Visible = False
                    Flexdetalle.Col = .Col
                    Flexdetalle.Row = .Row
                    .Focus()

                    If .Col = C_COLCUENTA Then
                        If Trim(.get_TextMatrix(.Row, C_COLCUENTA)) = "" Then
                            LimpiaDatosArticulo()
                            Exit Sub
                        End If

                        If CuentaRepetidaenGrid(Trim(.get_TextMatrix(.Row, C_COLCUENTA))) = True Then
                            .set_TextMatrix(.Row, C_COLCODIGO, Trim(.get_TextMatrix(.Row, C_COLCUENTA)))
                            'If TextMatrix(.Row, C_COLESTATUS) = "" Then
                            '   .TextMatrix(.Row, C_COLESTATUS) = "N"
                            'ElseIf Trim(.TextMatrix(.Row, C_COLESTATUS)) = "A" Then
                            '   .TextMatrix(.Row, C_COLESTATUS) = "M"
                            'End If
                            MsgBox("Existe una cuenta repetida en el grid de captura" & vbNewLine & "Verifique por favor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                            LimpiaDatosArticulo()
                            .Focus()
                            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                            Exit Sub
                        End If

                        AgregarFilaFinal()
                        'rowsiguiente = .Row + 1
                        ColSiguiente = C_COLESTADO
                        '.Row = rowsiguiente
                        .Col = ColSiguiente
                    End If
            End Select
        End With
    End Sub

    Private Sub txtDetalle_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDetalle.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'En este Evento se validan los datos que se introduzcan al control txtDetalle,dependiendo de la columan en que se esté editando
        If KeyAscii = 0 Or KeyAscii = System.Windows.Forms.Keys.Return Then GoTo EventExitSub
        With Flexdetalle
            If .Col = C_COLCUENTA Then txtDetalle.MaxLength = 255
        End With
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDetalle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetalle.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        'Aqui se Valida que los datos que se introdujeron en el Control Aux (txtDetalle), sea válido, dependiendo de la columna en que se está posicionado, y que se esté editando dicha columna
        'Para el Caso de la Columna de Codigo, al perder el foco el control, se mostrará los datos del Articulo
        txtDetalle_KeyDown(txtDetalle, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
        txtDetalle.Visible = False
        cmbEstatus.Visible = False
    End Sub

    Private Sub frmConfigCuentasNotific_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                                  Nuevo        Guardar       Cancelar       Eliminar         Buscar       Imprimir         Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmConfigCuentasNotific_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate, cmdAceptar.Click
        'Desactivar todas las opciones del Menu
        '                                  Nuevo        Guardar       Cancelar       Eliminar         Buscar       Imprimir      Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmConfigCuentasNotific_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                                 Nuevo        Guardar       Cancelar       Eliminar         Buscar       Imprimir         Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        If mintBorrados = 0 Then
            Encabezado()
            EncabezadoB()
            InicializaVariables()
            LlenaDatos()
        End If
    End Sub

    Private Sub frmConfigCuentasNotific_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        ' En este evento del formulario se valida la tecla presionada.
        ' Si es Enter se simula un tab(Avanza al siguiente control)
        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
            Case System.Windows.Forms.Keys.F2
                Encabezado()
                EncabezadoB()
                InicializaVariables()
                LlenaDatos()
        End Select
    End Sub

    Private Sub frmConfigCuentasNotific_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmConfigCuentasNotific_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''''    If Not mblnSalir Then
        ''''        'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        ''''        ModEstandar.RestaurarForma Me, False
        ''''        'Si se cierra el formulario y existio algun cambio en el registro se
        ''''        'informa al usuario del cabio y si desea guardar el registro, ya sea
        ''''        'que sea nuevo o un registro modificado
        ''''        If Cambios = True Then
        ''''            Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
        ''''                Case vbYes: 'Guardar el registro
        ''''                    If Guardar = False Then
        ''''                        Cancel = 1
        ''''                    End If
        ''''                Case vbNo: 'No hace nada y permite el cierre del formulario
        ''''                Case vbCancel: 'Cancela el cierre del formulario sin guardar
        ''''                    Cancel = 1
        ''''            End Select
        ''''        End If
        ''''    Else 'Se quiere salir con escape
        ''''        mblnSalir = False
        ''''        Select Case MsgBox(C_msgSALIR, vbQuestion + vbYesNo + vbDefaultButton2, gstrNombCortoEmpresa)
        ''''            Case vbYes: 'Sale del Formulario
        ''''                Cancel = 0
        ''''            Case vbNo: 'No sale del formulario
        ''''                Cancel = 1
        ''''        End Select
        ''''    End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmConfigCuentasNotific_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub InicializaVariables()
        mintRenglonB = 1
        mblnCambios = False
        mintBorrados = 0
    End Sub

    Private Sub LlenaDatos()
        Dim I As Integer
        On Error GoTo MErr

        frmConfigGralCorporativo.mintCuentas = 0
        gStrSql = "Select lTrim(rtrim(CuentaCorreo)) as CuentaCorreo, Case When Estatus = 'A' Then 'Activo' Else 'Suspendido' End as Estado, CodCuentaC, lTrim(rtrim(CuentaCorreo)) as CuentaTag, 'A' as EstatusEst From CatCuentasNotificaciones Order by CodCuentaC "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            Flexdetalle.Recordset = RsGral
            If RsGral.RecordCount < C_MaxPartidas Then Flexdetalle.Rows = C_MaxPartidas + 1
            With Flexdetalle
                .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            End With
            frmConfigGralCorporativo.mintCuentas = RsGral.RecordCount
        End If

MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub Encabezado()
        Dim LnContador As Integer

        With Flexdetalle
            .Clear()

            '''.FocusRect = flexFocusLight
            .set_Cols(0, 5)
            .WordWrap = True
            .FixedRows = 1
            .FixedCols = 0
            .set_RowHeight(0, 500)

            .set_ColWidth(C_COLCUENTA, 0, 4500)
            .set_ColWidth(C_COLESTADO, 0, 1000)
            .set_ColWidth(C_COLCODIGO, 0, 0)
            .set_ColWidth(C_COLCUENTATAG, 0, 0)
            .set_ColWidth(C_COLESTATUS, 0, 0)

            .set_TextMatrix(0, C_COLCUENTA, "CUENTA")
            .set_TextMatrix(0, C_COLESTADO, "ESTATUS")
            .set_TextMatrix(0, C_COLCODIGO, "CODIGO")
            .set_TextMatrix(0, C_COLCUENTATAG, "CUENTATAG")
            .set_TextMatrix(0, C_COLESTATUS, "ESTATUS_EST")

            .Row = 0
            For LnContador = 0 To C_COLESTATUS
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter
                .CellFontBold = True
            Next LnContador
            .Row = 1
            .Col = C_COLCUENTA
            .WordWrap = False 'Hacer esto , para que no se puedan escribir dos o mal lineas de texto en una  sola fila, solo se usa para el encabezado
        End With

    End Sub

    Private Sub EncabezadoB()
        Dim LnContador As Integer

        With FlexBorrados
            .Clear()

            '''.FocusRect = flexFocusLight
            .set_Cols(0, 3)
            .WordWrap = True
            .FixedRows = 1
            .FixedCols = 0
            .set_RowHeight(0, 500)

            .set_ColWidth(C_COLCUENTA, 0, 1000)
            .set_ColWidth(C_COLCODIGOB, 0, 1000)
            .set_ColWidth(C_COLESTATUSB, 0, 500)

            .set_TextMatrix(0, C_COLCUENTA, "CUENTA")
            .set_TextMatrix(0, C_COLCODIGOB, "CODIGO")
            .set_TextMatrix(0, C_COLESTATUSB, "ESTATUS")

            .Row = 0
            For LnContador = 0 To C_COLCODIGOB
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter
                .CellFontBold = True
            Next LnContador
            .Row = 1
            .Col = C_COLCUENTA
            .WordWrap = False 'Hacer esto , para que no se puedan escribir dos o mal lineas de texto en una  sola fila, solo se usa para el encabezado
        End With

    End Sub

    Private Function ValidaDatos() As Boolean
        Dim I As Integer

        With Flexdetalle
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, C_COLCUENTA) = "" And .get_TextMatrix(I, C_COLESTADO) = "" Then Exit For

                If .get_TextMatrix(I, C_COLCUENTA) = "" And Trim(.get_TextMatrix(I, C_COLESTADO)) <> "" Then
                    MsgBox("Debe especificar la cuenta de correo a registrar," & vbNewLine & "o borrar el registro", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                    .Col = 0
                    .Row = I
                    .Focus()
                    ValidaDatos = False
                    Exit Function
                End If

                If .get_TextMatrix(I, C_COLCUENTA) <> "" And Trim(.get_TextMatrix(I, C_COLESTADO)) = "" Then
                    MsgBox("Debe especificar el estatus de la cuenta de correo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                    .Col = 0
                    .Row = I
                    .Focus()
                    ValidaDatos = False
                    Exit Function
                End If

            Next I
        End With
        ValidaDatos = True

    End Function

    Public Function Cambios() As Boolean
        Dim I As Integer

        With Flexdetalle
            For I = 1 To .Rows - 1
                If IsNothing(.get_TextMatrix(I, C_COLCUENTA)) Then Exit For
                If .get_TextMatrix(I, C_COLESTATUS) = C_ELIMINADO Or .get_TextMatrix(I, C_COLESTATUS) = C_ACTIVO Then
                    'No hace nada
                ElseIf (Trim(.get_TextMatrix(I, C_COLCUENTA)) <> Trim(.get_TextMatrix(I, C_COLCUENTATAG)) And (.get_TextMatrix(I, C_COLCODIGO) <> "")) Then
                    .set_TextMatrix(I, C_COLESTATUS, C_MODIFICADO)
                    Cambios = True
                ElseIf Trim(.get_TextMatrix(I, C_COLCUENTA)) <> Trim(.get_TextMatrix(I, C_COLCUENTATAG)) And (.get_TextMatrix(I, C_COLCODIGO) = "") Then
                    .set_TextMatrix(I, C_COLESTATUS, C_NUEVO)
                    Cambios = True
                End If
            Next I
        End With
        If mintBorrados > 0 Then Cambios = True

        mblnCambios = Cambios
        Return mblnCambios
    End Function

    Public Function Guardar() As Boolean
        On Error GoTo MErr
        Dim blnTransaction As Boolean
        Dim I As Integer

        If txtDetalle.Visible Then txtDetalle_KeyDown(txtDetalle, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Return Or 0 * &H10000))
        If Not ValidaDatos() Then Exit Function
        If Not Cambios() Then Exit Function

        I = 0
        For I = 1 To Flexdetalle.Rows - 1
            '''LOS POSIBLES MODIFICADOS
            With Flexdetalle
                If (Trim(.get_TextMatrix(I, C_COLCUENTA)) <> "" And Trim(.get_TextMatrix(I, C_COLESTATUS)) <> "") Then
                    Select Case Trim(.get_TextMatrix(I, C_COLESTATUS)) '''Si es A - Activo no le hizo nada
                        Case C_MODIFICADO
                            ModStoredProcedures.PR_IME_CatCuentasNotifiaciones(.get_TextMatrix(I, C_COLCODIGO), .get_TextMatrix(I, C_COLCUENTA), IIf(Trim(.get_TextMatrix(I, C_COLESTADO)) = "Activo", "A", "S"), "M")
                            Cmd.Execute()
                        Case C_NUEVO
                            ModStoredProcedures.PR_IME_CatCuentasNotifiaciones(.get_TextMatrix(I, C_COLCODIGO), .get_TextMatrix(I, C_COLCUENTA), IIf(Trim(.get_TextMatrix(I, C_COLESTADO)) = "Activo", "A", "S"), "I")
                            Cmd.Execute()
                        Case C_ELIMINADO
                            ModStoredProcedures.PR_IME_CatCuentasNotifiaciones(.get_TextMatrix(I, C_COLCODIGO), .get_TextMatrix(I, C_COLCUENTA), IIf(Trim(.get_TextMatrix(I, C_COLESTADO)) = "Activo", "A", "S"), "E")
                            Cmd.Execute()
                    End Select
                End If
            End With
            '''LOS BORRADOS
            With FlexBorrados
                If (Trim(.get_TextMatrix(I, C_COLCUENTA)) <> "" And Trim(.get_TextMatrix(I, C_COLESTATUSB)) <> "") Then
                    Select Case Trim(.get_TextMatrix(I, C_COLESTATUSB))
                        Case C_ELIMINADO
                            ModStoredProcedures.PR_IME_CatCuentasNotifiaciones(Trim(.get_TextMatrix(I, C_COLCODIGOB)), Trim(.get_TextMatrix(I, C_COLCUENTA)), IIf(Trim(.get_TextMatrix(I, C_COLESTADO)) = "Activo", "A", "S"), "E")
                            Cmd.Execute()
                    End Select
                End If
            End With
        Next I
        Flexdetalle.TopRow = 1
        Flexdetalle.Col = 0
        Flexdetalle.Row = 1
        Guardar = True
        MsgBox(C_msgACTUALIZADO, vbInformation + vbOKOnly, ModVariables.gstrNombCortoEmpresa)
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub BorraGrid(ByRef Row As Integer)
        'Este Procediento borra un renglon del Grid
        'Si el Número de Filas que kedan en el grid, es menor de 8, se insertará una nueva fila al final del grid
        With Flexdetalle
            .RemoveItem(Row)
            'Si el número de filas es menor de 10 o esta posicionado en la utlima fila, entonces, agrega una fila
            If .Rows < 11 Or .Row = .Rows - 1 Then
                .AddItem("")
                .Row = .Row
            End If
        End With
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
        With Flexdetalle
            For I = 1 To .Rows
                If Trim(.get_TextMatrix(I, C_COLCUENTA)) <> "" Then
                    nCont = nCont + 1
                Else
                    Exit For
                End If
            Next I
            If nCont < 7 Then
                'Hay menos de 7 registros
                '            .TopRow = 7
                .Row = nCont + 1
                .Col = C_COLCUENTA

            Else
                'Hay 7 ó más registros, hay que recorrer el grid
                .TopRow = (nCont - nRen) + 2
                .Row = nCont + 1
                .Col = C_COLCUENTA
            End If
        End With
    End Sub

    Sub LimpiaDatosArticulo()
        With Flexdetalle
            .set_TextMatrix(.Row, C_COLCUENTA, "")
            .set_TextMatrix(.Row, C_COLESTADO, "")
            .set_TextMatrix(.Row, C_COLCODIGO, "")
            .set_TextMatrix(.Row, C_COLCUENTATAG, "")
            .set_TextMatrix(.Row, C_COLESTATUS, "")
        End With
    End Sub

    Function CuentaRepetidaenGrid(ByRef lCuenta As String) As Boolean
        Dim lRepetida As Integer
        Dim I As Integer

        lRepetida = 0
        CuentaRepetidaenGrid = False
        With Flexdetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCUENTA)) = lCuenta Then
                    lRepetida = lRepetida + 1
                    If lRepetida > 1 Then
                        CuentaRepetidaenGrid = True
                        Exit For
                    End If
                End If
            Next
        End With
        Exit Function

    End Function

    Sub AgregarFilaFinal()
        'Agrea una Fila al Final del Grid, Sólo cuando sea necesario hacerlo
        With Flexdetalle
            If .Row = .Rows - 1 And Trim(.get_TextMatrix(.Row, C_COLCUENTA)) <> "" Then
                ' Si se Presiono enter y estamos en la ultima fila, entonces se agregrara una nueva fila
                .AddItem("")
                FlexBorrados.AddItem("")
                ScrollGrid()
            End If
        End With
    End Sub

    Private Sub PasaElementoBorrado(ByRef lRen As Integer)
        If Trim(Flexdetalle.get_TextMatrix(lRen, C_COLESTATUS)) <> "" And Trim(Flexdetalle.get_TextMatrix(lRen, C_COLESTATUS)) <> "N" Then
            With FlexBorrados
                .set_TextMatrix(mintRenglonB, C_COLCUENTA, Trim(Flexdetalle.get_TextMatrix(lRen, C_COLCUENTA)))
                .set_TextMatrix(mintRenglonB, C_COLCODIGOB, Trim(Flexdetalle.get_TextMatrix(lRen, C_COLCODIGO)))
                .set_TextMatrix(mintRenglonB, C_COLESTATUSB, "E")
            End With
            mintRenglonB = mintRenglonB + 1
            mintBorrados = mintBorrados + 1
            frmConfigGralCorporativo.mintCuentas = frmConfigGralCorporativo.mintCuentas - 1
        End If
    End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfigCuentasNotific))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmbEstatus = New System.Windows.Forms.ComboBox()
        Me.txtDetalle = New System.Windows.Forms.TextBox()
        Me.cmdAceptar = New System.Windows.Forms.Button()
        Me.Flexdetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.FlexBorrados = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.lblMsg = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnGuardar = New System.Windows.Forms.Button()
        CType(Me.Flexdetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FlexBorrados, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbEstatus
        '
        Me.cmbEstatus.BackColor = System.Drawing.SystemColors.Window
        Me.cmbEstatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbEstatus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbEstatus.Items.AddRange(New Object() {"Activo", "Suspendido"})
        Me.cmbEstatus.Location = New System.Drawing.Point(84, 23)
        Me.cmbEstatus.Name = "cmbEstatus"
        Me.cmbEstatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbEstatus.Size = New System.Drawing.Size(66, 21)
        Me.cmbEstatus.TabIndex = 3
        Me.cmbEstatus.Text = "cmbEstatus"
        Me.ToolTip1.SetToolTip(Me.cmbEstatus, "Estatus")
        Me.cmbEstatus.Visible = False
        '
        'txtDetalle
        '
        Me.txtDetalle.AcceptsReturn = True
        Me.txtDetalle.BackColor = System.Drawing.SystemColors.Window
        Me.txtDetalle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDetalle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDetalle.Location = New System.Drawing.Point(19, 24)
        Me.txtDetalle.MaxLength = 0
        Me.txtDetalle.Name = "txtDetalle"
        Me.txtDetalle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDetalle.Size = New System.Drawing.Size(64, 20)
        Me.txtDetalle.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtDetalle, "Registro de cuenta de correo")
        Me.txtDetalle.Visible = False
        '
        'cmdAceptar
        '
        Me.cmdAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAceptar.Location = New System.Drawing.Point(353, 275)
        Me.cmdAceptar.Name = "cmdAceptar"
        Me.cmdAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAceptar.Size = New System.Drawing.Size(108, 31)
        Me.cmdAceptar.TabIndex = 6
        Me.cmdAceptar.Text = "&Aceptar"
        Me.cmdAceptar.UseVisualStyleBackColor = False
        '
        'Flexdetalle
        '
        Me.Flexdetalle.DataSource = Nothing
        Me.Flexdetalle.Location = New System.Drawing.Point(19, 12)
        Me.Flexdetalle.Name = "Flexdetalle"
        Me.Flexdetalle.OcxState = CType(resources.GetObject("Flexdetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Flexdetalle.Size = New System.Drawing.Size(400, 224)
        Me.Flexdetalle.TabIndex = 1
        '
        'FlexBorrados
        '
        Me.FlexBorrados.DataSource = Nothing
        Me.FlexBorrados.Location = New System.Drawing.Point(19, 12)
        Me.FlexBorrados.Name = "FlexBorrados"
        Me.FlexBorrados.OcxState = CType(resources.GetObject("FlexBorrados.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FlexBorrados.Size = New System.Drawing.Size(400, 224)
        Me.FlexBorrados.TabIndex = 5
        '
        'lblMsg
        '
        Me.lblMsg.BackColor = System.Drawing.SystemColors.Control
        Me.lblMsg.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMsg.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblMsg.Location = New System.Drawing.Point(13, 309)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMsg.Size = New System.Drawing.Size(422, 16)
        Me.lblMsg.TabIndex = 4
        Me.lblMsg.Text = "<F2>  Cargar Datos                    <Supr>     Eliminar renglón"
        Me.lblMsg.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmbEstatus)
        Me.Panel1.Controls.Add(Me.txtDetalle)
        Me.Panel1.Controls.Add(Me.Flexdetalle)
        Me.Panel1.Controls.Add(Me.FlexBorrados)
        Me.Panel1.Location = New System.Drawing.Point(16, 25)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(445, 244)
        Me.Panel1.TabIndex = 7
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(16, 275)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(108, 31)
        Me.btnGuardar.TabIndex = 8
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmConfigCuentasNotific
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(472, 332)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdAceptar)
        Me.Controls.Add(Me.lblMsg)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(254, 164)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmConfigCuentasNotific"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Cuentas de Notificaciones"
        CType(Me.Flexdetalle, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FlexBorrados, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub
End Class