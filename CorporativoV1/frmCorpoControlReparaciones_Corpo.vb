Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCorpoControlReparaciones_Corpo
    Inherits System.Windows.Forms.Form
    ''' ********************************************************************************************************************
    ''' Programa: Control de Reparaciones
    ''' Autor: Rosaura Torres López
    ''' Fecha de Creación: 01/Julio/2003
    '''
    ''' MODIFICACION DE LA FUNCION DE ESTATUS DE REPARACIONES - SE AGREGARON PARAMETROS
    ''' 15SEP2006 - MAVF
    ''' ********************************************************************************************************************


    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdRegistrar As System.Windows.Forms.Button
    Public WithEvents txtNuevoTexto As System.Windows.Forms.TextBox
    Public WithEvents fraNuevoTexto As System.Windows.Forms.GroupBox
    Public WithEvents Bitacora As System.Windows.Forms.RichTextBox
    Public WithEvents Frame3 As System.Windows.Forms.Panel
    Public WithEvents fraBitacora As System.Windows.Forms.GroupBox
    Public WithEvents txtDomicilio As System.Windows.Forms.TextBox
    Public WithEvents txtTelefono As System.Windows.Forms.TextBox
    Public WithEvents txtRFCCliente As System.Windows.Forms.TextBox
    Public WithEvents txtCodCliente As System.Windows.Forms.TextBox
    Public WithEvents txtCliente As System.Windows.Forms.TextBox
    Public WithEvents _lblVentas_4 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_3 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_2 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents fraDatosCliente As System.Windows.Forms.GroupBox
    Public WithEvents btnCatClientes As System.Windows.Forms.Button
    Public WithEvents chkConfirmacionCliente As System.Windows.Forms.CheckBox
    Public WithEvents chkEntregaCliente As System.Windows.Forms.CheckBox
    Public WithEvents dtpConfirmacionCliente As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaEntregaCliente As System.Windows.Forms.DateTimePicker
    Public WithEvents fraCliente As System.Windows.Forms.GroupBox
    Public WithEvents chkTallerEntrega As System.Windows.Forms.CheckBox
    Public WithEvents chkTallerRegreso As System.Windows.Forms.CheckBox
    Public WithEvents chkReparado As System.Windows.Forms.CheckBox
    Public WithEvents dtpEntregaTaller As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpRegresoTaller As System.Windows.Forms.DateTimePicker
    Public WithEvents fraTaller As System.Windows.Forms.GroupBox
    Public WithEvents txtDesArticulo As System.Windows.Forms.TextBox
    Public WithEvents fraMotivoReparacion As System.Windows.Forms.GroupBox
    Public WithEvents txtObservacionesTaller As System.Windows.Forms.TextBox
    Public WithEvents fraObservacionesTaller As System.Windows.Forms.GroupBox
    Public WithEvents dbcTipoReparacion As System.Windows.Forms.ComboBox
    Public WithEvents dbcTaller As System.Windows.Forms.ComboBox
    Public WithEvents _lblReparaciones_7 As System.Windows.Forms.Label
    Public WithEvents _lblReparaciones_5 As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents fraLinea As System.Windows.Forms.GroupBox
    Public WithEvents txtNuevoAbono As System.Windows.Forms.TextBox
    Public WithEvents txtAbonos As System.Windows.Forms.TextBox
    Public WithEvents txtImpReparacion As System.Windows.Forms.TextBox
    Public WithEvents txtAnticipo As System.Windows.Forms.TextBox
    Public WithEvents txtCosto As System.Windows.Forms.TextBox
    Public WithEvents chkCredito As System.Windows.Forms.CheckBox
    Public WithEvents lblNuevoAbono As System.Windows.Forms.Label
    Public WithEvents _lblReparaciones_11 As System.Windows.Forms.Label
    Public WithEvents _lblReparaciones_8 As System.Windows.Forms.Label
    Public WithEvents _lblReparaciones_6 As System.Windows.Forms.Label
    Public WithEvents _lblReparaciones_99 As System.Windows.Forms.Label
    Public WithEvents txtSaldo As System.Windows.Forms.Label
    Public WithEvents _lblReparaciones_9 As System.Windows.Forms.Label
    Public WithEvents Frame6 As System.Windows.Forms.GroupBox
    Public WithEvents txtFolio As System.Windows.Forms.TextBox
    Public WithEvents dtpFechaReparacion As System.Windows.Forms.DateTimePicker
    Public WithEvents Frame5 As System.Windows.Forms.Panel
    Public WithEvents txtTipoCambio As System.Windows.Forms.TextBox
    Public WithEvents txtDiasTranscurridos As System.Windows.Forms.TextBox
    Public WithEvents _lblReparaciones_2 As System.Windows.Forms.Label
    Public WithEvents _lblReparaciones_3 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.Panel
    Public WithEvents optMonedaDolar As System.Windows.Forms.RadioButton
    Public WithEvents optMonedaPeso As System.Windows.Forms.RadioButton
    Public WithEvents fraMoneda As System.Windows.Forms.GroupBox
    Public WithEvents dbcVendedor As System.Windows.Forms.ComboBox
    Public WithEvents _lblReparaciones_4 As System.Windows.Forms.Label
    Public WithEvents _lblReparaciones_0 As System.Windows.Forms.Label
    Public WithEvents _lblReparaciones_1 As System.Windows.Forms.Label
    Public WithEvents Frame7 As System.Windows.Forms.GroupBox
    Public WithEvents chkCorpoRegresa As System.Windows.Forms.CheckBox
    Public WithEvents chkCorpoEnvio As System.Windows.Forms.CheckBox
    Public WithEvents dtpCorpoEnvio As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpCorpoRegresa As System.Windows.Forms.DateTimePicker
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents lblEstatus As System.Windows.Forms.Label
    Public WithEvents lblCancelado As System.Windows.Forms.Label
    Public WithEvents lblLiquidado As System.Windows.Forms.Label
    Public WithEvents Marco As System.Windows.Forms.GroupBox
    Public WithEvents lblReparaciones As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Dim mblnCambiosEnCodigo As Boolean
    Dim FueraChange As Boolean
    Dim mblnSalir As Boolean 'Se utiliza para saber cuando el usuraio a presionado Escape estando en el primer control del form.
    Dim tecla As Integer ''''''
    Dim mblnNuevo As Boolean ''''''
    Dim I As Integer 'Para manejar el For
    Dim rsLocal As ADODB.Recordset ''''''
    Dim intCodCliente As Integer 'Esta Variable almacena el código del Cliente seleccionado en el DataCombo del Cliente
    Dim intCodVendedor As Integer 'Esta Variable almacena el código del Vendedor seleccionado en el DataCombo del Vendedor}
    Dim intCodGrupo As Integer
    Dim intCodTaller As Integer
    Dim mblnTecleoFechaI As Boolean
    Dim msglTiempoCambioI As Single
    Dim mstrFechaCaptRep As String '''Leyenda para que saber que el PtoVta captura algo en reparaciones
    Dim mblnUnaSolaFechaCaptRep As Boolean '''Solo debe aparecer una vez

    Dim mCurPorcIva As Decimal ''Variable para guardar el porcentaje de iva
    Public WithEvents btnBuscar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents btnGuardar As Button
    Const C_msgCANCELARFOLIO As String = "¿Está seguro de que desea CANCELAR este Folio?"
    Public bandera As Boolean = False
    Public strControlActual As String 'Nombre del control actual

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas

        Dim Columna As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control


        Select Case strControlActual
            Case "TXTFOLIO"
                FueraChange = True
                Dim FrmConsultasEspeciales As FrmConsultasEspeciales = New FrmConsultasEspeciales()
                FrmConsultasEspeciales.InitializeComponent()
                FrmConsultasEspeciales.strControlActual = strControlActual
                FrmConsultasEspeciales.strFormaActual = Me.Name
                FrmConsultasEspeciales.intCodSucursal = gintCodAlmacen
                FrmConsultasEspeciales.ShowDialog()
                FueraChange = False
            Case Else
                'Sale de este sub para ke no ejecute ninguna opcion
                Exit Sub
        End Select
        Exit Sub
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub
    Sub InicializaControles()
        fraMotivoReparacion.Enabled = True
        fraObservacionesTaller.Enabled = True
        fraDatosCliente.Enabled = True
        fraCliente.Enabled = True
        chkCredito.Enabled = True
        txtCliente.Enabled = False
        txtRFCCliente.Enabled = False
        txtDomicilio.Enabled = False
        txtTelefono.Enabled = False
        '''txtDesArticulo.Enabled = False
        chkConfirmacionCliente.Enabled = False
        chkEntregaCliente.Enabled = False
        dbcTaller.Enabled = True
        chkCredito.Enabled = False
        chkReparado.Enabled = False
        fraMoneda.Enabled = True
        optMonedaDolar.Enabled = True
        optMonedaPeso.Enabled = True
    End Sub

    Sub Nuevo()
        'Este procedimiento genera un nuevo registro para una venta
        'Se deben Limpiar todos los controles del formulario con excepcion del Control de la Llavve principal
        On Error GoTo Merr

        If (bandera = True) Then
            Exit Sub
        End If

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        lblCancelado.Visible = False
        FueraChange = True
        InicializaControles()
        txtCliente.Text = ""
        txtRFCCliente.Text = ""
        txtDomicilio.Text = ""
        txtTelefono.Text = ""
        dbcVendedor.Text = ""
        FueraChange = False
        lblEstatus.Visible = False
        txtTipoCambio.Text = ""
        txtDesArticulo.Text = ""
        txtCosto.Text = ""
        txtImpReparacion.Text = ""
        txtSaldo.Text = ""
        txtAbonos.Text = ""
        txtNuevoAbono.Visible = False
        lblNuevoAbono.Visible = False
        fraLinea.Visible = True
        txtAbonos.Enabled = False
        txtNuevoAbono.Text = ""
        txtAnticipo.Text = ""
        dtpFechaReparacion.Value = Today
        dtpEntregaTaller.Value = Today
        dtpConfirmacionCliente.Value = Today
        dtpRegresoTaller.Value = Today
        mstrFechaCaptRep = "ADMON - " & VB6.Format(Today, "dd/MMM/yyyy") & vbNewLine
        mblnUnaSolaFechaCaptRep = False
        Bitacora.Text = ""
        mCurPorcIva = 0
        fraMoneda.Enabled = True
        fraNuevoTexto.Enabled = True
        cmdRegistrar.Enabled = True

        chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkTallerEntrega.Tag = System.Windows.Forms.CheckState.Unchecked
        chkTallerEntrega.Enabled = True
        chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkTallerRegreso.Tag = System.Windows.Forms.CheckState.Unchecked
        chkTallerRegreso.Enabled = True
        chkCredito.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkCredito.Tag = System.Windows.Forms.CheckState.Unchecked
        chkConfirmacionCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkConfirmacionCliente.Tag = System.Windows.Forms.CheckState.Unchecked
        chkEntregaCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkEntregaCliente.Tag = System.Windows.Forms.CheckState.Unchecked
        chkCorpoEnvio.CheckState = System.Windows.Forms.CheckState.Unchecked
        dtpCorpoEnvio.Enabled = False
        dtpCorpoEnvio.Value = Today
        chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Unchecked
        dtpCorpoRegresa.Value = Today
        chkCorpoRegresa.Enabled = True
        dtpCorpoRegresa.Enabled = False

        dtpFechaEntregaCliente.Value = Today
        txtObservacionesTaller.Text = ""
        txtObservacionesTaller.Enabled = True
        txtDiasTranscurridos.Text = ""
        dbcTipoReparacion.Text = ""
        dbcTaller.Text = ""
        fraTaller.Enabled = True
        chkTallerEntrega.Enabled = True
        chkTallerRegreso.Enabled = True
        lblLiquidado.Visible = False
        lblCancelado.Visible = False
        chkReparado.CheckState = System.Windows.Forms.CheckState.Unchecked
        fraCliente.Enabled = True
        chkConfirmacionCliente.Enabled = False
        chkEntregaCliente.Enabled = False
        dtpConfirmacionCliente.Enabled = False
        dtpFechaEntregaCliente.Enabled = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function ValidaDatos() As Object
        'Esta Función valida que todos los datos e hayan introducido en el Form de Ventas , para poder procesar la venta
        On Error GoTo Merr

        If mblnTecleoFechaI Then
            Do While (VB.Timer() - msglTiempoCambioI) <= 2.1
            Loop
            mblnTecleoFechaI = False
        End If
        System.Windows.Forms.Application.DoEvents()
        If Trim(txtFolio.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Folio de la reparación", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            Me.txtFolio.Focus()
            Exit Function
        End If
        If mblnNuevo = False Then
            If lblCancelado.Visible = True Then
                MsgBox("La reparación ha sido cancelada." & vbNewLine & "No es Posible Realizar Modificaciones.", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                Me.txtFolio.Focus()
                Exit Function
            End If
            If chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                MsgBox(C_msgFALTADATO & "Fecha de entrega al taller.", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                Me.chkTallerEntrega.Focus()
                Exit Function
            End If
            If CDate(dtpEntregaTaller.Value) < CDate(dtpFechaReparacion.Value) Then
                MsgBox("La fecha de entrega al taller debe ser mayor o igual a la fecha de registro de la reparación." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                Me.dtpEntregaTaller.Focus()
                Exit Function
            End If
            If CDate(dtpRegresoTaller.Value) < CDate(dtpEntregaTaller.Value) And chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Checked Then
                MsgBox("La fecha de regreso del taller debe ser mayor o igual a la fecha de entrega al taller." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                Me.dtpRegresoTaller.Focus()
                Exit Function
            End If
            If CDbl(Trim(CStr(intCodTaller))) = 0 Then
                MsgBox(C_msgFALTADATO & "Nombre del taller.", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                Me.dbcTaller.Focus()
                Exit Function
            End If
            If (CDec(System.Math.Abs(CDbl(Numerico(txtAnticipo.Text)))) + CDec(Numerico(txtAbonos.Text))) = CDec(Numerico(txtImpReparacion.Text)) And mblnNuevo = False And lblLiquidado.Visible = True Then
                MsgBox("La reparación ha sido liquidada." & vbNewLine & "No es Posible Realizar Modificaciones.", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                Me.txtFolio.Focus()
                Exit Function
            End If
            '''        If CCur(Numerico(txtCosto.text)) > CCur(Numerico(txtImpReparacion.text)) Then
            '''            MsgBox C_msgFALTADATO & "Costo de la reparación.", vbExclamation, gstrCorpoNOMBREEMPRESA
            '''            Me.txtCosto.SetFocus
            '''            Exit Function
            '''        End If
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ValidaDatos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ValidaDatos = True
        Exit Function

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
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
                    MsgBox(C_msgSINAUTORIZACION & "Abandonar la captura sin guardar la información.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AVISO")
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
        bandera = True
        On Error GoTo Merr
        txtFolio.Text = ""
        Nuevo()
        'mbnlnNuevo se pone en Falso, para que cuando se cierre o se kiera salir de el sin guardar no pregunte nada.
        'Ya que unicamente se tomará como True cuando el txtFolio pierda el enfoque para agregar uno nuevo, en caso de que no sea consulta
        mblnNuevo = False
        mblnCambiosEnCodigo = False
        txtFolio.Focus()

        Exit Sub
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InicializaVariables()

        If (bandera = True) Then
            Exit Sub
        End If

        mblnNuevo = False
        mblnCambiosEnCodigo = False
        mblnSalir = False
        FueraChange = False
    End Sub

    Private Sub chkConfirmacionCliente_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkConfirmacionCliente.CheckStateChanged
        If FueraChange Then Exit Sub
        If chkConfirmacionCliente.CheckState = System.Windows.Forms.CheckState.Checked Then
            If chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                MsgBox("Antes de recibir la confirmación del cliente, debe capturarse la fecha de entrega al taller por parte del corporativo." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                chkConfirmacionCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
                Exit Sub
            End If
            dtpConfirmacionCliente.Enabled = True
            GeneraBitacora()
        Else
            If (((CDec(Numerico(txtAbonos.Text)) + CDec(Numerico(txtAnticipo.Text))) = 0) Or (CDec(Numerico(txtAnticipo.Text)) + CDec(Numerico(txtAbonos.Text)) - CDec(Numerico(txtImpReparacion.Text))) = 0) Then
                dtpConfirmacionCliente.Enabled = False
                If mblnNuevo = True Then
                    txtAnticipo.Text = "0.00"
                Else
                    txtNuevoAbono.Text = "0.00"
                End If
            Else
                FueraChange = True
                chkConfirmacionCliente.CheckState = System.Windows.Forms.CheckState.Checked
                MsgBox("Esta Reparación ya Tiene Abonos Registrados, No se Puede Desautorizar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                FueraChange = False
            End If
        End If
        GeneraBitacora()
    End Sub

    Private Sub chkConfirmacionCliente_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkConfirmacionCliente.Leave
        GeneraBitacora()
    End Sub

    Private Sub chkCorpoEnvio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCorpoEnvio.Leave
        GeneraBitacora()
    End Sub

    'UPGRADE_WARNING: Event chkCorpoRegresa.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkCorpoRegresa_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCorpoRegresa.CheckStateChanged
        If FueraChange Then Exit Sub
        If chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Checked Then
            If (chkCorpoEnvio.CheckState = System.Windows.Forms.CheckState.Unchecked) Then
                MsgBox("No es posible regresar la pieza al Punto de Venta si no se envió al Corporativo", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                If (chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Unchecked And dtpEntregaTaller.Enabled) And (chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Unchecked And dtpRegresoTaller.Enabled) Then
                    MsgBox("No es posible regresar la pieza al Punto de Venta si no ha sido enviada al Taller", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                    chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Unchecked
                Else
                    dtpCorpoRegresa.Enabled = True
                End If
                If (chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Checked And chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Unchecked) Or (chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Unchecked And chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Unchecked) Then
                    MsgBox("No es posible regresar la pieza al Punto de Venta si no ha regresado del Taller", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                    chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Unchecked
                Else
                    dtpCorpoRegresa.Enabled = True
                End If
            End If
        End If
        If chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Unchecked Then dtpCorpoRegresa.Enabled = False
        GeneraBitacora()
    End Sub

    Private Sub chkCorpoRegresa_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCorpoRegresa.Leave
        GeneraBitacora()
    End Sub

    Private Sub chkEntregaCliente_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkEntregaCliente.Leave
        GeneraBitacora()
    End Sub

    'UPGRADE_WARNING: Event chkReparado.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkReparado_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkReparado.CheckStateChanged
        If chkReparado.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            txtCosto.Text = "0.00"
            txtImpReparacion.Text = "0.00"
        End If
        GeneraBitacora()
    End Sub

    Private Sub chkTallerEntrega_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTallerEntrega.CheckStateChanged
        If FueraChange Then Exit Sub
        If chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Checked Then
            If (chkCorpoEnvio.CheckState = System.Windows.Forms.CheckState.Unchecked) Then
                MsgBox("No es posible mandar la reparación al taller" & vbNewLine & "si esta no ha sido enviada al Corporativo" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                dtpEntregaTaller.Enabled = True
            End If
        Else
            dtpEntregaTaller.Enabled = False
        End If
        GeneraBitacora()
    End Sub

    Private Sub chkTallerEntrega_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTallerEntrega.Leave
        GeneraBitacora()
    End Sub

    Private Sub chkTallerRegreso_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTallerRegreso.CheckStateChanged
        If FueraChange Then Exit Sub
        If chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Checked Then
            If (chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Unchecked And chkTallerEntrega.Enabled) And (Not dtpEntregaTaller.Enabled) Then
                MsgBox("No es posible registrar la fecha de regreso del taller si no ha sido enviado" & vbNewLine & "Favor de verificar...", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                FueraChange = True
                chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Unchecked
                FueraChange = False
            Else
                dtpRegresoTaller.Enabled = True
                chkReparado.Enabled = True
                chkReparado.CheckState = System.Windows.Forms.CheckState.Checked
            End If
        Else
            FueraChange = True
            dtpRegresoTaller.Enabled = False
            chkReparado.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkReparado.Enabled = False
            If chkCorpoRegresa.Enabled Then chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Unchecked
            FueraChange = False
        End If
        GeneraBitacora()
    End Sub

    Private Sub chkTallerRegreso_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTallerRegreso.Enter
        GeneraBitacora()
    End Sub

    Private Sub cmdRegistrar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRegistrar.Click

        If Not mblnUnaSolaFechaCaptRep Then
            If Trim(txtDesArticulo.Text) = "" Then
                txtDesArticulo.Text = vbNewLine & mstrFechaCaptRep & Trim(txtNuevoTexto.Text)
                txtDesArticulo.Tag = vbNewLine & mstrFechaCaptRep & Trim(txtNuevoTexto.Text)
            Else
                txtDesArticulo.Text = Trim(txtDesArticulo.Text) & vbNewLine & vbNewLine & mstrFechaCaptRep & Trim(txtNuevoTexto.Text)
                txtDesArticulo.Tag = Trim(txtDesArticulo.Text) & vbNewLine & vbNewLine & mstrFechaCaptRep & Trim(txtNuevoTexto.Text)
            End If
            mblnUnaSolaFechaCaptRep = True
            mblnCambiosEnCodigo = True
        Else
            txtDesArticulo.Text = Trim(txtDesArticulo.Text) & vbNewLine & Trim(txtNuevoTexto.Text)
            txtDesArticulo.Tag = Trim(txtDesArticulo.Text) & vbNewLine & Trim(txtNuevoTexto.Text)
            mblnCambiosEnCodigo = True
        End If

        txtDesArticulo.SelectionStart = Len(txtDesArticulo.Text)
        txtNuevoTexto.Text = ""
        txtNuevoTexto.SelectionStart = 0
        txtNuevoTexto.Focus()
    End Sub

    Private Sub dbcTaller_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcTaller.KeyUp
        gStrSql = "Select CodTaller , Ltrim(Rtrim(DescTaller)) As DescTaller From CatTalleres Where DescTaller LIKE '" & Trim(dbcTaller.Text) & "%' Order By DescTaller "
        ModDCombo.DCLostFocus(dbcTaller, gStrSql, intCodTaller)
    End Sub

    Private Sub dbcTaller_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcTaller.MouseUp
        'gStrSql = "Select CodTaller , Ltrim(Rtrim(DescTaller)) As DescTaller From CatTalleres Where DescTaller LIKE '" & Trim(dbcTaller.Text) & "%' Order By DescTaller "
        'ModDCombo.DCLostFocus(dbcTaller, gStrSql, intCodTaller)
    End Sub


    Private Sub dbcvendedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVendedor.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcVendedor.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT Codvendedor , Descvendedor=ltrim(rtrim(Descvendedor))  From Catvendedores WHERE Descvendedor LIKE '" & Trim(dbcVendedor.Text) & "%' ORDER BY Descvendedor"
        ModDCombo.DCChange(gStrSql, tecla)

    End Sub

    Private Sub dbcvendedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVendedor.Enter
        Pon_Tool()
        gStrSql = "SELECT Codvendedor , Descvendedor=ltrim(rtrim(Descvendedor))  From Catvendedores ORDER BY Descvendedor"
        ModDCombo.DCGotFocus(gStrSql, dbcVendedor)
    End Sub

    Private Sub dbcvendedor_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcVendedor.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
    End Sub

    Private Sub dbcvendedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVendedor.Leave
        intCodVendedor = 0
        gStrSql = "SELECT Codvendedor , Descvendedor=ltrim(rtrim(Descvendedor))  From Catvendedores WHERE Descvendedor LIKE '" & Trim(dbcVendedor.Text) & "%' ORDER BY Descvendedor"
        ModDCombo.DCLostFocus(dbcVendedor, gStrSql, intCodVendedor)
    End Sub

    Private Sub dbcTipoReparacioN_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcTipoReparacion.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcTipoReparacion.Name Then
        '    Exit Sub
        'End If
        gStrSql = "Select CodGrupo, DescGrupo=LTRIM(RTRIM(DescGrupo)) From CatGrupos Where DescGRupo LIKE '" & Trim(dbcTipoReparacion.Text) & "%' Order By DescGrupo"
        ModDCombo.DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcTipoReparacioN_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcTipoReparacion.Enter
        Pon_Tool()
        gStrSql = "Select CodGrupo, DescGrupo=LTRIM(RTRIM(DescGrupo)) From CatGrupos"
        ModDCombo.DCGotFocus(gStrSql, dbcTipoReparacion)
    End Sub

    Private Sub dbcTipoReparacioN_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcTipoReparacion.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
    End Sub

    Private Sub dbcTipoReparacioN_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcTipoReparacion.Leave
        intCodGrupo = 0
        gStrSql = "Select CodGrupo, DescGrupo=LTRIM(RTRIM(DescGrupo)) From CatGrupos Where DescGRupo LIKE '" & Trim(dbcTipoReparacion.Text) & "%' Order By DescGrupo"
        ModDCombo.DCLostFocus(dbcTipoReparacion, gStrSql, intCodGrupo)
    End Sub

    Private Sub dbcTaller_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcTaller.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcTaller.Name Then
        '    Exit Sub
        'End If
        gStrSql = "Select CodTaller , Ltrim(Rtrim(DescTaller)) As DescTaller From CatTalleres Where DescTaller LIKE '" & Trim(dbcTaller.Text) & "%' Order By DescTaller"
        ModDCombo.DCChange(gStrSql, tecla)
        GeneraBitacora()
    End Sub

    Private Sub dbcTaller_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcTaller.Enter
        Pon_Tool()
        gStrSql = "Select CodTaller , Ltrim(Rtrim(DescTaller)) As DescTaller From CatTalleres (Nolock) "
        ModDCombo.DCGotFocus(gStrSql, dbcTaller)
    End Sub

    Private Sub dbcTaller_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcTaller.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
    End Sub

    Private Sub dbcTaller_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcTaller.Leave
        intCodTaller = 0
        gStrSql = "Select CodTaller , Ltrim(Rtrim(DescTaller)) As DescTaller From CatTalleres (Nolock) Where DescTaller LIKE '" & Trim(dbcTaller.Text) & "%' Order By DescTaller"
        ModDCombo.DCLostFocus(dbcTaller, gStrSql, intCodTaller)
        GeneraBitacora()
    End Sub

    Private Sub dtpConfirmacionCliente_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpConfirmacionCliente.Leave
        If CDate(dtpConfirmacionCliente.Value) < CDate(dtpFechaReparacion.Value) Then
            MsgBox("La fecha de confirmacion del cliente debe ser mayor o igual a la fecha de registro de la reparación." & vbNewLine & "Verifique Por Favo.", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            Me.dtpConfirmacionCliente.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub dtpEntregaTaller_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dtpEntregaTaller.KeyPress
        mblnTecleoFechaI = True
        msglTiempoCambioI = VB.Timer()
    End Sub

    Private Sub dtpEntregaTaller_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpEntregaTaller.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        If CDate(dtpEntregaTaller.Value) < CDate(dtpFechaReparacion.Value) Then
            MsgBox("La fecha de entrega al taller debe ser mayor o igual a la fecha de registro de la reparación." & vbNewLine & "Verifique Por Favo.", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            Me.dtpEntregaTaller.Focus()
            Exit Sub
        End If
    End Sub


    Private Sub dtpRegresoTaller_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dtpRegresoTaller.KeyPress
        mblnTecleoFechaI = True
        msglTiempoCambioI = VB.Timer()
    End Sub

    Private Sub dtpRegresoTaller_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpRegresoTaller.Leave
        ' If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        If CDate(dtpRegresoTaller.Value) < CDate(dtpEntregaTaller.Value) Then
            MsgBox("La fecha de regreso del taller debe ser mayor o igual a la fecha de entrega al taller." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            Me.dtpRegresoTaller.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub frmCorpoControlReparaciones_Corpo_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub
    Private Sub frmCorpoControlReparaciones_Corpo_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmCorpoControlReparaciones_Corpo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        bandera = True
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
        Nuevo()
    End Sub

    Private Sub frmCorpoControlReparaciones_Corpo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        ' En este evento del formulario se valida la tecla presionada.
        ' Si es Enter se simula un tab(Avanza al siguiente control)
        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ' Si el control en que se presiono enter, es el Grid de Detalle de la venta que no se ejecute el avanzar tab
                'If ActiveControl.Name <> "msgDetalleApartado" Then
                '    ModEstandar.AvanzarTab(Me)
                'End If
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
                '        Case vbKeyF8
                ''            If mblnNuevo = True Then 'Unicamente si es un registro nuevo podrá procesaar un pago, en las consultas no es posible hacer bada.
                '                If txtFolio <> "" Then
                '                    If ValidaDatos = True Then
                '                        If (mblnNuevo = True And Numerico(txtAnticipo) > 0) Or (mblnNuevo = False And Numerico(txtNuevoAbono) > 0) Then
                '                            frmPagosSalMercancia.Show
                '                            PonerTotalesenFrmPagos
                '                        Else
                '                            Me.Guardar
                '                        End If
                '                    End If
                '                End If
                ''            End If
        End Select
    End Sub

    Private Sub frmCorpoControlReparaciones_Corpo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoControlReparaciones_Corpo_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSalir Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        '    'Si se cierra el formulario y existio algun cambio en el registro se
        '    'informa al usuario del cabio y si desea guardar el registro, ya sea
        '    'que sea nuevo o un registro modificado
        '    If mblnNuevo = True Then
        '        'Si desea Salir del Formulario sin haber guardaro datos, Verificar si se requiere autorización para hacerlo.
        '        If PuedeAbdandonarCaptIniciada() = True Then
        '            Cancel = 0
        '            Me.Close()
        '        Else
        '            Cancel = 1 'Para que no salga del Formulario hasta que guarde los datos, si no tiene premiso de hacerlo
        '        End If
        '    End If
        'Else 'Se quiere salir con escape
        '    mblnSalir = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
        '        Case MsgBoxResult.Yes 'Sale del Formulario, pero antes preguntar si desea grabar los datos registrados, solo cuando es nuevo
        '            'Si desea Salir del Formulario sin haber guardaro datos, Verificar si se requiere autorización para hacerlo.
        '            'Esto solo se valida cuando sea un nuevo registro, ya que si es consulta no puede hacerle modificaciones
        '            If mblnNuevo = True Then
        '                If PuedeAbdandonarCaptIniciada() = True Then
        '                    Cancel = 0
        '                Else
        '                    Cancel = 1 'Para que no salga del Formulario hasta que guarde los datos, si no tiene premiso de hacerlo
        '                End If
        '            Else
        '                Cancel = 0 'Sale de la Captura, Con 1: Sigue en la captura
        '            End If
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoControlReparaciones_Corpo_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        IsNothing(Me)
    End Sub

    Private Sub txtabonos_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAbonos.TextChanged
        '    Dim Abonos As Currency
        mblnCambiosEnCodigo = True
        '    Abonos = CCur(Numerico(txtAbonos)) + CCur(Numerico(txtAnticipo))
        '    txtSaldo = Format(Abonos - Numerico(txtImpReparacion), gstrFormatoCantidad)
    End Sub

    Private Sub txtabonos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAbonos.Enter
        SelTextoTxt(txtAbonos)
        Pon_Tool()
    End Sub

    Private Sub txtAnticipo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnticipo.TextChanged
        mblnCambiosEnCodigo = True
        txtSaldo.Text = VB6.Format(CDbl(Numerico(txtAnticipo.Text)) - CDbl(Numerico(txtImpReparacion.Text)), gstrFormatoCantidad)
    End Sub

    Private Sub txtAnticipo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnticipo.Enter
        SelTextoTxt(txtAnticipo)
        Pon_Tool()
    End Sub

    Private Sub LlenaDatosCliente()
        On Error GoTo Merr
        gStrSql = "SELECT   DescCliente,  Rfc,Ltrim(Rtrim(Domicilio))  as Domicilio , LTRIM(RTRIM(Colonia))  as Colonia , LTRIM(RTRIM(Ciudad)) as Ciudad,ltrim(rtrim(TelCasa)) + '     ' + ltrim(rtrim(TelOficina)) + '     ' + ltrim(rtrim(Fax)) as Telefonos FROM CATCLIENTES WHERE CodCLiente= " & Numerico(txtCodCliente.Text) & " "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            txtCliente.Text = Trim(RsGral.Fields("DescCliente").Value)
            txtRFCCliente.Text = RsGral.Fields("Rfc").Value
            txtRFCCliente.Tag = RsGral.Fields("Rfc").Value
            txtDomicilio.Text = RsGral.Fields("Domicilio").Value + Space(4) + RsGral.Fields("Colonia").Value + Space(4) + RsGral.Fields("Ciudad").Value
            txtDomicilio.Tag = RsGral.Fields("Domicilio").Value
            txtTelefono.Text = Trim(RsGral.Fields("Telefonos").Value)
            txtTelefono.Tag = Trim(RsGral.Fields("Telefonos").Value)
        Else
            txtRFCCliente.Text = ""
            txtRFCCliente.Tag = ""
            txtDomicilio.Text = ""
            txtDomicilio.Tag = ""
            txtTelefono.Text = ""
            txtTelefono.Tag = ""
            txtCliente.Text = ""
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatos()

        If (bandera = True) Then
            Exit Sub
        End If

        'Este Proc muestra los datos correspondientes a un Folio de venta dado
        On Error GoTo Merr
        Dim FechaApartado As String
        Dim FechaVencimiento As String
        Dim Estatus As String

        mCurPorcIva = 0

        gStrSql = "SELECT V.DescVendedor AS Vendedor, ISNULL(T.DescTaller,'') AS Taller, G.DescGrupo AS Grupo, Isnull(SUM(I.Total * I.TipoCambio ) ,0) AS Abonos, R.FolioReparacion, R.FechaReparacion, " & "R.Nombre, R.Rfc, R.Telefono, R.MotivoReparacion, R.ObservacionesTaller, R.TipoCambio, R.CostoReparacion, R.ImporteVta, R.Anticipo, R.Estatus, " & "R.FechaCancel, G.CodGrupo, ISNULL(T.CodTaller,0) as CodTaller, R.CodCliente, R.CodVendedor,R.FechaEntregaTaller,FechaRegreso,FechaConfirmacion,FechaEntregaCliente,Credito, R.Reparado, R.Moneda, " & "R.FechaCorpoEnvio, R.FechaCorpoRegresa,dbo.ReparacionesEstatus( R.CodSucursal, R.FechaReparacion, R.FolioReparacion ) AS EstatusRep,R.PorcIva " & "FROM dbo.Reparaciones R INNER JOIN " & "dbo.CatVendedores V ON R.CodVendedor = V.CodVendedor LEFT OUTER JOIN " & "dbo.CatTalleres T ON R.CodTaller = T.CodTaller INNER JOIN " & "dbo.CatGrupos G ON R.TipoReparacion = G.CodGrupo LEFT OUTER  JOIN " & "dbo.Ingresos I ON V.CodVendedor = I.CodVendedor AND R.FolioReparacion = I.FolioMovto AND (I.TipoIngreso <> 'A') " & "GROUP BY R.FolioReparacion, V.DescVendedor, T.DescTaller, G.DescGrupo, R.FechaReparacion, R.TipoMovto, R.CodSucursal, R.CodCaja, R.CodCliente, " & "R.Nombre, R.Rfc, R.Telefono, R.MotivoReparacion, R.ObservacionesTaller, R.CodTaller, R.FolioReparacion, R.Nombre, R.Rfc, R.Telefono, " & "R.MotivoReparacion , R.ObservacionesTaller, R.Tipocambio, R.CostoReparacion, R.ImporteVta, R.Anticipo, R.Estatus, R.FechaCancel, G.CodGRupo, R.CodCliente,R.CodVendedor, T.COdTaller, R.Moneda,  " & "R.FechaEntregaTaller,FechaRegreso,FechaConfirmacion,FechaEntregaCliente,Credito, R.Reparado, R.FechaCorpoEnvio, R.FechaCorpoRegresa,R.PorcIva " & "HAVING (R.FolioReparacion = '" & Trim(txtFolio.Text) & "')"

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then

            txtNuevoAbono.Visible = True
            lblNuevoAbono.Visible = True
            txtAbonos.Enabled = True
            fraLinea.Visible = False
            FueraChange = True
            dbcVendedor.Text = RsGral.Fields("Vendedor").Value
            intCodVendedor = RsGral.Fields("CodVendedor").Value
            dbcTipoReparacion.Text = RsGral.Fields("Grupo").Value
            intCodGrupo = RsGral.Fields("CodGrupo").Value
            dbcTaller.Text = Trim(RsGral.Fields("Taller").Value)
            dbcTaller.Tag = Trim(RsGral.Fields("Taller").Value)
            intCodTaller = RsGral.Fields("CodTaller").Value
            FueraChange = False
            lblEstatus.Visible = True
            lblEstatus.Text = RsGral.Fields("EstatusRep").Value
            Estatus = Trim(RsGral.Fields("Estatus").Value)

            If Trim(RsGral.Fields("Moneda").Value) = "P" Then
                optMonedaPeso.Checked = True
                txtTipoCambio.Text = VB6.Format(RsGral.Fields("TipoCambio").Value, gstrFormatoCantidad)
            ElseIf Trim(RsGral.Fields("Moneda").Value) = "D" Then
                optMonedaDolar.Checked = True
                txtTipoCambio.Text = VB6.Format(gcurCorpoTIPOCAMBIODOLAR, gstrFormatoCantidad)
            End If
            If RsGral.Fields("Reparado").Value = True Then
                chkReparado.CheckState = System.Windows.Forms.CheckState.Checked
                chkReparado.Tag = System.Windows.Forms.CheckState.Checked
            Else
                chkReparado.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkReparado.Tag = System.Windows.Forms.CheckState.Unchecked
            End If
            If RsGral.Fields("FechaCorpoEnvio").Value <> CDate("01/01/1900") Then
                FueraChange = True
                chkCorpoEnvio.CheckState = System.Windows.Forms.CheckState.Checked
                chkCorpoEnvio.Enabled = False
                FueraChange = False
                dtpCorpoEnvio.Value = VB6.Format(RsGral.Fields("FechaCorpoEnvio").Value, C_FORMATFECHAMOSTRAR)
            Else
                chkCorpoEnvio.CheckState = System.Windows.Forms.CheckState.Unchecked
                dtpCorpoEnvio.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
            End If
            If RsGral.Fields("FechaCorpoRegresa").Value <> CDate("01/01/1900") Then
                FueraChange = True
                chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Checked
                chkCorpoRegresa.Enabled = False
                FueraChange = False
                dtpCorpoRegresa.Value = VB6.Format(RsGral.Fields("FechaCorpoRegresa").Value, C_FORMATFECHAMOSTRAR)
                dtpCorpoRegresa.Enabled = False
            Else
                chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkCorpoRegresa.Enabled = True
                dtpCorpoRegresa.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
                dtpCorpoRegresa.Enabled = False
            End If

            If RsGral.Fields("FechaEntregaTaller").Value <> CDate("01/01/1900") Then
                chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Checked
                chkTallerEntrega.Tag = System.Windows.Forms.CheckState.Checked
                dtpEntregaTaller.Value = VB6.Format(RsGral.Fields("FechaEntregaTaller").Value, C_FORMATFECHAMOSTRAR)
                chkTallerEntrega.Enabled = False
                dtpEntregaTaller.Enabled = False
                dbcTaller.Enabled = False
            Else
                chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkTallerEntrega.Tag = System.Windows.Forms.CheckState.Unchecked
                chkTallerEntrega.Enabled = True
                dtpEntregaTaller.Enabled = False
                dtpEntregaTaller.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
            End If
            If RsGral.Fields("FechaRegreso").Value <> CDate("01/01/1900") Then
                chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Checked
                chkTallerRegreso.Tag = System.Windows.Forms.CheckState.Checked
                dtpRegresoTaller.Value = VB6.Format(RsGral.Fields("FechaRegreso").Value, C_FORMATFECHAMOSTRAR)
                chkTallerRegreso.Enabled = False
                dtpRegresoTaller.Enabled = False
                chkReparado.Enabled = False
            Else
                chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkTallerRegreso.Tag = System.Windows.Forms.CheckState.Unchecked
                chkTallerRegreso.Enabled = True
                dtpRegresoTaller.Enabled = False
                dtpRegresoTaller.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
            End If

            If RsGral.Fields("FechaConfirmacion").Value <> CDate("01/01/1900") Then
                chkConfirmacionCliente.CheckState = System.Windows.Forms.CheckState.Checked
                chkConfirmacionCliente.Tag = System.Windows.Forms.CheckState.Checked
                dtpConfirmacionCliente.Value = VB6.Format(RsGral.Fields("FechaConfirmacion").Value, C_FORMATFECHAMOSTRAR)
            Else
                chkConfirmacionCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkConfirmacionCliente.Tag = System.Windows.Forms.CheckState.Unchecked
                dtpConfirmacionCliente.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
            End If
            If RsGral.Fields("FechaEntregaCliente").Value <> CDate("01/01/1900") Then
                dtpFechaEntregaCliente.Value = VB6.Format(RsGral.Fields("FechaEntregaCliente").Value, C_FORMATFECHAMOSTRAR)
                chkEntregaCliente.CheckState = System.Windows.Forms.CheckState.Checked
                chkEntregaCliente.Tag = System.Windows.Forms.CheckState.Checked
            Else
                dtpFechaEntregaCliente.Value = Today
                chkEntregaCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkEntregaCliente.Tag = System.Windows.Forms.CheckState.Unchecked
            End If
            If RsGral.Fields("Credito").Value = True Then
                chkCredito.CheckState = System.Windows.Forms.CheckState.Checked
                chkCredito.Tag = System.Windows.Forms.CheckState.Checked
            Else
                chkCredito.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkCredito.Tag = System.Windows.Forms.CheckState.Unchecked
            End If
            dtpFechaReparacion.Value = VB6.Format(RsGral.Fields("FechaReparacion").Value, C_FORMATFECHAMOSTRAR)

            txtDesArticulo.Text = Trim(RsGral.Fields("MotivoReparacion").Value)
            txtDesArticulo.Tag = Trim(RsGral.Fields("MotivoReparacion").Value)
            txtDesArticulo.SelectionStart = 0
            txtDesArticulo.SelectionLength = 0

            txtObservacionesTaller.Text = RsGral.Fields("ObservacionesTaller").Value
            txtObservacionesTaller.Tag = RsGral.Fields("ObservacionesTaller").Value
            txtCosto.Text = VB6.Format(System.Math.Round(CDbl(Numerico(RsGral.Fields("CostoReparacion").Value)) * CDbl(Numerico(RsGral.Fields("TipoCambio").Value)), 1), gstrFormatoCantidad)
            txtCosto.Tag = VB6.Format(System.Math.Round(CDbl(Numerico(RsGral.Fields("CostoReparacion").Value)) * CDbl(Numerico(RsGral.Fields("TipoCambio").Value)), 1), gstrFormatoCantidad)
            txtImpReparacion.Text = VB6.Format(System.Math.Round(CDbl(Numerico(RsGral.Fields("ImporteVta").Value)) * CDbl(Numerico(RsGral.Fields("TipoCambio").Value)), 1), gstrFormatoCantidad)
            txtImpReparacion.Tag = VB6.Format(System.Math.Round(CDbl(Numerico(RsGral.Fields("ImporteVta").Value)) * CDbl(Numerico(RsGral.Fields("TipoCambio").Value)), 1), gstrFormatoCantidad)
            txtAnticipo.Text = VB6.Format(System.Math.Round(CDbl(Numerico(RsGral.Fields("Anticipo").Value)) * CDbl(Numerico(RsGral.Fields("TipoCambio").Value)), 1), gstrFormatoCantidad)
            txtAbonos.Text = VB6.Format(System.Math.Round(RsGral.Fields("Abonos").Value, 1), gstrFormatoCantidad)
            txtSaldo.Text = VB6.Format(System.Math.Round(CDec(Numerico(txtAnticipo.Text)) + CDec(Numerico(txtAbonos.Text)) - CDec(Numerico(txtImpReparacion.Text)), 1), gstrFormatoCantidad)
            txtNuevoAbono.Text = "0.00"

            mCurPorcIva = RsGral.Fields("PorcIva").Value

            FueraChange = False
            ' si la Reparacion fue liquidada, no mostrar el textbox de Nuevo Abono, y deshabilitar todos los controles
            lblLiquidado.Visible = False
            If (CDbl(Numerico(txtImpReparacion.Text)) > 0 And CDbl(Numerico(txtSaldo.Text)) = 0) Or (CDec(Numerico(txtSaldo.Text)) = 0 And CDec(Numerico(txtImpReparacion.Text)) = 0) Then
                If chkEntregaCliente.CheckState = System.Windows.Forms.CheckState.Checked Then
                    DesHabilitarControles()
                    If RsGral.Fields("Credito").Value = False Then lblLiquidado.Visible = True
                    If (CDec(Numerico((txtImpReparacion.Text))) > 0 And CDec(Numerico((txtSaldo.Text))) = 0) Then lblLiquidado.Visible = True

                    fraCliente.Enabled = False
                    fraTaller.Enabled = False
                    dtpRegresoTaller.Enabled = False
                    dtpEntregaTaller.Enabled = False
                    chkTallerEntrega.Enabled = False
                    chkTallerRegreso.Enabled = False
                    chkConfirmacionCliente.Enabled = False
                    chkCredito.Enabled = False
                    chkEntregaCliente.Enabled = False
                End If
            ElseIf lblCancelado.Visible = False Then
                HabilitarControles()
                lblLiquidado.Visible = False
                fraTaller.Enabled = True
            End If

            If RsGral.Fields("Estatus").Value = "C" Then
                lblCancelado.Visible = True
                DesHabilitarControlesFolioCancelado()
            Else
                lblCancelado.Visible = False
                HabilitarControles()
            End If
            If lblLiquidado.Visible Then
                fraNuevoTexto.Enabled = False
                cmdRegistrar.Enabled = False
                txtObservacionesTaller.Enabled = False
            End If

            txtDiasTranscurridos.Text = CalcularDiasTranscurridos()
            FueraChange = True
            txtCodCliente.Text = RsGral.Fields("CodCliente").Value
            LlenaDatosCliente()
            FueraChange = False

            Bitacora.Text = RegresaDatoBitacora(Trim(txtFolio.Text))

            If chkCorpoEnvio.CheckState = System.Windows.Forms.CheckState.Checked And chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                If Estatus <> "C" Then
                    fraCliente.Enabled = True
                    chkConfirmacionCliente.Enabled = True
                    chkEntregaCliente.Enabled = False
                    dtpConfirmacionCliente.Enabled = False
                    dtpFechaEntregaCliente.Enabled = False
                Else
                    fraCliente.Enabled = True
                    chkConfirmacionCliente.Enabled = False
                    chkEntregaCliente.Enabled = False
                    dtpConfirmacionCliente.Enabled = False
                    dtpFechaEntregaCliente.Enabled = False
                End If
            Else
                fraCliente.Enabled = True
                chkConfirmacionCliente.Enabled = False
                chkEntregaCliente.Enabled = False
                dtpConfirmacionCliente.Enabled = False
                dtpFechaEntregaCliente.Enabled = False
            End If

        Else
            MsjNoExiste("El Folio de reparación", gstrCorpoNOMBREEMPRESA)
            Limpiar()
        End If
        mblnCambiosEnCodigo = False
        mblnNuevo = False
        Exit Sub

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub PonerClienteDefault()
        FueraChange = True
        txtCodCliente.Text = CStr(1)
        LlenaDatosCliente()
        FueraChange = False
    End Sub

    Function FormateoDecimales(ByRef Cantidad As Object) As Object
        'Esta función formatea una cantidad, de acuerdo al numero de decimales especificados en la configuracion general.
        On Error GoTo Merr
        Dim F As Integer
        Dim Formato As String 'Contiene el Formato a dar a la cantidad (Por ejemplo: "0.00", "0.000",etc)
        FormateoDecimales = Cantidad
        If gbytCantidadDecimales = 0 Then Exit Function
        Formato = "0."
        For F = 1 To gbytCantidadDecimales
            Formato = Formato & "0"
        Next
        FormateoDecimales = VB6.Format(Cantidad, Formato)
        Exit Function
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Private Sub txtAnticipo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAnticipo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = MskCantidad(txtAnticipo.Text, KeyAscii, 8, CInt(gbytCantidadDecimales), (txtAnticipo.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtAnticipo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnticipo.Leave
        txtAnticipo.Text = FormateoDecimales(Numerico(txtAnticipo.Text))
        txtAnticipo.Text = VB6.Format(txtAnticipo.Text, gstrFormatoCantidad)
    End Sub

    Private Sub txtCliente_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCliente.TextChanged
        If FueraChange = True Then Exit Sub
        LimpiaDatosCliente()
    End Sub

    Private Sub txtCliente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCliente.Enter
        SelTextoTxt(txtCliente)
    End Sub

    Private Sub txtCosto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCosto.Enter
        SelTextoTxt(txtCosto)
        Pon_Tool()
    End Sub

    Private Sub txtCosto_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCosto.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtCosto.Text, KeyAscii, 8, 2, (txtCosto.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCosto_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCosto.Leave
        txtCosto.Text = FormateoDecimales(Numerico(txtCosto.Text))
        txtCosto.Text = VB6.Format(txtCosto.Text, gstrFormatoCantidad)
    End Sub

    Private Sub txtDesArticulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesArticulo.Leave
        '''    txtDesArticulo = ModEstandar.QuitaEnter(txtDesArticulo)
    End Sub

    Private Sub txtFolio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtFolio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.Enter
        strControlActual = UCase("txtFolio")
        SelTextoTxt(txtFolio)
        Pon_Tool()
    End Sub

    Private Sub txtFolio_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFolio.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            KeyCode = 0
        ElseIf KeyCode = System.Windows.Forms.Keys.Delete Then
            'sI La Tecla presionada fue SUPR, se borrará todo el contenido del form. ya que no es posible hacer modificaciones.
            'Unicamnete podran consultarse los datos.
            Nuevo()
            'Si la tecla presionada fue Delete y Hay cambios, pregunta si se desea guardar
            '        If Cambios = True And KeyCode = vbKeyDelete Then
            '            Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrCorpoNombreEmpresa)
            '                Case vbYes: 'Guardar el registro
            '                    If Guardar = False Then
            '                        KeyCode = 0
            '                        Exit Sub
            '                    End If
            '                Case vbNo: 'No hace nada y permite que se borre el contenido del text
            '                    Nuevo
            '                Case vbCancel: 'Cancela la captura
            '                    TxtFolio.SetFocus
            '                    KeyCode = 0
            '                    Exit Sub
            '            End Select
            '        End If
        End If
    End Sub

    Private Sub txtFolio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'Si la tecla presionada no es numero regresa un 0
        If Valida_Folio(txtFolio.Text, KeyAscii, Len(txtFolio.Text) + 1) = 0 And KeyAscii <> System.Windows.Forms.Keys.Back Then
            KeyAscii = 0
        Else
            'Cuando se edite el folio de la venta, se limpiaran todos los controles del formulario.
            Nuevo()
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.Leave
        'If System.Windows.Forms.Form.ActiveForm.Text <> Me.Text Then
        '    Exit Sub
        'End If
        Dim Prefijo As String
        Dim FechaApartado As Date
        If Trim(txtFolio.Text) = "" Then
            gStrSql = "Select  Prefijo, Consecutivo + 1 as Consecutivo From CatFolios Where DescFolio = 'REPARACIONES'"

            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            FechaApartado = Today
            Prefijo = Trim(RsGral.Fields("Prefijo").Value)
            txtFolio.Text = Prefijo & VB6.Format(CStr(gintCodAlmacen), "00") & VB6.Format(CStr(gintCodCaja), "00") & CStr(Year(FechaApartado)) & VB6.Format(CStr(Month(FechaApartado)), "00") & VB6.Format(CStr(VB.Day(FechaApartado)), "00") & "000000"

        End If
        LlenaDatos()
        txtDiasTranscurridos.Text = CalcularDiasTranscurridos()
        GeneraBitacora()
    End Sub

    Sub Guardar()
        'En las Reparaciones, los importes se dan en Pesos, pero para guardar, se hace en Dolares.
        On Error GoTo Merr
        Dim FolioReparacion As String 'Esta VArible alamcena el folio del Apartado que se esta dando de alta
        Dim FechaReparacion As Date
        Dim TipoCambioDolar As Decimal

        Dim Prefijo As String
        Dim Consecutivo As String
        Dim Moneda As String
        Dim Estatus As String
        Dim TipoMovto As String
        Dim TipoIngreso As String
        Dim mcurCostoReparacionD As Decimal
        Dim mcurSubTotalImporteD As Decimal
        Dim mcurIvaImporteD As Decimal
        Dim mcurImporteReparacionD As Decimal
        Dim mcurAnticipoD As Decimal
        Dim Reparado As Boolean
        If Cambios() = False Then Exit Sub

        If mblnTecleoFechaI Then
            Do While (VB.Timer() - msglTiempoCambioI) <= 2.1
            Loop
            mblnTecleoFechaI = False
        End If
        System.Windows.Forms.Application.DoEvents()

        'Indicar en que procedimiento de Guardar nos encontramos.
        gstrProcesoqueGeneraError = "frmcontrolReparaciones (Guardar) "

        dbcTaller_Leave(dbcTaller, New System.EventArgs())
        If Me.ValidaDatos = False Then
            Exit Sub
        End If
        mcurCostoReparacionD = CDbl(Numerico(txtCosto.Text)) / CDbl(Numerico(txtTipoCambio.Text))
        mcurImporteReparacionD = CDbl(Numerico(txtImpReparacion.Text)) / CDbl(Numerico(txtTipoCambio.Text))
        If mCurPorcIva = 0 Then
            mcurIvaImporteD = mcurImporteReparacionD * (gcurCorpoTASAIVA / 100) / (1 + (gcurCorpoTASAIVA / 100))
            mCurPorcIva = gcurCorpoTASAIVA
        Else
            mcurIvaImporteD = mcurImporteReparacionD * (mCurPorcIva / 100) / (1 + (mCurPorcIva / 100))
        End If
        mcurSubTotalImporteD = mcurImporteReparacionD - mcurIvaImporteD

        '    If mcurCostoReparacionD = 0 Then
        '        mCurPorcIva = 0
        '    End If

        intCodCliente = CInt(Numerico(txtCodCliente.Text))
        TipoCambioDolar = CDec(txtTipoCambio.Text)
        Estatus = "V"
        TipoMovto = "R"

        If optMonedaDolar.Checked = True Then
            Moneda = "D"
        ElseIf optMonedaPeso.Checked = True Then
            Moneda = "P"
        End If

        FolioReparacion = Trim(txtFolio.Text)
        FechaReparacion = Today
        If chkReparado.CheckState = System.Windows.Forms.CheckState.Checked Then
            Reparado = True
        Else
            Reparado = False
        End If

        'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Cnn.BeginTrans()
        ModStoredProcedures.PR_IME_ReparacionesCorpo(FolioReparacion, VB6.Format(FechaReparacion, C_FORMATFECHAGUARDAR), TipoMovto, CStr(gintCodAlmacen), CStr(gintCodCaja), CStr(intCodVendedor), CStr(intCodCliente), Trim(txtCliente.Text), Trim(txtRFCCliente.Text), txtTelefono.Text, txtDesArticulo.Text, (txtObservacionesTaller.Text), CStr(intCodTaller), CStr(intCodGrupo), Moneda, Trim(CStr(TipoCambioDolar)), CStr(mcurCostoReparacionD), CStr(mcurSubTotalImporteD), CStr(mcurIvaImporteD), CStr(mcurImporteReparacionD), CStr(mcurAnticipoD), Estatus, "01/01/1900", IIf((chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Checked), VB6.Format(dtpEntregaTaller.Value, C_FORMATFECHAGUARDAR), "01/01/1900"), IIf((chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Checked), VB6.Format(dtpRegresoTaller.Value, C_FORMATFECHAGUARDAR), "01/01/1900"), IIf((chkConfirmacionCliente.CheckState = System.Windows.Forms.CheckState.Checked), VB6.Format(dtpConfirmacionCliente.Value, C_FORMATFECHAGUARDAR), "01/01/1900"), IIf((chkEntregaCliente.CheckState = System.Windows.Forms.CheckState.Checked), VB6.Format(dtpFechaEntregaCliente.Value, C_FORMATFECHAGUARDAR), "01/01/1900"), CStr(False), CStr(Reparado), CStr(mCurPorcIva), "01/01/1900", IIf((chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Checked), VB6.Format(dtpCorpoRegresa.Value, C_FORMATFECHAGUARDAR), "01/01/1900"), Trim(Bitacora.Text), C_MODIFICACION, CStr(0))
        Cmd.Execute()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        If chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Checked And chkTallerEntrega.Enabled Then
            ImprimirTicket(2, FolioReparacion, True, "T")
        End If

        If chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Checked And chkCorpoRegresa.Enabled Then
            ImprimirTicket(2, FolioReparacion, False, "RP")
        End If

        If mblnNuevo Then
            MsgBox("La reparación ha sido grabada correctamente con el Código: " & FolioReparacion, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrCorpoNOMBREEMPRESA)
        End If
        'Dejar el Procedimiento Nuevo, sirve para que al usar limpiar,. no pregunte si se desea guardar cambios en el codigo
        Nuevo()
        Limpiar()
        'Descargar el Formulario de Pagos
        frmPagosSalMercancia.Close()
        'frmPVRegCheque.Close()
        'frmPVRegNotasCred.Close()
        'frmPVRegTarjeta_PV.Close()
        Exit Sub

Merr:
        If Err.Number <> 0 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Cnn.RollbackTrans()
            ModEstandar.MostrarError("Ocurrió un Error en el Formulario y Proceso: " & gstrProcesoqueGeneraError)
        End If
    End Sub

    Function Cambios() As Boolean
        Cambios = True
        If chkTallerEntrega.CheckState <> CDbl(chkTallerEntrega.Tag) Then Exit Function
        If chkTallerRegreso.CheckState <> CDbl(chkTallerRegreso.Tag) Then Exit Function
        If chkReparado.CheckState <> CDbl(chkReparado.Tag) Then Exit Function
        If Trim(txtDesArticulo.Text) <> Trim(txtDesArticulo.Tag) Then Exit Function
        If Trim(txtObservacionesTaller.Text) <> Trim(txtObservacionesTaller.Tag) Then Exit Function
        If Trim(dbcTaller.Text) <> Trim(dbcTaller.Tag) Then Exit Function
        If txtCosto.Text <> txtCosto.Tag Then Exit Function
        If txtImpReparacion.Text <> txtImpReparacion.Tag Then Exit Function
        If chkCorpoRegresa.CheckState <> Val(chkCorpoRegresa.Tag) Then Exit Function
        If chkConfirmacionCliente.CheckState <> CDbl(chkConfirmacionCliente.Tag) Then Exit Function
        If dtpConfirmacionCliente.Value <> CDate(dtpConfirmacionCliente.Tag) Then Exit Function
        Cambios = False
    End Function

    Sub Cancelar()
        On Error GoTo Merr
        '''NO DEBE CANCELARSE UNA REPARACION DESDE EL CONTROL ADMVO DEL CORPO

        '''    gStrSql = "SELECT * FROM Reparaciones WHERE FolioReparacion = '" & Trim(txtFolio) & "' "
        '''    ModEstandar.BorraCmd
        '''    Cmd.CommandText = "dbo.Up_Select_Datos"
        '''    Cmd.CommandType = adCmdStoredProc
        '''    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '''    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        '''    Set RsGral = Cmd.Execute
        '''    If RsGral.RecordCount = 0 Then
        '''        MsgBox "El Folio no existe." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, "Mensaje"
        '''        RsGral.Close
        '''        Exit Sub
        '''    End If
        '''    If RsGral!Estatus = "C" Then
        '''     MsgBox "El Folio ya ha sido Cancelado." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, "Mensaje"
        '''        RsGral.Close
        '''        Exit Sub
        '''    End If
        '''
        '''    gStrSql = "Select * From Ingresos Where FolioMovto = '" & Trim(txtFolio) & "' "
        '''    ModEstandar.BorraCmd
        '''    Cmd.CommandText = "dbo.Up_Select_Datos"
        '''    Cmd.CommandType = adCmdStoredProc
        '''    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '''    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        '''    Set RsGral = Cmd.Execute
        '''    If RsGral.RecordCount <> 0 Then
        '''        MsgBox "El Folio no puede ser cancelado. Existen abonos registrados." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, "Mensaje"
        '''        RsGral.Close
        '''        Exit Sub
        '''    End If
        '''    'Para Poder Cancelar, es Necesario que la Reparacion no haya sido confirmada o que no haya sido liquidada o entregada a crédito.
        '''    If chkConfirmacionCliente.Value = vbChecked Then
        '''        MsgBox "No es posible cancelar la reparación, existe confirmación por parte del cliente." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, "Mensaje"
        '''        RsGral.Close
        '''        Exit Sub
        '''    End If
        '''    If chkCredito.Value = vbChecked Then
        '''        MsgBox "No es posible cancelar la reparación, los artículos se han entregado a crédito al cliente." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, "Mensaje"
        '''        RsGral.Close
        '''        Exit Sub
        '''    End If
        '''    'Preguntar si desea borrar el registro
        '''    Select Case MsgBox(C_msgCANCELARFOLIO, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Mensaje")
        '''        Case vbNo
        '''          Exit Sub
        '''        Case vbCancel
        '''          Exit Sub
        '''    End Select
        '''
        '''    Screen.MousePointer = vbHourglass
        '''    Cnn.BeginTrans
        '''    ModStoredProcedures.PR_IME_ReparacionesCorpo txtFolio, CStr(dtpFechaReparacion), "", CStr(gintCodAlmacen), CStr(gintCodCaja), CStr(intCodVendedor), CStr(intCodCliente), Trim(txtCliente), Trim(txtRFCCliente), txtTelefono, txtDesArticulo, txtObservacionesTaller.text, CStr(intCodTaller), CStr(intCodGrupo), "", 0, CStr(txtCosto), 0, 0, 0, 0, "", Format(Date, C_FORMATFECHAGUARDAR), IIf((chkTallerEntrega.Value = vbChecked), CStr(dtpEntregaTaller), "01/01/1900"), IIf((chkTallerRegreso.Value = vbChecked), CStr(dtpRegresoTaller), "01/01/1900"), IIf((chkConfirmacionCliente.Value = vbChecked), CStr(dtpConfirmacionCliente), "01/01/1900"), IIf((dtpFechaEntregaCliente = ""), "01/01/1900", CStr(dtpFechaEntregaCliente)), False, False, C_ELIMINACION, 0
        '''    Cmd.Execute
        '''    Cnn.CommitTrans
        '''
        '''    Nuevo
        '''    Limpiar
        '''    Screen.MousePointer = vbDefault
        '''    Exit Sub

Merr:
        Cnn.RollbackTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub DesHabilitarControles()
        'Este Procedimiento deshabilita todos los controles del formulario de ventas. para cuando es una consulta, y no es posible modificar nada absolutamente.
        dbcVendedor.Enabled = False
        fraDatosCliente.Enabled = False
        '''fraMotivoReparacion.Enabled = False
        fraObservacionesTaller.Enabled = False
        btnCatClientes.Enabled = False
        dbcTipoReparacion.Enabled = False
        dbcTaller.Enabled = False
        txtCosto.Enabled = False
        txtAnticipo.Enabled = False
        txtImpReparacion.Enabled = False
        txtAbonos.Enabled = False
        txtNuevoAbono.Enabled = False
        fraMoneda.Enabled = False
        optMonedaDolar.Enabled = False
        optMonedaPeso.Enabled = False
        txtImpReparacion.Enabled = False
        txtCosto.Enabled = False
        dtpRegresoTaller.Enabled = False
        dtpEntregaTaller.Enabled = False
        chkTallerEntrega.Enabled = False
        chkTallerRegreso.Enabled = False
        chkConfirmacionCliente.Enabled = False
        chkCredito.Enabled = False
        chkEntregaCliente.Enabled = False
    End Sub

    Sub DesHabilitarControlesFolioCancelado()
        'Este Procedimiento deshabilita todos los controles del formulario de ventas. para cuando es una consulta, y no es posible modificar nada absolutamente.
        dbcVendedor.Enabled = False
        fraDatosCliente.Enabled = False
        '''fraMotivoReparacion.Enabled = False
        fraObservacionesTaller.Enabled = False
        btnCatClientes.Enabled = False
        dbcTipoReparacion.Enabled = False
        dbcTaller.Enabled = False
        txtCosto.Enabled = False
        txtAnticipo.Enabled = False
        txtImpReparacion.Enabled = False
        txtAbonos.Enabled = False
        chkConfirmacionCliente.Enabled = False
        chkCredito.Enabled = False
        chkEntregaCliente.Enabled = False
        chkReparado.Enabled = False
        chkTallerEntrega.Enabled = False
        chkTallerRegreso.Enabled = False
        chkCorpoEnvio.Enabled = False
        chkCorpoRegresa.Enabled = False
        txtNuevoAbono.Enabled = False
        fraMoneda.Enabled = False
    End Sub

    Sub HabilitarControles()
        dbcVendedor.Enabled = False
        fraDatosCliente.Enabled = False
        '''fraMotivoReparacion.Enabled = False
        fraObservacionesTaller.Enabled = True
        btnCatClientes.Enabled = False
        dbcTipoReparacion.Enabled = False
        txtCosto.Enabled = True
        txtAnticipo.Enabled = False
        txtImpReparacion.Enabled = True
        txtAbonos.Enabled = False
        txtNuevoAbono.Enabled = False
        fraMoneda.Enabled = False
        optMonedaDolar.Enabled = False
        optMonedaPeso.Enabled = False
        fraCliente.Enabled = False
        '    txtImpReparacion.Enabled = True
        '    txtCosto.Enabled = True
        '    dtpRegresoTaller.Enabled = True
        '    dtpEntregaTaller.Enabled = True
        '    chkTallerEntrega.Enabled = True
        '    chkTallerRegreso.Enabled = True
        '    chkConfirmacionCliente.Enabled = True
        '    chkCredito.Enabled = True
        '    chkEntregaCliente.Enabled = True
    End Sub

    Function CalcularDiasTranscurridos() As String
        'Obtiene el Número de Días que han transcurrido desde que se dejo el Artículo para la reparación hasta el día actual.
        Dim FechaReparacion As String
        FechaReparacion = CStr(CDate(dtpFechaReparacion.Value))
        'Si ya se ha dado la fecha de entega al Cliente, el número de días transcurridos se calcula con esa fecha.
        'De lo contrario, se calculará con la fecha actual
        If chkEntregaCliente.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            CalcularDiasTranscurridos = CStr(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(FechaReparacion), Today))
        Else
            CalcularDiasTranscurridos = CStr(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(FechaReparacion), CDate(dtpFechaEntregaCliente.Value)))
        End If
    End Function

    Sub PonerTotalesenFrmPagos()
        'Este Proc. Transporta los datos existentes en el Form de Control de Reparaciones, al Formulario de PAgos.
        'Los Datos que se pasan son: Subtotal, IVa,  Total,  Tpo Dólar.
        On Error GoTo Merr
        Dim SubtotalD As Decimal
        Dim ImpIva As Decimal
        Dim Descuento As Decimal
        Dim TipoCambioDolar As Decimal
        Dim TotalP As Decimal
        Dim TotalD As Decimal
        '    Dim PrecioListaSinIva As Currency
        '    Dim IvaReal As Currency

        ' En el Formulario de Pagos, existe un  Textbox apra alamacenar el Nombre del FOrmulario que ha invocado al FOrmulario de Pagos.
        ' Para posteriormente saber que función de Guardar se jecuta. (En este momento el FOrmulario de Pagos, se usa en en VEntas y Apartados.)
        '''frmPagosSalMercancia.txtFormaOrigen = frmControlReparaciones.Name

        'Si es un Folio nuevo, se Pasará el Anticipo, de lo contrario se pasará el Abono
        TipoCambioDolar = CDec(Numerico(txtTipoCambio.Text))
        If mblnNuevo = True And CDbl(Numerico(txtAnticipo.Text)) > 0 Then
            TotalP = FormateoDecimales(txtAnticipo)
            SubtotalD = FormateoDecimales(TotalP / TipoCambioDolar)
            ImpIva = FormateoDecimales(SubtotalD * (gcurCorpoTASAIVA / 100) / (1 + gcurCorpoTASAIVA / 100))
            SubtotalD = FormateoDecimales(SubtotalD - ImpIva)
            TotalD = FormateoDecimales(SubtotalD + ImpIva)
            Descuento = FormateoDecimales(0)
        Else
            TotalP = FormateoDecimales(txtNuevoAbono)
            SubtotalD = FormateoDecimales(TotalP / TipoCambioDolar)
            ImpIva = FormateoDecimales(SubtotalD * (gcurCorpoTASAIVA / 100) / (1 + gcurCorpoTASAIVA / 100))
            SubtotalD = FormateoDecimales(SubtotalD - ImpIva)
            TotalD = FormateoDecimales(SubtotalD + ImpIva)
            Descuento = FormateoDecimales(0)
        End If

        With frmPagosSalMercancia
            .txtSubtotal.Text = VB6.Format(FormateoDecimales(SubtotalD), gstrFormatoCantidad)
            .txtIVA.Text = VB6.Format(FormateoDecimales(ImpIva), gstrFormatoCantidad)
            .txtTotal.Text = VB6.Format(FormateoDecimales(TotalD), gstrFormatoCantidad)
            .txtDescuento.Text = VB6.Format(Descuento, gstrFormatoCantidad)
            .txtmnAPagar.Text = VB6.Format(FormateoDecimales(TotalP), gstrFormatoCantidad)
            .txtmnTotalPago.Text = VB6.Format(FormateoDecimales(0), gstrFormatoCantidad)
            .txtmnCambio.Text = VB6.Format(FormateoDecimales(CDbl(.txtmnTotalPago.Text) - CDbl(.txtmnAPagar.Text)), gstrFormatoCantidad)
            .txtdoAPagar.Text = VB6.Format(FormateoDecimales(TotalD), gstrFormatoCantidad)
            .txtdoTotalPago.Text = VB6.Format(FormateoDecimales(0), gstrFormatoCantidad)
            .txtdoCambio.Text = VB6.Format(FormateoDecimales(CDbl(.txtdoTotalPago.Text) - CDbl(.txtdoAPagar.Text)), gstrFormatoCantidad)
            .txtDolar.Text = VB6.Format(FormateoDecimales(TipoCambioDolar), gstrFormatoCantidad)
        End With
        'Validar el Importe del Pago que se ha hecho. Para calcular el Saldo y el CAmbio.
        frmPagosSalMercancia.ValidarImportedeFormaPago()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub
    Private Sub txtImpReparacion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImpReparacion.TextChanged
        Dim TotalAbonos As Decimal
        TotalAbonos = Val(Numerico(txtAbonos.Text)) + Val(Numerico(txtAnticipo.Text))
        txtSaldo.Text = VB6.Format(TotalAbonos - CDbl(Numerico(txtImpReparacion.Text)), gstrFormatoCantidad)
    End Sub

    Private Sub txtImpReparacion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImpReparacion.Enter
        SelTextoTxt(txtImpReparacion)
        Pon_Tool()
    End Sub

    Private Sub txtImpReparacion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImpReparacion.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtImpReparacion.Text, KeyAscii, 8, 2, (txtImpReparacion.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtImpReparacion_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImpReparacion.Leave
        If CDec(Numerico((txtImpReparacion.Text))) < CDec(Numerico((txtCosto.Text))) Then
            MsgBox("El importe de Venta de la Reparacion debe ser mayor o igual que el costo" & vbNewLine & "Favor de verficar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            txtImpReparacion.Text = "0.00"
        End If
        txtImpReparacion.Text = FormateoDecimales(Numerico(txtImpReparacion.Text))
        txtImpReparacion.Text = VB6.Format(txtImpReparacion.Text, gstrFormatoCantidad)
    End Sub

    Private Sub txtNuevoAbono_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNuevoAbono.TextChanged
        Dim Abonos As Decimal
        mblnCambiosEnCodigo = True
        Abonos = CDec(Numerico(txtNuevoAbono.Text)) + CDec(Numerico(txtAnticipo.Text)) + CDec(Numerico(txtAbonos.Text))
        txtSaldo.Text = VB6.Format(Abonos - CDbl(Numerico(txtImpReparacion.Text)), gstrFormatoCantidad)
    End Sub

    Private Sub txtnuevoabono_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNuevoAbono.Enter
        SelTextoTxt(txtNuevoAbono)
        Pon_Tool()
    End Sub

    Private Sub txtNuevoAbono_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNuevoAbono.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtNuevoAbono.Text, KeyAscii, 8, 2, (txtNuevoAbono.SelectionStart))
        'KeyAscii = ModEstandar.MskCantidad(txtAbono, KeyAscii, 8, 2, txtAbono.SelStart)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtNuevoAbono_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNuevoAbono.Leave
        txtNuevoAbono.Text = VB6.Format(txtNuevoAbono.Text, gstrFormatoCantidad)
    End Sub

    Private Sub txtNuevoTexto_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtNuevoTexto.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then cmdRegistrar_Click(cmdRegistrar, New System.EventArgs())
    End Sub

    Private Sub txtNuevoTexto_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNuevoTexto.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, ".,:\/-_()[]{}%$#&?!@<>*+")
        ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtNuevoTexto_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNuevoTexto.Leave
        txtNuevoTexto.Text = ModEstandar.QuitaEnter((txtNuevoTexto.Text))
    End Sub

    Private Sub txtObservacionesTaller_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtObservacionesTaller.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, ".,/-_:()[]{}")
        ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtObservacionesTaller_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtObservacionesTaller.Leave
        txtObservacionesTaller.Text = ModEstandar.QuitaEnter(txtObservacionesTaller.Text)
    End Sub

    Private Sub txtSaldo_Change()
        If CDbl(Numerico(txtSaldo.Text)) < 0 Then
            'Fondo Rojo
            txtSaldo.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
            txtSaldo.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
        Else
            txtSaldo.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000018)
            txtSaldo.ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
        End If
    End Sub

    Sub LimpiaDatosCliente()
        On Error GoTo Merr
        'Este Proc. Muestra los datos de un Cliente al seleccionarlo del DataCombo
        txtRFCCliente.Text = ""
        txtRFCCliente.Tag = ""
        txtDomicilio.Text = ""
        txtDomicilio.Tag = ""
        txtTelefono.Text = ""
        txtTelefono.Tag = ""
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Function RegresaDatoBitacora(ByRef FolioRep As String) As String
        Dim lSql As String
        Dim Rs As ADODB.Recordset
        On Error GoTo Merr

        RegresaDatoBitacora = ""
        lSql = "Select FolioReparacion, Bitacora From Reparaciones Where FolioReparacion = '" & FolioRep & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, lSql))
        Rs = Cmd.Execute

        If Rs.RecordCount > 0 Then
            RegresaDatoBitacora = Trim(Rs.Fields("Bitacora").Value)
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Private Sub GeneraBitacora()
        Dim lReparado As String

        lReparado = ""
        Bitacora.Text = ""
        Bitacora.Text = "PUNTO VENTA RECIBIO - " & UCase(VB6.Format(dtpFechaReparacion.Value, C_FORMATFECHAMOSTRAR)) & vbNewLine & vbNewLine

        If chkCorpoEnvio.CheckState = System.Windows.Forms.CheckState.Checked Then
            Bitacora.Text = Bitacora.Text & "ENVIO A CORPORATIVO - " & UCase(VB6.Format(dtpCorpoEnvio.Value, C_FORMATFECHAMOSTRAR)) & vbNewLine & vbNewLine
        End If
        If chkTallerEntrega.CheckState = System.Windows.Forms.CheckState.Checked Then
            Bitacora.Text = Bitacora.Text & "ENVIO A TALLER - " & UCase(VB6.Format(dtpEntregaTaller.Value, C_FORMATFECHAMOSTRAR)) & vbNewLine
            Bitacora.Text = Bitacora.Text & Trim(dbcTaller.Text) & vbNewLine & vbNewLine
        End If
        If chkConfirmacionCliente.CheckState = System.Windows.Forms.CheckState.Checked Then
            Bitacora.Text = Bitacora.Text & "CLIENTE AUTORIZA - " & UCase(VB6.Format(dtpConfirmacionCliente.Value, C_FORMATFECHAMOSTRAR)) & vbNewLine & vbNewLine
        End If
        If chkTallerRegreso.CheckState = System.Windows.Forms.CheckState.Checked Then
            Bitacora.Text = Bitacora.Text & "REGRESO DEL TALLER AL CORPORATIVO - " & UCase(VB6.Format(dtpRegresoTaller.Value, C_FORMATFECHAMOSTRAR)) & vbNewLine
            lReparado = IIf(chkReparado.CheckState = 1, "REPARADO", "NO REPARADO")
            Bitacora.Text = Bitacora.Text & lReparado & vbNewLine
        End If
        If chkCorpoRegresa.CheckState = System.Windows.Forms.CheckState.Checked Then
            Bitacora.Text = Bitacora.Text & "ENVIO A PUNTO VENTA - " & UCase(VB6.Format(dtpCorpoRegresa.Value, C_FORMATFECHAMOSTRAR)) & vbNewLine & vbNewLine
        End If
        If chkEntregaCliente.CheckState = System.Windows.Forms.CheckState.Checked Then
            Bitacora.Text = Bitacora.Text & "CLIENTE RECIBE - " & UCase(VB6.Format(dtpFechaEntregaCliente.Value, C_FORMATFECHAMOSTRAR)) & vbNewLine
            If chkCredito.CheckState = System.Windows.Forms.CheckState.Checked Then
                Bitacora.Text = Bitacora.Text & "CREDITO" & vbNewLine & vbNewLine
            Else
                Bitacora.Text = Bitacora.Text & vbNewLine
            End If
        End If

    End Sub

    Sub ImprimirTicket(ByRef NoImpresiones As Integer, ByRef FolioReparacion As String, ByRef lNuevo As Boolean, ByRef lTipoL As String)
        For I = 1 To NoImpresiones
            ModCorporativo.TicketReparacion(FolioReparacion, lNuevo, lTipoL)
        Next
    End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtNuevoTexto = New System.Windows.Forms.TextBox()
        Me.txtDomicilio = New System.Windows.Forms.TextBox()
        Me.txtTelefono = New System.Windows.Forms.TextBox()
        Me.txtRFCCliente = New System.Windows.Forms.TextBox()
        Me.txtCliente = New System.Windows.Forms.TextBox()
        Me.btnCatClientes = New System.Windows.Forms.Button()
        Me.txtDesArticulo = New System.Windows.Forms.TextBox()
        Me.txtObservacionesTaller = New System.Windows.Forms.TextBox()
        Me.txtNuevoAbono = New System.Windows.Forms.TextBox()
        Me.txtAbonos = New System.Windows.Forms.TextBox()
        Me.txtImpReparacion = New System.Windows.Forms.TextBox()
        Me.txtAnticipo = New System.Windows.Forms.TextBox()
        Me.txtCosto = New System.Windows.Forms.TextBox()
        Me.txtSaldo = New System.Windows.Forms.Label()
        Me.txtFolio = New System.Windows.Forms.TextBox()
        Me.txtTipoCambio = New System.Windows.Forms.TextBox()
        Me.txtDiasTranscurridos = New System.Windows.Forms.TextBox()
        Me.optMonedaDolar = New System.Windows.Forms.RadioButton()
        Me.optMonedaPeso = New System.Windows.Forms.RadioButton()
        Me.Marco = New System.Windows.Forms.GroupBox()
        Me.fraNuevoTexto = New System.Windows.Forms.GroupBox()
        Me.cmdRegistrar = New System.Windows.Forms.Button()
        Me.fraBitacora = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.Panel()
        Me.Bitacora = New System.Windows.Forms.RichTextBox()
        Me.fraDatosCliente = New System.Windows.Forms.GroupBox()
        Me.txtCodCliente = New System.Windows.Forms.TextBox()
        Me._lblVentas_4 = New System.Windows.Forms.Label()
        Me._lblVentas_3 = New System.Windows.Forms.Label()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me.fraCliente = New System.Windows.Forms.GroupBox()
        Me.chkConfirmacionCliente = New System.Windows.Forms.CheckBox()
        Me.chkEntregaCliente = New System.Windows.Forms.CheckBox()
        Me.dtpConfirmacionCliente = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaEntregaCliente = New System.Windows.Forms.DateTimePicker()
        Me.fraTaller = New System.Windows.Forms.GroupBox()
        Me.chkTallerEntrega = New System.Windows.Forms.CheckBox()
        Me.chkTallerRegreso = New System.Windows.Forms.CheckBox()
        Me.chkReparado = New System.Windows.Forms.CheckBox()
        Me.dtpEntregaTaller = New System.Windows.Forms.DateTimePicker()
        Me.dtpRegresoTaller = New System.Windows.Forms.DateTimePicker()
        Me.fraMotivoReparacion = New System.Windows.Forms.GroupBox()
        Me.fraObservacionesTaller = New System.Windows.Forms.GroupBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.dbcTipoReparacion = New System.Windows.Forms.ComboBox()
        Me.dbcTaller = New System.Windows.Forms.ComboBox()
        Me._lblReparaciones_7 = New System.Windows.Forms.Label()
        Me._lblReparaciones_5 = New System.Windows.Forms.Label()
        Me.Frame6 = New System.Windows.Forms.GroupBox()
        Me.fraLinea = New System.Windows.Forms.GroupBox()
        Me.chkCredito = New System.Windows.Forms.CheckBox()
        Me.lblNuevoAbono = New System.Windows.Forms.Label()
        Me._lblReparaciones_11 = New System.Windows.Forms.Label()
        Me._lblReparaciones_8 = New System.Windows.Forms.Label()
        Me._lblReparaciones_6 = New System.Windows.Forms.Label()
        Me._lblReparaciones_99 = New System.Windows.Forms.Label()
        Me._lblReparaciones_9 = New System.Windows.Forms.Label()
        Me.Frame7 = New System.Windows.Forms.GroupBox()
        Me.Frame5 = New System.Windows.Forms.Panel()
        Me.dtpFechaReparacion = New System.Windows.Forms.DateTimePicker()
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me._lblReparaciones_2 = New System.Windows.Forms.Label()
        Me._lblReparaciones_3 = New System.Windows.Forms.Label()
        Me.fraMoneda = New System.Windows.Forms.GroupBox()
        Me.dbcVendedor = New System.Windows.Forms.ComboBox()
        Me._lblReparaciones_4 = New System.Windows.Forms.Label()
        Me._lblReparaciones_0 = New System.Windows.Forms.Label()
        Me._lblReparaciones_1 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.chkCorpoRegresa = New System.Windows.Forms.CheckBox()
        Me.chkCorpoEnvio = New System.Windows.Forms.CheckBox()
        Me.dtpCorpoEnvio = New System.Windows.Forms.DateTimePicker()
        Me.dtpCorpoRegresa = New System.Windows.Forms.DateTimePicker()
        Me.lblEstatus = New System.Windows.Forms.Label()
        Me.lblCancelado = New System.Windows.Forms.Label()
        Me.lblLiquidado = New System.Windows.Forms.Label()
        Me.lblReparaciones = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Marco.SuspendLayout()
        Me.fraNuevoTexto.SuspendLayout()
        Me.fraBitacora.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.fraDatosCliente.SuspendLayout()
        Me.fraCliente.SuspendLayout()
        Me.fraTaller.SuspendLayout()
        Me.fraMotivoReparacion.SuspendLayout()
        Me.fraObservacionesTaller.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame6.SuspendLayout()
        Me.Frame7.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.fraMoneda.SuspendLayout()
        Me.Frame2.SuspendLayout()
        CType(Me.lblReparaciones, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtNuevoTexto
        '
        Me.txtNuevoTexto.AcceptsReturn = True
        Me.txtNuevoTexto.BackColor = System.Drawing.SystemColors.Window
        Me.txtNuevoTexto.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNuevoTexto.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNuevoTexto.Location = New System.Drawing.Point(10, 17)
        Me.txtNuevoTexto.MaxLength = 500
        Me.txtNuevoTexto.Multiline = True
        Me.txtNuevoTexto.Name = "txtNuevoTexto"
        Me.txtNuevoTexto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNuevoTexto.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtNuevoTexto.Size = New System.Drawing.Size(290, 47)
        Me.txtNuevoTexto.TabIndex = 30
        Me.ToolTip1.SetToolTip(Me.txtNuevoTexto, "Observaciones del Taller")
        '
        'txtDomicilio
        '
        Me.txtDomicilio.AcceptsReturn = True
        Me.txtDomicilio.BackColor = System.Drawing.SystemColors.Info
        Me.txtDomicilio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDomicilio.Enabled = False
        Me.txtDomicilio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDomicilio.Location = New System.Drawing.Point(77, 76)
        Me.txtDomicilio.MaxLength = 0
        Me.txtDomicilio.Multiline = True
        Me.txtDomicilio.Name = "txtDomicilio"
        Me.txtDomicilio.ReadOnly = True
        Me.txtDomicilio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDomicilio.Size = New System.Drawing.Size(300, 45)
        Me.txtDomicilio.TabIndex = 24
        Me.ToolTip1.SetToolTip(Me.txtDomicilio, "Domicilio del Cliente")
        '
        'txtTelefono
        '
        Me.txtTelefono.AcceptsReturn = True
        Me.txtTelefono.BackColor = System.Drawing.SystemColors.Info
        Me.txtTelefono.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTelefono.Enabled = False
        Me.txtTelefono.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTelefono.Location = New System.Drawing.Point(77, 127)
        Me.txtTelefono.MaxLength = 0
        Me.txtTelefono.Name = "txtTelefono"
        Me.txtTelefono.ReadOnly = True
        Me.txtTelefono.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTelefono.Size = New System.Drawing.Size(300, 20)
        Me.txtTelefono.TabIndex = 26
        Me.ToolTip1.SetToolTip(Me.txtTelefono, "Telefono del Cliente")
        '
        'txtRFCCliente
        '
        Me.txtRFCCliente.AcceptsReturn = True
        Me.txtRFCCliente.BackColor = System.Drawing.SystemColors.Info
        Me.txtRFCCliente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRFCCliente.Enabled = False
        Me.txtRFCCliente.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRFCCliente.Location = New System.Drawing.Point(77, 50)
        Me.txtRFCCliente.MaxLength = 0
        Me.txtRFCCliente.Name = "txtRFCCliente"
        Me.txtRFCCliente.ReadOnly = True
        Me.txtRFCCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRFCCliente.Size = New System.Drawing.Size(300, 20)
        Me.txtRFCCliente.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.txtRFCCliente, "RFC del Cliente")
        '
        'txtCliente
        '
        Me.txtCliente.AcceptsReturn = True
        Me.txtCliente.BackColor = System.Drawing.SystemColors.Window
        Me.txtCliente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCliente.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCliente.Location = New System.Drawing.Point(77, 24)
        Me.txtCliente.MaxLength = 0
        Me.txtCliente.Name = "txtCliente"
        Me.txtCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCliente.Size = New System.Drawing.Size(300, 20)
        Me.txtCliente.TabIndex = 20
        Me.ToolTip1.SetToolTip(Me.txtCliente, "Nombre del cliente")
        '
        'btnCatClientes
        '
        Me.btnCatClientes.BackColor = System.Drawing.SystemColors.Control
        Me.btnCatClientes.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCatClientes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCatClientes.Location = New System.Drawing.Point(563, 546)
        Me.btnCatClientes.Name = "btnCatClientes"
        Me.btnCatClientes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCatClientes.Size = New System.Drawing.Size(108, 37)
        Me.btnCatClientes.TabIndex = 75
        Me.btnCatClientes.Text = "ABC de &Clientes"
        Me.ToolTip1.SetToolTip(Me.btnCatClientes, "Registro de Clientes")
        Me.btnCatClientes.UseVisualStyleBackColor = False
        '
        'txtDesArticulo
        '
        Me.txtDesArticulo.AcceptsReturn = True
        Me.txtDesArticulo.BackColor = System.Drawing.SystemColors.Window
        Me.txtDesArticulo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDesArticulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDesArticulo.Location = New System.Drawing.Point(10, 21)
        Me.txtDesArticulo.MaxLength = 4000
        Me.txtDesArticulo.Multiline = True
        Me.txtDesArticulo.Name = "txtDesArticulo"
        Me.txtDesArticulo.ReadOnly = True
        Me.txtDesArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDesArticulo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDesArticulo.Size = New System.Drawing.Size(368, 80)
        Me.txtDesArticulo.TabIndex = 28
        Me.ToolTip1.SetToolTip(Me.txtDesArticulo, "Descripción de Artículos de la Reparación")
        '
        'txtObservacionesTaller
        '
        Me.txtObservacionesTaller.AcceptsReturn = True
        Me.txtObservacionesTaller.BackColor = System.Drawing.SystemColors.Window
        Me.txtObservacionesTaller.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtObservacionesTaller.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtObservacionesTaller.Location = New System.Drawing.Point(10, 22)
        Me.txtObservacionesTaller.MaxLength = 4000
        Me.txtObservacionesTaller.Multiline = True
        Me.txtObservacionesTaller.Name = "txtObservacionesTaller"
        Me.txtObservacionesTaller.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtObservacionesTaller.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtObservacionesTaller.Size = New System.Drawing.Size(368, 80)
        Me.txtObservacionesTaller.TabIndex = 33
        Me.ToolTip1.SetToolTip(Me.txtObservacionesTaller, "Observaciones del Taller")
        '
        'txtNuevoAbono
        '
        Me.txtNuevoAbono.AcceptsReturn = True
        Me.txtNuevoAbono.BackColor = System.Drawing.SystemColors.Window
        Me.txtNuevoAbono.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNuevoAbono.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNuevoAbono.Location = New System.Drawing.Point(107, 153)
        Me.txtNuevoAbono.MaxLength = 0
        Me.txtNuevoAbono.Name = "txtNuevoAbono"
        Me.txtNuevoAbono.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNuevoAbono.Size = New System.Drawing.Size(97, 20)
        Me.txtNuevoAbono.TabIndex = 70
        Me.txtNuevoAbono.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtNuevoAbono, "Nuevo Abono")
        Me.txtNuevoAbono.Visible = False
        '
        'txtAbonos
        '
        Me.txtAbonos.AcceptsReturn = True
        Me.txtAbonos.BackColor = System.Drawing.SystemColors.Window
        Me.txtAbonos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAbonos.Enabled = False
        Me.txtAbonos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAbonos.Location = New System.Drawing.Point(107, 125)
        Me.txtAbonos.MaxLength = 0
        Me.txtAbonos.Name = "txtAbonos"
        Me.txtAbonos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAbonos.Size = New System.Drawing.Size(97, 20)
        Me.txtAbonos.TabIndex = 69
        Me.txtAbonos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtAbonos, "Abonos")
        '
        'txtImpReparacion
        '
        Me.txtImpReparacion.AcceptsReturn = True
        Me.txtImpReparacion.BackColor = System.Drawing.SystemColors.Window
        Me.txtImpReparacion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImpReparacion.Enabled = False
        Me.txtImpReparacion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImpReparacion.Location = New System.Drawing.Point(107, 67)
        Me.txtImpReparacion.MaxLength = 0
        Me.txtImpReparacion.Name = "txtImpReparacion"
        Me.txtImpReparacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImpReparacion.Size = New System.Drawing.Size(97, 20)
        Me.txtImpReparacion.TabIndex = 67
        Me.txtImpReparacion.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtImpReparacion, "Importe de la Reparación")
        '
        'txtAnticipo
        '
        Me.txtAnticipo.AcceptsReturn = True
        Me.txtAnticipo.BackColor = System.Drawing.SystemColors.Window
        Me.txtAnticipo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAnticipo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnticipo.Location = New System.Drawing.Point(107, 96)
        Me.txtAnticipo.MaxLength = 0
        Me.txtAnticipo.Name = "txtAnticipo"
        Me.txtAnticipo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAnticipo.Size = New System.Drawing.Size(97, 20)
        Me.txtAnticipo.TabIndex = 68
        Me.txtAnticipo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtAnticipo, "Anticipo")
        '
        'txtCosto
        '
        Me.txtCosto.AcceptsReturn = True
        Me.txtCosto.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtCosto.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCosto.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCosto.Location = New System.Drawing.Point(107, 39)
        Me.txtCosto.MaxLength = 0
        Me.txtCosto.Name = "txtCosto"
        Me.txtCosto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCosto.Size = New System.Drawing.Size(97, 20)
        Me.txtCosto.TabIndex = 66
        Me.txtCosto.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCosto, "Costo de la Reparación")
        '
        'txtSaldo
        '
        Me.txtSaldo.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.txtSaldo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtSaldo.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtSaldo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtSaldo.Location = New System.Drawing.Point(107, 181)
        Me.txtSaldo.Name = "txtSaldo"
        Me.txtSaldo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSaldo.Size = New System.Drawing.Size(97, 21)
        Me.txtSaldo.TabIndex = 71
        Me.txtSaldo.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.txtSaldo, "Saldo de la Reparación")
        '
        'txtFolio
        '
        Me.txtFolio.AcceptsReturn = True
        Me.txtFolio.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolio.Location = New System.Drawing.Point(81, 17)
        Me.txtFolio.MaxLength = 19
        Me.txtFolio.Name = "txtFolio"
        Me.txtFolio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolio.Size = New System.Drawing.Size(145, 20)
        Me.txtFolio.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtFolio, "Folio de Reparación< ENTER = Nuevo >")
        '
        'txtTipoCambio
        '
        Me.txtTipoCambio.AcceptsReturn = True
        Me.txtTipoCambio.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.txtTipoCambio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambio.Location = New System.Drawing.Point(116, 33)
        Me.txtTipoCambio.MaxLength = 0
        Me.txtTipoCambio.Name = "txtTipoCambio"
        Me.txtTipoCambio.ReadOnly = True
        Me.txtTipoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambio.Size = New System.Drawing.Size(69, 20)
        Me.txtTipoCambio.TabIndex = 16
        Me.txtTipoCambio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambio, "Tipo de Cambio Actual")
        '
        'txtDiasTranscurridos
        '
        Me.txtDiasTranscurridos.AcceptsReturn = True
        Me.txtDiasTranscurridos.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.txtDiasTranscurridos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDiasTranscurridos.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtDiasTranscurridos.Location = New System.Drawing.Point(12, 33)
        Me.txtDiasTranscurridos.MaxLength = 0
        Me.txtDiasTranscurridos.Name = "txtDiasTranscurridos"
        Me.txtDiasTranscurridos.ReadOnly = True
        Me.txtDiasTranscurridos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDiasTranscurridos.Size = New System.Drawing.Size(92, 20)
        Me.txtDiasTranscurridos.TabIndex = 15
        Me.txtDiasTranscurridos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDiasTranscurridos, "Días Transcurridos")
        '
        'optMonedaDolar
        '
        Me.optMonedaDolar.BackColor = System.Drawing.SystemColors.Control
        Me.optMonedaDolar.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMonedaDolar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.optMonedaDolar.Location = New System.Drawing.Point(125, 24)
        Me.optMonedaDolar.Name = "optMonedaDolar"
        Me.optMonedaDolar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMonedaDolar.Size = New System.Drawing.Size(61, 17)
        Me.optMonedaDolar.TabIndex = 11
        Me.optMonedaDolar.TabStop = True
        Me.optMonedaDolar.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me.optMonedaDolar, "Moneda Dólares")
        Me.optMonedaDolar.UseVisualStyleBackColor = False
        '
        'optMonedaPeso
        '
        Me.optMonedaPeso.BackColor = System.Drawing.SystemColors.Control
        Me.optMonedaPeso.Checked = True
        Me.optMonedaPeso.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMonedaPeso.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.optMonedaPeso.Location = New System.Drawing.Point(29, 24)
        Me.optMonedaPeso.Name = "optMonedaPeso"
        Me.optMonedaPeso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMonedaPeso.Size = New System.Drawing.Size(76, 17)
        Me.optMonedaPeso.TabIndex = 10
        Me.optMonedaPeso.TabStop = True
        Me.optMonedaPeso.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me.optMonedaPeso, "Monea Pesos")
        Me.optMonedaPeso.UseVisualStyleBackColor = False
        '
        'Marco
        '
        Me.Marco.BackColor = System.Drawing.SystemColors.Control
        Me.Marco.Controls.Add(Me.fraNuevoTexto)
        Me.Marco.Controls.Add(Me.fraBitacora)
        Me.Marco.Controls.Add(Me.fraDatosCliente)
        Me.Marco.Controls.Add(Me.btnCatClientes)
        Me.Marco.Controls.Add(Me.fraCliente)
        Me.Marco.Controls.Add(Me.fraTaller)
        Me.Marco.Controls.Add(Me.fraMotivoReparacion)
        Me.Marco.Controls.Add(Me.fraObservacionesTaller)
        Me.Marco.Controls.Add(Me.Frame4)
        Me.Marco.Controls.Add(Me.Frame6)
        Me.Marco.Controls.Add(Me.Frame7)
        Me.Marco.Controls.Add(Me.Frame2)
        Me.Marco.Controls.Add(Me.lblEstatus)
        Me.Marco.Controls.Add(Me.lblCancelado)
        Me.Marco.Controls.Add(Me.lblLiquidado)
        Me.Marco.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Marco.Location = New System.Drawing.Point(8, 3)
        Me.Marco.Name = "Marco"
        Me.Marco.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Marco.Size = New System.Drawing.Size(904, 594)
        Me.Marco.TabIndex = 0
        Me.Marco.TabStop = False
        '
        'fraNuevoTexto
        '
        Me.fraNuevoTexto.BackColor = System.Drawing.SystemColors.Control
        Me.fraNuevoTexto.Controls.Add(Me.cmdRegistrar)
        Me.fraNuevoTexto.Controls.Add(Me.txtNuevoTexto)
        Me.fraNuevoTexto.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraNuevoTexto.Location = New System.Drawing.Point(12, 389)
        Me.fraNuevoTexto.Name = "fraNuevoTexto"
        Me.fraNuevoTexto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraNuevoTexto.Size = New System.Drawing.Size(389, 75)
        Me.fraNuevoTexto.TabIndex = 29
        Me.fraNuevoTexto.TabStop = False
        '
        'cmdRegistrar
        '
        Me.cmdRegistrar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRegistrar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRegistrar.Font = New System.Drawing.Font("Tahoma", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRegistrar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRegistrar.Location = New System.Drawing.Point(308, 17)
        Me.cmdRegistrar.Name = "cmdRegistrar"
        Me.cmdRegistrar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRegistrar.Size = New System.Drawing.Size(78, 47)
        Me.cmdRegistrar.TabIndex = 31
        Me.cmdRegistrar.Text = "REGISTRAR"
        Me.cmdRegistrar.UseVisualStyleBackColor = False
        '
        'fraBitacora
        '
        Me.fraBitacora.BackColor = System.Drawing.SystemColors.Control
        Me.fraBitacora.Controls.Add(Me.Frame3)
        Me.fraBitacora.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraBitacora.Location = New System.Drawing.Point(410, 203)
        Me.fraBitacora.Name = "fraBitacora"
        Me.fraBitacora.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraBitacora.Size = New System.Drawing.Size(259, 291)
        Me.fraBitacora.TabIndex = 39
        Me.fraBitacora.TabStop = False
        Me.fraBitacora.Text = " Bitácora "
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.Bitacora)
        Me.Frame3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame3.Enabled = False
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(8, 11)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(249, 277)
        Me.Frame3.TabIndex = 40
        '
        'Bitacora
        '
        Me.Bitacora.Location = New System.Drawing.Point(3, 6)
        Me.Bitacora.Name = "Bitacora"
        Me.Bitacora.Size = New System.Drawing.Size(243, 268)
        Me.Bitacora.TabIndex = 41
        Me.Bitacora.Text = ""
        '
        'fraDatosCliente
        '
        Me.fraDatosCliente.BackColor = System.Drawing.SystemColors.Control
        Me.fraDatosCliente.Controls.Add(Me.txtDomicilio)
        Me.fraDatosCliente.Controls.Add(Me.txtTelefono)
        Me.fraDatosCliente.Controls.Add(Me.txtRFCCliente)
        Me.fraDatosCliente.Controls.Add(Me.txtCodCliente)
        Me.fraDatosCliente.Controls.Add(Me.txtCliente)
        Me.fraDatosCliente.Controls.Add(Me._lblVentas_4)
        Me.fraDatosCliente.Controls.Add(Me._lblVentas_3)
        Me.fraDatosCliente.Controls.Add(Me._lblVentas_2)
        Me.fraDatosCliente.Controls.Add(Me._lblVentas_1)
        Me.fraDatosCliente.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraDatosCliente.Location = New System.Drawing.Point(12, 95)
        Me.fraDatosCliente.Name = "fraDatosCliente"
        Me.fraDatosCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraDatosCliente.Size = New System.Drawing.Size(389, 169)
        Me.fraDatosCliente.TabIndex = 17
        Me.fraDatosCliente.TabStop = False
        Me.fraDatosCliente.Text = " Datos Cliente "
        '
        'txtCodCliente
        '
        Me.txtCodCliente.AcceptsReturn = True
        Me.txtCodCliente.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodCliente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodCliente.Enabled = False
        Me.txtCodCliente.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodCliente.Location = New System.Drawing.Point(29, 24)
        Me.txtCodCliente.MaxLength = 0
        Me.txtCodCliente.Name = "txtCodCliente"
        Me.txtCodCliente.ReadOnly = True
        Me.txtCodCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodCliente.Size = New System.Drawing.Size(41, 20)
        Me.txtCodCliente.TabIndex = 18
        Me.txtCodCliente.Visible = False
        '
        '_lblVentas_4
        '
        Me._lblVentas_4.AutoSize = True
        Me._lblVentas_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblVentas_4.Location = New System.Drawing.Point(13, 127)
        Me._lblVentas_4.Name = "_lblVentas_4"
        Me._lblVentas_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_4.Size = New System.Drawing.Size(55, 13)
        Me._lblVentas_4.TabIndex = 25
        Me._lblVentas_4.Text = "Teléfono :"
        '
        '_lblVentas_3
        '
        Me._lblVentas_3.AutoSize = True
        Me._lblVentas_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblVentas_3.Location = New System.Drawing.Point(13, 76)
        Me._lblVentas_3.Name = "_lblVentas_3"
        Me._lblVentas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_3.Size = New System.Drawing.Size(58, 13)
        Me._lblVentas_3.TabIndex = 23
        Me._lblVentas_3.Text = "Domicilio : "
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblVentas_2.Location = New System.Drawing.Point(13, 50)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(40, 13)
        Me._lblVentas_2.TabIndex = 21
        Me._lblVentas_2.Text = "R.F.C :"
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblVentas_1.Location = New System.Drawing.Point(13, 24)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(50, 13)
        Me._lblVentas_1.TabIndex = 19
        Me._lblVentas_1.Text = "Nombre :"
        '
        'fraCliente
        '
        Me.fraCliente.BackColor = System.Drawing.SystemColors.Control
        Me.fraCliente.Controls.Add(Me.chkConfirmacionCliente)
        Me.fraCliente.Controls.Add(Me.chkEntregaCliente)
        Me.fraCliente.Controls.Add(Me.dtpConfirmacionCliente)
        Me.fraCliente.Controls.Add(Me.dtpFechaEntregaCliente)
        Me.fraCliente.Enabled = False
        Me.fraCliente.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraCliente.Location = New System.Drawing.Point(678, 286)
        Me.fraCliente.Name = "fraCliente"
        Me.fraCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCliente.Size = New System.Drawing.Size(214, 81)
        Me.fraCliente.TabIndex = 53
        Me.fraCliente.TabStop = False
        Me.fraCliente.Text = " Cliente "
        '
        'chkConfirmacionCliente
        '
        Me.chkConfirmacionCliente.BackColor = System.Drawing.SystemColors.Control
        Me.chkConfirmacionCliente.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkConfirmacionCliente.Enabled = False
        Me.chkConfirmacionCliente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkConfirmacionCliente.Location = New System.Drawing.Point(16, 22)
        Me.chkConfirmacionCliente.Name = "chkConfirmacionCliente"
        Me.chkConfirmacionCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkConfirmacionCliente.Size = New System.Drawing.Size(86, 21)
        Me.chkConfirmacionCliente.TabIndex = 54
        Me.chkConfirmacionCliente.Text = "Autorización"
        Me.chkConfirmacionCliente.UseVisualStyleBackColor = False
        '
        'chkEntregaCliente
        '
        Me.chkEntregaCliente.BackColor = System.Drawing.SystemColors.Control
        Me.chkEntregaCliente.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEntregaCliente.Enabled = False
        Me.chkEntregaCliente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEntregaCliente.Location = New System.Drawing.Point(16, 52)
        Me.chkEntregaCliente.Name = "chkEntregaCliente"
        Me.chkEntregaCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEntregaCliente.Size = New System.Drawing.Size(79, 16)
        Me.chkEntregaCliente.TabIndex = 56
        Me.chkEntregaCliente.Text = "Recibe"
        Me.chkEntregaCliente.UseVisualStyleBackColor = False
        '
        'dtpConfirmacionCliente
        '
        Me.dtpConfirmacionCliente.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpConfirmacionCliente.Location = New System.Drawing.Point(104, 23)
        Me.dtpConfirmacionCliente.Name = "dtpConfirmacionCliente"
        Me.dtpConfirmacionCliente.Size = New System.Drawing.Size(101, 20)
        Me.dtpConfirmacionCliente.TabIndex = 55
        '
        'dtpFechaEntregaCliente
        '
        Me.dtpFechaEntregaCliente.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaEntregaCliente.Location = New System.Drawing.Point(104, 48)
        Me.dtpFechaEntregaCliente.Name = "dtpFechaEntregaCliente"
        Me.dtpFechaEntregaCliente.Size = New System.Drawing.Size(101, 20)
        Me.dtpFechaEntregaCliente.TabIndex = 57
        '
        'fraTaller
        '
        Me.fraTaller.BackColor = System.Drawing.SystemColors.Control
        Me.fraTaller.Controls.Add(Me.chkTallerEntrega)
        Me.fraTaller.Controls.Add(Me.chkTallerRegreso)
        Me.fraTaller.Controls.Add(Me.chkReparado)
        Me.fraTaller.Controls.Add(Me.dtpEntregaTaller)
        Me.fraTaller.Controls.Add(Me.dtpRegresoTaller)
        Me.fraTaller.Enabled = False
        Me.fraTaller.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraTaller.Location = New System.Drawing.Point(678, 180)
        Me.fraTaller.Name = "fraTaller"
        Me.fraTaller.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTaller.Size = New System.Drawing.Size(214, 102)
        Me.fraTaller.TabIndex = 47
        Me.fraTaller.TabStop = False
        Me.fraTaller.Text = " Taller "
        '
        'chkTallerEntrega
        '
        Me.chkTallerEntrega.BackColor = System.Drawing.SystemColors.Control
        Me.chkTallerEntrega.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTallerEntrega.Enabled = False
        Me.chkTallerEntrega.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTallerEntrega.Location = New System.Drawing.Point(16, 24)
        Me.chkTallerEntrega.Name = "chkTallerEntrega"
        Me.chkTallerEntrega.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTallerEntrega.Size = New System.Drawing.Size(67, 17)
        Me.chkTallerEntrega.TabIndex = 48
        Me.chkTallerEntrega.Text = "Entrega"
        Me.chkTallerEntrega.UseVisualStyleBackColor = False
        '
        'chkTallerRegreso
        '
        Me.chkTallerRegreso.BackColor = System.Drawing.SystemColors.Control
        Me.chkTallerRegreso.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTallerRegreso.Enabled = False
        Me.chkTallerRegreso.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTallerRegreso.Location = New System.Drawing.Point(16, 50)
        Me.chkTallerRegreso.Name = "chkTallerRegreso"
        Me.chkTallerRegreso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTallerRegreso.Size = New System.Drawing.Size(67, 17)
        Me.chkTallerRegreso.TabIndex = 50
        Me.chkTallerRegreso.Text = "Regreso"
        Me.chkTallerRegreso.UseVisualStyleBackColor = False
        '
        'chkReparado
        '
        Me.chkReparado.BackColor = System.Drawing.SystemColors.Control
        Me.chkReparado.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkReparado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkReparado.Location = New System.Drawing.Point(36, 76)
        Me.chkReparado.Name = "chkReparado"
        Me.chkReparado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkReparado.Size = New System.Drawing.Size(73, 17)
        Me.chkReparado.TabIndex = 52
        Me.chkReparado.Text = "Reparado"
        Me.chkReparado.UseVisualStyleBackColor = False
        '
        'dtpEntregaTaller
        '
        Me.dtpEntregaTaller.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpEntregaTaller.Location = New System.Drawing.Point(102, 23)
        Me.dtpEntregaTaller.Name = "dtpEntregaTaller"
        Me.dtpEntregaTaller.Size = New System.Drawing.Size(101, 20)
        Me.dtpEntregaTaller.TabIndex = 49
        '
        'dtpRegresoTaller
        '
        Me.dtpRegresoTaller.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpRegresoTaller.Location = New System.Drawing.Point(102, 49)
        Me.dtpRegresoTaller.Name = "dtpRegresoTaller"
        Me.dtpRegresoTaller.Size = New System.Drawing.Size(101, 20)
        Me.dtpRegresoTaller.TabIndex = 51
        '
        'fraMotivoReparacion
        '
        Me.fraMotivoReparacion.BackColor = System.Drawing.SystemColors.Control
        Me.fraMotivoReparacion.Controls.Add(Me.txtDesArticulo)
        Me.fraMotivoReparacion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraMotivoReparacion.Location = New System.Drawing.Point(12, 268)
        Me.fraMotivoReparacion.Name = "fraMotivoReparacion"
        Me.fraMotivoReparacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMotivoReparacion.Size = New System.Drawing.Size(389, 116)
        Me.fraMotivoReparacion.TabIndex = 27
        Me.fraMotivoReparacion.TabStop = False
        Me.fraMotivoReparacion.Text = " Descripción de Artículos y Motivo de la Reparación "
        '
        'fraObservacionesTaller
        '
        Me.fraObservacionesTaller.BackColor = System.Drawing.SystemColors.Control
        Me.fraObservacionesTaller.Controls.Add(Me.txtObservacionesTaller)
        Me.fraObservacionesTaller.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraObservacionesTaller.Location = New System.Drawing.Point(12, 468)
        Me.fraObservacionesTaller.Name = "fraObservacionesTaller"
        Me.fraObservacionesTaller.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraObservacionesTaller.Size = New System.Drawing.Size(389, 116)
        Me.fraObservacionesTaller.TabIndex = 32
        Me.fraObservacionesTaller.TabStop = False
        Me.fraObservacionesTaller.Text = " Observaciones Taller "
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.dbcTipoReparacion)
        Me.Frame4.Controls.Add(Me.dbcTaller)
        Me.Frame4.Controls.Add(Me._lblReparaciones_7)
        Me.Frame4.Controls.Add(Me._lblReparaciones_5)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(410, 95)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(259, 104)
        Me.Frame4.TabIndex = 34
        Me.Frame4.TabStop = False
        '
        'dbcTipoReparacion
        '
        Me.dbcTipoReparacion.Location = New System.Drawing.Point(80, 23)
        Me.dbcTipoReparacion.Name = "dbcTipoReparacion"
        Me.dbcTipoReparacion.Size = New System.Drawing.Size(167, 21)
        Me.dbcTipoReparacion.TabIndex = 36
        '
        'dbcTaller
        '
        Me.dbcTaller.Location = New System.Drawing.Point(80, 67)
        Me.dbcTaller.Name = "dbcTaller"
        Me.dbcTaller.Size = New System.Drawing.Size(167, 21)
        Me.dbcTaller.TabIndex = 38
        '
        '_lblReparaciones_7
        '
        Me._lblReparaciones_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_7.ForeColor = System.Drawing.Color.Black
        Me._lblReparaciones_7.Location = New System.Drawing.Point(10, 59)
        Me._lblReparaciones_7.Name = "_lblReparaciones_7"
        Me._lblReparaciones_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_7.Size = New System.Drawing.Size(67, 28)
        Me._lblReparaciones_7.TabIndex = 37
        Me._lblReparaciones_7.Text = "Taller Asignado :"
        '
        '_lblReparaciones_5
        '
        Me._lblReparaciones_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_5.ForeColor = System.Drawing.Color.Black
        Me._lblReparaciones_5.Location = New System.Drawing.Point(10, 16)
        Me._lblReparaciones_5.Name = "_lblReparaciones_5"
        Me._lblReparaciones_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_5.Size = New System.Drawing.Size(77, 27)
        Me._lblReparaciones_5.TabIndex = 35
        Me._lblReparaciones_5.Text = "Tipo de Reparación :"
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.fraLinea)
        Me.Frame6.Controls.Add(Me.txtNuevoAbono)
        Me.Frame6.Controls.Add(Me.txtAbonos)
        Me.Frame6.Controls.Add(Me.txtImpReparacion)
        Me.Frame6.Controls.Add(Me.txtAnticipo)
        Me.Frame6.Controls.Add(Me.txtCosto)
        Me.Frame6.Controls.Add(Me.chkCredito)
        Me.Frame6.Controls.Add(Me.lblNuevoAbono)
        Me.Frame6.Controls.Add(Me._lblReparaciones_11)
        Me.Frame6.Controls.Add(Me._lblReparaciones_8)
        Me.Frame6.Controls.Add(Me._lblReparaciones_6)
        Me.Frame6.Controls.Add(Me._lblReparaciones_99)
        Me.Frame6.Controls.Add(Me.txtSaldo)
        Me.Frame6.Controls.Add(Me._lblReparaciones_9)
        Me.Frame6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame6.Location = New System.Drawing.Point(678, 371)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(214, 212)
        Me.Frame6.TabIndex = 58
        Me.Frame6.TabStop = False
        Me.Frame6.Text = " Importe "
        '
        'fraLinea
        '
        Me.fraLinea.BackColor = System.Drawing.SystemColors.Control
        Me.fraLinea.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraLinea.Location = New System.Drawing.Point(87, 161)
        Me.fraLinea.Name = "fraLinea"
        Me.fraLinea.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraLinea.Size = New System.Drawing.Size(119, 2)
        Me.fraLinea.TabIndex = 72
        Me.fraLinea.TabStop = False
        '
        'chkCredito
        '
        Me.chkCredito.BackColor = System.Drawing.SystemColors.Control
        Me.chkCredito.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCredito.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCredito.Location = New System.Drawing.Point(14, 13)
        Me.chkCredito.Name = "chkCredito"
        Me.chkCredito.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCredito.Size = New System.Drawing.Size(200, 20)
        Me.chkCredito.TabIndex = 59
        Me.chkCredito.Text = "Entregar articulo con pago a crédito"
        Me.chkCredito.UseVisualStyleBackColor = False
        '
        'lblNuevoAbono
        '
        Me.lblNuevoAbono.AutoSize = True
        Me.lblNuevoAbono.BackColor = System.Drawing.SystemColors.Control
        Me.lblNuevoAbono.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNuevoAbono.ForeColor = System.Drawing.Color.Black
        Me.lblNuevoAbono.Location = New System.Drawing.Point(21, 159)
        Me.lblNuevoAbono.Name = "lblNuevoAbono"
        Me.lblNuevoAbono.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNuevoAbono.Size = New System.Drawing.Size(54, 13)
        Me.lblNuevoAbono.TabIndex = 64
        Me.lblNuevoAbono.Text = "Su Pago :"
        Me.lblNuevoAbono.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblNuevoAbono.Visible = False
        '
        '_lblReparaciones_11
        '
        Me._lblReparaciones_11.AutoSize = True
        Me._lblReparaciones_11.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_11.ForeColor = System.Drawing.Color.Black
        Me._lblReparaciones_11.Location = New System.Drawing.Point(21, 131)
        Me._lblReparaciones_11.Name = "_lblReparaciones_11"
        Me._lblReparaciones_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_11.Size = New System.Drawing.Size(49, 13)
        Me._lblReparaciones_11.TabIndex = 63
        Me._lblReparaciones_11.Text = "Abonos :"
        Me._lblReparaciones_11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblReparaciones_8
        '
        Me._lblReparaciones_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblReparaciones_8.Location = New System.Drawing.Point(21, 60)
        Me._lblReparaciones_8.Name = "_lblReparaciones_8"
        Me._lblReparaciones_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_8.Size = New System.Drawing.Size(81, 33)
        Me._lblReparaciones_8.TabIndex = 61
        Me._lblReparaciones_8.Text = "Importe Reparación  :"
        '
        '_lblReparaciones_6
        '
        Me._lblReparaciones_6.AutoSize = True
        Me._lblReparaciones_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_6.ForeColor = System.Drawing.Color.Black
        Me._lblReparaciones_6.Location = New System.Drawing.Point(21, 43)
        Me._lblReparaciones_6.Name = "_lblReparaciones_6"
        Me._lblReparaciones_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_6.Size = New System.Drawing.Size(40, 13)
        Me._lblReparaciones_6.TabIndex = 60
        Me._lblReparaciones_6.Text = "Costo :"
        '
        '_lblReparaciones_99
        '
        Me._lblReparaciones_99.AutoSize = True
        Me._lblReparaciones_99.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_99.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_99.ForeColor = System.Drawing.Color.Black
        Me._lblReparaciones_99.Location = New System.Drawing.Point(21, 101)
        Me._lblReparaciones_99.Name = "_lblReparaciones_99"
        Me._lblReparaciones_99.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_99.Size = New System.Drawing.Size(51, 13)
        Me._lblReparaciones_99.TabIndex = 62
        Me._lblReparaciones_99.Text = "Anticipo :"
        Me._lblReparaciones_99.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblReparaciones_9
        '
        Me._lblReparaciones_9.AutoSize = True
        Me._lblReparaciones_9.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_9.ForeColor = System.Drawing.Color.Black
        Me._lblReparaciones_9.Location = New System.Drawing.Point(21, 186)
        Me._lblReparaciones_9.Name = "_lblReparaciones_9"
        Me._lblReparaciones_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_9.Size = New System.Drawing.Size(40, 13)
        Me._lblReparaciones_9.TabIndex = 65
        Me._lblReparaciones_9.Text = "Saldo :"
        Me._lblReparaciones_9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame7
        '
        Me.Frame7.BackColor = System.Drawing.SystemColors.Control
        Me.Frame7.Controls.Add(Me.txtFolio)
        Me.Frame7.Controls.Add(Me.Frame5)
        Me.Frame7.Controls.Add(Me.Frame1)
        Me.Frame7.Controls.Add(Me.fraMoneda)
        Me.Frame7.Controls.Add(Me.dbcVendedor)
        Me.Frame7.Controls.Add(Me._lblReparaciones_4)
        Me.Frame7.Controls.Add(Me._lblReparaciones_0)
        Me.Frame7.Controls.Add(Me._lblReparaciones_1)
        Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame7.Location = New System.Drawing.Point(12, 12)
        Me.Frame7.Name = "Frame7"
        Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame7.Size = New System.Drawing.Size(877, 79)
        Me.Frame7.TabIndex = 1
        Me.Frame7.TabStop = False
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.dtpFechaReparacion)
        Me.Frame5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame5.Enabled = False
        Me.Frame5.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.Frame5.Location = New System.Drawing.Point(273, 9)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(121, 33)
        Me.Frame5.TabIndex = 6
        '
        'dtpFechaReparacion
        '
        Me.dtpFechaReparacion.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaReparacion.Location = New System.Drawing.Point(8, 8)
        Me.dtpFechaReparacion.Name = "dtpFechaReparacion"
        Me.dtpFechaReparacion.Size = New System.Drawing.Size(105, 20)
        Me.dtpFechaReparacion.TabIndex = 7
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtTipoCambio)
        Me.Frame1.Controls.Add(Me.txtDiasTranscurridos)
        Me.Frame1.Controls.Add(Me._lblReparaciones_2)
        Me.Frame1.Controls.Add(Me._lblReparaciones_3)
        Me.Frame1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame1.Enabled = False
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(673, 14)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(192, 57)
        Me.Frame1.TabIndex = 12
        '
        '_lblReparaciones_2
        '
        Me._lblReparaciones_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblReparaciones_2.Location = New System.Drawing.Point(12, 1)
        Me._lblReparaciones_2.Name = "_lblReparaciones_2"
        Me._lblReparaciones_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_2.Size = New System.Drawing.Size(92, 30)
        Me._lblReparaciones_2.TabIndex = 13
        Me._lblReparaciones_2.Text = "Días Transcurridos"
        Me._lblReparaciones_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblReparaciones_3
        '
        Me._lblReparaciones_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblReparaciones_3.Location = New System.Drawing.Point(98, 13)
        Me._lblReparaciones_3.Name = "_lblReparaciones_3"
        Me._lblReparaciones_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_3.Size = New System.Drawing.Size(86, 14)
        Me._lblReparaciones_3.TabIndex = 14
        Me._lblReparaciones_3.Text = "Tipo Cambio"
        Me._lblReparaciones_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'fraMoneda
        '
        Me.fraMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.fraMoneda.Controls.Add(Me.optMonedaDolar)
        Me.fraMoneda.Controls.Add(Me.optMonedaPeso)
        Me.fraMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraMoneda.Location = New System.Drawing.Point(398, 11)
        Me.fraMoneda.Name = "fraMoneda"
        Me.fraMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMoneda.Size = New System.Drawing.Size(207, 57)
        Me.fraMoneda.TabIndex = 9
        Me.fraMoneda.TabStop = False
        Me.fraMoneda.Text = " Moneda "
        '
        'dbcVendedor
        '
        Me.dbcVendedor.Location = New System.Drawing.Point(81, 46)
        Me.dbcVendedor.Name = "dbcVendedor"
        Me.dbcVendedor.Size = New System.Drawing.Size(305, 21)
        Me.dbcVendedor.TabIndex = 8
        '
        '_lblReparaciones_4
        '
        Me._lblReparaciones_4.AutoSize = True
        Me._lblReparaciones_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblReparaciones_4.Location = New System.Drawing.Point(10, 51)
        Me._lblReparaciones_4.Name = "_lblReparaciones_4"
        Me._lblReparaciones_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_4.Size = New System.Drawing.Size(59, 13)
        Me._lblReparaciones_4.TabIndex = 3
        Me._lblReparaciones_4.Text = "Vendedor :"
        '
        '_lblReparaciones_0
        '
        Me._lblReparaciones_0.AutoSize = True
        Me._lblReparaciones_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblReparaciones_0.Location = New System.Drawing.Point(11, 22)
        Me._lblReparaciones_0.Name = "_lblReparaciones_0"
        Me._lblReparaciones_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_0.Size = New System.Drawing.Size(35, 13)
        Me._lblReparaciones_0.TabIndex = 2
        Me._lblReparaciones_0.Text = "Folio :"
        '
        '_lblReparaciones_1
        '
        Me._lblReparaciones_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblReparaciones_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblReparaciones_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblReparaciones_1.Location = New System.Drawing.Point(237, 21)
        Me._lblReparaciones_1.Name = "_lblReparaciones_1"
        Me._lblReparaciones_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblReparaciones_1.Size = New System.Drawing.Size(39, 15)
        Me._lblReparaciones_1.TabIndex = 4
        Me._lblReparaciones_1.Text = "Fecha :"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.chkCorpoRegresa)
        Me.Frame2.Controls.Add(Me.chkCorpoEnvio)
        Me.Frame2.Controls.Add(Me.dtpCorpoEnvio)
        Me.Frame2.Controls.Add(Me.dtpCorpoRegresa)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(678, 95)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(214, 81)
        Me.Frame2.TabIndex = 42
        Me.Frame2.TabStop = False
        Me.Frame2.Text = " Corporativo "
        '
        'chkCorpoRegresa
        '
        Me.chkCorpoRegresa.BackColor = System.Drawing.SystemColors.Control
        Me.chkCorpoRegresa.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCorpoRegresa.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCorpoRegresa.Location = New System.Drawing.Point(12, 47)
        Me.chkCorpoRegresa.Name = "chkCorpoRegresa"
        Me.chkCorpoRegresa.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCorpoRegresa.Size = New System.Drawing.Size(83, 25)
        Me.chkCorpoRegresa.TabIndex = 45
        Me.chkCorpoRegresa.Text = "Regreso"
        Me.chkCorpoRegresa.UseVisualStyleBackColor = False
        '
        'chkCorpoEnvio
        '
        Me.chkCorpoEnvio.BackColor = System.Drawing.SystemColors.Control
        Me.chkCorpoEnvio.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCorpoEnvio.Enabled = False
        Me.chkCorpoEnvio.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCorpoEnvio.Location = New System.Drawing.Point(12, 22)
        Me.chkCorpoEnvio.Name = "chkCorpoEnvio"
        Me.chkCorpoEnvio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCorpoEnvio.Size = New System.Drawing.Size(83, 25)
        Me.chkCorpoEnvio.TabIndex = 43
        Me.chkCorpoEnvio.Text = "Envio"
        Me.chkCorpoEnvio.UseVisualStyleBackColor = False
        '
        'dtpCorpoEnvio
        '
        Me.dtpCorpoEnvio.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpCorpoEnvio.Location = New System.Drawing.Point(101, 20)
        Me.dtpCorpoEnvio.Name = "dtpCorpoEnvio"
        Me.dtpCorpoEnvio.Size = New System.Drawing.Size(101, 20)
        Me.dtpCorpoEnvio.TabIndex = 44
        '
        'dtpCorpoRegresa
        '
        Me.dtpCorpoRegresa.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpCorpoRegresa.Location = New System.Drawing.Point(101, 48)
        Me.dtpCorpoRegresa.Name = "dtpCorpoRegresa"
        Me.dtpCorpoRegresa.Size = New System.Drawing.Size(101, 20)
        Me.dtpCorpoRegresa.TabIndex = 46
        '
        'lblEstatus
        '
        Me.lblEstatus.BackColor = System.Drawing.SystemColors.Window
        Me.lblEstatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEstatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEstatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblEstatus.Location = New System.Drawing.Point(416, 504)
        Me.lblEstatus.Name = "lblEstatus"
        Me.lblEstatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEstatus.Size = New System.Drawing.Size(254, 33)
        Me.lblEstatus.TabIndex = 76
        Me.lblEstatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblEstatus.Visible = False
        '
        'lblCancelado
        '
        Me.lblCancelado.BackColor = System.Drawing.SystemColors.Window
        Me.lblCancelado.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCancelado.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCancelado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblCancelado.Location = New System.Drawing.Point(417, 543)
        Me.lblCancelado.Name = "lblCancelado"
        Me.lblCancelado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCancelado.Size = New System.Drawing.Size(121, 33)
        Me.lblCancelado.TabIndex = 73
        Me.lblCancelado.Text = "Cancelado"
        Me.lblCancelado.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblCancelado.Visible = False
        '
        'lblLiquidado
        '
        Me.lblLiquidado.BackColor = System.Drawing.SystemColors.Window
        Me.lblLiquidado.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLiquidado.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLiquidado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblLiquidado.Location = New System.Drawing.Point(417, 543)
        Me.lblLiquidado.Name = "lblLiquidado"
        Me.lblLiquidado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLiquidado.Size = New System.Drawing.Size(121, 33)
        Me.lblLiquidado.TabIndex = 74
        Me.lblLiquidado.Text = "Liquidado"
        Me.lblLiquidado.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblLiquidado.Visible = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(310, 603)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(93, 35)
        Me.btnBuscar.TabIndex = 75
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Location = New System.Drawing.Point(210, 603)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(93, 35)
        Me.btnLimpiar.TabIndex = 74
        Me.btnLimpiar.Text = "Nuevo"
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(111, 603)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(93, 35)
        Me.btnEliminar.TabIndex = 73
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.Location = New System.Drawing.Point(12, 603)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(93, 35)
        Me.btnGuardar.TabIndex = 72
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmCorpoControlReparaciones_Corpo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(918, 643)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Marco)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(74, 95)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoControlReparaciones_Corpo"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Administración de Reparaciones"
        Me.Marco.ResumeLayout(False)
        Me.fraNuevoTexto.ResumeLayout(False)
        Me.fraNuevoTexto.PerformLayout()
        Me.fraBitacora.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.fraDatosCliente.ResumeLayout(False)
        Me.fraDatosCliente.PerformLayout()
        Me.fraCliente.ResumeLayout(False)
        Me.fraTaller.ResumeLayout(False)
        Me.fraMotivoReparacion.ResumeLayout(False)
        Me.fraMotivoReparacion.PerformLayout()
        Me.fraObservacionesTaller.ResumeLayout(False)
        Me.fraObservacionesTaller.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame6.ResumeLayout(False)
        Me.Frame6.PerformLayout()
        Me.Frame7.ResumeLayout(False)
        Me.Frame7.PerformLayout()
        Me.Frame5.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.fraMoneda.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        CType(Me.lblReparaciones, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click

    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnCatClientes_Click(sender As Object, e As EventArgs) Handles btnCatClientes.Click

    End Sub
End Class