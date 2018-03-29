'**********************************************************************************************************************'
'*PROGRAMA: ABC CLIENTES  
'*AUTOR: MIGUEL ANGEL GARCIA WHA   
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCorpoABCClientes


    Inherits System.Windows.Forms.Form

    'Variables
    Public mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Public mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Public mblnVigente As Boolean 'Estatus Vigente
    Public mblnCancelado As Boolean 'Estatus Cancelado
    Public mblnSuspension As Boolean 'Estatus Suspensión
    Public mblnSalir As Boolean 'Para Salir Con el Esc
    Public intVendExterno As Integer
    Public sglTiempoCambio As Single  'Para Esperar un Tiempo
    Public intCodAlmacen As Integer
    Public CodAlmacen As Integer
    Public tecla As Integer
    Public FueraChange As Boolean
    Public intCodSucursal As Integer

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdRptCtes As System.Windows.Forms.Button
    Public WithEvents RTObservaciones As System.Windows.Forms.RichTextBox
    Public WithEvents chkVendExterno As System.Windows.Forms.CheckBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents _optSuspension_2 As System.Windows.Forms.RadioButton
    Public WithEvents _optCancelado_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optVigente_0 As System.Windows.Forms.RadioButton
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents txtConyuge As System.Windows.Forms.TextBox
    Public WithEvents chkFechaNacimiento As System.Windows.Forms.CheckBox
    Public WithEvents txtCiudad As System.Windows.Forms.TextBox
    Public WithEvents dtpAniversarioBodasTag As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpCumpleañosConyugeTag As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFecNacimientoTag As System.Windows.Forms.DateTimePicker
    Public WithEvents chkAniversario As System.Windows.Forms.CheckBox
    Public WithEvents chkCumpleaños As System.Windows.Forms.CheckBox
    Public WithEvents dtpFechaNacimiento As System.Windows.Forms.DateTimePicker
    Public WithEvents cboTipoCliente As System.Windows.Forms.ComboBox
    Public WithEvents txtEmail As System.Windows.Forms.TextBox
    Public WithEvents txtCodPostal As System.Windows.Forms.TextBox
    Public WithEvents txtFax As System.Windows.Forms.TextBox
    Public WithEvents txtTelOficina As System.Windows.Forms.TextBox
    Public WithEvents txtTelCasa As System.Windows.Forms.TextBox
    Public WithEvents txtRFC As System.Windows.Forms.TextBox
    Public WithEvents txtColonia As System.Windows.Forms.TextBox
    Public WithEvents txtDomicilio As System.Windows.Forms.TextBox
    Public WithEvents txtNombre As System.Windows.Forms.TextBox
    Public WithEvents dtpCumpleaños As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpAniversario As System.Windows.Forms.DateTimePicker
    Public WithEvents _Label1_14 As System.Windows.Forms.Label
    Public WithEvents _Label1_13 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_12 As System.Windows.Forms.Label
    Public WithEvents _Label1_11 As System.Windows.Forms.Label
    Public WithEvents _Label1_10 As System.Windows.Forms.Label
    Public WithEvents _Label1_9 As System.Windows.Forms.Label
    Public WithEvents _Label1_8 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_4 As System.Windows.Forms.Label
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optCancelado As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optSuspension As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents txtCodigo As TextBox
    Public WithEvents dtpFechaAlta As System.Windows.Forms.DateTimePicker
    Public WithEvents dbcAlmacen As System.Windows.Forms.ComboBox
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents txtFormQueEjecuta As TextBox
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents Panel1 As Panel
    Public WithEvents Panel2 As Panel
    Public WithEvents Panel3 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents optVigente As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray

    Sub frmCorpoABCClientes()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        InicializaVariables()
        Nuevo()
        HabilitarDesabilitarBotones()
    End Sub
    Private Sub frmCorpoABCClientes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeComponent()
        frmCorpoABCClientes()
    End Sub

    Public Sub HabilitarDesabilitarBotones()
        btnGuardar.Enabled = True
        btnEliminar.Enabled = False
        btnLimpiar.Enabled = True
        btnBuscar.Enabled = True
        btnSalir.Enabled = True
    End Sub

    Sub Buscar()
        'On Local Error GoTo Merr
        Try
            Dim strSQL As String
            Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
            Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
            Dim strControlActual As String 'Nombre del control actual

            'If Trim(dbcSucursales.Text) = "" Then
            '    MsgBox("Proporcione la sucursal donde se buscarán los clientes", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA)
            '    dbcSucursales_GotFocus(New Object, New EventArgs)
            '    Exit Sub
            'End If

            If (txtCodigo.Text = "") Then
                strControlActual = UCase(txtCodigo.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
                strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control
            ElseIf (txtNombre.Text = "") Then
                strControlActual = UCase(txtNombre.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
                strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control 
            End If


            'Select Case strControlActual
            '    Case "TXTCODIGO"
            '        strCaptionForm = "Consulta de Clientes"
            '        gStrSql = "SELECT RIGHT('00000'+LTRIM(CodCliente),5) AS CODIGO, DescCliente AS NOMBRE FROM CatClientes  Where  CodAlmacen = " & Trim(txtCodigo.Text) & " or isnull(CodAlmacen,0) = 0 ORDER BY CodCliente"
            '    Case "TXTNOMBRE"
            '        strCaptionForm = "Consulta de Clientes"
            '        gStrSql = "SELECT DescCliente AS NOMBRE, RIGHT('00000'+LTRIM(CodCliente),5) AS CODIGO FROM CatClientes WHERE DescCliente LIKE '" & Trim(txtNombre.Text) & "%' and (CodAlmacen = " & Trim(txtCodigo.Text) & " or isnull(CodAlmacen,0) = 0 )  ORDER BY DescCliente"
            '    Case Else
            '        'Sale de este sub para QUE no ejecute ninguna opcion
            '        Exit Sub
            'End Select
            'strSQL = gStrSql 'Se hace uso de una variable temporal para el query
            ''Si hubo cambios y es una modificacion entonces preguntara que si desea gravar los cambios
            'If Cambios() = True And mblnNuevo = False Then
            '    Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
            '        Case vbYes  'Guardar el registro
            '            If Guardar() = False Then
            '                Exit Sub
            '            End If
            '        Case vbNo  'No hace nada y permite que se carguela consulta
            '        Case vbCancel  'Cancela la consulta
            '            Exit Sub
            '    End Select
            'End If
            'gStrSql = strSQL 'Se regresa el valor de la variavle temporal a la variable original
            'ModEstandar.BorraCmd()
            'Cmd.CommandText = "dbo.Up_Select_Datos"
            'Cmd.CommandType = CommandTypeEnum.adCmdStoredProc
            'Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamReturnValue))
            'Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", DataTypeEnum.adChar, ParameterDirectionEnum.adParamInput, 800, gStrSql))
            'RsGral = Cmd.Execute
            ''Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
            'If RsGral.RecordCount = 0 Then
            '    MsjNoExiste(C_msgSINDATOS, gstrNombCortoEmpresa)
            '    Exit Sub
            'End If
            ''Carga el formulario de consulta
            'ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

            'With FrmConsultas.Flexdet
            '    Select Case strControlActual
            '        Case "TXTCODIGO"
            '            .set_ColWidth(0, 0, 900) 'Columna del Código
            '            .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
            '        Case "TXTNOMBRE"
            '            .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
            '            .set_ColWidth(1, 0, 900) 'Columna del Código
            '    End Select
            'End With

            If strControlActual = "TXTCODIGO" Or strControlActual = "TXTNOMBRE" Then
                FueraChange = True
                Dim FrmConsultasClientes As FrmConsultasClientes = New FrmConsultasClientes()
                FrmConsultasClientes.InitializeComponent()
                FrmConsultasClientes.dbcSucursales.Text = dbcSucursales.Text
                FrmConsultasClientes.strFormaActual = UCase(Me.Name)
                FrmConsultasClientes.strControlActual = strControlActual
                FrmConsultasClientes.intCodSucursal = intCodSucursal
                FrmConsultasClientes.Show()
                FueraChange = False
            End If
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Sub DesHabilitarControles()
        txtConyuge.Enabled = False
        txtDomicilio.Enabled = False
        txtColonia.Enabled = False
        txtCiudad.Enabled = False
        txtCodPostal.Enabled = False
        txtRFC.Enabled = False
        txtTelCasa.Enabled = False
        txtTelOficina.Enabled = False
        txtFax.Enabled = False
        txtEmail.Enabled = False
        _optVigente_0.Enabled = False
        _optCancelado_1.Enabled = False
        _optSuspension_2.Enabled = False
        cboTipoCliente.Enabled = False
        chkFechaNacimiento.Enabled = False
        chkAniversario.Enabled = False
        chkCumpleaños.Enabled = False
        'chkVendExterno.Enabled = False
        RTObservaciones.Enabled = True
        dtpAniversario.Enabled = False
        dtpFechaNacimiento.Enabled = False
        dtpCumpleaños.Enabled = False
        dbcSucursales.Enabled = False
        dtpFechaAlta.Enabled = False


    End Sub

    Sub HabilitarControles()
        txtDomicilio.Enabled = True
        txtColonia.Enabled = True
        txtCiudad.Enabled = True
        txtCodPostal.Enabled = True
        txtRFC.Enabled = True
        txtTelCasa.Enabled = True
        txtTelOficina.Enabled = True
        txtFax.Enabled = True
        txtEmail.Enabled = True
        _optVigente_0.Enabled = True
        _optCancelado_1.Enabled = True
        _optSuspension_2.Enabled = True
        cboTipoCliente.Enabled = True
        chkFechaNacimiento.Enabled = True
        chkAniversario.Enabled = True
        chkCumpleaños.Enabled = True
        chkVendExterno.Enabled = True
        RTObservaciones.Enabled = True
        dtpAniversario.Enabled = True
        dtpFechaNacimiento.Enabled = True
        dtpCumpleaños.Enabled = True

    End Sub

    Sub Eliminar()
        '    On Local Error GoTo MErr
        '    Dim blnTransaccion As Boolean
        '    Dim strEstatus As String
        '    gStrSql = "SELECT DescCliente FROM CatClientes WHERE CodCliente=" & Val(txtCodigo)
        '    ModEstandar.BorraCmd
        '    Cmd.CommandText = "dbo.Up_Select_Datos"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, gStrSql)
        '    Set RsGral = Cmd.Execute
        '    If RsGral.RecordCount = 0 Then
        '        MsgBox "Proporcione un código valido para eliminar.", vbInformation + vbOKOnly, "Mensaje"
        '        Exit Sub
        '    End If
        '    'Preguntar si desea borrar el registro
        '    Select Case MsgBox(C_msgBORRAR, vbQuestion + vbYesNoCancel + vbDefaultButton3, "")
        '        Case vbNo
        '            Exit Sub
        '        Case vbCancel
        '            Exit Sub
        '    End Select
        '    Cnn.BeginTrans
        '    Me.MousePointer = vbHourglass
        '    blnTransaccion = True
        '    If _optVigente_0.Value Then
        '        strEstatus = "V"
        '    ElseIf _optCancelado_1.Value Then
        '        strEstatus = "C"
        '    ElseIf _optSuspension_2.Value Then
        '        strEstatus = "S"
        '    End If
        '    ModStoredProcedures.PR_IMECatClientes txtCodigo, txtNombre, txtRFC, txtDomicilio, txtColonia, txtCiudad, _
        '            txtTelCasa, txtTelOficina, txtFax, txtCodPostal, txtEmail, cboTipoCliente, Format(dtpFechaNacimiento, C_FORMATFECHAGUARDAR), _
        '            Format(dtpCumpleaños, C_FORMATFECHAGUARDAR), Format(dtpAniversario, C_FORMATFECHAGUARDAR), strEstatus, _
        '            txtObservaciones, Format(dtpFechaAlta, C_FORMATFECHAGUARDAR), "", "0", 0, C_ELIMINACION, 0
        '    Cmd.Execute
        '    Me.MousePointer = vbDefault
        '    Cnn.CommitTrans
        '    blnTransaccion = False
        '    Limpiar
        'MErr:
        '    If Err.Number <> 0 Then
        '        If blnTransaccion = True Then Cnn.RollbackTrans
        '        Me.MousePointer = vbDefault
        '        ModEstandar.MostrarError
        '    End If
    End Sub

    Public Function Guardar()
        'On Local Error GoTo Merr
        Dim blnTransaccion As Boolean
        Dim strEstatus As String
        '    Guardar = False
        '    Do While (Timer - sglTiempoCambio) <= 2.1
        '    Loop
        '    DoEvents

        'If Cambios() = False Then
        '    Limpiar()
        '    Exit Function
        'End If

        'Valida si todos los datos han sido llenados para poder ser guardados
        If ValidaDatos() = False Then
            Exit Function
        End If
        '    If Val(txtCodigo) = 0 Then
        '        mblnNuevo = True
        '    End If
        'cnn.BeginTrans()
        'Me.MousePointer = vbHourglass
        blnTransaccion = True
        If _optVigente_0.Checked Then
            strEstatus = "V"
        ElseIf _optCancelado_1.Checked Then
            strEstatus = "C"
        ElseIf _optSuspension_2.Checked Then
            strEstatus = "S"
        End If

        Dim fechaNacimiento = AgregarHoraAFecha(dtpFechaNacimiento.Value)
        Dim fechaCumpleaños = AgregarHoraAFecha(dtpCumpleaños.Value)
        Dim fechaAniversario = AgregarHoraAFecha(dtpAniversario.Value)
        Dim fechaAlta = AgregarHoraAFecha(dtpFechaAlta.Value)

        If mblnNuevo Then
            ModStoredProcedures.PR_IMECatClientes(Val(txtCodigo), txtNombre.Text, txtRFC.Text, txtDomicilio.Text, txtColonia.Text, txtCiudad.Text,
            txtTelCasa.Text, txtTelOficina.Text, txtFax.Text, txtCodPostal.Text, txtEmail.Text, IIf(cboTipoCliente.Text = "ESPECIAL", "E", cboTipoCliente.Text), fechaNacimiento,
            fechaCumpleaños, fechaAniversario, strEstatus, RTObservaciones.Text, fechaAlta, CStr(intCodAlmacen), CStr(intCodSucursal), "", C_INSERCION, 0)
            Cmd.Execute()
            txtCodigo.Text = Format(Cmd("ID"), "00000")
        Else
            ModStoredProcedures.PR_IMECatClientes(txtCodigo.Text, txtNombre.Text, txtRFC.Text, txtDomicilio.Text, txtColonia.Text, txtCiudad.Text,
            txtTelCasa.Text, txtTelOficina.Text, txtFax.Text, txtCodPostal.Text, txtEmail.Text, IIf(cboTipoCliente.Text = "ESPECIAL", "E", cboTipoCliente.Text), fechaNacimiento,
            fechaCumpleaños, fechaAniversario, strEstatus,
            RTObservaciones.Text, fechaAlta, CStr(intCodAlmacen), CStr(intCodSucursal), "", C_MODIFICACION, 0)
            Cmd.Execute()
        End If

        dbcAlmacen_Leave(New Object, New EventArgs)
        'ModStoredProcedures.PR_IMECatClientes txtCodigo, "", "", "", "", "", "", "", "", "", "", "", "01/01/1900", "01/01/1900", "01/01/1900", "", "", "01/01/1900", CStr(intCodAlmacen), CStr(intCodSucursal), "", C_MODIFICACION, 1
        Cmd.Execute()
        'Me.MousePointer = vbDefault
        'Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox(C_msgACTUALIZADO, vbInformation + vbOKOnly, ModVariables.gstrNombCortoEmpresa)
        Nuevo()
        Guardar = True
        Limpiar()
        Exit Function
Merr:
        If blnTransaccion = True Then
            Cnn.RollbackTrans()
        End If
        'Me.MousePointer = vbDefault
        ModEstandar.MostrarError()
        Return Guardar()
    End Function

    Sub Limpiar()
        'On Error GoTo Merr
        Try
            FueraChange = True
            txtCodigo.Text = ""
            dbcSucursales.Text = ""
            Nuevo()
            FueraChange = False
            'txtCodigo.SetFocus    
            '''dbcSucursales.SetFocus
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Public Function LlenaDatos()
        'On Local Error GoTo Merr
        Try
            'txtCodigo.Text = String.Format(txtCodigo, "00000") 
            For i = 1 To 5 - (txtCodigo.TextLength)
                txtCodigo.Text = String.Concat("0" + txtCodigo.Text)
            Next i

            gStrSql = "SELECT C.CodCliente,C.DescCliente,C.Rfc,C.Conyuge,C.Domicilio,C.Colonia,C.Ciudad,C.TelCasa,C.TelOficina,
    C.Fax, C.CP, C.Email, C.TipoCte, C.FechaNacimiento, C.FechaNacimientoConyuge, C.AniversarioBodas, C.Estatus,
    C.Observaciones, C.FechaAlta, C.AlmacenVExt, A.CodAlmacen, A.DescAlmacen 
    From CatClientes C 
	Left OUTER JOIN CatAlmacen A ON C.AlmacenVExt = A.CodAlmacen
    WHERE C.CodCliente =" & CLng(txtCodigo.Text)


            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", DataTypeEnum.adChar, ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute

            If Not RsGral.EOF Then

                txtNombre.Text = Trim(RsGral.Fields("DescCliente").Value.ToString())
                txtNombre.Tag = Trim(RsGral.Fields("DescCliente").Value.ToString())

                txtRFC.Text = Trim(RsGral.Fields("Rfc").Value.ToString())
                txtRFC.Tag = Trim(RsGral.Fields("Rfc").Value.ToString())

                txtConyuge.Text = Trim(RsGral.Fields("Conyuge").Value.ToString())
                txtConyuge.Tag = Trim(RsGral.Fields("Conyuge").Value.ToString())

                txtDomicilio.Text = Trim(RsGral.Fields("Domicilio").Value.ToString())
                txtDomicilio.Tag = Trim(RsGral.Fields("Domicilio").Value.ToString())

                txtColonia.Text = Trim(RsGral.Fields("Colonia").Value.ToString())
                txtColonia.Tag = Trim(RsGral.Fields("Colonia").Value.ToString())

                txtCiudad.Text = Trim(RsGral.Fields("Ciudad").Value.ToString())
                txtCiudad.Tag = Trim(RsGral.Fields("Ciudad").Value.ToString())

                txtTelCasa.Text = Trim(RsGral.Fields("TelCasa").Value.ToString())
                txtTelCasa.Tag = Trim(RsGral.Fields("TelCasa").Value.ToString())

                txtTelOficina.Text = Trim(RsGral.Fields("TelOficina").Value.ToString())
                txtTelOficina.Tag = Trim(RsGral.Fields("TelOficina").Value.ToString())

                txtFax.Text = Trim(RsGral.Fields("Fax").Value.ToString())
                txtFax.Tag = Trim(RsGral.Fields("Fax").Value.ToString())

                txtCodPostal.Text = Trim(RsGral.Fields("CP").Value.ToString())
                txtCodPostal.Tag = Trim(RsGral.Fields("CP").Value.ToString())

                txtEmail.Text = Trim(RsGral.Fields("Email").Value.ToString())
                txtEmail.Tag = Trim(RsGral.Fields("Email").Value.ToString())

                'txtEmail.FontUnderline = True
                txtEmail.Font = New Font(txtEmail.Font, FontStyle.Underline)


                If Trim(RsGral.Fields("TipoCte").Value.ToString()) = "E" Then
                    cboTipoCliente.Text = "ESPECIAL"
                    cboTipoCliente.Tag = "ESPECIAL"
                Else
                    cboTipoCliente.Text = Trim(RsGral.Fields("TipoCte").Value.ToString())
                    cboTipoCliente.Tag = Trim(RsGral.Fields("TipoCte").Value.ToString())
                End If

                If Year(Convert.ToDateTime(RsGral.Fields("FechaNacimiento").Value)) <> Year(dtpFechaNacimiento.MinDate) Then
                    dtpFechaNacimiento.Value = Format(Convert.ToDateTime(RsGral.Fields("FechaNacimiento").Value), C_FORMATFECHAMOSTRAR)
                    chkFechaNacimiento.Checked = True
                Else
                    dtpFecNacimientoTag.Value = dtpFechaNacimiento.MinDate
                End If

                If Year(Convert.ToDateTime(RsGral.Fields("FechaNacimientoConyuge").Value)) <> Year(dtpCumpleaños.MinDate) Then
                    dtpCumpleaños.Value = Format(Convert.ToDateTime(RsGral.Fields("FechaNacimientoConyuge").Value), C_FORMATFECHAMOSTRAR)
                    chkCumpleaños.Checked = True
                Else
                    dtpCumpleañosConyugeTag.Value = dtpCumpleaños.MinDate
                End If

                If Year(Convert.ToDateTime(RsGral.Fields("AniversarioBodas").Value)) <> Year(dtpAniversario.MinDate) Then
                    dtpAniversario.Value = Format(Convert.ToDateTime(RsGral.Fields("AniversarioBodas").Value), C_FORMATFECHAMOSTRAR)
                    chkAniversario.Checked = True
                Else
                    dtpAniversarioBodasTag.Value = dtpAniversario.MinDate
                End If

                dtpFechaNacimiento.Enabled = False
                dtpCumpleaños.Enabled = False
                dtpAniversario.Enabled = False

                DesHabilitarControles()

                If (RsGral.Fields("Estatus").Value.ToString()) = "V" Then
                    _optVigente_0.Checked = True
                    mblnVigente = True
                    mblnCancelado = False
                    mblnSuspension = False
                ElseIf (RsGral.Fields("Estatus").Value.ToString()) = "C" Then
                    _optCancelado_1.Checked = True
                    mblnVigente = False
                    mblnCancelado = True
                    mblnSuspension = False
                ElseIf (RsGral.Fields("Estatus").Value.ToString()) = "S" Then
                    _optSuspension_2.Checked = True
                    mblnVigente = False
                    mblnCancelado = False
                    mblnSuspension = True
                End If

                RTObservaciones.Text = Trim(RsGral.Fields("Observaciones").Value.ToString())
                RTObservaciones.Tag = Trim(RsGral.Fields("Observaciones").Value.ToString())
                dtpFechaAlta.Value = Format(Convert.ToDateTime(RsGral.Fields("FechaAlta").Value), C_FORMATFECHAMOSTRAR)


                If (RsGral.Fields("AlmacenVExt").Value) IsNot DBNull.Value Then
                    chkVendExterno.Checked = True
                    intVendExterno = 1
                    dbcAlmacen.Text = RsGral.Fields("DescAlmacen").Value.ToString()
                    dbcAlmacen.Tag = RsGral.Fields("DescAlmacen").Value.ToString()
                    intCodAlmacen = RsGral.Fields("AlmacenVExt").Value()
                    CodAlmacen = RsGral.Fields("AlmacenVExt").Value()
                    chkVendExterno.Enabled = True
                Else
                    dbcAlmacen.Text = ""
                    'chkVendExterno.Enabled = False
                End If

                If (txtCodigo.Text > 1) Then
                    chkVendExterno.Enabled = True
                Else
                    chkVendExterno.Enabled = False
                End If

                gStrSql = "Select A.CodAlmacen, A.DescAlmacen " &
                    "FROM CatClientes As C INNER JOIN CatAlmacen As A On C.CodAlmacen = A.CodAlmacen " &
                    "WHERE C.CodCliente = " & CLng(Numerico(txtCodigo.Text)) & " And C.CodAlmacen = " & intCodSucursal

                    gStrSql = "Select A.CodAlmacen, IsNull(A.DescAlmacen,'') as DescAlmacen FROM CatClientes AS C (NOLOCK) Left Outer Join CatAlmacen AS A (NOLOCK) ON C.CodAlmacen = A.CodAlmacen 
                WHERE C.CodCliente =" & CLng(txtCodigo.Text)

                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", DataTypeEnum.adChar, ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute

                If RsGral.EOF Then
                    Err.Raise(vbObjectError + 1, , "No se encontró el cliente " + txtCodigo.Text + " en la sucursal " & intCodSucursal)
                End If
                dbcSucursales.Text = Trim(RsGral.Fields("DescAlmacen").Value().ToString())
            Else
            MsjNoExiste("El Cliente", gstrNombCortoEmpresa)
            Nuevo()
                txtCodigo.Text = ""
                txtCodigo.Focus()
            End If

            mblnCambiosEnCodigo = False
            mblnNuevo = False
            'Exit Sub
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
        Return True
    End Function

    Sub Nuevo()
        '        On Local Error GoTo Merr
        Try
            txtCodigo.Text = ""
            txtNombre.Text = ""
            txtNombre.Tag = ""
            txtDomicilio.Text = ""
            txtDomicilio.Tag = ""
            txtColonia.Text = ""
            txtColonia.Tag = ""
            txtCiudad.Text = ""
            txtCiudad.Tag = ""
            txtConyuge.Text = ""
            txtConyuge.Tag = ""
            txtRFC.Text = ""
            txtRFC.Tag = ""
            txtTelCasa.Text = ""
            txtTelCasa.Tag = ""
            txtTelOficina.Text = ""
            txtTelOficina.Tag = ""
            txtFax.Text = ""
            txtFax.Tag = ""
            txtCodPostal.Text = ""
            txtCodPostal.Tag = ""
            txtEmail.Text = ""
            txtEmail.Tag = ""
            'txtEmail.Font.Underline = False
            'cboTipoCliente.SelectedIndex = 0
            'cboTipoCliente.Tag = cboTipoCliente.Text 
            'dtpFechaNacimiento.Enabled = True
            'dtpCumpleaños.Enabled = True
            'dtpAniversario.Enabled = True
            'dtpFechaNacimiento.Value = dtpFechaNacimiento.MinDate
            'dtpAniversario.Value = dtpAniversario.MinDate
            'dtpCumpleaños.Value = dtpCumpleaños.MinDate
            dbcSucursales.Text = ""
            'chkAniversario.Text = 0
            'chkCumpleaños.Text = 0
            'chkFechaNacimiento.Text = 0
            '_optVigente_0.Text = True
            RTObservaciones.Text = ""
            RTObservaciones.Tag = ""
            RTObservaciones.Enabled = False
            'dtpFechaAlta.Value = ""
            dbcAlmacen.Text = ""
            dbcAlmacen.Tag = ""
            dbcAlmacen.Enabled = False
            chkVendExterno.Enabled = False
            'InicializaVariables()
            DesHabilitarControles()
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Function Cambios() As Boolean
        Cambios = True
        If chkVendExterno.Checked <> intVendExterno Then Exit Function
        If Trim(dbcAlmacen.Text) <> dbcAlmacen.Tag Then Exit Function
        Cambios = False
    End Function

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Trim(txtFormQueEjecuta.Text) = "frmVentasSalMercancia" Then
            ValidaDatos = ValidaDatosDesdeVenta()
            Exit Function
        End If
        If chkVendExterno.Checked = 1 Then
            If Len(Trim(dbcAlmacen.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Descripción del almacen", vbInformation + vbOKOnly, gstrNombCortoEmpresa)
                dbcAlmacen.Focus()
                Exit Function
            End If
        End If
        ValidaDatos = True
    End Function

    Function ValidaDatosDesdeVenta() As Boolean
        'Esta Función de Validar Datos es Especial y
        'sólo se Aplica cuando el Formulario de Clientes se ha mandado llamar desde el Formulario de Ventas Salida de Mercancía.
        'Unicamente será obligatorio Proporcionar :
        ' -Nombre del CLiente
        ' -Fecha de Nacimiento
        ' -Fecha de Nac. de su Conyuge
        ' -Fecha de Aniversario
        ' -y Telefonos.

        'ValidaDatosDesdeVenta = False
        'If Len(Trim(txtNombre)) = 0 Then
        '    MsgBox C_msgFALTADATO & "Nombre del cliente", vbInformation, gstrNombCortoEmpresa
        'txtNombre.SetFocus
        '    Exit Function
        'End If
        'If Len(Trim(txtTelCasa)) = 0 Then
        '    MsgBox C_msgFALTADATO & "Teléfono de casa", vbInformation, gstrNombCortoEmpresa
        'txtTelCasa.SetFocus
        '    Exit Function
        'End If
        'If chkFechaNacimiento.Value = vbUnchecked Then
        '    MsgBox C_msgFALTADATO & "Fecha de nacimiento", vbInformation, gstrNombCortoEmpresa
        'chkFechaNacimiento.SetFocus
        '    Exit Function
        'End If
        'If chkCumpleaños.Value = vbUnchecked Then
        '    MsgBox C_msgFALTADATO & "Cumpleaños conyuge", vbInformation, gstrNombCortoEmpresa
        'chkCumpleaños.SetFocus
        '    Exit Function
        'End If
        'If chkAniversario.Value = vbUnchecked Then
        '    MsgBox C_msgFALTADATO & "Fecha de aniversario de bodas", vbInformation, gstrNombCortoEmpresa
        'chkAniversario.SetFocus
        '    Exit Function
        'End If


        'ValidaDatosDesdeVenta = True
    End Function

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnVigente = True
        mblnCancelado = False
        mblnSuspension = False
        mblnSalir = False
        FueraChange = False
        intVendExterno = 0
        'dtpAniversarioBodasTag = dtpAniversario.MinDate
        'dtpCumpleañosConyugeTag = dtpCumpleaños.MinDate
        'dtpFecNacimientoTag = dtpFechaNacimiento.MinDate
    End Sub

    Private Sub cboTipoCliente_GotFocus()
        Pon_Tool()
    End Sub

    Private Sub chkAniversario_Click()
        'If chkAniversario Then
        '    dtpAniversario.Enabled = True
        '    dtpCumpleaños.Enabled = True
        '    dtpAniversario.Year = Year("1 /1 /2000")
        '    dtpCumpleaños.Year = Year("1 /1 /2000")
        '    chkCumpleaños.Value = 1
        'Else
        '    dtpAniversario.Enabled = False
        '    dtpCumpleaños.Enabled = False
        '    dtpCumpleaños.Value = dtpCumpleaños.MinDate
        '    dtpAniversario.Value = dtpAniversario.MinDate
        '    chkCumpleaños.Value = 0
        'End If
    End Sub

    Private Sub chkAniversario_GotFocus()
        Pon_Tool()
    End Sub

    Private Sub chkCumpleaños_Click()
        'If chkCumpleaños Then
        '    dtpCumpleaños.Enabled = True
        '    dtpAniversario.Enabled = True
        '    dtpAniversario.Year = Year("1 /1 /2000")
        '    dtpCumpleaños.Year = Year("1 /1 /2000")
        '    chkAniversario.Value = 1
        'Else
        '    dtpCumpleaños.Enabled = False
        '    dtpAniversario.Enabled = False
        '    dtpCumpleaños.Value = dtpCumpleaños.MinDate
        '    dtpAniversario.Value = dtpAniversario.MinDate
        '    chkAniversario.Value = 0
        'End If
    End Sub

    Private Sub chkCumpleaños_GotFocus()
        Pon_Tool()
    End Sub

    Private Sub chkFechaNacimiento_Click()
        'If chkFechaNacimiento Then
        '    dtpFechaNacimiento.Enabled = True
        '    dtpFechaNacimiento.Year = Year("1 /1 /2000")
        '    chkFechaNacimiento.Value = 1
        'Else
        '    dtpFechaNacimiento.Enabled = False
        '    dtpFechaNacimiento.Value = dtpAniversario.MinDate
        '    chkFechaNacimiento.Value = 0
        'End If
    End Sub


    Private Sub chkVendExterno_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkVendExterno.CheckStateChanged
        If chkVendExterno.CheckState = 1 Then
            If CDbl(Numerico(txtCodigo.Text)) = 1 Then
                MsgBox("El cliente público en general" & vbNewLine & "No puede ser vendedor externo" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                chkVendExterno.CheckState = System.Windows.Forms.CheckState.Unchecked
                Exit Sub
            End If
            dbcAlmacen.Enabled = True
        ElseIf chkVendExterno.CheckState = 0 Then
            dbcAlmacen.Text = ""
            intCodAlmacen = 0
            dbcAlmacen.Enabled = False
        End If
    End Sub

    Private Sub dbcAlmacen_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcAlmacen.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcAlmacen" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen, LTRIM(RTRIM(DescAlmacen)) AS DescAlmacen From CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcAlmacen.Text) & "%' AND TipoAlmacen = 'V' AND (CODALMACEN NOT IN(SELECT ALMACENVEXT FROM CATCLIENTES WHERE ISNULL(ALMACENVEXT,0) > 0 )) ORDER BY DescAlmacen "

        DCChange(gStrSql, tecla)
        intCodAlmacen = 0
    End Sub

    Private Sub dbcAlmacen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcAlmacen.Enter
        gStrSql = "SELECT CodAlmacen, LTRIM(RTRIM(DescAlmacen)) AS DescAlmacen From CatAlmacen WHERE TipoAlmacen = 'V' AND (CODALMACEN NOT IN(SELECT ALMACENVEXT FROM CATCLIENTES WHERE ISNULL(ALMACENVEXT,0) > 0 )) ORDER BY DescAlmacen "
        DCGotFocus(gStrSql, dbcAlmacen)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcAlmacen_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcAlmacen.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkVendExterno.Focus()
        End If
    End Sub

    Private Sub dbcAlmacen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dbcAlmacen.KeyPress
        'eventArgs.keyAscii = ModEstandar.gp_CampoMayusculas(eventArgs.keyAscii)
    End Sub

    Private Sub dbcAlmacen_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcAlmacen.KeyUp
        Dim Aux As String
        Aux = dbcAlmacen.Text
        'If dbcAlmacen.SelectedItem <> 0 Then
        '    dbcAlmacen_Leave(dbcAlmacen, New System.EventArgs())
        'End If
        dbcAlmacen.Text = Aux
    End Sub

    Private Sub dbcAlmacen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcAlmacen.Leave
        gStrSql = "SELECT CodAlmacen, LTRIM(RTRIM(DescAlmacen)) AS DescAlmacen From CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcAlmacen.Text) & "%' AND TipoAlmacen = 'V' AND (CODALMACEN NOT IN(SELECT ALMACENVEXT FROM CATCLIENTES WHERE ISNULL(ALMACENVEXT,0) > 0 )) ORDER BY DescAlmacen "
        DCLostFocus(dbcAlmacen, gStrSql, intCodAlmacen)
    End Sub

    Private Sub dbcAlmacen_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcAlmacen.MouseUp
        Dim Aux As String
        Aux = dbcAlmacen.Text
        'If dbcAlmacen.SelectedItem <> 0 Then
        '    dbcAlmacen_Leave(dbcAlmacen, New System.EventArgs())
        'End If
        dbcAlmacen.Text = Aux
    End Sub



    Private Sub dtpAniversario_Click()
        'sglTiempoCambio = Timer
    End Sub

    Private Sub dtpAniversario_GotFocus()
        Pon_Tool()
    End Sub

    Private Sub dtpAniversario_KeyPress(KeyAscii As Integer)
        'sglTiempoCambio = Timer
    End Sub

    Private Sub dtpCumpleaños_Click()
        'sglTiempoCambio = Timer
    End Sub

    Private Sub dtpCumpleaños_GotFocus()
        Pon_Tool()
    End Sub

    Private Sub dtpCumpleaños_KeyPress(KeyAscii As Integer)
        'sglTiempoCambio = Timer
    End Sub

    Private Sub dtpFechaNacimiento_Click()
        'sglTiempoCambio = Timer
    End Sub

    Private Sub dtpFechaNacimiento_GotFocus()
        Pon_Tool()
    End Sub

    Private Sub dtpFechaNacimiento_KeyPress(KeyAscii As Integer)
        'sglTiempoCambio = Timer
    End Sub

    Private Sub Form_Activate()
        ' ModEstandar.ActivaMenu C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO
        'Me.ZOrder
    End Sub

    Private Sub Form_Deactivate()
        'ModEstandar.ActivaMenu C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO
    End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        '    Select Case KeyCode
        '        Case vbKeyReturn
        '            AvanzarTab Me
        '    Case vbKeyEscape
        '            RetrocederTab Me
        'End Select
    End Sub

    Private Sub Form_KeyPress(KeyAscii As Integer)
        'If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        'If Screen.ActiveForm.ActiveControl.Name <> "txtEmail" Then
        '    KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        'End If
    End Sub


    Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        '    ModEstandar.RestaurarForma Me, False
        'If mblnSalir Then
        '        Select Case MsgBox(C_msgSALIR, vbYesNo + vbQuestion + vbDefaultButton2, gstrNombCortoEmpresa)
        '            Case vbYes
        '                Cancel = 0
        '            Case vbNo
        '                mblnSalir = False
        '                Cancel = 1
        '        End Select
        '    End If
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        '        ModEstandar.ActivaMenu C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO
        '    ModEstandar.LimpiaDescBarraEstado()
        '        Set frmCorpoABCClientes = Nothing
        ''    MenuPrincipal.mnuCatalogosOpc(0).Enabled = True
    End Sub

    Private Sub optCancelado_GotFocus(Index As Integer)
        Pon_Tool()
    End Sub

    Private Sub optSuspension_GotFocus(Index As Integer)
        Pon_Tool()
    End Sub

    Private Sub optVigente_GotFocus(Index As Integer)
        Pon_Tool()
    End Sub

    Private Sub txtCiudad_GotFocus()
        'SelTextoTxt txtCiudad
        Pon_Tool()
    End Sub

    Private Sub txtCiudad_KeyPress(KeyAscii As Integer)
        'ModEstandar.gp_CampoLetras KeyAscii
    End Sub

    Private Sub txtCodigo_GotFocus()
        'SelTextoTxt txtCodigo
        Pon_Tool()
    End Sub

    Private Sub txtCodigo_Change()
        If FueraChange = True Then Exit Sub
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
        '    'Pregunta solo si existieron cambios
        '    If Cambios = True And KeyCode = vbKeyDelete Then
        '        Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
        '            Case vbYes: 'Guardar el registro
        '                If Guardar = False Then
        '                    KeyCode = 0
        '                    Exit Sub
        '                End If
        '            Case vbNo: 'No hace nada y permite que se borre el contenido del text
        '            Case vbCancel: 'Cancela la captura
        '                txtCodigo.SetFocus
        '                KeyCode = 0
        '                Exit Sub
        '        End Select
        '    End If
    End Sub

    'Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    '    'If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    '    '    KeyAscii = 0
    '    'Else
    '    '    'Pregunta solo si existieron cambios
    '    '    If Cambios() = True And mblnNuevo = False Then
    '    '        Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
    '    '            Case vbYes 'Guardar el registro
    '    '                If Guardar() = False Then
    '    '                    KeyAscii = 0
    '    '                    Exit Sub
    '    '                End If
    '    '            Case vbNo 'No hace nada y permite que se teclee y borre
    '    '            Case vbCancel 'Cancela la captura
    '    '                txtCodigo.SetFocus
    '    '                KeyAscii = 0
    '    '                Exit Sub
    '    '        End Select
    '    '    End If
    '    'End If
    'End Sub

    'Private Sub txtCodigo_LostFocus()
    '    If Screen.ActiveForm.Caption <> Me.Caption Then
    '        Exit Sub
    '    End If
    '    If mblnCambiosEnCodigo = True And txtCodigo <> "" Then 'si hubo cambios en el codigo hace la consulta
    '        LlenaDatos()
    '    End If
    'End Sub

    Private Sub txtCodPostal_GotFocus()
        ' SelTextoTxt txtCodPostal
        Pon_Tool()
    End Sub

    Private Sub txtCodPostal_KeyPress(KeyAscii As Integer)
        ' ModEstandar.gp_CampoNumerico KeyAscii
    End Sub

    Private Sub txtColonia_GotFocus()
        'SelTextoTxt txtColonia
        Pon_Tool()
    End Sub

    Private Sub txtDomicilio_GotFocus()
        ' SelTextoTxt txtDomicilio
        Pon_Tool()
    End Sub

    Private Sub txtEmail_GotFocus()
        'txtEmail.FontUnderline = False
        'SelTextoTxt txtEmail
        Pon_Tool()
    End Sub

    Private Sub txtEmail_KeyPress(KeyAscii As Integer)
        'ModEstandar.gp_CampoAlfanumerico KeyAscii, "@_!#$%&/()=?'¡¿*-+\,.:;|°~"
    End Sub

    Private Sub txtEmail_LostFocus()
        ' txtEmail.FontUnderline = True
    End Sub

    Private Sub txtFax_GotFocus()
        'SelTextoTxt txtFax
        Pon_Tool()
    End Sub

    Private Sub txtFax_KeyPress(KeyAscii As Integer)
        ' ModEstandar.gp_CampoNumerico KeyAscii, "-"
    End Sub

    Private Sub txtNombre_GotFocus()
        '  SelTextoTxt txtNombre
        Pon_Tool()
    End Sub

    Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
        'ModEstandar.gp_CampoLetras KeyAscii
    End Sub

    Private Sub rtObservaciones_GotFocus()
        'RTObservaciones.SelStart = 0
        'RTObservaciones.SelLength = 0
        '''SelTextoTxt RTObservaciones
        Pon_Tool()
    End Sub

    Private Sub txtRFC_GotFocus()
        'SelTextoTxt txtRFC
        Pon_Tool()
    End Sub

    Private Sub txtRFC_KeyPress(KeyAscii As Integer)
        'KeyAscii = ModEstandar.Valida_RFC(txtRFC, KeyAscii, Len(txtRFC) + 1)
    End Sub

    Private Sub txtTelCasa_GotFocus()
        'SelTextoTxt txtTelCasa
        Pon_Tool()
    End Sub

    Private Sub txtTelCasa_KeyPress(KeyAscii As Integer)
        'ModEstandar.gp_CampoNumerico KeyAscii, "-"
    End Sub

    Private Sub txtTelOficina_GotFocus()
        ' SelTextoTxt txtTelOficina
        Pon_Tool()
    End Sub

    Private Sub txtTelOficina_KeyPress(KeyAscii As Integer)
        'ModEstandar.gp_CampoNumerico KeyAscii, "-"
    End Sub

    Sub PonerCodigoSucursal()
        '    gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales) & "' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        '    ModEstandar.BorraCmd()
        '    Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        'Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        'Set RsGral = Cmd.Execute

        'If RsGral.RecordCount = 0 Then
        '        intCodSucursal = 0
        '    Else
        '        intCodSucursal = RsGral!CodAlmacen
        '    End If

    End Sub

    Private Sub dbcSucursales_Change()
        If FueraChange = True Then Exit Sub
        'If Screen.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE TipoAlmacen ='P' and  DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        Nuevo()
        txtCodigo.Text = ""
    End Sub

    Private Sub dbcSucursales_GotFocus(sender As Object, e As EventArgs) Handles dbcSucursales.GotFocus
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen  FROM CatAlmacen where  TipoAlmacen ='P'  ORDER BY DescAlmacen"
        DCGotFocus(gStrSql)
    End Sub

    Private Sub dbcSucursales_KeyDown(KeyCode As Integer, Shift As Integer)
        tecla = KeyCode
        '    If KeyCode = vbKeyEscape Then
        '        mblnSalir = True
        '        Unload Me
        'End If
    End Sub

    Private Sub dbcSucursales_KeyUp(KeyCode As Integer, Shift As Integer)
        'If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        '    PonerCodigoSucursal()
        '    '        Buscar
        '    Exit Sub
        'End If
    End Sub

    Private Sub dbcSucursales_LostFocus(sender As Object, e As EventArgs) Handles dbcSucursales.LostFocus
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE  TipoAlmacen ='P' and  DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
    End Sub

    Private Sub dbcSucursales_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        PonerCodigoSucursal()
        '    Buscar
    End Sub

    Private Sub _Label1_0_Click(sender As Object, e As EventArgs) Handles _Label1_0.Click

    End Sub

    Private Sub txtCodigo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCodigo.KeyPress

    End Sub

    Private Sub txtCodigo_ChangeUICues(sender As Object, e As UICuesEventArgs) Handles txtCodigo.ChangeUICues
        'LlenaDatos()
        'Else
        '    MessageBox.Show("Ingrese un codigo de cliente!")
        'End If
    End Sub

    Private Sub txtCodigo_LostFocus(sender As Object, e As EventArgs) Handles txtCodigo.LostFocus
        'If mblnCambiosEnCodigo = True And txtCodigo.Text <> "" Then 
        If (txtCodigo.Text <> "") Then
            LlenaDatos()
        End If
    End Sub

    Private Sub txtCodigo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCodigo.KeyDown
        If (txtCodigo.Text <> "" And e.KeyCode = Keys.Enter) Then
            txtCodigo_LostFocus(New Object, New EventArgs)
        End If
    End Sub


    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        'Eliminar()
    End Sub

    Private Sub cmdRptCtes_Click(sender As Object, e As EventArgs) Handles cmdRptCtes.Click
        Dim f1 As frmCorpoRptClientes = New frmCorpoRptClientes()
        f1.Show()
        'f1.ZOrder()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._optSuspension_2 = New System.Windows.Forms.RadioButton()
        Me._optCancelado_1 = New System.Windows.Forms.RadioButton()
        Me._optVigente_0 = New System.Windows.Forms.RadioButton()
        Me.txtConyuge = New System.Windows.Forms.TextBox()
        Me.txtCiudad = New System.Windows.Forms.TextBox()
        Me.cboTipoCliente = New System.Windows.Forms.ComboBox()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.txtCodPostal = New System.Windows.Forms.TextBox()
        Me.txtFax = New System.Windows.Forms.TextBox()
        Me.txtTelOficina = New System.Windows.Forms.TextBox()
        Me.txtTelCasa = New System.Windows.Forms.TextBox()
        Me.txtRFC = New System.Windows.Forms.TextBox()
        Me.txtColonia = New System.Windows.Forms.TextBox()
        Me.txtDomicilio = New System.Windows.Forms.TextBox()
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me.txtCodigo = New System.Windows.Forms.TextBox()
        Me.cmdRptCtes = New System.Windows.Forms.Button()
        Me.RTObservaciones = New System.Windows.Forms.RichTextBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.dbcAlmacen = New System.Windows.Forms.ComboBox()
        Me.chkVendExterno = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtFormQueEjecuta = New System.Windows.Forms.TextBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me.dtpFechaAlta = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaNacimiento = New System.Windows.Forms.DateTimePicker()
        Me.dtpCumpleaños = New System.Windows.Forms.DateTimePicker()
        Me.dtpAniversario = New System.Windows.Forms.DateTimePicker()
        Me.chkFechaNacimiento = New System.Windows.Forms.CheckBox()
        Me.chkAniversario = New System.Windows.Forms.CheckBox()
        Me.chkCumpleaños = New System.Windows.Forms.CheckBox()
        Me._Label1_14 = New System.Windows.Forms.Label()
        Me._Label1_13 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_12 = New System.Windows.Forms.Label()
        Me._Label1_11 = New System.Windows.Forms.Label()
        Me._Label1_10 = New System.Windows.Forms.Label()
        Me._Label1_9 = New System.Windows.Forms.Label()
        Me._Label1_8 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_4 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optCancelado = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optSuspension = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optVigente = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optCancelado, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optSuspension, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optVigente, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        '_optSuspension_2
        '
        Me._optSuspension_2.BackColor = System.Drawing.Color.Silver
        Me._optSuspension_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optSuspension_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optSuspension_2.Location = New System.Drawing.Point(265, 17)
        Me._optSuspension_2.Margin = New System.Windows.Forms.Padding(2)
        Me._optSuspension_2.Name = "_optSuspension_2"
        Me._optSuspension_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optSuspension_2.Size = New System.Drawing.Size(92, 21)
        Me._optSuspension_2.TabIndex = 22
        Me._optSuspension_2.TabStop = True
        Me._optSuspension_2.Text = "Suspendido"
        Me.ToolTip1.SetToolTip(Me._optSuspension_2, "Estatus Suspendido")
        Me._optSuspension_2.UseVisualStyleBackColor = False
        '
        '_optCancelado_1
        '
        Me._optCancelado_1.BackColor = System.Drawing.Color.Silver
        Me._optCancelado_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optCancelado_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optCancelado_1.Location = New System.Drawing.Point(136, 17)
        Me._optCancelado_1.Margin = New System.Windows.Forms.Padding(2)
        Me._optCancelado_1.Name = "_optCancelado_1"
        Me._optCancelado_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optCancelado_1.Size = New System.Drawing.Size(80, 21)
        Me._optCancelado_1.TabIndex = 21
        Me._optCancelado_1.TabStop = True
        Me._optCancelado_1.Text = "Cancelado"
        Me.ToolTip1.SetToolTip(Me._optCancelado_1, "Estatus Cancelado")
        Me._optCancelado_1.UseVisualStyleBackColor = False
        '
        '_optVigente_0
        '
        Me._optVigente_0.BackColor = System.Drawing.Color.Silver
        Me._optVigente_0.Checked = True
        Me._optVigente_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optVigente_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optVigente_0.Location = New System.Drawing.Point(26, 17)
        Me._optVigente_0.Margin = New System.Windows.Forms.Padding(2)
        Me._optVigente_0.Name = "_optVigente_0"
        Me._optVigente_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optVigente_0.Size = New System.Drawing.Size(64, 21)
        Me._optVigente_0.TabIndex = 20
        Me._optVigente_0.TabStop = True
        Me._optVigente_0.Text = "Vigente"
        Me.ToolTip1.SetToolTip(Me._optVigente_0, "Estatus Vigente")
        Me._optVigente_0.UseVisualStyleBackColor = False
        '
        'txtConyuge
        '
        Me.txtConyuge.AcceptsReturn = True
        Me.txtConyuge.BackColor = System.Drawing.SystemColors.Window
        Me.txtConyuge.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConyuge.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConyuge.Location = New System.Drawing.Point(87, 87)
        Me.txtConyuge.Margin = New System.Windows.Forms.Padding(2)
        Me.txtConyuge.MaxLength = 40
        Me.txtConyuge.Name = "txtConyuge"
        Me.txtConyuge.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConyuge.Size = New System.Drawing.Size(420, 20)
        Me.txtConyuge.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtConyuge, "Domicilio")
        '
        'txtCiudad
        '
        Me.txtCiudad.AcceptsReturn = True
        Me.txtCiudad.BackColor = System.Drawing.SystemColors.Window
        Me.txtCiudad.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCiudad.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCiudad.Location = New System.Drawing.Point(361, 135)
        Me.txtCiudad.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCiudad.MaxLength = 30
        Me.txtCiudad.Name = "txtCiudad"
        Me.txtCiudad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCiudad.Size = New System.Drawing.Size(146, 20)
        Me.txtCiudad.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtCiudad, "Ciudad")
        '
        'cboTipoCliente
        '
        Me.cboTipoCliente.BackColor = System.Drawing.SystemColors.Window
        Me.cboTipoCliente.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboTipoCliente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTipoCliente.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboTipoCliente.Items.AddRange(New Object() {"A", "B", "C", "D", "ESPECIAL"})
        Me.cboTipoCliente.Location = New System.Drawing.Point(653, 98)
        Me.cboTipoCliente.Margin = New System.Windows.Forms.Padding(2)
        Me.cboTipoCliente.Name = "cboTipoCliente"
        Me.cboTipoCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboTipoCliente.Size = New System.Drawing.Size(113, 21)
        Me.cboTipoCliente.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.cboTipoCliente, "Tipo de Cliente")
        '
        'txtEmail
        '
        Me.txtEmail.AcceptsReturn = True
        Me.txtEmail.BackColor = System.Drawing.SystemColors.Window
        Me.txtEmail.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtEmail.ForeColor = System.Drawing.Color.Blue
        Me.txtEmail.Location = New System.Drawing.Point(361, 206)
        Me.txtEmail.Margin = New System.Windows.Forms.Padding(2)
        Me.txtEmail.MaxLength = 50
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEmail.Size = New System.Drawing.Size(146, 20)
        Me.txtEmail.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.txtEmail, "Email")
        '
        'txtCodPostal
        '
        Me.txtCodPostal.AcceptsReturn = True
        Me.txtCodPostal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodPostal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodPostal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodPostal.Location = New System.Drawing.Point(87, 158)
        Me.txtCodPostal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodPostal.MaxLength = 10
        Me.txtCodPostal.Name = "txtCodPostal"
        Me.txtCodPostal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodPostal.Size = New System.Drawing.Size(201, 20)
        Me.txtCodPostal.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtCodPostal, "Codigo Postal")
        '
        'txtFax
        '
        Me.txtFax.AcceptsReturn = True
        Me.txtFax.BackColor = System.Drawing.SystemColors.Window
        Me.txtFax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFax.Location = New System.Drawing.Point(87, 207)
        Me.txtFax.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFax.MaxLength = 20
        Me.txtFax.Name = "txtFax"
        Me.txtFax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFax.Size = New System.Drawing.Size(201, 20)
        Me.txtFax.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.txtFax, "Fax")
        '
        'txtTelOficina
        '
        Me.txtTelOficina.AcceptsReturn = True
        Me.txtTelOficina.BackColor = System.Drawing.SystemColors.Window
        Me.txtTelOficina.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTelOficina.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTelOficina.Location = New System.Drawing.Point(361, 183)
        Me.txtTelOficina.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTelOficina.MaxLength = 20
        Me.txtTelOficina.Name = "txtTelOficina"
        Me.txtTelOficina.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTelOficina.Size = New System.Drawing.Size(146, 20)
        Me.txtTelOficina.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txtTelOficina, "Telefono de la Oficina")
        '
        'txtTelCasa
        '
        Me.txtTelCasa.AcceptsReturn = True
        Me.txtTelCasa.BackColor = System.Drawing.SystemColors.Window
        Me.txtTelCasa.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTelCasa.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTelCasa.Location = New System.Drawing.Point(87, 183)
        Me.txtTelCasa.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTelCasa.MaxLength = 20
        Me.txtTelCasa.Name = "txtTelCasa"
        Me.txtTelCasa.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTelCasa.Size = New System.Drawing.Size(201, 20)
        Me.txtTelCasa.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.txtTelCasa, "Telefono de la Casa")
        '
        'txtRFC
        '
        Me.txtRFC.AcceptsReturn = True
        Me.txtRFC.BackColor = System.Drawing.SystemColors.Window
        Me.txtRFC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRFC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRFC.Location = New System.Drawing.Point(361, 159)
        Me.txtRFC.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRFC.MaxLength = 15
        Me.txtRFC.Name = "txtRFC"
        Me.txtRFC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRFC.Size = New System.Drawing.Size(146, 20)
        Me.txtRFC.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtRFC, "RFC")
        '
        'txtColonia
        '
        Me.txtColonia.AcceptsReturn = True
        Me.txtColonia.BackColor = System.Drawing.SystemColors.Window
        Me.txtColonia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtColonia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtColonia.Location = New System.Drawing.Point(87, 135)
        Me.txtColonia.Margin = New System.Windows.Forms.Padding(2)
        Me.txtColonia.MaxLength = 30
        Me.txtColonia.Name = "txtColonia"
        Me.txtColonia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtColonia.Size = New System.Drawing.Size(201, 20)
        Me.txtColonia.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtColonia, "Colonia")
        '
        'txtDomicilio
        '
        Me.txtDomicilio.AcceptsReturn = True
        Me.txtDomicilio.BackColor = System.Drawing.SystemColors.Window
        Me.txtDomicilio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDomicilio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDomicilio.Location = New System.Drawing.Point(87, 111)
        Me.txtDomicilio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDomicilio.MaxLength = 65
        Me.txtDomicilio.Name = "txtDomicilio"
        Me.txtDomicilio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDomicilio.Size = New System.Drawing.Size(420, 20)
        Me.txtDomicilio.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtDomicilio, "Domicilio")
        '
        'txtNombre
        '
        Me.txtNombre.AcceptsReturn = True
        Me.txtNombre.BackColor = System.Drawing.SystemColors.Window
        Me.txtNombre.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombre.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombre.Location = New System.Drawing.Point(87, 58)
        Me.txtNombre.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNombre.MaxLength = 40
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombre.Size = New System.Drawing.Size(420, 20)
        Me.txtNombre.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtNombre, "Nombre")
        '
        'txtCodigo
        '
        Me.txtCodigo.AcceptsReturn = True
        Me.txtCodigo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodigo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodigo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodigo.Location = New System.Drawing.Point(13, 57)
        Me.txtCodigo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodigo.MaxLength = 5
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodigo.Size = New System.Drawing.Size(61, 20)
        Me.txtCodigo.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCodigo, "Codigo del Cliente")
        '
        'cmdRptCtes
        '
        Me.cmdRptCtes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRptCtes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRptCtes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRptCtes.Location = New System.Drawing.Point(275, 31)
        Me.cmdRptCtes.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdRptCtes.Name = "cmdRptCtes"
        Me.cmdRptCtes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRptCtes.Size = New System.Drawing.Size(107, 40)
        Me.cmdRptCtes.TabIndex = 50
        Me.cmdRptCtes.Text = "Repor&te de Clientes"
        Me.cmdRptCtes.UseVisualStyleBackColor = False
        '
        'RTObservaciones
        '
        Me.RTObservaciones.BackColor = System.Drawing.Color.White
        Me.RTObservaciones.Location = New System.Drawing.Point(14, 31)
        Me.RTObservaciones.Margin = New System.Windows.Forms.Padding(2)
        Me.RTObservaciones.Name = "RTObservaciones"
        Me.RTObservaciones.ReadOnly = True
        Me.RTObservaciones.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
        Me.RTObservaciones.Size = New System.Drawing.Size(348, 46)
        Me.RTObservaciones.TabIndex = 49
        Me.RTObservaciones.Text = ""
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.Color.Silver
        Me.Frame3.Controls.Add(Me.dbcAlmacen)
        Me.Frame3.Controls.Add(Me.chkVendExterno)
        Me.Frame3.Controls.Add(Me.Label3)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(421, 272)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(377, 57)
        Me.Frame3.TabIndex = 45
        Me.Frame3.TabStop = False
        '
        'dbcAlmacen
        '
        Me.dbcAlmacen.FormattingEnabled = True
        Me.dbcAlmacen.Location = New System.Drawing.Point(98, 25)
        Me.dbcAlmacen.Name = "dbcAlmacen"
        Me.dbcAlmacen.Size = New System.Drawing.Size(264, 21)
        Me.dbcAlmacen.TabIndex = 58
        '
        'chkVendExterno
        '
        Me.chkVendExterno.BackColor = System.Drawing.Color.Silver
        Me.chkVendExterno.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVendExterno.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkVendExterno.Location = New System.Drawing.Point(6, 0)
        Me.chkVendExterno.Margin = New System.Windows.Forms.Padding(2)
        Me.chkVendExterno.Name = "chkVendExterno"
        Me.chkVendExterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVendExterno.Size = New System.Drawing.Size(121, 25)
        Me.chkVendExterno.TabIndex = 23
        Me.chkVendExterno.Text = "Vendedor Externo"
        Me.chkVendExterno.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Silver
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(37, 29)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(62, 14)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Almacen :"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.Color.Silver
        Me.Frame2.Controls.Add(Me._optSuspension_2)
        Me.Frame2.Controls.Add(Me._optCancelado_1)
        Me.Frame2.Controls.Add(Me._optVigente_0)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(17, 272)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(398, 57)
        Me.Frame2.TabIndex = 38
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Estatus "
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.Color.Silver
        Me.Frame1.Controls.Add(Me.txtFormQueEjecuta)
        Me.Frame1.Controls.Add(Me.dbcSucursales)
        Me.Frame1.Controls.Add(Me.dtpFechaAlta)
        Me.Frame1.Controls.Add(Me.dtpFechaNacimiento)
        Me.Frame1.Controls.Add(Me.dtpCumpleaños)
        Me.Frame1.Controls.Add(Me.dtpAniversario)
        Me.Frame1.Controls.Add(Me.txtConyuge)
        Me.Frame1.Controls.Add(Me.chkFechaNacimiento)
        Me.Frame1.Controls.Add(Me.txtFax)
        Me.Frame1.Controls.Add(Me.txtCiudad)
        Me.Frame1.Controls.Add(Me.chkAniversario)
        Me.Frame1.Controls.Add(Me.chkCumpleaños)
        Me.Frame1.Controls.Add(Me.cboTipoCliente)
        Me.Frame1.Controls.Add(Me.txtEmail)
        Me.Frame1.Controls.Add(Me.txtCodPostal)
        Me.Frame1.Controls.Add(Me.txtTelOficina)
        Me.Frame1.Controls.Add(Me.txtTelCasa)
        Me.Frame1.Controls.Add(Me.txtRFC)
        Me.Frame1.Controls.Add(Me.txtColonia)
        Me.Frame1.Controls.Add(Me.txtDomicilio)
        Me.Frame1.Controls.Add(Me.txtNombre)
        Me.Frame1.Controls.Add(Me.txtCodigo)
        Me.Frame1.Controls.Add(Me._Label1_14)
        Me.Frame1.Controls.Add(Me._Label1_13)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Controls.Add(Me._Label1_12)
        Me.Frame1.Controls.Add(Me._Label1_11)
        Me.Frame1.Controls.Add(Me._Label1_10)
        Me.Frame1.Controls.Add(Me._Label1_9)
        Me.Frame1.Controls.Add(Me._Label1_8)
        Me.Frame1.Controls.Add(Me._Label1_7)
        Me.Frame1.Controls.Add(Me._Label1_6)
        Me.Frame1.Controls.Add(Me._Label1_5)
        Me.Frame1.Controls.Add(Me._Label1_4)
        Me.Frame1.Controls.Add(Me._Label1_3)
        Me.Frame1.Controls.Add(Me._Label1_2)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(17, 19)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(781, 249)
        Me.Frame1.TabIndex = 25
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Datos del Cliente "
        '
        'txtFormQueEjecuta
        '
        Me.txtFormQueEjecuta.Location = New System.Drawing.Point(584, 132)
        Me.txtFormQueEjecuta.Name = "txtFormQueEjecuta"
        Me.txtFormQueEjecuta.Size = New System.Drawing.Size(182, 20)
        Me.txtFormQueEjecuta.TabIndex = 58
        Me.txtFormQueEjecuta.Visible = False
        '
        'dbcSucursales
        '
        Me.dbcSucursales.FormattingEnabled = True
        Me.dbcSucursales.Location = New System.Drawing.Point(616, 12)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(150, 21)
        Me.dbcSucursales.TabIndex = 57
        '
        'dtpFechaAlta
        '
        Me.dtpFechaAlta.CustomFormat = ""
        Me.dtpFechaAlta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaAlta.Location = New System.Drawing.Point(653, 73)
        Me.dtpFechaAlta.Name = "dtpFechaAlta"
        Me.dtpFechaAlta.Size = New System.Drawing.Size(113, 20)
        Me.dtpFechaAlta.TabIndex = 56
        '
        'dtpFechaNacimiento
        '
        Me.dtpFechaNacimiento.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaNacimiento.Location = New System.Drawing.Point(675, 161)
        Me.dtpFechaNacimiento.Name = "dtpFechaNacimiento"
        Me.dtpFechaNacimiento.Size = New System.Drawing.Size(91, 20)
        Me.dtpFechaNacimiento.TabIndex = 55
        '
        'dtpCumpleaños
        '
        Me.dtpCumpleaños.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpCumpleaños.Location = New System.Drawing.Point(675, 187)
        Me.dtpCumpleaños.Name = "dtpCumpleaños"
        Me.dtpCumpleaños.Size = New System.Drawing.Size(91, 20)
        Me.dtpCumpleaños.TabIndex = 54
        '
        'dtpAniversario
        '
        Me.dtpAniversario.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpAniversario.Location = New System.Drawing.Point(675, 213)
        Me.dtpAniversario.Name = "dtpAniversario"
        Me.dtpAniversario.Size = New System.Drawing.Size(91, 20)
        Me.dtpAniversario.TabIndex = 53
        '
        'chkFechaNacimiento
        '
        Me.chkFechaNacimiento.BackColor = System.Drawing.Color.Silver
        Me.chkFechaNacimiento.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkFechaNacimiento.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkFechaNacimiento.Location = New System.Drawing.Point(564, 161)
        Me.chkFechaNacimiento.Margin = New System.Windows.Forms.Padding(2)
        Me.chkFechaNacimiento.Name = "chkFechaNacimiento"
        Me.chkFechaNacimiento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkFechaNacimiento.Size = New System.Drawing.Size(117, 20)
        Me.chkFechaNacimiento.TabIndex = 14
        Me.chkFechaNacimiento.Text = "Fecha Nacimiento"
        Me.chkFechaNacimiento.UseVisualStyleBackColor = False
        '
        'chkAniversario
        '
        Me.chkAniversario.BackColor = System.Drawing.Color.Silver
        Me.chkAniversario.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAniversario.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAniversario.Location = New System.Drawing.Point(565, 210)
        Me.chkAniversario.Margin = New System.Windows.Forms.Padding(2)
        Me.chkAniversario.Name = "chkAniversario"
        Me.chkAniversario.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAniversario.Size = New System.Drawing.Size(106, 32)
        Me.chkAniversario.TabIndex = 18
        Me.chkAniversario.Text = "Aniversario de  Bodas"
        Me.chkAniversario.UseVisualStyleBackColor = False
        '
        'chkCumpleaños
        '
        Me.chkCumpleaños.BackColor = System.Drawing.Color.Silver
        Me.chkCumpleaños.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCumpleaños.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCumpleaños.Location = New System.Drawing.Point(565, 188)
        Me.chkCumpleaños.Margin = New System.Windows.Forms.Padding(2)
        Me.chkCumpleaños.Name = "chkCumpleaños"
        Me.chkCumpleaños.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCumpleaños.Size = New System.Drawing.Size(84, 18)
        Me.chkCumpleaños.TabIndex = 16
        Me.chkCumpleaños.Text = "Cumpleaños Conyuge"
        Me.chkCumpleaños.UseVisualStyleBackColor = False
        '
        '_Label1_14
        '
        Me._Label1_14.BackColor = System.Drawing.Color.Silver
        Me._Label1_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_14.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_14.Location = New System.Drawing.Point(10, 87)
        Me._Label1_14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_14.Name = "_Label1_14"
        Me._Label1_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_14.Size = New System.Drawing.Size(56, 15)
        Me._Label1_14.TabIndex = 51
        Me._Label1_14.Text = "Conyuge :"
        '
        '_Label1_13
        '
        Me._Label1_13.BackColor = System.Drawing.Color.Silver
        Me._Label1_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_13.Location = New System.Drawing.Point(561, 15)
        Me._Label1_13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_13.Name = "_Label1_13"
        Me._Label1_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_13.Size = New System.Drawing.Size(52, 19)
        Me._Label1_13.TabIndex = 48
        Me._Label1_13.Text = "Sucursal"
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.Color.Silver
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(303, 138)
        Me._Label1_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(54, 20)
        Me._Label1_1.TabIndex = 44
        Me._Label1_1.Text = "Ciudad :"
        '
        '_Label1_12
        '
        Me._Label1_12.BackColor = System.Drawing.Color.Silver
        Me._Label1_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_12.Location = New System.Drawing.Point(581, 101)
        Me._Label1_12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_12.Name = "_Label1_12"
        Me._Label1_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_12.Size = New System.Drawing.Size(98, 12)
        Me._Label1_12.TabIndex = 37
        Me._Label1_12.Text = "Tipo Cliente :"
        '
        '_Label1_11
        '
        Me._Label1_11.BackColor = System.Drawing.Color.Silver
        Me._Label1_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_11.Location = New System.Drawing.Point(303, 209)
        Me._Label1_11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_11.Name = "_Label1_11"
        Me._Label1_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_11.Size = New System.Drawing.Size(54, 18)
        Me._Label1_11.TabIndex = 36
        Me._Label1_11.Text = "Email :"
        '
        '_Label1_10
        '
        Me._Label1_10.BackColor = System.Drawing.Color.Silver
        Me._Label1_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_10.Location = New System.Drawing.Point(11, 155)
        Me._Label1_10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_10.Name = "_Label1_10"
        Me._Label1_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_10.Size = New System.Drawing.Size(43, 20)
        Me._Label1_10.TabIndex = 35
        Me._Label1_10.Text = "C.P."
        '
        '_Label1_9
        '
        Me._Label1_9.BackColor = System.Drawing.Color.Silver
        Me._Label1_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_9.Location = New System.Drawing.Point(10, 204)
        Me._Label1_9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_9.Name = "_Label1_9"
        Me._Label1_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_9.Size = New System.Drawing.Size(43, 17)
        Me._Label1_9.TabIndex = 34
        Me._Label1_9.Text = "Fax :"
        '
        '_Label1_8
        '
        Me._Label1_8.BackColor = System.Drawing.Color.Silver
        Me._Label1_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_8.Location = New System.Drawing.Point(303, 186)
        Me._Label1_8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_8.Name = "_Label1_8"
        Me._Label1_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_8.Size = New System.Drawing.Size(54, 20)
        Me._Label1_8.TabIndex = 33
        Me._Label1_8.Text = "Oficina :"
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.Color.Silver
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Location = New System.Drawing.Point(11, 180)
        Me._Label1_7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(50, 15)
        Me._Label1_7.TabIndex = 32
        Me._Label1_7.Text = "Tel. Casa :"
        '
        '_Label1_6
        '
        Me._Label1_6.BackColor = System.Drawing.Color.Silver
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_6.Location = New System.Drawing.Point(303, 162)
        Me._Label1_6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(43, 20)
        Me._Label1_6.TabIndex = 31
        Me._Label1_6.Text = "RFC :"
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.Color.Silver
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_5.Location = New System.Drawing.Point(11, 135)
        Me._Label1_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(56, 20)
        Me._Label1_5.TabIndex = 30
        Me._Label1_5.Text = "Colonia :"
        '
        '_Label1_4
        '
        Me._Label1_4.BackColor = System.Drawing.Color.Silver
        Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_4.Location = New System.Drawing.Point(10, 111)
        Me._Label1_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_4.Size = New System.Drawing.Size(56, 20)
        Me._Label1_4.TabIndex = 29
        Me._Label1_4.Text = "Domicilio :"
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.Color.Silver
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_3.Location = New System.Drawing.Point(84, 39)
        Me._Label1_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(61, 17)
        Me._Label1_3.TabIndex = 28
        Me._Label1_3.Text = "Nombre :"
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.Color.Silver
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_2.Location = New System.Drawing.Point(11, 36)
        Me._Label1_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(50, 17)
        Me._Label1_2.TabIndex = 27
        Me._Label1_2.Text = "Codigo"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.Color.Silver
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(581, 79)
        Me._Label1_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(70, 12)
        Me._Label1_0.TabIndex = 26
        Me._Label1_0.Text = "Fecha Alta :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Silver
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(16, 12)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(104, 17)
        Me.Label2.TabIndex = 39
        Me.Label2.Text = "Observaciones :"
        '
        'btnGuardar
        '
        Me.btnGuardar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.grabar
        Me.btnGuardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnGuardar.Location = New System.Drawing.Point(12, 31)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(50, 42)
        Me.btnGuardar.TabIndex = 64
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.nuevo
        Me.btnLimpiar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnLimpiar.Location = New System.Drawing.Point(111, 31)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(50, 42)
        Me.btnLimpiar.TabIndex = 66
        Me.btnLimpiar.Text = " "
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.Eliminar
        Me.btnEliminar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnEliminar.Location = New System.Drawing.Point(62, 31)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(50, 42)
        Me.btnEliminar.TabIndex = 65
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.buscar
        Me.btnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBuscar.Location = New System.Drawing.Point(161, 31)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(50, 42)
        Me.btnBuscar.TabIndex = 67
        Me.btnBuscar.Text = " "
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Frame1)
        Me.Panel1.Controls.Add(Me.Frame3)
        Me.Panel1.Controls.Add(Me.Frame2)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(815, 442)
        Me.Panel1.TabIndex = 68
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Silver
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.RTObservaciones)
        Me.Panel2.Location = New System.Drawing.Point(421, 334)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(377, 96)
        Me.Panel2.TabIndex = 51
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.cmdRptCtes)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(17, 334)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(398, 96)
        Me.Panel3.TabIndex = 68
        '
        'btnSalir
        '
        Me.btnSalir.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.salir
        Me.btnSalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSalir.Location = New System.Drawing.Point(209, 31)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(50, 42)
        Me.btnSalir.TabIndex = 70
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'frmCorpoABCClientes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(838, 462)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(165, 125)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoABCClientes"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Abc a Clientes"
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optCancelado, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optSuspension, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optVigente, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub


End Class
