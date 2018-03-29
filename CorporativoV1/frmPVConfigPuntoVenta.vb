Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmPVConfigPuntoVenta
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA : CONFIGURACION DEL PUNTO DE VENTA                                                             *'
    '*AUTOR : JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO : LUNES 26 DE MAYO DE 2003                                                                     *'
    '*FECHA DE TERMINACION : MIERCOLES 28 DE MAYO DE 2003                                                                 *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkMostrarCodigoViejo As System.Windows.Forms.CheckBox
    Public WithEvents chkTransferenciaTicket As System.Windows.Forms.CheckBox
    Public WithEvents chkAplicarSucursales As System.Windows.Forms.CheckBox
    Public WithEvents cboCentavos As System.Windows.Forms.ComboBox
    Public WithEvents _optElectronica_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optDiskette_1 As System.Windows.Forms.RadioButton
    Public WithEvents Frame10 As System.Windows.Forms.GroupBox
    Public WithEvents chkCapturarCantArticulos As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents txtEfectivoMaxi As System.Windows.Forms.TextBox
    Public WithEvents txtSimbolo As System.Windows.Forms.TextBox
    Public WithEvents txtTipoCambioDolar As System.Windows.Forms.TextBox
    Public WithEvents txtPosicionesDeci As System.Windows.Forms.TextBox
    Public WithEvents _Label1_9 As System.Windows.Forms.Label
    Public WithEvents _Label1_18 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_8 As System.Windows.Forms.Label
    Public WithEvents txtMensajeFiscal As System.Windows.Forms.TextBox
    Public WithEvents txtMensajeDevoluciones As System.Windows.Forms.TextBox
    Public WithEvents txtMensajeVentas As System.Windows.Forms.TextBox
    Public WithEvents txtMensajeNormal As System.Windows.Forms.TextBox
    Public WithEvents _Label1_25 As System.Windows.Forms.Label
    Public WithEvents _Label1_13 As System.Windows.Forms.Label
    Public WithEvents _Label1_12 As System.Windows.Forms.Label
    Public WithEvents _Label1_11 As System.Windows.Forms.Label
    Public WithEvents txtTasaIVA As System.Windows.Forms.TextBox
    Public WithEvents chk_AutPModifDesctos As System.Windows.Forms.CheckBox
    Public WithEvents chk_AutPConsFolVta As System.Windows.Forms.CheckBox
    Public WithEvents chk_IndCuandoProdNoSopDescto As System.Windows.Forms.CheckBox
    Public WithEvents chk_AutAbandCaptIni As System.Windows.Forms.CheckBox
    Public WithEvents chk_AutSupLinCapt As System.Windows.Forms.CheckBox
    Public WithEvents chk_AutCambCodCapt As System.Windows.Forms.CheckBox
    Public WithEvents chk_ConsDesc As System.Windows.Forms.CheckBox
    Public WithEvents chk_PermVtaSinExist As System.Windows.Forms.CheckBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optDiskette As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optElectronica As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray



    'Variables
    Dim mblnSalir As Boolean 'Para Salir sin Preguntar por los Cambios
    Dim mblnSaliryGrabar As Boolean
    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo
    Dim tecla As Integer 'Para Saber que Tecla se Tecleo
    Dim FueraChange As Boolean
    Dim intCodSucursal As Integer
    Dim intCodCaja As Integer
    Dim mbytCapturarCantidadArticulos As Byte
    Dim mbytPermitirVtasSinExistencias As Byte
    Dim mbytConsultasPorDescripcion As Byte
    Dim mbytAutCambiarCodigoCapturado As Byte
    Dim mbytAutSuprimirLineaCapturada As Byte
    Dim mbytAutAbandonarCapturaIniciada As Byte
    Dim mbytIndicarProdNoSoportaDescto As Byte
    Dim mbytAutParaConsFoliosdeVenta As Byte
    Dim mbytAutparaModificarDesctos As Byte
    Dim mbytAumentarDescto As Byte
    Dim mbytDisminuirDescto As Byte
    Dim mbytAplicarIEPS As Byte
    Dim mbytAplicarSucursales As Byte
    Dim mblnElectronica As Boolean
    Dim mblnDiskette As Boolean
    Dim RsAux As ADODB.Recordset

    'Constantes
    Private Const C_TITLEIMAGES As String = "Seleccione la Ruta del Archivo de Imagen ..."
    Friend WithEvents Panel2 As Panel
    Public WithEvents sstConfig As TabControl
    Public WithEvents Frame11 As GroupBox
    Public WithEvents Frame9 As GroupBox
    Public WithEvents _sstConfig_TabPage0 As TabPage
    Friend WithEvents Panel1 As Panel
    Public WithEvents Frame5 As GroupBox
    Public WithEvents Frame7 As GroupBox
    Public WithEvents _sstConfig_TabPage1 As TabPage
    Public WithEvents btnNuevo As Button
    Public WithEvents btnGuardar As Button
    Private Const C_TITLEINVELECTRONICO As String = "Seleccione la Ruta del Archivo de Inventario Electrónico ..."

    Function Guardar() As Boolean
        On Error GoTo Merr
        Dim blnTransaccion As Boolean
        Dim ObjArchivo As Object
        Dim ArchivoTxt As Object
        Dim I As Integer
        '    Dim Sucursal As Integer
        Guardar = False
        If Cambios() = False And mblnNuevo = False Then 'Si No Hay Cambios Limpia Y se Sale
            Me.Close()
            mblnSaliryGrabar = True
            Exit Function
        End If
        'Valida si todos los datos han sido llenados para poder ser guardados
        If ValidaDatos() = False Then
            Exit Function
        End If
        txtTipoCambioDolar.Text = VB6.Format(txtTipoCambioDolar.Text, "##0.00")
        txtEfectivoMaxi.Text = VB6.Format(txtEfectivoMaxi.Text, "##0.00")

        'Si se Seleccionó la Opcion de Todas las Sucursales
        If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = "SELECT CodAlmacen FROM CatAlmacen Where TipoAlmacen = 'P' "
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
        End If
        Cnn.BeginTrans()
        'blnTransaccion = True
        'Cursor = System.Windows.Forms.Cursors.WaitCursor

        ModStoredProcedures.PR_IMEConfiguracionGralPV(Str(intCodSucursal), IIf(chkCapturarCantArticulos.CheckState = 1, 1, 0), cboCentavos.Text, IIf(_optElectronica_0.Checked, "E", "D"), CStr(Val(txtPosicionesDeci.Text)), txtSimbolo.Text, CStr(Val(txtEfectivoMaxi.Text)), "", "", "", CStr(0), IIf(chk_PermVtaSinExist.CheckState = 1, 1, 0), IIf(chk_ConsDesc.CheckState = 1, 1, 0), IIf(chk_AutCambCodCapt.CheckState = 1, 1, 0), IIf(chk_AutSupLinCapt.CheckState = 1, 1, 0), IIf(chk_AutAbandCaptIni.CheckState = 1, 1, 0), IIf(chk_IndCuandoProdNoSopDescto.CheckState = 1, 1, 0), IIf(chk_AutPConsFolVta.CheckState = 1, 1, 0), IIf(chk_AutPModifDesctos.CheckState = 1, 1, 0), CStr(0), txtMensajeFiscal.Text, txtMensajeNormal.Text, txtMensajeVentas.Text, txtMensajeDevoluciones.Text, IIf((chkTransferenciaTicket.CheckState = System.Windows.Forms.CheckState.Checked), "T", "N"), txtTasaIVA.Text, (chkMostrarCodigoViejo.CheckState), C_ELIMINACION, CStr(0))
        Cmd.Execute()

        If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            RsGral.MoveFirst()
            For I = 1 To RsGral.RecordCount
                ModStoredProcedures.PR_IMEConfiguracionGralPV(Str(RsGral.Fields("CodAlmacen").Value), IIf(chkCapturarCantArticulos.CheckState = 1, 1, 0), cboCentavos.Text, IIf(_optElectronica_0.Checked, "E", "D"), CStr(Val(txtPosicionesDeci.Text)), txtSimbolo.Text, CStr(Val(txtEfectivoMaxi.Text)), "", "", "", CStr(0), IIf(chk_PermVtaSinExist.CheckState = 1, 1, 0), IIf(chk_ConsDesc.CheckState = 1, 1, 0), IIf(chk_AutCambCodCapt.CheckState = 1, 1, 0), IIf(chk_AutSupLinCapt.CheckState = 1, 1, 0), IIf(chk_AutAbandCaptIni.CheckState = 1, 1, 0), IIf(chk_IndCuandoProdNoSopDescto.CheckState = 1, 1, 0), IIf(chk_AutPConsFolVta.CheckState = 1, 1, 0), IIf(chk_AutPModifDesctos.CheckState = 1, 1, 0), CStr(0), txtMensajeFiscal.Text, txtMensajeNormal.Text, txtMensajeVentas.Text, txtMensajeDevoluciones.Text, IIf((chkTransferenciaTicket.CheckState = System.Windows.Forms.CheckState.Checked), "T", "N"), txtTasaIVA.Text, (chkMostrarCodigoViejo.CheckState), C_ELIMINACION, CStr(0))
                Cmd.Execute()
                RsGral.MoveNext()
            Next
        End If
        '    If mblnNuevo Then
        ModStoredProcedures.PR_IMEConfiguracionGralPV(Str(intCodSucursal), IIf(chkCapturarCantArticulos.CheckState = 1, 1, 0), cboCentavos.Text, IIf(_optElectronica_0.Checked, "E", "D"), CStr(Val(txtPosicionesDeci.Text)), txtSimbolo.Text, CStr(Val(txtEfectivoMaxi.Text)), "", "", "", CStr(0), IIf(chk_PermVtaSinExist.CheckState = 1, 1, 0), IIf(chk_ConsDesc.CheckState = 1, 1, 0), IIf(chk_AutCambCodCapt.CheckState = 1, 1, 0), IIf(chk_AutSupLinCapt.CheckState = 1, 1, 0), IIf(chk_AutAbandCaptIni.CheckState = 1, 1, 0), IIf(chk_IndCuandoProdNoSopDescto.CheckState = 1, 1, 0), IIf(chk_AutPConsFolVta.CheckState = 1, 1, 0), IIf(chk_AutPModifDesctos.CheckState = 1, 1, 0), CStr(0), txtMensajeFiscal.Text, txtMensajeNormal.Text, txtMensajeVentas.Text, txtMensajeDevoluciones.Text, IIf((chkTransferenciaTicket.CheckState = System.Windows.Forms.CheckState.Checked), "T", "N"), txtTasaIVA.Text, (chkMostrarCodigoViejo.CheckState), C_INSERCION, CStr(0))
        Cmd.Execute()

        'Si se requiere aplicar a todas las sucursales
        If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            RsGral.MoveFirst()
            For I = 1 To RsGral.RecordCount
                If RsGral.Fields("CodAlmacen").Value <> intCodSucursal Then
                    ModStoredProcedures.PR_IMEConfiguracionGralPV(Str(RsGral.Fields("CodAlmacen").Value), IIf(chkCapturarCantArticulos.CheckState = 1, 1, 0), cboCentavos.Text, IIf(_optElectronica_0.Checked, "E", "D"), CStr(Val(txtPosicionesDeci.Text)), txtSimbolo.Text, CStr(Val(txtEfectivoMaxi.Text)), "", "", "", CStr(0), IIf(chk_PermVtaSinExist.CheckState = 1, 1, 0), IIf(chk_ConsDesc.CheckState = 1, 1, 0), IIf(chk_AutCambCodCapt.CheckState = 1, 1, 0), IIf(chk_AutSupLinCapt.CheckState = 1, 1, 0), IIf(chk_AutAbandCaptIni.CheckState = 1, 1, 0), IIf(chk_IndCuandoProdNoSopDescto.CheckState = 1, 1, 0), IIf(chk_AutPConsFolVta.CheckState = 1, 1, 0), IIf(chk_AutPModifDesctos.CheckState = 1, 1, 0), CStr(0), txtMensajeFiscal.Text, txtMensajeNormal.Text, txtMensajeVentas.Text, txtMensajeDevoluciones.Text, IIf((chkTransferenciaTicket.CheckState = System.Windows.Forms.CheckState.Checked), "T", "N"), txtTasaIVA.Text, (chkMostrarCodigoViejo.CheckState), C_INSERCION, CStr(0))
                    Cmd.Execute()
                End If
                RsGral.MoveNext()
            Next
        End If

        Cnn.CommitTrans()
        Cursor = System.Windows.Forms.Cursors.Default
        blnTransaccion = False
        If mblnNuevo Then
            MsgBox("La configuración del Punto de Venta  para la" & vbNewLine & "sucursal " & CStr(intCodSucursal) & " ha sido grabada correctamente ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
        End If
        'Carga la configuración despues de haberla guardado
        CargarDatosConfiguracionPV()
        Guardar = True
        mblnSaliryGrabar = True
        Nuevo()
        Limpiar()
        Exit Function
Merr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Trim(dbcSucursales.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Sucursal", MsgBoxStyle.Information, gstrNombCortoEmpresa)
            sstConfig.SelectedIndex = 0
            dbcSucursales.Focus()
            Exit Function
        End If
        If CDbl(Numerico(Trim(txtTasaIVA.Text))) > 100 Or Numerico(Trim(txtTasaIVA.Text)) < CStr(0.01) Then
            MsgBox("La tasa de iva debe ser un valor comprendido entre 0.01 y 100." & vbNewLine & "Verifique Por favor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtTasaIVA.Focus()
            Exit Function
        End If

        ValidaDatos = True
    End Function

    Function Cambios() As Boolean
        Cambios = True
        If cboCentavos.Text <> cboCentavos.Tag Then Exit Function
        If chkCapturarCantArticulos.CheckState <> mbytCapturarCantidadArticulos Then Exit Function
        If _optElectronica_0.Checked <> mblnElectronica Then Exit Function
        If _optDiskette_1.Checked <> mblnDiskette Then Exit Function
        If txtTipoCambioDolar.Text <> txtTipoCambioDolar.Tag Then Exit Function
        If txtPosicionesDeci.Text <> txtPosicionesDeci.Tag Then Exit Function
        If txtSimbolo.Text <> txtSimbolo.Tag Then Exit Function
        If txtEfectivoMaxi.Text <> txtEfectivoMaxi.Tag Then Exit Function
        If chk_PermVtaSinExist.CheckState <> mbytPermitirVtasSinExistencias Then Exit Function
        If chk_ConsDesc.CheckState <> mbytConsultasPorDescripcion Then Exit Function
        If chk_AutCambCodCapt.CheckState <> mbytAutCambiarCodigoCapturado Then Exit Function
        If chk_AutSupLinCapt.CheckState <> mbytAutSuprimirLineaCapturada Then Exit Function
        If chk_AutAbandCaptIni.CheckState <> mbytAutAbandonarCapturaIniciada Then Exit Function
        If chk_IndCuandoProdNoSopDescto.CheckState <> mbytIndicarProdNoSoportaDescto Then Exit Function
        If chk_AutPConsFolVta.CheckState <> mbytAutParaConsFoliosdeVenta Then Exit Function
        If chk_AutPModifDesctos.CheckState <> mbytAutparaModificarDesctos Then Exit Function
        If chkAplicarSucursales.CheckState <> mbytAplicarSucursales Then Exit Function
        '''If txtUtilidadOperacio <> txtUtilidadOperacio.Tag Then Exit Function
        If txtMensajeFiscal.Text <> txtMensajeFiscal.Tag Then Exit Function
        If txtMensajeNormal.Text <> txtMensajeNormal.Tag Then Exit Function
        If txtMensajeVentas.Text <> txtMensajeVentas.Tag Then Exit Function
        If txtMensajeDevoluciones.Text <> txtMensajeDevoluciones.Tag Then Exit Function
        If chkTransferenciaTicket.CheckState <> CDbl(chkTransferenciaTicket.Tag) Then Exit Function
        If Numerico(Trim(txtTasaIVA.Text)) <> Numerico(Trim(txtTasaIVA.Tag)) Then Exit Function
        If CDbl(chkMostrarCodigoViejo.Tag) <> chkMostrarCodigoViejo.CheckState Then Exit Function
        Cambios = False
    End Function

    Sub LlenaDatos()
        On Error GoTo Merr
        FueraChange = True

        gStrSql = "SELECT     dbo.CatAlmacen.DescAlmacen AS DescAlmacen, dbo.ConfiguracionGralPV.*, dbo.ConfiguracionGeneral.TipoCambio AS TipoCambio " & "FROM         dbo.ConfiguracionGralPV INNER JOIN " & "dbo.CatAlmacen ON dbo.ConfiguracionGralPV.CodAlmacen = dbo.CatAlmacen.CodAlmacen CROSS JOIN " & "dbo.ConfiguracionGeneral " & "WHERE     dbo.ConfiguracionGralPV.CodAlmacen = " & intCodSucursal
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            mblnNuevo = False
            txtTipoCambioDolar.Text = VB6.Format(RsGral.Fields("TipoCambio").Value, "###,###.00")
            txtTipoCambioDolar.Tag = VB6.Format(RsGral.Fields("TipoCambio").Value, "###,###.00")
            If RsGral.Fields("CapturarCantArts").Value = True Then
                chkCapturarCantArticulos.CheckState = System.Windows.Forms.CheckState.Checked
                mbytCapturarCantidadArticulos = 1
            Else
                chkCapturarCantArticulos.CheckState = System.Windows.Forms.CheckState.Unchecked
                mbytCapturarCantidadArticulos = 0
            End If
            cboCentavos.Text = VB6.Format(RsGral.Fields("Redondeo").Value, "0.00")
            cboCentavos.Tag = VB6.Format(RsGral.Fields("Redondeo").Value, "0.00")
            If RsGral.Fields("Transferencia").Value = "E" Then
                _optElectronica_0.Checked = True
                mblnElectronica = True
                mblnDiskette = False
            ElseIf RsGral.Fields("Transferencia").Value = "D" Then
                _optDiskette_1.Checked = True
                mblnDiskette = True
                mblnElectronica = False
            End If
            txtPosicionesDeci.Text = RsGral.Fields("PosicionDecimal").Value
            txtPosicionesDeci.Tag = RsGral.Fields("PosicionDecimal").Value
            txtSimbolo.Text = Trim(RsGral.Fields("SimboloMonedaNac").Value)
            txtSimbolo.Tag = Trim(RsGral.Fields("SimboloMonedaNac").Value)
            txtEfectivoMaxi.Text = VB6.Format(RsGral.Fields("EfectivoMaximo").Value, "###,##0.00")
            txtEfectivoMaxi.Tag = VB6.Format(RsGral.Fields("EfectivoMaximo").Value, "###,##0.00")
            If RsGral.Fields("PermitirVtaSinExistencia").Value = True Then
                chk_PermVtaSinExist.CheckState = System.Windows.Forms.CheckState.Checked
                mbytPermitirVtasSinExistencias = 1
            Else
                chk_PermVtaSinExist.CheckState = System.Windows.Forms.CheckState.Unchecked
                mbytPermitirVtasSinExistencias = 0
            End If
            If RsGral.Fields("ConsultarXDescrip").Value = True Then
                chk_ConsDesc.CheckState = System.Windows.Forms.CheckState.Checked
                mbytConsultasPorDescripcion = 1
            Else
                chk_ConsDesc.CheckState = System.Windows.Forms.CheckState.Unchecked
                mbytConsultasPorDescripcion = 0
            End If
            If RsGral.Fields("AutorizCambiarCodigoCapt").Value = True Then
                chk_AutCambCodCapt.CheckState = System.Windows.Forms.CheckState.Checked
                mbytAutCambiarCodigoCapturado = 1
            Else
                chk_AutCambCodCapt.CheckState = System.Windows.Forms.CheckState.Unchecked
                mbytAutCambiarCodigoCapturado = 0
            End If
            If RsGral.Fields("AutorizCambiarLineaCapt").Value = True Then
                chk_AutSupLinCapt.CheckState = System.Windows.Forms.CheckState.Checked
                mbytAutSuprimirLineaCapturada = 1
            Else
                chk_AutSupLinCapt.CheckState = System.Windows.Forms.CheckState.Unchecked
                mbytAutSuprimirLineaCapturada = 0
            End If
            If RsGral.Fields("AutorizAbandonarCaptIni").Value = True Then
                chk_AutAbandCaptIni.CheckState = System.Windows.Forms.CheckState.Checked
                mbytAutAbandonarCapturaIniciada = 1
            Else
                chk_AutAbandCaptIni.CheckState = System.Windows.Forms.CheckState.Unchecked
                mbytAutAbandonarCapturaIniciada = 0
            End If
            If RsGral.Fields("IndicarSiProdNoSoportaDescto").Value = True Then
                chk_IndCuandoProdNoSopDescto.CheckState = System.Windows.Forms.CheckState.Checked
                mbytIndicarProdNoSoportaDescto = 1
            Else
                chk_IndCuandoProdNoSopDescto.CheckState = System.Windows.Forms.CheckState.Unchecked
                mbytIndicarProdNoSoportaDescto = 0
            End If
            If RsGral.Fields("AutorizConsultarFoliosVta").Value = True Then
                chk_AutPConsFolVta.CheckState = System.Windows.Forms.CheckState.Checked
                mbytAutParaConsFoliosdeVenta = 1
            Else
                chk_AutPConsFolVta.CheckState = System.Windows.Forms.CheckState.Unchecked
                mbytAutParaConsFoliosdeVenta = 0
            End If
            If RsGral.Fields("AutorizModificarDesctos").Value = True Then
                chk_AutPModifDesctos.CheckState = System.Windows.Forms.CheckState.Checked
                mbytAutparaModificarDesctos = 1
            Else
                chk_AutPModifDesctos.CheckState = System.Windows.Forms.CheckState.Unchecked
                mbytAutparaModificarDesctos = 0
            End If
            If Trim(RsGral.Fields("ImpresionTransferencias").Value) = "T" Then
                chkTransferenciaTicket.CheckState = System.Windows.Forms.CheckState.Checked
                chkTransferenciaTicket.Tag = System.Windows.Forms.CheckState.Checked
            Else
                chkTransferenciaTicket.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkTransferenciaTicket.Tag = System.Windows.Forms.CheckState.Unchecked
            End If
            txtMensajeFiscal.Text = Trim(RsGral.Fields("MsgFiscal").Value)
            txtMensajeFiscal.Tag = Trim(RsGral.Fields("MsgFiscal").Value)
            txtMensajeNormal.Text = Trim(RsGral.Fields("MsgNormal").Value)
            txtMensajeNormal.Tag = Trim(RsGral.Fields("MsgNormal").Value)
            txtMensajeVentas.Text = Trim(RsGral.Fields("MsgCredito").Value)
            txtMensajeVentas.Tag = Trim(RsGral.Fields("MsgCredito").Value)
            txtMensajeDevoluciones.Text = Trim(RsGral.Fields("MsgDevoluciones").Value)
            txtMensajeDevoluciones.Tag = Trim(RsGral.Fields("MsgDevoluciones").Value)
            txtTasaIVA.Text = VB6.Format(Numerico(RsGral.Fields("TasaIVA").Value), "0.00")
            txtTasaIVA.Tag = VB6.Format(Numerico(RsGral.Fields("TasaIVA").Value), "0.00")
            chkMostrarCodigoViejo.CheckState = IIf(RsGral.Fields("CodigoViejo").Value, 1, 0)
            chkMostrarCodigoViejo.Tag = chkMostrarCodigoViejo.CheckState
            FueraChange = False
        Else
            mblnNuevo = True
        End If
        Exit Sub
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
        mblnSaliryGrabar = False
        mblnNuevo = True
        intCodSucursal = 0
        mbytCapturarCantidadArticulos = 0
        mbytPermitirVtasSinExistencias = 0
        mbytConsultasPorDescripcion = 0
        mbytAutCambiarCodigoCapturado = 0
        mbytAutSuprimirLineaCapturada = 0
        mbytAutAbandonarCapturaIniciada = 0
        mbytIndicarProdNoSoportaDescto = 0
        mbytAutParaConsFoliosdeVenta = 0
        mbytAutparaModificarDesctos = 0
        mbytAumentarDescto = 0
        mbytDisminuirDescto = 0
        mbytAplicarIEPS = 0
        mbytAplicarSucursales = 0
        mblnElectronica = True
        mblnDiskette = False
    End Sub

    Sub Limpiar()
        On Error Resume Next
        'Validar si hubo cambios que pregunte si decea guardar
        If Cambios() = True And mblnNuevo = False Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        'Cancel = 1
                    End If
                Case MsgBoxResult.No 'No hace nada y permite el cierre del formulario
                Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
                    'Cancel = 1
            End Select
        End If
        dbcSucursales.Text = ""
        Nuevo()
        sstConfig.SelectedIndex = 0
        dbcSucursales.Focus()
    End Sub

    Sub Nuevo()
        On Error GoTo Merr
        cboCentavos.SelectedIndex = 0
        cboCentavos.Tag = cboCentavos.Text
        chkCapturarCantArticulos.CheckState = System.Windows.Forms.CheckState.Unchecked
        _optElectronica_0.Checked = True
        _optDiskette_1.Checked = False
        txtTipoCambioDolar.Text = ""
        txtTipoCambioDolar.Tag = ""
        txtPosicionesDeci.Text = ""
        txtPosicionesDeci.Tag = ""
        txtSimbolo.Text = ""
        txtSimbolo.Tag = ""
        txtEfectivoMaxi.Text = ""
        txtEfectivoMaxi.Tag = ""
        chk_PermVtaSinExist.CheckState = System.Windows.Forms.CheckState.Unchecked
        chk_ConsDesc.CheckState = System.Windows.Forms.CheckState.Unchecked
        chk_AutCambCodCapt.CheckState = System.Windows.Forms.CheckState.Unchecked
        chk_AutSupLinCapt.CheckState = System.Windows.Forms.CheckState.Unchecked
        chk_AutAbandCaptIni.CheckState = System.Windows.Forms.CheckState.Unchecked
        chk_IndCuandoProdNoSopDescto.CheckState = System.Windows.Forms.CheckState.Unchecked
        chk_AutPConsFolVta.CheckState = System.Windows.Forms.CheckState.Unchecked
        chk_AutPModifDesctos.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtMensajeFiscal.Text = ""
        txtMensajeFiscal.Tag = ""
        txtMensajeNormal.Text = ""
        txtMensajeNormal.Tag = ""
        txtMensajeVentas.Text = ""
        txtMensajeVentas.Tag = ""
        txtMensajeDevoluciones.Text = ""
        txtMensajeDevoluciones.Tag = ""
        chkTransferenciaTicket.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkTransferenciaTicket.Tag = System.Windows.Forms.CheckState.Unchecked
        chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkMostrarCodigoViejo.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkMostrarCodigoViejo.Tag = System.Windows.Forms.CheckState.Unchecked
        InicializaVariables()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub btnArchInvElect_GotFocus()
        Pon_Tool()
    End Sub

    Private Sub cboCentavos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCentavos.Enter
        Pon_Tool()
    End Sub

    Private Sub cboCentavos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cboCentavos.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape And (Not dbcSucursales.Enabled) Then
            sstConfig.Focus()
        End If
    End Sub

    Private Sub cboEspacios_GotFocus()
        sstConfig.TabIndex = 19
        Pon_Tool()
    End Sub

    Private Sub chk_AutAbandCaptIni_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk_AutAbandCaptIni.Enter
        Pon_Tool()
    End Sub

    Private Sub chk_AutCambCodCapt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk_AutCambCodCapt.Enter
        Pon_Tool()
    End Sub

    Private Sub chk_AutPConsFolVta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk_AutPConsFolVta.Enter
        Pon_Tool()
    End Sub

    Private Sub chk_AutPModifDesctos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk_AutPModifDesctos.Enter
        Pon_Tool()
    End Sub

    Private Sub chk_AutSupLinCapt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk_AutSupLinCapt.Enter
        Pon_Tool()
    End Sub

    Private Sub chk_ConsDesc_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk_ConsDesc.Enter
        Pon_Tool()
    End Sub

    Private Sub chk_IndCuandoProdNoSopDescto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk_IndCuandoProdNoSopDescto.Enter
        Pon_Tool()
    End Sub

    Private Sub chk_PermVtaSinExist_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chk_PermVtaSinExist.Enter
        Pon_Tool()
    End Sub

    Private Sub chk_PermVtaSinExist_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chk_PermVtaSinExist.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            sstConfig.Focus()
        End If
    End Sub

    Private Sub chkCapturarCantArticulos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCapturarCantArticulos.Enter
        Pon_Tool()
    End Sub

    Private Sub chkModificaDescuentos_GotFocus(ByRef Index As Integer)
        Select Case Index
            Case 0
                Pon_Tool()
            Case 1
                Pon_Tool()
        End Select
    End Sub


    Private Sub dbcSucursales_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If
        Nuevo()
        gStrSql = "SELECT CodAlmacen,Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        'If dbcSucursales.SelectedItem <> "" Then
        dbcSucursales_Leave(dbcSucursales, New System.EventArgs())
        'End If
        mblnNuevo = True
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursales.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen  FROM CatAlmacen   Where TipoAlmacen ='P'  ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub


    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            eventSender.KeyCode = 0
        End If
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        intCodSucursal = 0
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%'  and TipoAlmacen ='P'  ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        LlenaDatos()
    End Sub

    Private Sub frmPVConfigPuntoVenta_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmPVConfigPuntoVenta_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmPVConfigPuntoVenta_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmPVConfigPuntoVenta_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmPVConfigPuntoVenta_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        sstConfig.SelectedIndex = 0
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        'InicializaVariables()
        'Nuevo()
        CentrarForma(Me)
    End Sub

    Private Sub frmPVConfigPuntoVenta_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir And Not mblnSaliryGrabar Then
        '    If Cambios() = True And mblnNuevo = False Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes 'Guardar el registro
        '                If Guardar() = False Then
        '                    Cancel = 1
        '                End If
        '            Case MsgBoxResult.No 'No hace nada y permite el cierre del formulario
        '            Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
        '                Cancel = 1
        '        End Select
        '    End If
        'Else
        '    If mblnSaliryGrabar Then
        '        Cancel = 0
        '        Exit Sub
        '    End If
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0
        '        Case MsgBoxResult.No
        '            mblnSalir = False
        '            Cancel = 1
        '            sstConfig.Focus()
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmPVConfigPuntoVenta_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
    End Sub

    Private Sub optDiskette_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim Index As Integer = optDiskette.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub optElectronica_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim Index As Integer = optElectronica.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub sstConfig_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sstConfig.Enter
        Pon_Tool()
        If sstConfig.SelectedIndex = 0 Then
            sstConfig.TabIndex = 0
        ElseIf sstConfig.SelectedIndex = 1 Then
            sstConfig.TabIndex = 19
        End If
    End Sub

    Private Sub sstConfig_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles sstConfig.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return And sstConfig.SelectedIndex = 0 Then
            If dbcSucursales.Enabled Then
                dbcSucursales.Focus()
            End If
        ElseIf KeyCode = System.Windows.Forms.Keys.Return And sstConfig.SelectedIndex = 1 Then
            chk_PermVtaSinExist.Focus()
        End If
        If KeyCode = System.Windows.Forms.Keys.Escape And sstConfig.SelectedIndex = 0 Then
            mblnSalir = True
            Me.Close()
        End If
    End Sub

    Private Sub sstConfig_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles sstConfig.KeyUp
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Left Then
            If sstConfig.SelectedIndex = 0 Then
                sstConfig.TabIndex = 0
            ElseIf sstConfig.SelectedIndex = 1 Then
                sstConfig.TabIndex = 19
            End If
        End If
        If KeyCode = System.Windows.Forms.Keys.Right Then
            If sstConfig.SelectedIndex = 0 Then
                sstConfig.TabIndex = 0
            ElseIf sstConfig.SelectedIndex = 1 Then
                sstConfig.TabIndex = 19
            End If
        End If
    End Sub

    Private Sub txtArchivoInvElect_GotFocus()
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtEfectivoMaxi_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEfectivoMaxi.Enter
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtEfectivoMaxi_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEfectivoMaxi.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtEfectivoMaxi.Text, KeyAscii, 9, 2, (txtEfectivoMaxi.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtEfectivoMaxi_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEfectivoMaxi.Leave
        txtEfectivoMaxi.Text = VB6.Format(txtEfectivoMaxi.Text, "###,##0.00")
    End Sub

    Private Sub txtMensajeDevoluciones_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensajeDevoluciones.Enter
        sstConfig.TabIndex = 37
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtMensajeFiscal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensajeFiscal.Enter
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtMensajeNormal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensajeNormal.Enter
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtMensajeVentas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensajeVentas.Enter
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtPosicionesDeci_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPosicionesDeci.Enter
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtPosicionesDeci_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPosicionesDeci.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtPosicionesDeci.Text, KeyAscii, 1, 0, (txtPosicionesDeci.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtRutaInvElect_GotFocus()
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtSeparador_GotFocus()
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtSimbolo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSimbolo.Enter
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtSimbolo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSimbolo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoLetras(KeyAscii, "$!""#%&/()=?¡'¿/*-+@")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTasaIVA_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTasaIVA.Enter
        SelTextoTxt(txtTasaIVA)
        Pon_Tool()
    End Sub

    Private Sub txtTasaIVA_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTasaIVA.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad(txtTasaIVA.Text, KeyAscii, 3, 2, (txtTasaIVA.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTasaIVA_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTasaIVA.Leave
        txtTasaIVA.Text = VB6.Format(Numerico(txtTasaIVA.Text), "0.00")
        If CDbl(Numerico(Trim(txtTasaIVA.Text))) > 100 Then
            MsgBox("La Tasa de Iva debe ser menor o igual a 100." & vbNewLine & "Verifique Porfavor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Me.txtTasaIVA.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub txtTipoCambioDolar_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioDolar.Enter
        Pon_Tool()
        SelTxt()
    End Sub

    Private Sub txtTipoCambioDolar_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambioDolar.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtTipoCambioDolar.Text, KeyAscii, 4, 2, (txtTipoCambioDolar.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambioDolar_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioDolar.Leave
        txtTipoCambioDolar.Text = VB6.Format(txtTipoCambioDolar.Text, "##0.00")
    End Sub

    Private Sub txtUtilidadOperacio_GotFocus()
        Pon_Tool()
        SelTxt()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkTransferenciaTicket = New System.Windows.Forms.CheckBox()
        Me.cboCentavos = New System.Windows.Forms.ComboBox()
        Me._optElectronica_0 = New System.Windows.Forms.RadioButton()
        Me._optDiskette_1 = New System.Windows.Forms.RadioButton()
        Me.chkCapturarCantArticulos = New System.Windows.Forms.CheckBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.txtEfectivoMaxi = New System.Windows.Forms.TextBox()
        Me.txtSimbolo = New System.Windows.Forms.TextBox()
        Me.txtTipoCambioDolar = New System.Windows.Forms.TextBox()
        Me.txtPosicionesDeci = New System.Windows.Forms.TextBox()
        Me.txtMensajeFiscal = New System.Windows.Forms.TextBox()
        Me.txtMensajeDevoluciones = New System.Windows.Forms.TextBox()
        Me.txtMensajeVentas = New System.Windows.Forms.TextBox()
        Me.txtMensajeNormal = New System.Windows.Forms.TextBox()
        Me.chkAplicarSucursales = New System.Windows.Forms.CheckBox()
        Me.chkMostrarCodigoViejo = New System.Windows.Forms.CheckBox()
        Me._Label1_9 = New System.Windows.Forms.Label()
        Me._Label1_18 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_8 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Frame11 = New System.Windows.Forms.GroupBox()
        Me._Label1_25 = New System.Windows.Forms.Label()
        Me._Label1_13 = New System.Windows.Forms.Label()
        Me._Label1_12 = New System.Windows.Forms.Label()
        Me._Label1_11 = New System.Windows.Forms.Label()
        Me.Frame9 = New System.Windows.Forms.GroupBox()
        Me.txtTasaIVA = New System.Windows.Forms.TextBox()
        Me.chk_AutPModifDesctos = New System.Windows.Forms.CheckBox()
        Me.chk_AutPConsFolVta = New System.Windows.Forms.CheckBox()
        Me.chk_IndCuandoProdNoSopDescto = New System.Windows.Forms.CheckBox()
        Me.chk_AutAbandCaptIni = New System.Windows.Forms.CheckBox()
        Me.chk_AutSupLinCapt = New System.Windows.Forms.CheckBox()
        Me.chk_AutCambCodCapt = New System.Windows.Forms.CheckBox()
        Me.chk_ConsDesc = New System.Windows.Forms.CheckBox()
        Me.chk_PermVtaSinExist = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Frame10 = New System.Windows.Forms.GroupBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me.sstConfig = New System.Windows.Forms.TabControl()
        Me._sstConfig_TabPage0 = New System.Windows.Forms.TabPage()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.Frame7 = New System.Windows.Forms.GroupBox()
        Me._sstConfig_TabPage1 = New System.Windows.Forms.TabPage()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Panel2.SuspendLayout()
        Me.Frame11.SuspendLayout()
        Me.Frame9.SuspendLayout()
        Me.Frame10.SuspendLayout()
        Me.sstConfig.SuspendLayout()
        Me._sstConfig_TabPage0.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frame7.SuspendLayout()
        Me._sstConfig_TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkTransferenciaTicket
        '
        Me.chkTransferenciaTicket.BackColor = System.Drawing.SystemColors.Control
        Me.chkTransferenciaTicket.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTransferenciaTicket.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTransferenciaTicket.Location = New System.Drawing.Point(22, 253)
        Me.chkTransferenciaTicket.Name = "chkTransferenciaTicket"
        Me.chkTransferenciaTicket.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTransferenciaTicket.Size = New System.Drawing.Size(238, 21)
        Me.chkTransferenciaTicket.TabIndex = 40
        Me.chkTransferenciaTicket.Text = "Imprimir transfererencias en ticket"
        Me.ToolTip1.SetToolTip(Me.chkTransferenciaTicket, "Capturar Cantidad de Articulos.")
        Me.chkTransferenciaTicket.UseVisualStyleBackColor = False
        '
        'cboCentavos
        '
        Me.cboCentavos.BackColor = System.Drawing.SystemColors.Window
        Me.cboCentavos.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCentavos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCentavos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCentavos.Items.AddRange(New Object() {"0.00", "0.10", "0.20", "0.30", "0.40", "0.50", "0.60", "0.70", "0.80", "0.90", "1.00", "2.00", "3.00", "4.00", "5.00"})
        Me.cboCentavos.Location = New System.Drawing.Point(218, 72)
        Me.cboCentavos.Name = "cboCentavos"
        Me.cboCentavos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCentavos.Size = New System.Drawing.Size(53, 21)
        Me.cboCentavos.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.cboCentavos, "Cantidad de Redondeo para  los Totales.")
        '
        '_optElectronica_0
        '
        Me._optElectronica_0.BackColor = System.Drawing.SystemColors.Control
        Me._optElectronica_0.Checked = True
        Me._optElectronica_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optElectronica_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optElectronica_0.Location = New System.Drawing.Point(80, 16)
        Me._optElectronica_0.Name = "_optElectronica_0"
        Me._optElectronica_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optElectronica_0.Size = New System.Drawing.Size(89, 17)
        Me._optElectronica_0.TabIndex = 7
        Me._optElectronica_0.TabStop = True
        Me._optElectronica_0.Text = "Electrónica"
        Me.ToolTip1.SetToolTip(Me._optElectronica_0, "Transferencia Electrónica.")
        Me._optElectronica_0.UseVisualStyleBackColor = False
        '
        '_optDiskette_1
        '
        Me._optDiskette_1.BackColor = System.Drawing.SystemColors.Control
        Me._optDiskette_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optDiskette_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optDiskette_1.Location = New System.Drawing.Point(192, 16)
        Me._optDiskette_1.Name = "_optDiskette_1"
        Me._optDiskette_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optDiskette_1.Size = New System.Drawing.Size(105, 17)
        Me._optDiskette_1.TabIndex = 8
        Me._optDiskette_1.TabStop = True
        Me._optDiskette_1.Text = "Diskette"
        Me.ToolTip1.SetToolTip(Me._optDiskette_1, "Transferencia por Diskette.")
        Me._optDiskette_1.UseVisualStyleBackColor = False
        '
        'chkCapturarCantArticulos
        '
        Me.chkCapturarCantArticulos.BackColor = System.Drawing.SystemColors.Control
        Me.chkCapturarCantArticulos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCapturarCantArticulos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCapturarCantArticulos.Location = New System.Drawing.Point(38, 104)
        Me.chkCapturarCantArticulos.Name = "chkCapturarCantArticulos"
        Me.chkCapturarCantArticulos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCapturarCantArticulos.Size = New System.Drawing.Size(222, 21)
        Me.chkCapturarCantArticulos.TabIndex = 5
        Me.chkCapturarCantArticulos.Text = "Capturar Cantidad de Articulos"
        Me.ToolTip1.SetToolTip(Me.chkCapturarCantArticulos, "Capturar Cantidad de Articulos.")
        Me.chkCapturarCantArticulos.UseVisualStyleBackColor = False
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.Color.Black
        Me._Label1_1.Location = New System.Drawing.Point(35, 35)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(60, 17)
        Me._Label1_1.TabIndex = 1
        Me._Label1_1.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_1, "Nombre de la Farmacia Actual")
        '
        'txtEfectivoMaxi
        '
        Me.txtEfectivoMaxi.AcceptsReturn = True
        Me.txtEfectivoMaxi.BackColor = System.Drawing.SystemColors.Window
        Me.txtEfectivoMaxi.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEfectivoMaxi.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtEfectivoMaxi.Location = New System.Drawing.Point(192, 146)
        Me.txtEfectivoMaxi.MaxLength = 0
        Me.txtEfectivoMaxi.Name = "txtEfectivoMaxi"
        Me.txtEfectivoMaxi.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEfectivoMaxi.Size = New System.Drawing.Size(73, 20)
        Me.txtEfectivoMaxi.TabIndex = 16
        Me.txtEfectivoMaxi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtEfectivoMaxi, "Efectivo Máximo que Debe de Tener la Caja.")
        '
        'txtSimbolo
        '
        Me.txtSimbolo.AcceptsReturn = True
        Me.txtSimbolo.BackColor = System.Drawing.SystemColors.Window
        Me.txtSimbolo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSimbolo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtSimbolo.Location = New System.Drawing.Point(192, 114)
        Me.txtSimbolo.MaxLength = 3
        Me.txtSimbolo.Name = "txtSimbolo"
        Me.txtSimbolo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSimbolo.Size = New System.Drawing.Size(73, 20)
        Me.txtSimbolo.TabIndex = 14
        Me.txtSimbolo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtSimbolo, "Símbolo para los Totales")
        '
        'txtTipoCambioDolar
        '
        Me.txtTipoCambioDolar.AcceptsReturn = True
        Me.txtTipoCambioDolar.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipoCambioDolar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambioDolar.Enabled = False
        Me.txtTipoCambioDolar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtTipoCambioDolar.Location = New System.Drawing.Point(192, 51)
        Me.txtTipoCambioDolar.MaxLength = 7
        Me.txtTipoCambioDolar.Name = "txtTipoCambioDolar"
        Me.txtTipoCambioDolar.ReadOnly = True
        Me.txtTipoCambioDolar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambioDolar.Size = New System.Drawing.Size(73, 20)
        Me.txtTipoCambioDolar.TabIndex = 10
        Me.txtTipoCambioDolar.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambioDolar, "Valor del Dólar.")
        '
        'txtPosicionesDeci
        '
        Me.txtPosicionesDeci.AcceptsReturn = True
        Me.txtPosicionesDeci.BackColor = System.Drawing.SystemColors.Window
        Me.txtPosicionesDeci.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPosicionesDeci.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtPosicionesDeci.Location = New System.Drawing.Point(192, 82)
        Me.txtPosicionesDeci.MaxLength = 1
        Me.txtPosicionesDeci.Name = "txtPosicionesDeci"
        Me.txtPosicionesDeci.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPosicionesDeci.Size = New System.Drawing.Size(73, 20)
        Me.txtPosicionesDeci.TabIndex = 12
        Me.txtPosicionesDeci.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtPosicionesDeci, "Posiciones Decimales para las Cantidades.")
        '
        'txtMensajeFiscal
        '
        Me.txtMensajeFiscal.AcceptsReturn = True
        Me.txtMensajeFiscal.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensajeFiscal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensajeFiscal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtMensajeFiscal.Location = New System.Drawing.Point(167, 14)
        Me.txtMensajeFiscal.MaxLength = 50
        Me.txtMensajeFiscal.Name = "txtMensajeFiscal"
        Me.txtMensajeFiscal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensajeFiscal.Size = New System.Drawing.Size(490, 20)
        Me.txtMensajeFiscal.TabIndex = 26
        Me.ToolTip1.SetToolTip(Me.txtMensajeFiscal, "Mensaje Fiscal para los Tickets")
        '
        'txtMensajeDevoluciones
        '
        Me.txtMensajeDevoluciones.AcceptsReturn = True
        Me.txtMensajeDevoluciones.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensajeDevoluciones.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensajeDevoluciones.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtMensajeDevoluciones.Location = New System.Drawing.Point(167, 91)
        Me.txtMensajeDevoluciones.MaxLength = 50
        Me.txtMensajeDevoluciones.Name = "txtMensajeDevoluciones"
        Me.txtMensajeDevoluciones.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensajeDevoluciones.Size = New System.Drawing.Size(490, 20)
        Me.txtMensajeDevoluciones.TabIndex = 29
        Me.ToolTip1.SetToolTip(Me.txtMensajeDevoluciones, "Mensaje para las impresiones de devoluciones")
        '
        'txtMensajeVentas
        '
        Me.txtMensajeVentas.AcceptsReturn = True
        Me.txtMensajeVentas.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensajeVentas.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensajeVentas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtMensajeVentas.Location = New System.Drawing.Point(167, 66)
        Me.txtMensajeVentas.MaxLength = 50
        Me.txtMensajeVentas.Name = "txtMensajeVentas"
        Me.txtMensajeVentas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensajeVentas.Size = New System.Drawing.Size(490, 20)
        Me.txtMensajeVentas.TabIndex = 28
        Me.ToolTip1.SetToolTip(Me.txtMensajeVentas, "Mensaje para las impresiones de las ventas a crédito.")
        '
        'txtMensajeNormal
        '
        Me.txtMensajeNormal.AcceptsReturn = True
        Me.txtMensajeNormal.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensajeNormal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensajeNormal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtMensajeNormal.Location = New System.Drawing.Point(167, 40)
        Me.txtMensajeNormal.MaxLength = 50
        Me.txtMensajeNormal.Name = "txtMensajeNormal"
        Me.txtMensajeNormal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensajeNormal.Size = New System.Drawing.Size(490, 20)
        Me.txtMensajeNormal.TabIndex = 27
        Me.ToolTip1.SetToolTip(Me.txtMensajeNormal, "Mensaje Normal para los Tickets")
        '
        'chkAplicarSucursales
        '
        Me.chkAplicarSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkAplicarSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAplicarSucursales.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkAplicarSucursales.Location = New System.Drawing.Point(520, 263)
        Me.chkAplicarSucursales.Name = "chkAplicarSucursales"
        Me.chkAplicarSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAplicarSucursales.Size = New System.Drawing.Size(179, 36)
        Me.chkAplicarSucursales.TabIndex = 39
        Me.chkAplicarSucursales.Text = "Aplicar esta configuración a todas las sucursales"
        Me.chkAplicarSucursales.UseVisualStyleBackColor = False
        '
        'chkMostrarCodigoViejo
        '
        Me.chkMostrarCodigoViejo.BackColor = System.Drawing.SystemColors.Control
        Me.chkMostrarCodigoViejo.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMostrarCodigoViejo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMostrarCodigoViejo.Location = New System.Drawing.Point(22, 280)
        Me.chkMostrarCodigoViejo.Name = "chkMostrarCodigoViejo"
        Me.chkMostrarCodigoViejo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMostrarCodigoViejo.Size = New System.Drawing.Size(209, 16)
        Me.chkMostrarCodigoViejo.TabIndex = 43
        Me.chkMostrarCodigoViejo.Text = "Mostrar código anterior"
        Me.chkMostrarCodigoViejo.UseVisualStyleBackColor = False
        '
        '_Label1_9
        '
        Me._Label1_9.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_9.ForeColor = System.Drawing.Color.Black
        Me._Label1_9.Location = New System.Drawing.Point(16, 150)
        Me._Label1_9.Name = "_Label1_9"
        Me._Label1_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_9.Size = New System.Drawing.Size(179, 14)
        Me._Label1_9.TabIndex = 15
        Me._Label1_9.Text = "Efectivo Máximo en Caja ($) :"
        '
        '_Label1_18
        '
        Me._Label1_18.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_18.ForeColor = System.Drawing.Color.Black
        Me._Label1_18.Location = New System.Drawing.Point(16, 119)
        Me._Label1_18.Name = "_Label1_18"
        Me._Label1_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_18.Size = New System.Drawing.Size(160, 15)
        Me._Label1_18.TabIndex = 13
        Me._Label1_18.Text = "Símbolo Pesos p/Totales :"
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.ForeColor = System.Drawing.Color.Black
        Me._Label1_5.Location = New System.Drawing.Point(13, 55)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(193, 17)
        Me._Label1_5.TabIndex = 9
        Me._Label1_5.Text = "Tipo de Cambio de Dólar ($) :"
        '
        '_Label1_8
        '
        Me._Label1_8.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_8.ForeColor = System.Drawing.Color.Black
        Me._Label1_8.Location = New System.Drawing.Point(16, 87)
        Me._Label1_8.Name = "_Label1_8"
        Me._Label1_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_8.Size = New System.Drawing.Size(160, 14)
        Me._Label1_8.TabIndex = 11
        Me._Label1_8.Text = "Posiciones  Decim. Cant. :"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Frame11)
        Me.Panel2.Controls.Add(Me.Frame9)
        Me.Panel2.Location = New System.Drawing.Point(12, 11)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(758, 362)
        Me.Panel2.TabIndex = 18
        '
        'Frame11
        '
        Me.Frame11.BackColor = System.Drawing.SystemColors.Control
        Me.Frame11.Controls.Add(Me.txtMensajeFiscal)
        Me.Frame11.Controls.Add(Me.txtMensajeDevoluciones)
        Me.Frame11.Controls.Add(Me.txtMensajeVentas)
        Me.Frame11.Controls.Add(Me.txtMensajeNormal)
        Me.Frame11.Controls.Add(Me._Label1_25)
        Me.Frame11.Controls.Add(Me._Label1_13)
        Me.Frame11.Controls.Add(Me._Label1_12)
        Me.Frame11.Controls.Add(Me._Label1_11)
        Me.Frame11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame11.Location = New System.Drawing.Point(14, 228)
        Me.Frame11.Name = "Frame11"
        Me.Frame11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame11.Size = New System.Drawing.Size(673, 121)
        Me.Frame11.TabIndex = 34
        Me.Frame11.TabStop = False
        '
        '_Label1_25
        '
        Me._Label1_25.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_25.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_25.ForeColor = System.Drawing.Color.Black
        Me._Label1_25.Location = New System.Drawing.Point(15, 17)
        Me._Label1_25.Name = "_Label1_25"
        Me._Label1_25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_25.Size = New System.Drawing.Size(116, 17)
        Me._Label1_25.TabIndex = 35
        Me._Label1_25.Text = "Mensaje Fiscal  :"
        '
        '_Label1_13
        '
        Me._Label1_13.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_13.ForeColor = System.Drawing.Color.Black
        Me._Label1_13.Location = New System.Drawing.Point(15, 94)
        Me._Label1_13.Name = "_Label1_13"
        Me._Label1_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_13.Size = New System.Drawing.Size(146, 17)
        Me._Label1_13.TabIndex = 38
        Me._Label1_13.Text = "Mensaje Devoluciones :"
        '
        '_Label1_12
        '
        Me._Label1_12.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_12.ForeColor = System.Drawing.Color.Black
        Me._Label1_12.Location = New System.Drawing.Point(15, 69)
        Me._Label1_12.Name = "_Label1_12"
        Me._Label1_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_12.Size = New System.Drawing.Size(116, 17)
        Me._Label1_12.TabIndex = 37
        Me._Label1_12.Text = "Mensaje Crédito :"
        '
        '_Label1_11
        '
        Me._Label1_11.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_11.ForeColor = System.Drawing.Color.Black
        Me._Label1_11.Location = New System.Drawing.Point(15, 44)
        Me._Label1_11.Name = "_Label1_11"
        Me._Label1_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_11.Size = New System.Drawing.Size(116, 17)
        Me._Label1_11.TabIndex = 36
        Me._Label1_11.Text = "Mensaje Normal :"
        '
        'Frame9
        '
        Me.Frame9.BackColor = System.Drawing.SystemColors.Control
        Me.Frame9.Controls.Add(Me.txtTasaIVA)
        Me.Frame9.Controls.Add(Me.chk_AutPModifDesctos)
        Me.Frame9.Controls.Add(Me.chk_AutPConsFolVta)
        Me.Frame9.Controls.Add(Me.chk_IndCuandoProdNoSopDescto)
        Me.Frame9.Controls.Add(Me.chk_AutAbandCaptIni)
        Me.Frame9.Controls.Add(Me.chk_AutSupLinCapt)
        Me.Frame9.Controls.Add(Me.chk_AutCambCodCapt)
        Me.Frame9.Controls.Add(Me.chk_ConsDesc)
        Me.Frame9.Controls.Add(Me.chk_PermVtaSinExist)
        Me.Frame9.Controls.Add(Me.Label2)
        Me.Frame9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame9.Location = New System.Drawing.Point(14, 16)
        Me.Frame9.Name = "Frame9"
        Me.Frame9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame9.Size = New System.Drawing.Size(673, 206)
        Me.Frame9.TabIndex = 33
        Me.Frame9.TabStop = False
        '
        'txtTasaIVA
        '
        Me.txtTasaIVA.AcceptsReturn = True
        Me.txtTasaIVA.BackColor = System.Drawing.SystemColors.Window
        Me.txtTasaIVA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTasaIVA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTasaIVA.Location = New System.Drawing.Point(468, 36)
        Me.txtTasaIVA.MaxLength = 0
        Me.txtTasaIVA.Name = "txtTasaIVA"
        Me.txtTasaIVA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTasaIVA.Size = New System.Drawing.Size(49, 20)
        Me.txtTasaIVA.TabIndex = 42
        '
        'chk_AutPModifDesctos
        '
        Me.chk_AutPModifDesctos.BackColor = System.Drawing.SystemColors.Control
        Me.chk_AutPModifDesctos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chk_AutPModifDesctos.ForeColor = System.Drawing.Color.Black
        Me.chk_AutPModifDesctos.Location = New System.Drawing.Point(14, 176)
        Me.chk_AutPModifDesctos.Name = "chk_AutPModifDesctos"
        Me.chk_AutPModifDesctos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chk_AutPModifDesctos.Size = New System.Drawing.Size(248, 19)
        Me.chk_AutPModifDesctos.TabIndex = 25
        Me.chk_AutPModifDesctos.Text = "Autorización para Modificar Desctos."
        Me.chk_AutPModifDesctos.UseVisualStyleBackColor = False
        '
        'chk_AutPConsFolVta
        '
        Me.chk_AutPConsFolVta.BackColor = System.Drawing.SystemColors.Control
        Me.chk_AutPConsFolVta.Cursor = System.Windows.Forms.Cursors.Default
        Me.chk_AutPConsFolVta.ForeColor = System.Drawing.Color.Black
        Me.chk_AutPConsFolVta.Location = New System.Drawing.Point(14, 153)
        Me.chk_AutPConsFolVta.Name = "chk_AutPConsFolVta"
        Me.chk_AutPConsFolVta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chk_AutPConsFolVta.Size = New System.Drawing.Size(259, 19)
        Me.chk_AutPConsFolVta.TabIndex = 24
        Me.chk_AutPConsFolVta.Text = "Autoriz. para Consultar Folios de Venta"
        Me.chk_AutPConsFolVta.UseVisualStyleBackColor = False
        '
        'chk_IndCuandoProdNoSopDescto
        '
        Me.chk_IndCuandoProdNoSopDescto.BackColor = System.Drawing.SystemColors.Control
        Me.chk_IndCuandoProdNoSopDescto.Cursor = System.Windows.Forms.Cursors.Default
        Me.chk_IndCuandoProdNoSopDescto.ForeColor = System.Drawing.Color.Black
        Me.chk_IndCuandoProdNoSopDescto.Location = New System.Drawing.Point(14, 130)
        Me.chk_IndCuandoProdNoSopDescto.Name = "chk_IndCuandoProdNoSopDescto"
        Me.chk_IndCuandoProdNoSopDescto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chk_IndCuandoProdNoSopDescto.Size = New System.Drawing.Size(271, 19)
        Me.chk_IndCuandoProdNoSopDescto.TabIndex = 23
        Me.chk_IndCuandoProdNoSopDescto.Text = "Indicar Cuando Prod. No Soporta Descto."
        Me.chk_IndCuandoProdNoSopDescto.UseVisualStyleBackColor = False
        '
        'chk_AutAbandCaptIni
        '
        Me.chk_AutAbandCaptIni.BackColor = System.Drawing.SystemColors.Control
        Me.chk_AutAbandCaptIni.Cursor = System.Windows.Forms.Cursors.Default
        Me.chk_AutAbandCaptIni.ForeColor = System.Drawing.Color.Black
        Me.chk_AutAbandCaptIni.Location = New System.Drawing.Point(14, 107)
        Me.chk_AutAbandCaptIni.Name = "chk_AutAbandCaptIni"
        Me.chk_AutAbandCaptIni.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chk_AutAbandCaptIni.Size = New System.Drawing.Size(248, 19)
        Me.chk_AutAbandCaptIni.TabIndex = 22
        Me.chk_AutAbandCaptIni.Text = "Autoriz. Abandonar Captura Iniciada"
        Me.chk_AutAbandCaptIni.UseVisualStyleBackColor = False
        '
        'chk_AutSupLinCapt
        '
        Me.chk_AutSupLinCapt.BackColor = System.Drawing.SystemColors.Control
        Me.chk_AutSupLinCapt.Cursor = System.Windows.Forms.Cursors.Default
        Me.chk_AutSupLinCapt.ForeColor = System.Drawing.Color.Black
        Me.chk_AutSupLinCapt.Location = New System.Drawing.Point(14, 84)
        Me.chk_AutSupLinCapt.Name = "chk_AutSupLinCapt"
        Me.chk_AutSupLinCapt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chk_AutSupLinCapt.Size = New System.Drawing.Size(223, 19)
        Me.chk_AutSupLinCapt.TabIndex = 21
        Me.chk_AutSupLinCapt.Text = "Autoriz. Suprimir Linea Capturada"
        Me.chk_AutSupLinCapt.UseVisualStyleBackColor = False
        '
        'chk_AutCambCodCapt
        '
        Me.chk_AutCambCodCapt.BackColor = System.Drawing.SystemColors.Control
        Me.chk_AutCambCodCapt.Cursor = System.Windows.Forms.Cursors.Default
        Me.chk_AutCambCodCapt.ForeColor = System.Drawing.Color.Black
        Me.chk_AutCambCodCapt.Location = New System.Drawing.Point(14, 61)
        Me.chk_AutCambCodCapt.Name = "chk_AutCambCodCapt"
        Me.chk_AutCambCodCapt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chk_AutCambCodCapt.Size = New System.Drawing.Size(223, 19)
        Me.chk_AutCambCodCapt.TabIndex = 20
        Me.chk_AutCambCodCapt.Text = "Autoriz. Cambiar Código Capturado"
        Me.chk_AutCambCodCapt.UseVisualStyleBackColor = False
        '
        'chk_ConsDesc
        '
        Me.chk_ConsDesc.BackColor = System.Drawing.SystemColors.Control
        Me.chk_ConsDesc.Cursor = System.Windows.Forms.Cursors.Default
        Me.chk_ConsDesc.ForeColor = System.Drawing.Color.Black
        Me.chk_ConsDesc.Location = New System.Drawing.Point(14, 38)
        Me.chk_ConsDesc.Name = "chk_ConsDesc"
        Me.chk_ConsDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chk_ConsDesc.Size = New System.Drawing.Size(181, 19)
        Me.chk_ConsDesc.TabIndex = 19
        Me.chk_ConsDesc.Text = "Consultas por Descripción"
        Me.chk_ConsDesc.UseVisualStyleBackColor = False
        '
        'chk_PermVtaSinExist
        '
        Me.chk_PermVtaSinExist.BackColor = System.Drawing.SystemColors.Control
        Me.chk_PermVtaSinExist.Cursor = System.Windows.Forms.Cursors.Default
        Me.chk_PermVtaSinExist.ForeColor = System.Drawing.Color.Black
        Me.chk_PermVtaSinExist.Location = New System.Drawing.Point(14, 15)
        Me.chk_PermVtaSinExist.Name = "chk_PermVtaSinExist"
        Me.chk_PermVtaSinExist.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chk_PermVtaSinExist.Size = New System.Drawing.Size(181, 19)
        Me.chk_PermVtaSinExist.TabIndex = 18
        Me.chk_PermVtaSinExist.Text = "Permitir Vtas sin Existencia"
        Me.chk_PermVtaSinExist.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(364, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(98, 13)
        Me.Label2.TabIndex = 41
        Me.Label2.Text = "% Tasa de IVA :"
        '
        'Frame10
        '
        Me.Frame10.BackColor = System.Drawing.SystemColors.Control
        Me.Frame10.Controls.Add(Me._optElectronica_0)
        Me.Frame10.Controls.Add(Me._optDiskette_1)
        Me.Frame10.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame10.Location = New System.Drawing.Point(46, 136)
        Me.Frame10.Name = "Frame10"
        Me.Frame10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame10.Size = New System.Drawing.Size(330, 37)
        Me.Frame10.TabIndex = 6
        Me.Frame10.TabStop = False
        Me.Frame10.Text = "Tipo de Transferencia"
        Me.Frame10.Visible = False
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(101, 32)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(275, 21)
        Me.dbcSucursales.TabIndex = 2
        '
        '_Label1_3
        '
        Me._Label1_3.AutoSize = True
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.ForeColor = System.Drawing.Color.Black
        Me._Label1_3.Location = New System.Drawing.Point(46, 76)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(173, 13)
        Me._Label1_3.TabIndex = 3
        Me._Label1_3.Text = "Redondeo de Montos en ($) :"
        '
        'sstConfig
        '
        Me.sstConfig.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.sstConfig.Controls.Add(Me._sstConfig_TabPage0)
        Me.sstConfig.Controls.Add(Me._sstConfig_TabPage1)
        Me.sstConfig.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sstConfig.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.sstConfig.ItemSize = New System.Drawing.Size(42, 18)
        Me.sstConfig.Location = New System.Drawing.Point(5, 8)
        Me.sstConfig.Name = "sstConfig"
        Me.sstConfig.SelectedIndex = 0
        Me.sstConfig.Size = New System.Drawing.Size(797, 419)
        Me.sstConfig.TabIndex = 17
        '
        '_sstConfig_TabPage0
        '
        Me._sstConfig_TabPage0.Controls.Add(Me.Panel1)
        Me._sstConfig_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._sstConfig_TabPage0.Name = "_sstConfig_TabPage0"
        Me._sstConfig_TabPage0.Size = New System.Drawing.Size(789, 393)
        Me._sstConfig_TabPage0.TabIndex = 0
        Me._sstConfig_TabPage0.Text = "General"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.chkAplicarSucursales)
        Me.Panel1.Controls.Add(Me.chkMostrarCodigoViejo)
        Me.Panel1.Controls.Add(Me.Frame5)
        Me.Panel1.Controls.Add(Me.chkTransferenciaTicket)
        Me.Panel1.Controls.Add(Me.Frame7)
        Me.Panel1.Location = New System.Drawing.Point(13, 14)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(751, 334)
        Me.Panel1.TabIndex = 18
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.cboCentavos)
        Me.Frame5.Controls.Add(Me.dbcSucursales)
        Me.Frame5.Controls.Add(Me.Frame10)
        Me.Frame5.Controls.Add(Me._Label1_3)
        Me.Frame5.Controls.Add(Me._Label1_1)
        Me.Frame5.Controls.Add(Me.chkCapturarCantArticulos)
        Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame5.Location = New System.Drawing.Point(22, 24)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(393, 205)
        Me.Frame5.TabIndex = 0
        Me.Frame5.TabStop = False
        '
        'Frame7
        '
        Me.Frame7.BackColor = System.Drawing.SystemColors.Control
        Me.Frame7.Controls.Add(Me.txtEfectivoMaxi)
        Me.Frame7.Controls.Add(Me.txtSimbolo)
        Me.Frame7.Controls.Add(Me.txtTipoCambioDolar)
        Me.Frame7.Controls.Add(Me.txtPosicionesDeci)
        Me.Frame7.Controls.Add(Me._Label1_9)
        Me.Frame7.Controls.Add(Me._Label1_18)
        Me.Frame7.Controls.Add(Me._Label1_5)
        Me.Frame7.Controls.Add(Me._Label1_8)
        Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame7.Location = New System.Drawing.Point(434, 24)
        Me.Frame7.Name = "Frame7"
        Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame7.Size = New System.Drawing.Size(273, 205)
        Me.Frame7.TabIndex = 31
        Me.Frame7.TabStop = False
        '
        '_sstConfig_TabPage1
        '
        Me._sstConfig_TabPage1.Controls.Add(Me.Panel2)
        Me._sstConfig_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._sstConfig_TabPage1.Name = "_sstConfig_TabPage1"
        Me._sstConfig_TabPage1.Size = New System.Drawing.Size(789, 393)
        Me._sstConfig_TabPage1.TabIndex = 1
        Me._sstConfig_TabPage1.Text = "Captura de Ventas"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(127, 436)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 78
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(12, 436)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 77
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmPVConfigPuntoVenta
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(812, 485)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.sstConfig)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(153, 130)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPVConfigPuntoVenta"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Configuración del Sistema"
        Me.Panel2.ResumeLayout(False)
        Me.Frame11.ResumeLayout(False)
        Me.Frame11.PerformLayout()
        Me.Frame9.ResumeLayout(False)
        Me.Frame9.PerformLayout()
        Me.Frame10.ResumeLayout(False)
        Me.sstConfig.ResumeLayout(False)
        Me._sstConfig_TabPage0.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Frame5.ResumeLayout(False)
        Me.Frame5.PerformLayout()
        Me.Frame7.ResumeLayout(False)
        Me.Frame7.PerformLayout()
        Me._sstConfig_TabPage1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub
End Class