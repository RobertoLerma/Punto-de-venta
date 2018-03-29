Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmPVConfigFacturacion
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             CONFIGURACION DE LA FACTURA                                                                  *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      SABADO 17 DE MAYO DE 2003                                                                    *'
    '*FECHA DE TERMINACION : MIERCOLES 21 DE MAYO DE 2003                                                                 *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents _optTicket_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optTicket_1 As System.Windows.Forms.RadioButton
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents chkAplicarSucursales As System.Windows.Forms.CheckBox
    Public WithEvents btnLong As System.Windows.Forms.Button
    Public WithEvents txtLeyenda As System.Windows.Forms.TextBox
    Public WithEvents cboLetra As System.Windows.Forms.ComboBox
    Public WithEvents TxtCoordLeyenda As System.Windows.Forms.TextBox
    Public WithEvents txtColIVAporPartida As System.Windows.Forms.TextBox
    Public WithEvents txtColDesctoDetalle As System.Windows.Forms.TextBox
    Public WithEvents txtColDescripcion As System.Windows.Forms.TextBox
    Public WithEvents txtRenXdetalle As System.Windows.Forms.TextBox
    Public WithEvents txtCoordLugarExped As System.Windows.Forms.TextBox
    Public WithEvents txtCoordCP As System.Windows.Forms.TextBox
    Public WithEvents txtCoordFolio As System.Windows.Forms.TextBox
    Public WithEvents txtImpteConLetra As System.Windows.Forms.TextBox
    Public WithEvents txtPrimerPartida As System.Windows.Forms.TextBox
    Public WithEvents txtPartidasXfactura As System.Windows.Forms.TextBox
    Public WithEvents txtCoordTotal As System.Windows.Forms.TextBox
    Public WithEvents txtCoordIVA As System.Windows.Forms.TextBox
    Public WithEvents txtCoordSubTotal As System.Windows.Forms.TextBox
    Public WithEvents txtCoordImporte As System.Windows.Forms.TextBox
    Public WithEvents txtCoordPrecioVta As System.Windows.Forms.TextBox
    Public WithEvents txtCoordPromocion As System.Windows.Forms.TextBox
    Public WithEvents txtCoordDesctos As System.Windows.Forms.TextBox
    Public WithEvents txtCoordCantidad As System.Windows.Forms.TextBox
    Public WithEvents txtCoordCodigo As System.Windows.Forms.TextBox
    Public WithEvents txtCoordTelefono As System.Windows.Forms.TextBox
    Public WithEvents txtCoordEstado As System.Windows.Forms.TextBox
    Public WithEvents txtCoordCiudad As System.Windows.Forms.TextBox
    Public WithEvents txtCoordColonia As System.Windows.Forms.TextBox
    Public WithEvents txtCoordCalle As System.Windows.Forms.TextBox
    Public WithEvents txtCoordFecha As System.Windows.Forms.TextBox
    Public WithEvents txtCoordRFC As System.Windows.Forms.TextBox
    Public WithEvents txtCoordEmpresa As System.Windows.Forms.TextBox
    Public WithEvents txtRenXFactura As System.Windows.Forms.TextBox
    Public WithEvents _Line1_1 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents _Line1_0 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_100 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_0 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents txtDatosPartida As System.Windows.Forms.Label
    Public WithEvents TxtDatosGenerales As System.Windows.Forms.Label
    Public WithEvents _Label2_28 As System.Windows.Forms.Label
    Public WithEvents _Label2_27 As System.Windows.Forms.Label
    Public WithEvents _Label2_26 As System.Windows.Forms.Label
    Public WithEvents _Label2_25 As System.Windows.Forms.Label
    Public WithEvents _Label2_24 As System.Windows.Forms.Label
    Public WithEvents _Label2_158 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_153 As System.Windows.Forms.Label
    Public WithEvents _Label2_21 As System.Windows.Forms.Label
    Public WithEvents _Label2_20 As System.Windows.Forms.Label
    Public WithEvents _Label2_19 As System.Windows.Forms.Label
    Public WithEvents _Label2_18 As System.Windows.Forms.Label
    Public WithEvents _Label2_16 As System.Windows.Forms.Label
    Public WithEvents _Label2_15 As System.Windows.Forms.Label
    Public WithEvents _Label2_160 As System.Windows.Forms.Label
    Public WithEvents _Label2_161 As System.Windows.Forms.Label
    Public WithEvents _Label2_12 As System.Windows.Forms.Label
    Public WithEvents _Label2_11 As System.Windows.Forms.Label
    Public WithEvents _Label2_10 As System.Windows.Forms.Label
    Public WithEvents _Label2_9 As System.Windows.Forms.Label
    Public WithEvents _Label2_159 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_157 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_156 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_155 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_154 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_152 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_151 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_150 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_1 As System.Windows.Forms.Label
    Public WithEvents Marco As System.Windows.Forms.GroupBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents mskFecha As System.Windows.Forms.MaskedTextBox
    Public WithEvents mskHora As System.Windows.Forms.MaskedTextBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents Line1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblEtiqueta As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optTicket As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray


    Dim mblnNuevo As Boolean
    Dim mblnSALIR As Boolean 'Para Salir de la Captura Sin Preguntar Por Cambios
    Dim mblnSaliryGrabar As Boolean
    Dim mblnTicket As Boolean
    Dim mblnFormato As Boolean
    Dim mstrCadena1 As String
    Dim mstrCadena2 As String
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim rsAux As ADODB.Recordset
    Dim FueraChange As Boolean
    Public WithEvents Panel3 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Dim I As Integer
    Function Guardar() As Boolean
        On Error GoTo Merr
        Dim blnTransaccion As Boolean
        Dim LonCliente As String
        Dim LonDireccion As String
        Dim LonColonia As String
        Dim LonCiudad As String
        Dim LonEstado As String
        Dim LonProducto As String
        Dim LonLeyenda As String
        Guardar = False
        If Cambios() = False And chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            mblnSaliryGrabar = True
            Me.Close()
            Exit Function
        End If
        ValidaDatos()

        frmPVConfigLongitudDeDatosFactura.InitializeComponent()
        With frmPVConfigLongitudDeDatosFactura.FlexDetalle
            .Col = 1
            .Row = 1
            LonCliente = .Text
            .Row = 2
            LonDireccion = .Text
            .Row = 3
            LonColonia = .Text
            .Row = 4
            LonCiudad = .Text
            .Row = 5
            LonEstado = .Text
            .Row = 6
            LonLeyenda = .Text
            .Row = 7
            LonProducto = .Text
        End With
        'Si se selecciono la Opción de Aplicar a todas las Sucursales
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
        blnTransaccion = True

        'Eliminar todos los datos de Configuración generalde la tabla, par posteiormente solo insertar
        ''    gStrSql = "Delete FROM ConfigFactura Where CodAlmacen = " & intCodSucursal
        ''    ModEstandar.BorraCmd
        ''    Cmd.CommandText = "dbo.Up_Select_Datos"
        ''    Cmd.CommandType = adCmdStoredProc
        ''    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, gStrSql)
        ''    Cmd.Execute

        ModStoredProcedures.PR_IMEConfigFactura(CStr(intCodSucursal), txtRenXFactura.Text, BuscaRenglon(txtCoordEmpresa), BuscaColumna(txtCoordEmpresa), BuscaRenglon(txtCoordRFC), BuscaColumna(txtCoordRFC), BuscaRenglon(txtCoordFecha), BuscaColumna(txtCoordFecha), BuscaRenglon(txtCoordFolio), BuscaColumna(txtCoordFolio), BuscaRenglon(txtCoordCalle), BuscaColumna(txtCoordCalle), BuscaRenglon(txtCoordColonia), BuscaColumna(txtCoordColonia), BuscaRenglon(txtCoordCiudad), BuscaColumna(txtCoordCiudad), BuscaRenglon(txtCoordEstado), BuscaColumna(txtCoordEstado), BuscaRenglon(txtCoordCP), BuscaColumna(txtCoordCP), BuscaRenglon(txtCoordTelefono), BuscaColumna(txtCoordTelefono), BuscaRenglon(txtCoordSubTotal), BuscaColumna(txtCoordSubTotal), BuscaRenglon(txtCoordIVA), BuscaColumna(txtCoordIVA), BuscaRenglon(txtCoordTotal), BuscaColumna(txtCoordTotal), BuscaRenglon(txtCoordDesctos), BuscaColumna(txtCoordDesctos), BuscaRenglon(txtImpteConLetra), BuscaColumna(txtImpteConLetra), BuscaRenglon(txtCoordLugarExped), BuscaColumna(txtCoordLugarExped), txtRenXdetalle.Text, txtPrimerPartida.Text, txtCoordCodigo.Text, txtCoordCantidad.Text, txtColDescripcion.Text, txtColDesctoDetalle.Text, txtCoordPromocion.Text, txtCoordPrecioVta.Text, txtCoordImporte.Text, txtColIVAporPartida.Text, IIf(_optTicket_0.Checked, "True", "False"), BuscaColumna(TxtCoordLeyenda), BuscaRenglon(TxtCoordLeyenda), txtLeyenda.Text, (cboLetra.Text), LonCliente, LonDireccion, LonColonia, LonCiudad, LonEstado, LonLeyenda, LonProducto, C_ELIMINACION, CStr(0))
        Cmd.Execute()

        If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            RsGral.MoveFirst()
            For I = 1 To RsGral.RecordCount
                If RsGral.Fields("CodAlmacen").Value <> intCodSucursal Then
                    '                gStrSql = "Delete FROM ConfigFactura Where CodAlmacen = " & RsGral!CodAlmacen
                    '                ModEstandar.BorraCmd
                    '                Cmd.CommandText = "dbo.Up_Select_Datos"
                    '                Cmd.CommandType = adCmdStoredProc
                    '                Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
                    '                Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, gStrSql)
                    '                Cmd.Execute
                    ModStoredProcedures.PR_IMEConfigFactura(CStr(RsGral.Fields("CodAlmacen").Value), txtRenXFactura.Text, BuscaRenglon(txtCoordEmpresa), BuscaColumna(txtCoordEmpresa), BuscaRenglon(txtCoordRFC), BuscaColumna(txtCoordRFC), BuscaRenglon(txtCoordFecha), BuscaColumna(txtCoordFecha), BuscaRenglon(txtCoordFolio), BuscaColumna(txtCoordFolio), BuscaRenglon(txtCoordCalle), BuscaColumna(txtCoordCalle), BuscaRenglon(txtCoordColonia), BuscaColumna(txtCoordColonia), BuscaRenglon(txtCoordCiudad), BuscaColumna(txtCoordCiudad), BuscaRenglon(txtCoordEstado), BuscaColumna(txtCoordEstado), BuscaRenglon(txtCoordCP), BuscaColumna(txtCoordCP), BuscaRenglon(txtCoordTelefono), BuscaColumna(txtCoordTelefono), BuscaRenglon(txtCoordSubTotal), BuscaColumna(txtCoordSubTotal), BuscaRenglon(txtCoordIVA), BuscaColumna(txtCoordIVA), BuscaRenglon(txtCoordTotal), BuscaColumna(txtCoordTotal), BuscaRenglon(txtCoordDesctos), BuscaColumna(txtCoordDesctos), BuscaRenglon(txtImpteConLetra), BuscaColumna(txtImpteConLetra), BuscaRenglon(txtCoordLugarExped), BuscaColumna(txtCoordLugarExped), txtRenXdetalle.Text, txtPrimerPartida.Text, txtCoordCodigo.Text, txtCoordCantidad.Text, txtColDescripcion.Text, txtColDesctoDetalle.Text, txtCoordPromocion.Text, txtCoordPrecioVta.Text, txtCoordImporte.Text, txtColIVAporPartida.Text, IIf(_optTicket_0.Checked, "True", "False"), BuscaColumna(TxtCoordLeyenda), BuscaRenglon(TxtCoordLeyenda), txtLeyenda.Text, (cboLetra.Text), LonCliente, LonDireccion, LonColonia, LonCiudad, LonEstado, LonLeyenda, LonProducto, C_ELIMINACION, CStr(0))
                    Cmd.Execute()
                End If
                RsGral.MoveNext()
            Next
        End If

        '    If mblnNuevo Then
        ModStoredProcedures.PR_IMEConfigFactura(CStr(intCodSucursal), txtRenXFactura.Text, BuscaRenglon(txtCoordEmpresa), BuscaColumna(txtCoordEmpresa), BuscaRenglon(txtCoordRFC), BuscaColumna(txtCoordRFC), BuscaRenglon(txtCoordFecha), BuscaColumna(txtCoordFecha), BuscaRenglon(txtCoordFolio), BuscaColumna(txtCoordFolio), BuscaRenglon(txtCoordCalle), BuscaColumna(txtCoordCalle), BuscaRenglon(txtCoordColonia), BuscaColumna(txtCoordColonia), BuscaRenglon(txtCoordCiudad), BuscaColumna(txtCoordCiudad), BuscaRenglon(txtCoordEstado), BuscaColumna(txtCoordEstado), BuscaRenglon(txtCoordCP), BuscaColumna(txtCoordCP), BuscaRenglon(txtCoordTelefono), BuscaColumna(txtCoordTelefono), BuscaRenglon(txtCoordSubTotal), BuscaColumna(txtCoordSubTotal), BuscaRenglon(txtCoordIVA), BuscaColumna(txtCoordIVA), BuscaRenglon(txtCoordTotal), BuscaColumna(txtCoordTotal), BuscaRenglon(txtCoordDesctos), BuscaColumna(txtCoordDesctos), BuscaRenglon(txtImpteConLetra), BuscaColumna(txtImpteConLetra), BuscaRenglon(txtCoordLugarExped), BuscaColumna(txtCoordLugarExped), txtRenXdetalle.Text, txtPrimerPartida.Text, txtCoordCodigo.Text, txtCoordCantidad.Text, txtColDescripcion.Text, txtColDesctoDetalle.Text, txtCoordPromocion.Text, txtCoordPrecioVta.Text, txtCoordImporte.Text, txtColIVAporPartida.Text, IIf(_optTicket_0.Checked, "True", "False"), BuscaColumna(TxtCoordLeyenda), BuscaRenglon(TxtCoordLeyenda), txtLeyenda.Text, (cboLetra.Text), LonCliente, LonDireccion, LonColonia, LonCiudad, LonEstado, LonLeyenda, LonProducto, C_INSERCION, CStr(0))
        Cmd.Execute()
        If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            RsGral.MoveFirst()
            For I = 1 To RsGral.RecordCount
                If RsGral.Fields("CodAlmacen").Value <> intCodSucursal Then
                    ModStoredProcedures.PR_IMEConfigFactura(CStr(RsGral.Fields("CodAlmacen").Value), txtRenXFactura.Text, BuscaRenglon(txtCoordEmpresa), BuscaColumna(txtCoordEmpresa), BuscaRenglon(txtCoordRFC), BuscaColumna(txtCoordRFC), BuscaRenglon(txtCoordFecha), BuscaColumna(txtCoordFecha), BuscaRenglon(txtCoordFolio), BuscaColumna(txtCoordFolio), BuscaRenglon(txtCoordCalle), BuscaColumna(txtCoordCalle), BuscaRenglon(txtCoordColonia), BuscaColumna(txtCoordColonia), BuscaRenglon(txtCoordCiudad), BuscaColumna(txtCoordCiudad), BuscaRenglon(txtCoordEstado), BuscaColumna(txtCoordEstado), BuscaRenglon(txtCoordCP), BuscaColumna(txtCoordCP), BuscaRenglon(txtCoordTelefono), BuscaColumna(txtCoordTelefono), BuscaRenglon(txtCoordSubTotal), BuscaColumna(txtCoordSubTotal), BuscaRenglon(txtCoordIVA), BuscaColumna(txtCoordIVA), BuscaRenglon(txtCoordTotal), BuscaColumna(txtCoordTotal), BuscaRenglon(txtCoordDesctos), BuscaColumna(txtCoordDesctos), BuscaRenglon(txtImpteConLetra), BuscaColumna(txtImpteConLetra), BuscaRenglon(txtCoordLugarExped), BuscaColumna(txtCoordLugarExped), txtRenXdetalle.Text, txtPrimerPartida.Text, txtCoordCodigo.Text, txtCoordCantidad.Text, txtColDescripcion.Text, txtColDesctoDetalle.Text, txtCoordPromocion.Text, txtCoordPrecioVta.Text, txtCoordImporte.Text, txtColIVAporPartida.Text, IIf(_optTicket_0.Checked, "True", "False"), BuscaColumna(TxtCoordLeyenda), BuscaRenglon(TxtCoordLeyenda), txtLeyenda.Text, (cboLetra.Text), LonCliente, LonDireccion, LonColonia, LonCiudad, LonEstado, LonLeyenda, LonProducto, C_INSERCION, CStr(0))
                    Cmd.Execute()
                End If
                RsGral.MoveNext()
            Next
        End If
        Cnn.CommitTrans()
        blnTransaccion = False
        If mblnNuevo Then
            MsgBox("La Configuración de la Factura Ha sido Grabada Correctamente ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
        End If
        InicializaVariables()
        Guardar = True
        mblnSaliryGrabar = True
        Nuevo()
        Limpiar()
Merr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Function

    Sub LlenaDatos()
        On Error GoTo Merr
        gStrSql = "SELECT     dbo.CatAlmacen.DescAlmacen AS DescAlmacen, dbo.ConfigFactura.* " & "FROM         dbo.CatAlmacen INNER JOIN " & "dbo.ConfigFactura ON dbo.CatAlmacen.CodAlmacen = dbo.ConfigFactura.CodAlmacen " & "WHERE     (dbo.CatAlmacen.CodAlmacen = " & intCodSucursal & ")"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        InicializaVariables()
        If RsGral.RecordCount > 0 Then
            mblnNuevo = False
            cboLetra.Text = RsGral.Fields("TamLetra").Value
            cboLetra.Tag = RsGral.Fields("TamLetra").Value
            If RsGral.Fields("RenTotales").Value <> 0 Then
                txtRenXFactura.Text = RsGral.Fields("RenTotales").Value
                txtRenXFactura.Tag = RsGral.Fields("RenTotales").Value
            End If
            If RsGral.Fields("RenFolio").Value <> 0 And RsGral.Fields("ColFolio").Value <> 0 Then
                txtCoordFolio.Text = RsGral.Fields("RenFolio").Value & "," & RsGral.Fields("ColFolio").Value
                txtCoordFolio.Tag = RsGral.Fields("RenFolio").Value & "," & RsGral.Fields("ColFolio").Value
            End If
            If RsGral.Fields("RenFecha").Value <> 0 And RsGral.Fields("ColFecha").Value <> 0 Then
                txtCoordFecha.Text = RsGral.Fields("RenFecha").Value & "," & RsGral.Fields("ColFecha").Value
                txtCoordFecha.Tag = RsGral.Fields("RenFecha").Value & "," & RsGral.Fields("ColFecha").Value
            End If
            If RsGral.Fields("RenEmpresa").Value <> 0 And RsGral.Fields("ColEmpresa").Value <> 0 Then
                txtCoordEmpresa.Text = RsGral.Fields("RenEmpresa").Value & "," & RsGral.Fields("ColEmpresa").Value
                txtCoordEmpresa.Tag = RsGral.Fields("RenEmpresa").Value & "," & RsGral.Fields("ColEmpresa").Value
            End If
            If RsGral.Fields("RenRFC").Value <> 0 And RsGral.Fields("ColRFC").Value <> 0 Then
                txtCoordRFC.Text = RsGral.Fields("RenRFC").Value & "," & RsGral.Fields("ColRFC").Value
                txtCoordRFC.Tag = RsGral.Fields("RenRFC").Value & "," & RsGral.Fields("ColRFC").Value
            End If
            If RsGral.Fields("RenCalle").Value <> 0 And RsGral.Fields("ColCalle").Value <> 0 Then
                txtCoordCalle.Text = RsGral.Fields("RenCalle").Value & "," & RsGral.Fields("ColCalle").Value
                txtCoordCalle.Tag = RsGral.Fields("RenCalle").Value & "," & RsGral.Fields("ColCalle").Value
            End If
            If RsGral.Fields("RenColonia").Value <> 0 And RsGral.Fields("ColColonia").Value <> 0 Then
                txtCoordColonia.Text = RsGral.Fields("RenColonia").Value & "," & RsGral.Fields("ColColonia").Value
                txtCoordColonia.Tag = RsGral.Fields("RenColonia").Value & "," & RsGral.Fields("ColColonia").Value
            End If
            If RsGral.Fields("RenCP").Value <> 0 And RsGral.Fields("ColCP").Value <> 0 Then
                txtCoordCP.Text = RsGral.Fields("RenCP").Value & "," & RsGral.Fields("ColCP").Value
                txtCoordCP.Tag = RsGral.Fields("RenCP").Value & "," & RsGral.Fields("ColCP").Value
            End If
            If RsGral.Fields("RenTelefono").Value <> 0 And RsGral.Fields("ColTelefono").Value <> 0 Then
                txtCoordTelefono.Text = RsGral.Fields("RenTelefono").Value & "," & RsGral.Fields("ColTelefono").Value
                txtCoordTelefono.Tag = RsGral.Fields("RenTelefono").Value & "," & RsGral.Fields("ColTelefono").Value
            End If
            If RsGral.Fields("RenCiudad").Value <> 0 And RsGral.Fields("ColCiudad").Value <> 0 Then
                txtCoordCiudad.Text = RsGral.Fields("RenCiudad").Value & "," & RsGral.Fields("ColCiudad").Value
                txtCoordCiudad.Tag = RsGral.Fields("RenCiudad").Value & "," & RsGral.Fields("ColCiudad").Value
            End If
            If RsGral.Fields("RenEstado").Value <> 0 And RsGral.Fields("ColEstado").Value <> 0 Then
                txtCoordEstado.Text = RsGral.Fields("RenEstado").Value & "," & RsGral.Fields("ColEstado").Value
                txtCoordEstado.Tag = RsGral.Fields("RenEstado").Value & "," & RsGral.Fields("ColEstado").Value
            End If
            If RsGral.Fields("RenSubTotal").Value <> 0 And RsGral.Fields("ColSubTotal").Value <> 0 Then
                txtCoordSubTotal.Text = RsGral.Fields("RenSubTotal").Value & "," & RsGral.Fields("ColSubTotal").Value
                txtCoordSubTotal.Tag = RsGral.Fields("RenSubTotal").Value & "," & RsGral.Fields("ColSubTotal").Value
            End If
            If RsGral.Fields("RenDescto").Value <> 0 And RsGral.Fields("ColDescto").Value <> 0 Then
                txtCoordDesctos.Text = RsGral.Fields("RenDescto").Value & "," & RsGral.Fields("ColDescto").Value
                txtCoordDesctos.Tag = RsGral.Fields("RenDescto").Value & "," & RsGral.Fields("ColDescto").Value
            End If
            If RsGral.Fields("RenIva").Value <> 0 And RsGral.Fields("ColIva").Value <> 0 Then
                txtCoordIVA.Text = RsGral.Fields("RenIva").Value & "," & RsGral.Fields("ColIva").Value
                txtCoordIVA.Tag = RsGral.Fields("RenIva").Value & "," & RsGral.Fields("ColIva").Value
            End If
            If RsGral.Fields("RenTotal").Value <> 0 And RsGral.Fields("ColTotal").Value <> 0 Then
                txtCoordTotal.Text = RsGral.Fields("RenTotal").Value & "," & RsGral.Fields("ColTotal").Value
                txtCoordTotal.Tag = RsGral.Fields("RenTotal").Value & "," & RsGral.Fields("ColTotal").Value
            End If
            If RsGral.Fields("RenImpLetra").Value <> 0 And RsGral.Fields("ColImpLetra").Value <> 0 Then
                txtImpteConLetra.Text = RsGral.Fields("RenImpLetra").Value & "," & RsGral.Fields("ColImpLetra").Value
                txtImpteConLetra.Tag = RsGral.Fields("RenImpLetra").Value & "," & RsGral.Fields("ColImpLetra").Value
            End If
            If RsGral.Fields("RenLugarExped").Value <> 0 And RsGral.Fields("ColLugarExped").Value <> 0 Then
                txtCoordLugarExped.Text = RsGral.Fields("RenLugarExped").Value & "," & RsGral.Fields("ColLugarExped").Value
                txtCoordLugarExped.Tag = RsGral.Fields("RenLugarExped").Value & "," & RsGral.Fields("ColLugarExped").Value
            End If
            If RsGral.Fields("RenLeyenda").Value <> 0 And RsGral.Fields("ColLeyenda").Value <> 0 Then
                TxtCoordLeyenda.Text = RsGral.Fields("RenLeyenda").Value & "," & RsGral.Fields("ColLeyenda").Value
                TxtCoordLeyenda.Tag = RsGral.Fields("RenLeyenda").Value & "," & RsGral.Fields("ColLeyenda").Value
            End If
            If RsGral.Fields("CantRenXDet").Value <> 0 Then
                txtRenXdetalle.Text = RsGral.Fields("CantRenXDet").Value
                txtRenXdetalle.Tag = RsGral.Fields("CantRenXDet").Value
            End If
            If RsGral.Fields("RenPrimerPartida").Value <> 0 Then
                txtPrimerPartida.Text = RsGral.Fields("RenPrimerPartida").Value
                txtPrimerPartida.Tag = RsGral.Fields("RenPrimerPartida").Value
            End If
            If RsGral.Fields("ColCantidad").Value <> 0 Then
                txtCoordCantidad.Text = RsGral.Fields("ColCantidad").Value
                txtCoordCantidad.Tag = RsGral.Fields("ColCantidad").Value
            End If
            If RsGral.Fields("ColCodigo").Value <> 0 Then
                txtCoordCodigo.Text = RsGral.Fields("ColCodigo").Value
                txtCoordCodigo.Tag = RsGral.Fields("ColCodigo").Value
            End If
            If RsGral.Fields("ColDescripcion").Value <> 0 Then
                txtColDescripcion.Text = RsGral.Fields("ColDescripcion").Value
                txtColDescripcion.Tag = RsGral.Fields("ColDescripcion").Value
            End If
            If RsGral.Fields("ColDesctoDetalle").Value <> 0 Then
                txtColDesctoDetalle.Text = RsGral.Fields("ColDesctoDetalle").Value
                txtColDesctoDetalle.Tag = RsGral.Fields("ColDesctoDetalle").Value
            End If
            If RsGral.Fields("ColPromocion").Value <> 0 Then
                txtCoordPromocion.Text = RsGral.Fields("ColPromocion").Value
                txtCoordPromocion.Tag = RsGral.Fields("ColPromocion").Value
            End If
            If RsGral.Fields("ColIVAPartida").Value <> 0 Then
                txtColIVAporPartida.Text = RsGral.Fields("ColIVAPartida").Value
                txtColIVAporPartida.Tag = RsGral.Fields("ColIVAPartida").Value
            End If
            If RsGral.Fields("ColPrecioVenta").Value <> 0 Then
                txtCoordPrecioVta.Text = RsGral.Fields("ColPrecioVenta").Value
                txtCoordPrecioVta.Tag = RsGral.Fields("ColPrecioVenta").Value
            End If
            If RsGral.Fields("ColImporte").Value <> 0 Then
                txtCoordImporte.Text = RsGral.Fields("ColImporte").Value
                txtCoordImporte.Tag = RsGral.Fields("ColImporte").Value
            End If
            txtLeyenda.Text = Trim(RsGral.Fields("Leyenda").Value)
            txtLeyenda.Tag = Trim(RsGral.Fields("Leyenda").Value)

            frmPVConfigLongitudDeDatosFactura.InitializeComponent()
            With frmPVConfigLongitudDeDatosFactura.FlexDetalle
                .Col = 1
                .Row = 1
                If RsGral.Fields("LongCliente").Value <> 0 Then
                    .Text = RsGral.Fields("LongCliente").Value
                    gintLonCliente = RsGral.Fields("LongCliente").Value
                End If
                .Row = 2
                If RsGral.Fields("LongDireccion").Value <> 0 Then
                    .Text = RsGral.Fields("LongDireccion").Value
                    gintLonDireccion = RsGral.Fields("LongDireccion").Value
                End If
                .Row = 3
                If RsGral.Fields("LongColonia").Value <> 0 Then
                    .Text = RsGral.Fields("LongColonia").Value
                    gintLonColonia = RsGral.Fields("LongColonia").Value
                End If
                .Row = 4
                If RsGral.Fields("LongCiudad").Value <> 0 Then
                    .Text = RsGral.Fields("LongCiudad").Value
                    gintLonCiudad = RsGral.Fields("LongCiudad").Value
                End If
                .Row = 5
                If RsGral.Fields("LongEstado").Value <> 0 Then
                    .Text = RsGral.Fields("LongEstado").Value
                    gintLonEstado = RsGral.Fields("LongEstado").Value
                End If
                .Row = 6
                If RsGral.Fields("LongLeyenda").Value <> 0 Then
                    .Text = RsGral.Fields("LongLeyenda").Value
                    gintLonLeyenda = RsGral.Fields("LongLeyenda").Value
                End If
                .Row = 7
                If RsGral.Fields("LongProducto").Value <> 0 Then
                    .Text = RsGral.Fields("LongProducto").Value
                    gintLonDescProducto = RsGral.Fields("LongProducto").Value
                End If
                .Row = 1
                .Col = 1
            End With
            If RsGral.Fields("Ticket").Value Then
                _optTicket_0.Checked = True
                mblnTicket = True
                mblnFormato = False
            Else
                _optTicket_1.Checked = True
                mblnTicket = False
                mblnFormato = True
            End If
        Else
            cboLetra.SelectedIndex = 0
            cboLetra.Tag = cboLetra.Text
            mblnNuevo = True
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function Cambios() As Boolean
        Cambios = True
        If cboLetra.Text <> cboLetra.Tag Then Exit Function
        If Trim(txtRenXFactura.Text) <> txtRenXFactura.Tag Then Exit Function
        If Trim(txtCoordFolio.Text) <> txtCoordFolio.Tag Then Exit Function
        If Trim(txtCoordFecha.Text) <> txtCoordFecha.Tag Then Exit Function
        If Trim(txtCoordEmpresa.Text) <> txtCoordEmpresa.Tag Then Exit Function
        If Trim(txtCoordRFC.Text) <> txtCoordRFC.Tag Then Exit Function
        If Trim(txtCoordCalle.Text) <> txtCoordCalle.Tag Then Exit Function
        If Trim(txtCoordColonia.Text) <> txtCoordColonia.Tag Then Exit Function
        If Trim(txtCoordCP.Text) <> txtCoordCP.Tag Then Exit Function
        If Trim(txtCoordTelefono.Text) <> txtCoordTelefono.Tag Then Exit Function
        If Trim(txtCoordCiudad.Text) <> txtCoordCiudad.Tag Then Exit Function
        If Trim(txtCoordEstado.Text) <> txtCoordEstado.Tag Then Exit Function
        If Trim(txtCoordSubTotal.Text) <> txtCoordSubTotal.Tag Then Exit Function
        If Trim(txtCoordDesctos.Text) <> txtCoordDesctos.Tag Then Exit Function
        If Trim(txtCoordIVA.Text) <> txtCoordIVA.Tag Then Exit Function
        If Trim(txtCoordTotal.Text) <> txtCoordTotal.Tag Then Exit Function
        If Trim(txtImpteConLetra.Text) <> txtImpteConLetra.Tag Then Exit Function
        If Trim(txtCoordLugarExped.Text) <> txtCoordLugarExped.Tag Then Exit Function
        If Trim(TxtCoordLeyenda.Text) <> TxtCoordLeyenda.Tag Then Exit Function
        If Trim(txtRenXdetalle.Text) <> txtRenXdetalle.Tag Then Exit Function
        If Trim(txtPrimerPartida.Text) <> txtPrimerPartida.Tag Then Exit Function
        If Trim(txtCoordCantidad.Text) <> txtCoordCantidad.Tag Then Exit Function
        If Trim(txtCoordCodigo.Text) <> txtCoordCodigo.Tag Then Exit Function
        If Trim(txtColDescripcion.Text) <> txtColDescripcion.Tag Then Exit Function
        If Trim(txtColDesctoDetalle.Text) <> txtColDesctoDetalle.Tag Then Exit Function
        If Trim(txtCoordPromocion.Text) <> txtCoordPromocion.Tag Then Exit Function
        If Trim(txtColIVAporPartida.Text) <> txtColIVAporPartida.Tag Then Exit Function
        If Trim(txtCoordPrecioVta.Text) <> txtCoordPrecioVta.Tag Then Exit Function
        If Trim(txtCoordImporte.Text) <> txtCoordImporte.Tag Then Exit Function
        If Trim(txtLeyenda.Text) <> txtLeyenda.Tag Then Exit Function
        If _optTicket_0.Checked <> mblnTicket Then Exit Function
        If _optTicket_1.Checked <> mblnFormato Then Exit Function
        With frmPVConfigLongitudDeDatosFactura.FlexDetalle
            .Col = 1
            .Row = 1
            If Val(.Text) <> gintLonCliente Then Exit Function
            .Row = 2
            If Val(.Text) <> gintLonDireccion Then Exit Function
            .Row = 3
            If Val(.Text) <> gintLonColonia Then Exit Function
            .Row = 4
            If Val(.Text) <> gintLonCiudad Then Exit Function
            .Row = 5
            If Val(.Text) <> gintLonEstado Then Exit Function
            .Row = 6
            If Val(.Text) <> gintLonLeyenda Then Exit Function
            .Row = 7
            If Val(.Text) <> gintLonDescProducto Then Exit Function
        End With
        Cambios = False
    End Function

    Sub InicializaVariables()
        mblnSALIR = False
        mblnTicket = True
        mblnFormato = False
        mblnSaliryGrabar = False
    End Sub

    Sub ValidaDatos()
        Dim LonCliente As Integer
        Dim LonDireccion As Integer
        Dim LonColonia As Integer
        Dim LonCiudad As Integer
        Dim LonEstado As Integer
        Dim LonLeyenda As Integer
        Dim LonProducto As Integer
        If txtRenXFactura.Text = "" Then txtRenXFactura.Text = CStr(0)
        If Not BuscaCoordenadas(txtCoordFolio) Then txtCoordFolio.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordFecha) Then txtCoordFecha.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordEmpresa) Then txtCoordEmpresa.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordRFC) Then txtCoordRFC.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordCalle) Then txtCoordCalle.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordColonia) Then txtCoordColonia.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordCP) Then txtCoordCP.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordTelefono) Then txtCoordTelefono.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordCiudad) Then txtCoordCiudad.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordEstado) Then txtCoordEstado.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordSubTotal) Then txtCoordSubTotal.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordDesctos) Then txtCoordDesctos.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordIVA) Then txtCoordIVA.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordTotal) Then txtCoordTotal.Text = "0,0"
        If Not BuscaCoordenadas(txtImpteConLetra) Then txtImpteConLetra.Text = "0,0"
        If Not BuscaCoordenadas(txtCoordLugarExped) Then txtCoordLugarExped.Text = "0,0"
        If Not BuscaCoordenadas(TxtCoordLeyenda) Then TxtCoordLeyenda.Text = "0,0"
        If txtRenXdetalle.Text = "" Then txtRenXdetalle.Text = CStr(0)
        If txtPrimerPartida.Text = "" Then txtPrimerPartida.Text = CStr(0)
        If txtCoordCantidad.Text = "" Then txtCoordCantidad.Text = CStr(0)
        If txtCoordCodigo.Text = "" Then txtCoordCodigo.Text = CStr(0)
        If txtColDescripcion.Text = "" Then txtColDescripcion.Text = CStr(0)
        If txtColDesctoDetalle.Text = "" Then txtColDesctoDetalle.Text = CStr(0)
        If txtCoordPromocion.Text = "" Then txtCoordPromocion.Text = CStr(0)
        If txtColIVAporPartida.Text = "" Then txtColIVAporPartida.Text = CStr(0)
        If txtCoordPrecioVta.Text = "" Then txtCoordPrecioVta.Text = CStr(0)
        If txtCoordImporte.Text = "" Then txtCoordImporte.Text = CStr(0)

        frmPVConfigLongitudDeDatosFactura.InitializeComponent()
        With frmPVConfigLongitudDeDatosFactura.FlexDetalle
            .Col = 1
            .Row = 1
            If Trim(.Text) = "" Then .Text = CStr(0)
            LonCliente = CInt(.Text)
            .Row = 2
            If Trim(.Text) = "" Then .Text = CStr(0)
            LonDireccion = CInt(.Text)
            .Row = 3
            If Trim(.Text) = "" Then .Text = CStr(0)
            LonColonia = CInt(.Text)
            .Row = 4
            If Trim(.Text) = "" Then .Text = CStr(0)
            LonCiudad = CInt(.Text)
            .Row = 5
            If Trim(.Text) = "" Then .Text = CStr(0)
            LonEstado = CInt(.Text)
            .Row = 6
            If Trim(.Text) = "" Then .Text = CStr(0)
            LonLeyenda = CInt(.Text)
            .Row = 7
            If Trim(.Text) = "" Then .Text = CStr(0)
            LonProducto = CInt(.Text)
        End With
        If Val(txtRenXFactura.Text) = 0 Or Val(BuscaRenglon(txtCoordFolio)) = 0 Or Val(BuscaColumna(txtCoordFolio)) = 0 Or Val(BuscaRenglon(txtCoordFecha)) = 0 Or Val(BuscaColumna(txtCoordFecha)) = 0 Or Val(BuscaRenglon(txtCoordEmpresa)) = 0 Or Val(BuscaColumna(txtCoordEmpresa)) = 0 Or Val(BuscaRenglon(txtCoordRFC)) = 0 Or Val(BuscaColumna(txtCoordRFC)) = 0 Or Val(BuscaRenglon(txtCoordCalle)) = 0 Or Val(BuscaColumna(txtCoordCalle)) = 0 Or Val(BuscaRenglon(txtCoordColonia)) = 0 Or Val(BuscaColumna(txtCoordColonia)) = 0 Or Val(BuscaRenglon(txtCoordCP)) = 0 Or Val(BuscaColumna(txtCoordCP)) = 0 Or Val(BuscaRenglon(txtCoordTelefono)) = 0 Or Val(BuscaColumna(txtCoordTelefono)) = 0 Or Val(BuscaRenglon(txtCoordCiudad)) = 0 Or Val(BuscaColumna(txtCoordCiudad)) = 0 Or Val(BuscaRenglon(txtCoordEstado)) = 0 Or Val(BuscaColumna(txtCoordEstado)) = 0 Or Val(BuscaRenglon(txtCoordSubTotal)) = 0 Or Val(BuscaColumna(txtCoordSubTotal)) = 0 Or Val(BuscaRenglon(txtCoordDesctos)) = 0 Or Val(BuscaColumna(txtCoordDesctos)) = 0 Or Val(BuscaRenglon(txtCoordIVA)) = 0 Or Val(BuscaColumna(txtCoordIVA)) = 0 Or Val(BuscaRenglon(txtCoordTotal)) = 0 Or Val(BuscaColumna(txtCoordTotal)) = 0 Or Val(BuscaRenglon(txtImpteConLetra)) = 0 Or Val(BuscaColumna(txtImpteConLetra)) = 0 Or Val(BuscaRenglon(txtCoordLugarExped)) = 0 Or Val(BuscaColumna(txtCoordLugarExped)) = 0 Or Val(BuscaRenglon(TxtCoordLeyenda)) = 0 Or Val(BuscaColumna(TxtCoordLeyenda)) = 0 Or Val(txtRenXdetalle.Text) = 0 Or Val(txtPrimerPartida.Text) = 0 Or Val(txtCoordCantidad.Text) = 0 Or Val(txtCoordCodigo.Text) = 0 Or Val(txtColDescripcion.Text) = 0 Or Val(txtColDesctoDetalle.Text) = 0 Or Val(txtCoordPromocion.Text) = 0 Or Val(txtColIVAporPartida.Text) = 0 Or Val(txtCoordPrecioVta.Text) = 0 Or Val(txtCoordImporte.Text) = 0 Or LonCliente = 0 Or LonDireccion = 0 Or LonColonia = 0 Or LonCiudad = 0 Or LonEstado = 0 Or LonLeyenda = 0 Or LonProducto = 0 Then
            MsgBox("Algunos Parametros de Configuración de la Factura No Han Sido Llenados" & Chr(13) & "La Configuración Se Guardara Con Los Parametros Que Se Han Indicado ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        End If
    End Sub

    Private Sub btnLong_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnLong.Click
        'frmPVConfigLongitudDeDatosFactura.InitializeComponent()
        frmPVConfigLongitudDeDatosFactura.ShowDialog()
    End Sub

    Private Sub btnLong_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnLong.Enter
        Pon_Tool()
    End Sub

    Private Sub CboLetra_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLetra.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursales.Name Then
        '    Exit Sub
        'End If
        Nuevo()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        If dbcSucursales.SelectedItem <> "" Then
            Call dbcSucursales_Leave(dbcSucursales, New System.EventArgs())
        End If
        mblnNuevo = True
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursales.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen Where TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSALIR = True
            Me.Close()
            '    ElseIf KeyCode = vbKeyReturn Then
            '        cboLetra.SetFocus
        End If
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        intCodSucursal = 0
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' And TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        LlenaDatos()
        '    ValidarSucursalyMostrarDatos
    End Sub

    Private Sub frmPVConfigFacturacion_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmPVConfigFacturacion_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmPVConfigFacturacion_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optTicket" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSALIR = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmPVConfigFacturacion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmPVConfigFacturacion_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ModEstandar.CentrarForma(Me)
        Nuevo()
        InicializaVariables()
        '     gStrSql = "SELECT  CodAlmacen, DescAlmacen FROM CatAlmacen  " & _
        ''            "Where codAlmacen in(select min(CodAlmacen)  from CatAlmacen)"
        '    ModEstandar.BorraCmd
        '    Cmd.CommandText = "dbo.Up_Select_Datos"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        '    Set RsAux = Cmd.Execute
        '
        '    If RsAux.RecordCount > 0 Then
        '        dbcSucursales = Trim(RsAux!DescAlmacen)
        '        dbcSucursales_LostFocus
        '    End If
    End Sub

    Private Sub frmPVConfigFacturacion_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSALIR And Not mblnSaliryGrabar Then
        '    If Cambios() = True Then
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
        '            mblnSALIR = False
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmPVConfigFacturacion_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        frmPVConfigLongitudDeDatosFactura.Close()
    End Sub

    Private Sub _optTicket_0_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optTicket_0.Enter
        Dim Index As Integer
        '= _optTicket_0.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub txtColDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtColDescripcion.Enter
        SelTextoTxt(txtColDescripcion)
        Pon_Tool()
    End Sub

    Private Sub txtColDescripcion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtColDescripcion.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtColDesctoDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtColDesctoDetalle.Enter
        SelTextoTxt(txtColDesctoDetalle)
        Pon_Tool()
    End Sub

    Private Sub txtColDesctoDetalle_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtColDesctoDetalle.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtColIVAporPartida_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtColIVAporPartida.Enter
        SelTextoTxt(txtColIVAporPartida)
        Pon_Tool()
    End Sub

    Private Sub txtColIVAporPartida_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtColIVAporPartida.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordCalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordCalle.Enter
        SelTextoTxt(txtCoordCalle)
        Pon_Tool()
    End Sub

    Private Sub txtCoordCalle_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordCalle.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordCalle.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordCalle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordCalle.Leave
        If txtCoordCalle.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordCalle) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordCalle.Text = ""
        End If
    End Sub

    Private Sub txtCoordCantidad_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordCantidad.Enter
        SelTextoTxt(txtCoordCantidad)
        Pon_Tool()
    End Sub

    Private Sub txtCoordCantidad_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordCantidad.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordCiudad_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordCiudad.Enter
        SelTextoTxt(txtCoordCiudad)
        Pon_Tool()
    End Sub

    Private Sub txtCoordCiudad_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordCiudad.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordCiudad.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordCiudad_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordCiudad.Leave
        If txtCoordCiudad.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordCiudad) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordCiudad.Text = ""
        End If
    End Sub

    Private Sub txtCoordCodigo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordCodigo.Enter
        SelTextoTxt(txtCoordCodigo)
        Pon_Tool()
    End Sub

    Private Sub txtCoordCodigo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordCodigo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordColonia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordColonia.Enter
        SelTextoTxt(txtCoordColonia)
        Pon_Tool()
    End Sub

    Private Sub txtCoordColonia_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordColonia.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordColonia.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordColonia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordColonia.Leave
        If txtCoordColonia.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordColonia) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordColonia.Text = ""
        End If
    End Sub

    Private Sub txtCoordCP_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordCP.Enter
        SelTextoTxt(txtCoordCP)
        Pon_Tool()
    End Sub

    Private Sub txtCoordCP_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordCP.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordCP.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordCP_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordCP.Leave
        If txtCoordCP.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordCP) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordCP.Text = ""
        End If
    End Sub

    Private Sub txtCoordDesctos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordDesctos.Enter
        SelTextoTxt(txtCoordDesctos)
        Pon_Tool()
    End Sub

    Private Sub txtCoordDesctos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordDesctos.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordDesctos.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordDesctos_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordDesctos.Leave
        If txtCoordDesctos.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordDesctos) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordDesctos.Text = ""
        End If
    End Sub

    Private Sub txtCoordEmpresa_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordEmpresa.Enter
        SelTextoTxt(txtCoordEmpresa)
        Pon_Tool()
    End Sub

    Private Sub txtCoordEmpresa_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordEmpresa.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordEmpresa.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordEmpresa_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordEmpresa.Leave
        If txtCoordEmpresa.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordEmpresa) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordEmpresa.Text = ""
        End If
    End Sub

    Private Sub txtCoordEstado_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordEstado.Enter
        SelTextoTxt(txtCoordEstado)
        Pon_Tool()
    End Sub

    Private Sub txtCoordEstado_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordEstado.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordEstado.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordEstado_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordEstado.Leave
        If txtCoordEstado.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordEstado) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordEstado.Text = ""
        End If
    End Sub

    Private Sub txtCoordFecha_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordFecha.Enter
        SelTextoTxt(txtCoordFecha)
        Pon_Tool()
    End Sub

    Private Sub txtCoordFecha_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordFecha.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordFecha.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordFecha_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordFecha.Leave
        If txtCoordFecha.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordFecha) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordFecha.Text = ""
        End If
    End Sub

    Private Sub txtCoordFolio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordFolio.Enter
        SelTextoTxt(txtCoordFolio)
        Pon_Tool()
    End Sub

    Private Sub txtCoordFolio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordFolio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordFolio.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordFolio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordFolio.Leave
        If txtCoordFolio.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordFolio) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordFolio.Text = ""
        End If
    End Sub

    Private Sub txtCoordImporte_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordImporte.Enter
        SelTextoTxt(txtCoordImporte)
        Pon_Tool()
    End Sub

    Private Sub txtCoordImporte_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordImporte.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordIVA_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordIVA.Enter
        SelTextoTxt(txtCoordIVA)
        Pon_Tool()
    End Sub

    Private Sub txtCoordIVA_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordIVA.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordIVA.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordIVA_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordIVA.Leave
        If txtCoordIVA.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordIVA) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordIVA.Text = ""
        End If
    End Sub

    Private Sub TxtCoordLeyenda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtCoordLeyenda.Enter
        SelTextoTxt(TxtCoordLeyenda)
        Pon_Tool()
    End Sub

    Private Sub TxtCoordLeyenda_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtCoordLeyenda.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(TxtCoordLeyenda.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub TxtCoordLeyenda_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtCoordLeyenda.Leave
        If TxtCoordLeyenda.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(TxtCoordLeyenda) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            TxtCoordLeyenda.Text = ""
        End If
    End Sub

    Private Sub txtCoordLugarExped_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordLugarExped.Enter
        SelTextoTxt(txtCoordLugarExped)
        Pon_Tool()
    End Sub

    Private Sub txtCoordLugarExped_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordLugarExped.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordLugarExped.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordLugarExped_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordLugarExped.Leave
        If txtCoordLugarExped.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordLugarExped) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordLugarExped.Text = ""
        End If
    End Sub

    Private Sub txtCoordPrecioVta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordPrecioVta.Enter
        SelTextoTxt(txtCoordPrecioVta)
        Pon_Tool()
    End Sub

    Private Sub txtCoordPrecioVta_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordPrecioVta.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordPromocion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordPromocion.Enter
        SelTextoTxt(txtCoordPromocion)
        Pon_Tool()
    End Sub

    Private Sub txtCoordPromocion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordPromocion.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordRFC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordRFC.Enter
        SelTextoTxt(txtCoordRFC)
        Pon_Tool()
    End Sub

    Private Sub txtCoordRFC_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordRFC.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordRFC.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordRFC_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordRFC.Leave
        If txtCoordRFC.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordRFC) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordRFC.Text = ""
        End If
    End Sub

    Private Sub txtCoordSubTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordSubTotal.Enter
        SelTextoTxt(txtCoordSubTotal)
        Pon_Tool()
    End Sub

    Private Sub txtCoordSubTotal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordSubTotal.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordSubTotal.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordSubTotal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordSubTotal.Leave
        If txtCoordSubTotal.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordSubTotal) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordSubTotal.Text = ""
        End If
    End Sub

    Private Sub txtCoordTelefono_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordTelefono.Enter
        SelTextoTxt(txtCoordTelefono)
        Pon_Tool()
    End Sub

    Private Sub txtCoordTelefono_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordTelefono.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordTelefono.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordTelefono_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordTelefono.Leave
        If txtCoordTelefono.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordTelefono) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordTelefono.Text = ""
        End If
    End Sub

    Private Sub txtCoordTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordTotal.Enter
        SelTextoTxt(txtCoordTotal)
        Pon_Tool()
    End Sub

    Private Sub txtCoordTotal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCoordTotal.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtCoordTotal.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCoordTotal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoordTotal.Leave
        If txtCoordTotal.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtCoordTotal) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCoordTotal.Text = ""
        End If
    End Sub

    Private Sub txtImpteConLetra_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImpteConLetra.Enter
        SelTextoTxt(txtImpteConLetra)
        Pon_Tool()
    End Sub

    Private Sub txtImpteConLetra_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImpteConLetra.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Valida_Coordenadas(txtImpteConLetra.Text, KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtImpteConLetra_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImpteConLetra.Leave
        If txtImpteConLetra.Text = "" Then Exit Sub
        If Not BuscaCoordenadas(txtImpteConLetra) Then
            MsgBox(C_msgCOORDNOVALIDA, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtImpteConLetra.Text = ""
        End If
    End Sub

    Private Sub TxtLeyenda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeyenda.Enter
        SelTextoTxt(txtLeyenda)
        Pon_Tool()
    End Sub

    Private Sub txtPrimerPartida_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrimerPartida.Enter
        SelTextoTxt(txtPrimerPartida)
        Pon_Tool()
    End Sub

    Private Sub txtPrimerPartida_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPrimerPartida.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtRenXdetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRenXdetalle.Enter
        SelTextoTxt(txtRenXdetalle)
        Pon_Tool()
    End Sub

    Private Sub txtRenXdetalle_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRenXdetalle.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtRenXFactura_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRenXFactura.Enter
        SelTextoTxt(txtRenXFactura)
        Pon_Tool()
    End Sub

    Private Sub txtRenXFactura_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRenXFactura.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub Nuevo()
        cboLetra.SelectedIndex = 0
        cboLetra.Tag = cboLetra.Text
        txtRenXFactura.Text = ""
        txtRenXFactura.Tag = ""
        txtCoordFolio.Text = ""
        txtCoordFolio.Tag = ""
        txtCoordFecha.Text = ""
        txtCoordFecha.Tag = ""
        txtCoordEmpresa.Text = ""
        txtCoordEmpresa.Tag = ""
        txtCoordRFC.Text = ""
        txtCoordRFC.Tag = ""
        txtCoordCalle.Text = ""
        txtCoordCalle.Tag = ""
        txtCoordColonia.Text = ""
        txtCoordColonia.Tag = ""
        txtCoordCP.Text = ""
        txtCoordCP.Tag = ""
        txtCoordTelefono.Text = ""
        txtCoordTelefono.Tag = ""
        txtCoordCiudad.Text = ""
        txtCoordCiudad.Tag = ""
        txtCoordEstado.Text = ""
        txtCoordEstado.Tag = ""
        txtCoordSubTotal.Text = ""
        txtCoordSubTotal.Tag = ""
        txtCoordDesctos.Text = ""
        txtCoordDesctos.Tag = ""
        txtCoordIVA.Text = ""
        txtCoordIVA.Tag = ""
        txtCoordTotal.Text = ""
        txtCoordTotal.Tag = ""
        txtCoordImporte.Text = ""
        txtCoordImporte.Tag = ""
        txtCoordLugarExped.Text = ""
        txtCoordLugarExped.Tag = ""
        TxtCoordLeyenda.Text = ""
        TxtCoordLeyenda.Tag = ""
        txtRenXdetalle.Text = ""
        txtRenXdetalle.Tag = ""
        txtPrimerPartida.Text = ""
        txtPrimerPartida.Tag = ""
        txtCoordCantidad.Text = ""
        txtCoordCantidad.Tag = ""
        txtCoordCodigo.Text = ""
        txtCoordCodigo.Tag = ""
        txtColDescripcion.Text = ""
        txtColDescripcion.Tag = ""
        txtColDesctoDetalle.Text = ""
        txtColDesctoDetalle.Tag = ""
        txtCoordPromocion.Text = ""
        txtCoordPromocion.Tag = ""
        txtColIVAporPartida.Text = ""
        txtColIVAporPartida.Tag = ""
        txtCoordPrecioVta.Text = ""
        txtCoordPrecioVta.Tag = ""
        txtCoordImporte.Text = ""
        txtCoordImporte.Tag = ""
        txtImpteConLetra.Text = ""
        txtImpteConLetra.Tag = ""
        txtLeyenda.Text = ""
        txtLeyenda.Tag = ""
        chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked
        gintLonCiudad = 0
        gintLonCliente = 0
        gintLonColonia = 0
        gintLonDescProducto = 0
        gintLonDireccion = 0
        gintLonEstado = 0
        gintLonLeyenda = 0
        'frmPVConfigLongitudDeDatosFactura.FlexDetalle.Clear()
        'frmPVConfigLongitudDeDatosFactura.Encabezado()
    End Sub

    Sub Limpiar()
        Nuevo()
        dbcSucursales.Text = ""
        dbcSucursales.Focus()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._optTicket_0 = New System.Windows.Forms.RadioButton()
        Me._optTicket_1 = New System.Windows.Forms.RadioButton()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me.btnLong = New System.Windows.Forms.Button()
        Me.txtLeyenda = New System.Windows.Forms.TextBox()
        Me.cboLetra = New System.Windows.Forms.ComboBox()
        Me.TxtCoordLeyenda = New System.Windows.Forms.TextBox()
        Me.txtColIVAporPartida = New System.Windows.Forms.TextBox()
        Me.txtColDesctoDetalle = New System.Windows.Forms.TextBox()
        Me.txtColDescripcion = New System.Windows.Forms.TextBox()
        Me.txtRenXdetalle = New System.Windows.Forms.TextBox()
        Me.txtCoordLugarExped = New System.Windows.Forms.TextBox()
        Me.txtCoordCP = New System.Windows.Forms.TextBox()
        Me.txtCoordFolio = New System.Windows.Forms.TextBox()
        Me.txtImpteConLetra = New System.Windows.Forms.TextBox()
        Me.txtPrimerPartida = New System.Windows.Forms.TextBox()
        Me.txtPartidasXfactura = New System.Windows.Forms.TextBox()
        Me.txtCoordTotal = New System.Windows.Forms.TextBox()
        Me.txtCoordIVA = New System.Windows.Forms.TextBox()
        Me.txtCoordSubTotal = New System.Windows.Forms.TextBox()
        Me.txtCoordImporte = New System.Windows.Forms.TextBox()
        Me.txtCoordPrecioVta = New System.Windows.Forms.TextBox()
        Me.txtCoordPromocion = New System.Windows.Forms.TextBox()
        Me.txtCoordDesctos = New System.Windows.Forms.TextBox()
        Me.txtCoordCantidad = New System.Windows.Forms.TextBox()
        Me.txtCoordCodigo = New System.Windows.Forms.TextBox()
        Me.txtCoordTelefono = New System.Windows.Forms.TextBox()
        Me.txtCoordEstado = New System.Windows.Forms.TextBox()
        Me.txtCoordCiudad = New System.Windows.Forms.TextBox()
        Me.txtCoordColonia = New System.Windows.Forms.TextBox()
        Me.txtCoordCalle = New System.Windows.Forms.TextBox()
        Me.txtCoordFecha = New System.Windows.Forms.TextBox()
        Me.txtCoordRFC = New System.Windows.Forms.TextBox()
        Me.txtCoordEmpresa = New System.Windows.Forms.TextBox()
        Me.txtRenXFactura = New System.Windows.Forms.TextBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me.chkAplicarSucursales = New System.Windows.Forms.CheckBox()
        Me.Marco = New System.Windows.Forms.GroupBox()
        Me._Line1_1 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me._Line1_0 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_100 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_0 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDatosPartida = New System.Windows.Forms.Label()
        Me.TxtDatosGenerales = New System.Windows.Forms.Label()
        Me._Label2_28 = New System.Windows.Forms.Label()
        Me._Label2_27 = New System.Windows.Forms.Label()
        Me._Label2_26 = New System.Windows.Forms.Label()
        Me._Label2_25 = New System.Windows.Forms.Label()
        Me._Label2_24 = New System.Windows.Forms.Label()
        Me._Label2_158 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_153 = New System.Windows.Forms.Label()
        Me._Label2_21 = New System.Windows.Forms.Label()
        Me._Label2_20 = New System.Windows.Forms.Label()
        Me._Label2_19 = New System.Windows.Forms.Label()
        Me._Label2_18 = New System.Windows.Forms.Label()
        Me._Label2_16 = New System.Windows.Forms.Label()
        Me._Label2_15 = New System.Windows.Forms.Label()
        Me._Label2_160 = New System.Windows.Forms.Label()
        Me._Label2_161 = New System.Windows.Forms.Label()
        Me._Label2_12 = New System.Windows.Forms.Label()
        Me._Label2_11 = New System.Windows.Forms.Label()
        Me._Label2_10 = New System.Windows.Forms.Label()
        Me._Label2_9 = New System.Windows.Forms.Label()
        Me._Label2_159 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_157 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_156 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_155 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_154 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_152 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_151 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_150 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_1 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.mskFecha = New System.Windows.Forms.MaskedTextBox()
        Me.mskHora = New System.Windows.Forms.MaskedTextBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Line1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblEtiqueta = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optTicket = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Marco.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblEtiqueta, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optTicket, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        '_optTicket_0
        '
        Me._optTicket_0.BackColor = System.Drawing.SystemColors.Control
        Me._optTicket_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTicket_0.ForeColor = System.Drawing.Color.Black
        Me._optTicket_0.Location = New System.Drawing.Point(376, 20)
        Me._optTicket_0.Name = "_optTicket_0"
        Me._optTicket_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTicket_0.Size = New System.Drawing.Size(70, 17)
        Me._optTicket_0.TabIndex = 76
        Me._optTicket_0.TabStop = True
        Me._optTicket_0.Text = "Ticket"
        Me.ToolTip1.SetToolTip(Me._optTicket_0, "Tipo de Facturación.")
        Me._optTicket_0.UseVisualStyleBackColor = False
        '
        '_optTicket_1
        '
        Me._optTicket_1.BackColor = System.Drawing.SystemColors.Control
        Me._optTicket_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTicket_1.ForeColor = System.Drawing.Color.Black
        Me._optTicket_1.Location = New System.Drawing.Point(552, 20)
        Me._optTicket_1.Name = "_optTicket_1"
        Me._optTicket_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTicket_1.Size = New System.Drawing.Size(78, 17)
        Me._optTicket_1.TabIndex = 78
        Me._optTicket_1.TabStop = True
        Me._optTicket_1.Text = "Formato"
        Me.ToolTip1.SetToolTip(Me._optTicket_1, "Tipo de Facturación.")
        Me._optTicket_1.UseVisualStyleBackColor = False
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.ForeColor = System.Drawing.Color.Black
        Me._Label1_2.Location = New System.Drawing.Point(16, 16)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(60, 17)
        Me._Label1_2.TabIndex = 2
        Me._Label1_2.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_2, "Nombre de la Farmacia Actual")
        '
        'btnLong
        '
        Me.btnLong.BackColor = System.Drawing.SystemColors.Control
        Me.btnLong.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLong.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLong.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLong.Location = New System.Drawing.Point(343, 265)
        Me.btnLong.Name = "btnLong"
        Me.btnLong.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLong.Size = New System.Drawing.Size(57, 21)
        Me.btnLong.TabIndex = 45
        Me.btnLong.Text = ". . ."
        Me.ToolTip1.SetToolTip(Me.btnLong, "Longitudes de Datos")
        Me.btnLong.UseVisualStyleBackColor = False
        '
        'txtLeyenda
        '
        Me.txtLeyenda.AcceptsReturn = True
        Me.txtLeyenda.BackColor = System.Drawing.SystemColors.Window
        Me.txtLeyenda.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLeyenda.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLeyenda.Location = New System.Drawing.Point(140, 304)
        Me.txtLeyenda.MaxLength = 200
        Me.txtLeyenda.Name = "txtLeyenda"
        Me.txtLeyenda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLeyenda.Size = New System.Drawing.Size(501, 20)
        Me.txtLeyenda.TabIndex = 68
        Me.ToolTip1.SetToolTip(Me.txtLeyenda, "Leyenda de Factura")
        '
        'cboLetra
        '
        Me.cboLetra.BackColor = System.Drawing.SystemColors.Window
        Me.cboLetra.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLetra.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLetra.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLetra.Items.AddRange(New Object() {"8", "9", "10", "11", "12", "14", "16", "18", "24"})
        Me.cboLetra.Location = New System.Drawing.Point(140, 51)
        Me.cboLetra.Name = "cboLetra"
        Me.cboLetra.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLetra.Size = New System.Drawing.Size(57, 21)
        Me.cboLetra.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.cboLetra, "Tamaño de Letra a Imprimir")
        '
        'TxtCoordLeyenda
        '
        Me.TxtCoordLeyenda.AcceptsReturn = True
        Me.TxtCoordLeyenda.BackColor = System.Drawing.SystemColors.Window
        Me.TxtCoordLeyenda.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtCoordLeyenda.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtCoordLeyenda.Location = New System.Drawing.Point(343, 241)
        Me.TxtCoordLeyenda.MaxLength = 7
        Me.TxtCoordLeyenda.Name = "TxtCoordLeyenda"
        Me.TxtCoordLeyenda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtCoordLeyenda.Size = New System.Drawing.Size(57, 20)
        Me.TxtCoordLeyenda.TabIndex = 44
        Me.TxtCoordLeyenda.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.TxtCoordLeyenda, "Coord. para imprimir Leyenda Factura (Ren, Col).")
        '
        'txtColIVAporPartida
        '
        Me.txtColIVAporPartida.AcceptsReturn = True
        Me.txtColIVAporPartida.BackColor = System.Drawing.SystemColors.Window
        Me.txtColIVAporPartida.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtColIVAporPartida.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtColIVAporPartida.Location = New System.Drawing.Point(584, 217)
        Me.txtColIVAporPartida.MaxLength = 4
        Me.txtColIVAporPartida.Name = "txtColIVAporPartida"
        Me.txtColIVAporPartida.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtColIVAporPartida.Size = New System.Drawing.Size(57, 20)
        Me.txtColIVAporPartida.TabIndex = 64
        Me.txtColIVAporPartida.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtColIVAporPartida, "Columna para Imprimir IVA")
        '
        'txtColDesctoDetalle
        '
        Me.txtColDesctoDetalle.AcceptsReturn = True
        Me.txtColDesctoDetalle.BackColor = System.Drawing.SystemColors.Window
        Me.txtColDesctoDetalle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtColDesctoDetalle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtColDesctoDetalle.Location = New System.Drawing.Point(584, 169)
        Me.txtColDesctoDetalle.MaxLength = 4
        Me.txtColDesctoDetalle.Name = "txtColDesctoDetalle"
        Me.txtColDesctoDetalle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtColDesctoDetalle.Size = New System.Drawing.Size(57, 20)
        Me.txtColDesctoDetalle.TabIndex = 62
        Me.txtColDesctoDetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtColDesctoDetalle, "Columna para Imprimir Descuento")
        '
        'txtColDescripcion
        '
        Me.txtColDescripcion.AcceptsReturn = True
        Me.txtColDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.txtColDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtColDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtColDescripcion.Location = New System.Drawing.Point(584, 145)
        Me.txtColDescripcion.MaxLength = 4
        Me.txtColDescripcion.Name = "txtColDescripcion"
        Me.txtColDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtColDescripcion.Size = New System.Drawing.Size(57, 20)
        Me.txtColDescripcion.TabIndex = 61
        Me.txtColDescripcion.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtColDescripcion, "Columna para Imprimir Descripción de Producto")
        '
        'txtRenXdetalle
        '
        Me.txtRenXdetalle.AcceptsReturn = True
        Me.txtRenXdetalle.BackColor = System.Drawing.SystemColors.Window
        Me.txtRenXdetalle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRenXdetalle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRenXdetalle.Location = New System.Drawing.Point(584, 51)
        Me.txtRenXdetalle.MaxLength = 4
        Me.txtRenXdetalle.Name = "txtRenXdetalle"
        Me.txtRenXdetalle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRenXdetalle.Size = New System.Drawing.Size(57, 20)
        Me.txtRenXdetalle.TabIndex = 57
        Me.txtRenXdetalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtRenXdetalle, "Total de Partidas de la Factura")
        '
        'txtCoordLugarExped
        '
        Me.txtCoordLugarExped.AcceptsReturn = True
        Me.txtCoordLugarExped.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordLugarExped.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordLugarExped.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordLugarExped.Location = New System.Drawing.Point(343, 217)
        Me.txtCoordLugarExped.MaxLength = 7
        Me.txtCoordLugarExped.Name = "txtCoordLugarExped"
        Me.txtCoordLugarExped.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordLugarExped.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordLugarExped.TabIndex = 43
        Me.txtCoordLugarExped.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordLugarExped, "Coord. para Imprimir Lugar de Expedición (Ren, Col).")
        '
        'txtCoordCP
        '
        Me.txtCoordCP.AcceptsReturn = True
        Me.txtCoordCP.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordCP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordCP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordCP.Location = New System.Drawing.Point(140, 241)
        Me.txtCoordCP.MaxLength = 7
        Me.txtCoordCP.Name = "txtCoordCP"
        Me.txtCoordCP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordCP.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordCP.TabIndex = 24
        Me.txtCoordCP.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordCP, "Coord. para Imprimir Código Postal (Ren, Col).")
        '
        'txtCoordFolio
        '
        Me.txtCoordFolio.AcceptsReturn = True
        Me.txtCoordFolio.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordFolio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordFolio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordFolio.Location = New System.Drawing.Point(140, 97)
        Me.txtCoordFolio.MaxLength = 7
        Me.txtCoordFolio.Name = "txtCoordFolio"
        Me.txtCoordFolio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordFolio.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordFolio.TabIndex = 18
        Me.txtCoordFolio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordFolio, "Coord. para Imprimir Folio (Ren, Col).")
        '
        'txtImpteConLetra
        '
        Me.txtImpteConLetra.AcceptsReturn = True
        Me.txtImpteConLetra.BackColor = System.Drawing.SystemColors.Window
        Me.txtImpteConLetra.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImpteConLetra.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImpteConLetra.Location = New System.Drawing.Point(343, 193)
        Me.txtImpteConLetra.MaxLength = 7
        Me.txtImpteConLetra.Name = "txtImpteConLetra"
        Me.txtImpteConLetra.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImpteConLetra.Size = New System.Drawing.Size(57, 20)
        Me.txtImpteConLetra.TabIndex = 42
        Me.txtImpteConLetra.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtImpteConLetra, "Coord. para Imprimir Importe con Letra (Ren, Col).")
        '
        'txtPrimerPartida
        '
        Me.txtPrimerPartida.AcceptsReturn = True
        Me.txtPrimerPartida.BackColor = System.Drawing.SystemColors.Window
        Me.txtPrimerPartida.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrimerPartida.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrimerPartida.Location = New System.Drawing.Point(584, 74)
        Me.txtPrimerPartida.MaxLength = 4
        Me.txtPrimerPartida.Name = "txtPrimerPartida"
        Me.txtPrimerPartida.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrimerPartida.Size = New System.Drawing.Size(57, 20)
        Me.txtPrimerPartida.TabIndex = 58
        Me.txtPrimerPartida.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtPrimerPartida, "Renglón para imprimir la primer Partida")
        '
        'txtPartidasXfactura
        '
        Me.txtPartidasXfactura.AcceptsReturn = True
        Me.txtPartidasXfactura.BackColor = System.Drawing.SystemColors.Window
        Me.txtPartidasXfactura.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPartidasXfactura.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPartidasXfactura.Location = New System.Drawing.Point(668, 297)
        Me.txtPartidasXfactura.MaxLength = 4
        Me.txtPartidasXfactura.Name = "txtPartidasXfactura"
        Me.txtPartidasXfactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPartidasXfactura.Size = New System.Drawing.Size(57, 20)
        Me.txtPartidasXfactura.TabIndex = 69
        Me.txtPartidasXfactura.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtPartidasXfactura, "Productos totales por Factura.")
        Me.txtPartidasXfactura.Visible = False
        '
        'txtCoordTotal
        '
        Me.txtCoordTotal.AcceptsReturn = True
        Me.txtCoordTotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordTotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordTotal.Location = New System.Drawing.Point(343, 169)
        Me.txtCoordTotal.MaxLength = 7
        Me.txtCoordTotal.Name = "txtCoordTotal"
        Me.txtCoordTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordTotal.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordTotal.TabIndex = 41
        Me.txtCoordTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordTotal, "Coord. para Imprimir Total (Ren, Col).")
        '
        'txtCoordIVA
        '
        Me.txtCoordIVA.AcceptsReturn = True
        Me.txtCoordIVA.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordIVA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordIVA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordIVA.Location = New System.Drawing.Point(343, 145)
        Me.txtCoordIVA.MaxLength = 7
        Me.txtCoordIVA.Name = "txtCoordIVA"
        Me.txtCoordIVA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordIVA.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordIVA.TabIndex = 40
        Me.txtCoordIVA.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordIVA, "Coord. para Imprimir IVA (Ren, Col).")
        '
        'txtCoordSubTotal
        '
        Me.txtCoordSubTotal.AcceptsReturn = True
        Me.txtCoordSubTotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordSubTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordSubTotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordSubTotal.Location = New System.Drawing.Point(343, 97)
        Me.txtCoordSubTotal.MaxLength = 7
        Me.txtCoordSubTotal.Name = "txtCoordSubTotal"
        Me.txtCoordSubTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordSubTotal.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordSubTotal.TabIndex = 38
        Me.txtCoordSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordSubTotal, "Coord. para Imprimir SubTotal (Ren, Col).")
        '
        'txtCoordImporte
        '
        Me.txtCoordImporte.AcceptsReturn = True
        Me.txtCoordImporte.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordImporte.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordImporte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordImporte.Location = New System.Drawing.Point(584, 265)
        Me.txtCoordImporte.MaxLength = 4
        Me.txtCoordImporte.Name = "txtCoordImporte"
        Me.txtCoordImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordImporte.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordImporte.TabIndex = 66
        Me.txtCoordImporte.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordImporte, "Columna para Imprimir Importe")
        '
        'txtCoordPrecioVta
        '
        Me.txtCoordPrecioVta.AcceptsReturn = True
        Me.txtCoordPrecioVta.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordPrecioVta.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordPrecioVta.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordPrecioVta.Location = New System.Drawing.Point(584, 241)
        Me.txtCoordPrecioVta.MaxLength = 4
        Me.txtCoordPrecioVta.Name = "txtCoordPrecioVta"
        Me.txtCoordPrecioVta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordPrecioVta.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordPrecioVta.TabIndex = 65
        Me.txtCoordPrecioVta.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordPrecioVta, "Columna para Imprimir Precio de Venta")
        '
        'txtCoordPromocion
        '
        Me.txtCoordPromocion.AcceptsReturn = True
        Me.txtCoordPromocion.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordPromocion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordPromocion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordPromocion.Location = New System.Drawing.Point(584, 193)
        Me.txtCoordPromocion.MaxLength = 4
        Me.txtCoordPromocion.Name = "txtCoordPromocion"
        Me.txtCoordPromocion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordPromocion.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordPromocion.TabIndex = 63
        Me.txtCoordPromocion.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordPromocion, "Columna para Imprimir Promoción")
        '
        'txtCoordDesctos
        '
        Me.txtCoordDesctos.AcceptsReturn = True
        Me.txtCoordDesctos.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordDesctos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordDesctos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordDesctos.Location = New System.Drawing.Point(343, 121)
        Me.txtCoordDesctos.MaxLength = 7
        Me.txtCoordDesctos.Name = "txtCoordDesctos"
        Me.txtCoordDesctos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordDesctos.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordDesctos.TabIndex = 39
        Me.txtCoordDesctos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordDesctos, "Coord. para Imprimir Descuento (Ren, Col).")
        '
        'txtCoordCantidad
        '
        Me.txtCoordCantidad.AcceptsReturn = True
        Me.txtCoordCantidad.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordCantidad.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordCantidad.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordCantidad.Location = New System.Drawing.Point(584, 97)
        Me.txtCoordCantidad.MaxLength = 4
        Me.txtCoordCantidad.Name = "txtCoordCantidad"
        Me.txtCoordCantidad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordCantidad.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordCantidad.TabIndex = 59
        Me.txtCoordCantidad.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordCantidad, "Columna para Imprimir Cantidad")
        '
        'txtCoordCodigo
        '
        Me.txtCoordCodigo.AcceptsReturn = True
        Me.txtCoordCodigo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordCodigo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordCodigo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordCodigo.Location = New System.Drawing.Point(584, 121)
        Me.txtCoordCodigo.MaxLength = 4
        Me.txtCoordCodigo.Name = "txtCoordCodigo"
        Me.txtCoordCodigo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordCodigo.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtCoordCodigo.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordCodigo.TabIndex = 60
        Me.txtCoordCodigo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordCodigo, "Columna para Imprimir Código de Producto")
        Me.txtCoordCodigo.WordWrap = False
        '
        'txtCoordTelefono
        '
        Me.txtCoordTelefono.AcceptsReturn = True
        Me.txtCoordTelefono.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordTelefono.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordTelefono.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordTelefono.Location = New System.Drawing.Point(140, 265)
        Me.txtCoordTelefono.MaxLength = 7
        Me.txtCoordTelefono.Name = "txtCoordTelefono"
        Me.txtCoordTelefono.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordTelefono.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordTelefono.TabIndex = 25
        Me.txtCoordTelefono.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordTelefono, "Coord. para Imprimir Teléfono (Ren, Col).")
        '
        'txtCoordEstado
        '
        Me.txtCoordEstado.AcceptsReturn = True
        Me.txtCoordEstado.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordEstado.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordEstado.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordEstado.Location = New System.Drawing.Point(343, 74)
        Me.txtCoordEstado.MaxLength = 7
        Me.txtCoordEstado.Name = "txtCoordEstado"
        Me.txtCoordEstado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordEstado.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordEstado.TabIndex = 37
        Me.txtCoordEstado.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordEstado, "Coord. para Imprimir Estado (Ren, Col).")
        '
        'txtCoordCiudad
        '
        Me.txtCoordCiudad.AcceptsReturn = True
        Me.txtCoordCiudad.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordCiudad.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordCiudad.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordCiudad.Location = New System.Drawing.Point(343, 51)
        Me.txtCoordCiudad.MaxLength = 7
        Me.txtCoordCiudad.Name = "txtCoordCiudad"
        Me.txtCoordCiudad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordCiudad.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordCiudad.TabIndex = 36
        Me.txtCoordCiudad.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordCiudad, "Coord. para Imprimir Ciudad (Ren, Col).")
        '
        'txtCoordColonia
        '
        Me.txtCoordColonia.AcceptsReturn = True
        Me.txtCoordColonia.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordColonia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordColonia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordColonia.Location = New System.Drawing.Point(140, 217)
        Me.txtCoordColonia.MaxLength = 7
        Me.txtCoordColonia.Name = "txtCoordColonia"
        Me.txtCoordColonia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordColonia.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordColonia.TabIndex = 23
        Me.txtCoordColonia.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordColonia, "Coord. para Imprimir Colonia (Ren, Col).")
        '
        'txtCoordCalle
        '
        Me.txtCoordCalle.AcceptsReturn = True
        Me.txtCoordCalle.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordCalle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordCalle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordCalle.Location = New System.Drawing.Point(140, 193)
        Me.txtCoordCalle.MaxLength = 7
        Me.txtCoordCalle.Name = "txtCoordCalle"
        Me.txtCoordCalle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordCalle.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordCalle.TabIndex = 22
        Me.txtCoordCalle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordCalle, "Coord. para Imprimir Dirección(Ren, Col).")
        '
        'txtCoordFecha
        '
        Me.txtCoordFecha.AcceptsReturn = True
        Me.txtCoordFecha.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordFecha.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordFecha.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordFecha.Location = New System.Drawing.Point(140, 121)
        Me.txtCoordFecha.MaxLength = 7
        Me.txtCoordFecha.Name = "txtCoordFecha"
        Me.txtCoordFecha.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordFecha.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordFecha.TabIndex = 19
        Me.txtCoordFecha.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordFecha, "Coord. para Imprimir Fecha (Ren, Col).")
        '
        'txtCoordRFC
        '
        Me.txtCoordRFC.AcceptsReturn = True
        Me.txtCoordRFC.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordRFC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordRFC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordRFC.Location = New System.Drawing.Point(140, 169)
        Me.txtCoordRFC.MaxLength = 7
        Me.txtCoordRFC.Name = "txtCoordRFC"
        Me.txtCoordRFC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordRFC.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordRFC.TabIndex = 21
        Me.txtCoordRFC.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordRFC, "Coord. para Imprimir RFC (Ren, Col).")
        '
        'txtCoordEmpresa
        '
        Me.txtCoordEmpresa.AcceptsReturn = True
        Me.txtCoordEmpresa.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoordEmpresa.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoordEmpresa.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoordEmpresa.Location = New System.Drawing.Point(140, 145)
        Me.txtCoordEmpresa.MaxLength = 7
        Me.txtCoordEmpresa.Name = "txtCoordEmpresa"
        Me.txtCoordEmpresa.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoordEmpresa.Size = New System.Drawing.Size(57, 20)
        Me.txtCoordEmpresa.TabIndex = 20
        Me.txtCoordEmpresa.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCoordEmpresa, "Coord. para Imprimir Nombre Cliente (Ren, Col).")
        '
        'txtRenXFactura
        '
        Me.txtRenXFactura.AcceptsReturn = True
        Me.txtRenXFactura.BackColor = System.Drawing.SystemColors.Window
        Me.txtRenXFactura.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRenXFactura.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRenXFactura.Location = New System.Drawing.Point(140, 74)
        Me.txtRenXFactura.MaxLength = 4
        Me.txtRenXFactura.Name = "txtRenXFactura"
        Me.txtRenXFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRenXFactura.Size = New System.Drawing.Size(57, 20)
        Me.txtRenXFactura.TabIndex = 17
        Me.txtRenXFactura.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtRenXFactura, "Renglones que debe  tener una factura")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.Panel3)
        Me.Frame2.Controls.Add(Me.Frame3)
        Me.Frame2.Controls.Add(Me.chkAplicarSucursales)
        Me.Frame2.Controls.Add(Me.Marco)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(9, 6)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(697, 505)
        Me.Frame2.TabIndex = 0
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Tipo de Facturación"
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me._optTicket_0)
        Me.Frame3.Controls.Add(Me._optTicket_1)
        Me.Frame3.Controls.Add(Me.dbcSucursales)
        Me.Frame3.Controls.Add(Me._Label1_2)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(16, 16)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(665, 49)
        Me.Frame3.TabIndex = 1
        Me.Frame3.TabStop = False
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(80, 16)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(155, 21)
        Me.dbcSucursales.TabIndex = 3
        '
        'chkAplicarSucursales
        '
        Me.chkAplicarSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkAplicarSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAplicarSucursales.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkAplicarSucursales.Location = New System.Drawing.Point(511, 412)
        Me.chkAplicarSucursales.Name = "chkAplicarSucursales"
        Me.chkAplicarSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAplicarSucursales.Size = New System.Drawing.Size(170, 39)
        Me.chkAplicarSucursales.TabIndex = 77
        Me.chkAplicarSucursales.Text = "Aplicar esta configuración a todas las sucursales"
        Me.chkAplicarSucursales.UseVisualStyleBackColor = False
        '
        'Marco
        '
        Me.Marco.BackColor = System.Drawing.SystemColors.Control
        Me.Marco.Controls.Add(Me.btnLong)
        Me.Marco.Controls.Add(Me.txtLeyenda)
        Me.Marco.Controls.Add(Me.cboLetra)
        Me.Marco.Controls.Add(Me.TxtCoordLeyenda)
        Me.Marco.Controls.Add(Me.txtColIVAporPartida)
        Me.Marco.Controls.Add(Me.txtColDesctoDetalle)
        Me.Marco.Controls.Add(Me.txtColDescripcion)
        Me.Marco.Controls.Add(Me.txtRenXdetalle)
        Me.Marco.Controls.Add(Me.txtCoordLugarExped)
        Me.Marco.Controls.Add(Me.txtCoordCP)
        Me.Marco.Controls.Add(Me.txtCoordFolio)
        Me.Marco.Controls.Add(Me.txtImpteConLetra)
        Me.Marco.Controls.Add(Me.txtPrimerPartida)
        Me.Marco.Controls.Add(Me.txtPartidasXfactura)
        Me.Marco.Controls.Add(Me.txtCoordTotal)
        Me.Marco.Controls.Add(Me.txtCoordIVA)
        Me.Marco.Controls.Add(Me.txtCoordSubTotal)
        Me.Marco.Controls.Add(Me.txtCoordImporte)
        Me.Marco.Controls.Add(Me.txtCoordPrecioVta)
        Me.Marco.Controls.Add(Me.txtCoordPromocion)
        Me.Marco.Controls.Add(Me.txtCoordDesctos)
        Me.Marco.Controls.Add(Me.txtCoordCantidad)
        Me.Marco.Controls.Add(Me.txtCoordCodigo)
        Me.Marco.Controls.Add(Me.txtCoordTelefono)
        Me.Marco.Controls.Add(Me.txtCoordEstado)
        Me.Marco.Controls.Add(Me.txtCoordCiudad)
        Me.Marco.Controls.Add(Me.txtCoordColonia)
        Me.Marco.Controls.Add(Me.txtCoordCalle)
        Me.Marco.Controls.Add(Me.txtCoordFecha)
        Me.Marco.Controls.Add(Me.txtCoordRFC)
        Me.Marco.Controls.Add(Me.txtCoordEmpresa)
        Me.Marco.Controls.Add(Me.txtRenXFactura)
        Me.Marco.Controls.Add(Me._Line1_1)
        Me.Marco.Controls.Add(Me.Label5)
        Me.Marco.Controls.Add(Me._Line1_0)
        Me.Marco.Controls.Add(Me._lblEtiqueta_100)
        Me.Marco.Controls.Add(Me._lblEtiqueta_0)
        Me.Marco.Controls.Add(Me.Label3)
        Me.Marco.Controls.Add(Me.txtDatosPartida)
        Me.Marco.Controls.Add(Me.TxtDatosGenerales)
        Me.Marco.Controls.Add(Me._Label2_28)
        Me.Marco.Controls.Add(Me._Label2_27)
        Me.Marco.Controls.Add(Me._Label2_26)
        Me.Marco.Controls.Add(Me._Label2_25)
        Me.Marco.Controls.Add(Me._Label2_24)
        Me.Marco.Controls.Add(Me._Label2_158)
        Me.Marco.Controls.Add(Me._lblEtiqueta_153)
        Me.Marco.Controls.Add(Me._Label2_21)
        Me.Marco.Controls.Add(Me._Label2_20)
        Me.Marco.Controls.Add(Me._Label2_19)
        Me.Marco.Controls.Add(Me._Label2_18)
        Me.Marco.Controls.Add(Me._Label2_16)
        Me.Marco.Controls.Add(Me._Label2_15)
        Me.Marco.Controls.Add(Me._Label2_160)
        Me.Marco.Controls.Add(Me._Label2_161)
        Me.Marco.Controls.Add(Me._Label2_12)
        Me.Marco.Controls.Add(Me._Label2_11)
        Me.Marco.Controls.Add(Me._Label2_10)
        Me.Marco.Controls.Add(Me._Label2_9)
        Me.Marco.Controls.Add(Me._Label2_159)
        Me.Marco.Controls.Add(Me._lblEtiqueta_157)
        Me.Marco.Controls.Add(Me._lblEtiqueta_156)
        Me.Marco.Controls.Add(Me._lblEtiqueta_155)
        Me.Marco.Controls.Add(Me._lblEtiqueta_154)
        Me.Marco.Controls.Add(Me._lblEtiqueta_152)
        Me.Marco.Controls.Add(Me._lblEtiqueta_151)
        Me.Marco.Controls.Add(Me._lblEtiqueta_150)
        Me.Marco.Controls.Add(Me._lblEtiqueta_1)
        Me.Marco.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Marco.Location = New System.Drawing.Point(16, 69)
        Me.Marco.Name = "Marco"
        Me.Marco.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Marco.Size = New System.Drawing.Size(665, 335)
        Me.Marco.TabIndex = 4
        Me.Marco.TabStop = False
        '
        '_Line1_1
        '
        Me._Line1_1.BackColor = System.Drawing.Color.Black
        Me._Line1_1.Location = New System.Drawing.Point(30, 296)
        Me._Line1_1.Name = "_Line1_1"
        Me._Line1_1.Size = New System.Drawing.Size(614, 1)
        Me._Line1_1.TabIndex = 70
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(222, 266)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(116, 21)
        Me.Label5.TabIndex = 35
        Me.Label5.Text = "Longitudes de Datos :"
        '
        '_Line1_0
        '
        Me._Line1_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._Line1_0.Location = New System.Drawing.Point(30, 295)
        Me._Line1_0.Name = "_Line1_0"
        Me._Line1_0.Size = New System.Drawing.Size(612, 1)
        Me._Line1_0.TabIndex = 71
        '
        '_lblEtiqueta_100
        '
        Me._lblEtiqueta_100.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_100.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_100.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_100.Location = New System.Drawing.Point(31, 304)
        Me._lblEtiqueta_100.Name = "_lblEtiqueta_100"
        Me._lblEtiqueta_100.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_100.Size = New System.Drawing.Size(80, 21)
        Me._lblEtiqueta_100.TabIndex = 67
        Me._lblEtiqueta_100.Text = "Leyeda Factura"
        '
        '_lblEtiqueta_0
        '
        Me._lblEtiqueta_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_0.Location = New System.Drawing.Point(31, 50)
        Me._lblEtiqueta_0.Name = "_lblEtiqueta_0"
        Me._lblEtiqueta_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_0.Size = New System.Drawing.Size(100, 21)
        Me._lblEtiqueta_0.TabIndex = 6
        Me._lblEtiqueta_0.Text = "Tamaño de Letra :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(222, 242)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(127, 21)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Coord. Leyenda Factura :"
        '
        'txtDatosPartida
        '
        Me.txtDatosPartida.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.txtDatosPartida.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDatosPartida.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDatosPartida.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtDatosPartida.Location = New System.Drawing.Point(431, 24)
        Me.txtDatosPartida.Name = "txtDatosPartida"
        Me.txtDatosPartida.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDatosPartida.Size = New System.Drawing.Size(210, 19)
        Me.txtDatosPartida.TabIndex = 46
        Me.txtDatosPartida.Text = "Datos de Partidas"
        Me.txtDatosPartida.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TxtDatosGenerales
        '
        Me.TxtDatosGenerales.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.TxtDatosGenerales.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtDatosGenerales.Cursor = System.Windows.Forms.Cursors.Default
        Me.TxtDatosGenerales.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.TxtDatosGenerales.Location = New System.Drawing.Point(29, 24)
        Me.TxtDatosGenerales.Name = "TxtDatosGenerales"
        Me.TxtDatosGenerales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtDatosGenerales.Size = New System.Drawing.Size(372, 19)
        Me.TxtDatosGenerales.TabIndex = 5
        Me.TxtDatosGenerales.Text = "Datos Generales"
        Me.TxtDatosGenerales.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label2_28
        '
        Me._Label2_28.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_28.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_28.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_28.Location = New System.Drawing.Point(432, 217)
        Me._Label2_28.Name = "_Label2_28"
        Me._Label2_28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_28.Size = New System.Drawing.Size(153, 21)
        Me._Label2_28.TabIndex = 54
        Me._Label2_28.Text = "Columna IVA :"
        '
        '_Label2_27
        '
        Me._Label2_27.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_27.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_27.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_27.Location = New System.Drawing.Point(431, 170)
        Me._Label2_27.Name = "_Label2_27"
        Me._Label2_27.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_27.Size = New System.Drawing.Size(153, 21)
        Me._Label2_27.TabIndex = 52
        Me._Label2_27.Text = "Columna Descuento :"
        '
        '_Label2_26
        '
        Me._Label2_26.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_26.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_26.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_26.Location = New System.Drawing.Point(431, 146)
        Me._Label2_26.Name = "_Label2_26"
        Me._Label2_26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_26.Size = New System.Drawing.Size(153, 21)
        Me._Label2_26.TabIndex = 51
        Me._Label2_26.Text = "Columna Descripción :"
        '
        '_Label2_25
        '
        Me._Label2_25.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_25.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_25.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_25.Location = New System.Drawing.Point(431, 50)
        Me._Label2_25.Name = "_Label2_25"
        Me._Label2_25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_25.Size = New System.Drawing.Size(153, 21)
        Me._Label2_25.TabIndex = 47
        Me._Label2_25.Text = "Total de Partidas :"
        '
        '_Label2_24
        '
        Me._Label2_24.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_24.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_24.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_24.Location = New System.Drawing.Point(222, 217)
        Me._Label2_24.Name = "_Label2_24"
        Me._Label2_24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_24.Size = New System.Drawing.Size(127, 21)
        Me._Label2_24.TabIndex = 33
        Me._Label2_24.Text = "Coord. Lugar Expedición  "
        '
        '_Label2_158
        '
        Me._Label2_158.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_158.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_158.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_158.Location = New System.Drawing.Point(31, 242)
        Me._Label2_158.Name = "_Label2_158"
        Me._Label2_158.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_158.Size = New System.Drawing.Size(107, 21)
        Me._Label2_158.TabIndex = 14
        Me._Label2_158.Text = "Coord. Código Postal :"
        '
        '_lblEtiqueta_153
        '
        Me._lblEtiqueta_153.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_153.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_153.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_153.Location = New System.Drawing.Point(31, 98)
        Me._lblEtiqueta_153.Name = "_lblEtiqueta_153"
        Me._lblEtiqueta_153.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_153.Size = New System.Drawing.Size(107, 21)
        Me._lblEtiqueta_153.TabIndex = 8
        Me._lblEtiqueta_153.Text = "Coord. Folio :   "
        '
        '_Label2_21
        '
        Me._Label2_21.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_21.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_21.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_21.Location = New System.Drawing.Point(222, 170)
        Me._Label2_21.Name = "_Label2_21"
        Me._Label2_21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_21.Size = New System.Drawing.Size(127, 21)
        Me._Label2_21.TabIndex = 31
        Me._Label2_21.Text = "Coord. Total :"
        '
        '_Label2_20
        '
        Me._Label2_20.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_20.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_20.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_20.Location = New System.Drawing.Point(485, 304)
        Me._Label2_20.Name = "_Label2_20"
        Me._Label2_20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_20.Size = New System.Drawing.Size(178, 17)
        Me._Label2_20.TabIndex = 74
        Me._Label2_20.Text = "Partidas por Factura                    :"
        Me._Label2_20.Visible = False
        '
        '_Label2_19
        '
        Me._Label2_19.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_19.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_19.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_19.Location = New System.Drawing.Point(431, 74)
        Me._Label2_19.Name = "_Label2_19"
        Me._Label2_19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_19.Size = New System.Drawing.Size(153, 21)
        Me._Label2_19.TabIndex = 48
        Me._Label2_19.Text = "Renglón para la Primer Partida :"
        '
        '_Label2_18
        '
        Me._Label2_18.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_18.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_18.Location = New System.Drawing.Point(222, 192)
        Me._Label2_18.Name = "_Label2_18"
        Me._Label2_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_18.Size = New System.Drawing.Size(127, 21)
        Me._Label2_18.TabIndex = 32
        Me._Label2_18.Text = "Coord. Importe con Letra "
        '
        '_Label2_16
        '
        Me._Label2_16.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_16.Location = New System.Drawing.Point(431, 242)
        Me._Label2_16.Name = "_Label2_16"
        Me._Label2_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_16.Size = New System.Drawing.Size(153, 21)
        Me._Label2_16.TabIndex = 55
        Me._Label2_16.Text = "Columna Precio Venta :"
        '
        '_Label2_15
        '
        Me._Label2_15.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_15.Location = New System.Drawing.Point(431, 266)
        Me._Label2_15.Name = "_Label2_15"
        Me._Label2_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_15.Size = New System.Drawing.Size(153, 21)
        Me._Label2_15.TabIndex = 56
        Me._Label2_15.Text = "Columna Importe :"
        '
        '_Label2_160
        '
        Me._Label2_160.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_160.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_160.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_160.Location = New System.Drawing.Point(222, 98)
        Me._Label2_160.Name = "_Label2_160"
        Me._Label2_160.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_160.Size = New System.Drawing.Size(127, 21)
        Me._Label2_160.TabIndex = 28
        Me._Label2_160.Text = "Coord. SubTotal :"
        '
        '_Label2_161
        '
        Me._Label2_161.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_161.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_161.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_161.Location = New System.Drawing.Point(222, 146)
        Me._Label2_161.Name = "_Label2_161"
        Me._Label2_161.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_161.Size = New System.Drawing.Size(127, 21)
        Me._Label2_161.TabIndex = 30
        Me._Label2_161.Text = "Coord. IVA :"
        '
        '_Label2_12
        '
        Me._Label2_12.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_12.Location = New System.Drawing.Point(431, 194)
        Me._Label2_12.Name = "_Label2_12"
        Me._Label2_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_12.Size = New System.Drawing.Size(153, 21)
        Me._Label2_12.TabIndex = 53
        Me._Label2_12.Text = "Columna Promoción :"
        '
        '_Label2_11
        '
        Me._Label2_11.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_11.Location = New System.Drawing.Point(222, 122)
        Me._Label2_11.Name = "_Label2_11"
        Me._Label2_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_11.Size = New System.Drawing.Size(127, 21)
        Me._Label2_11.TabIndex = 29
        Me._Label2_11.Text = "Coord. Descuento :"
        '
        '_Label2_10
        '
        Me._Label2_10.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_10.Location = New System.Drawing.Point(431, 98)
        Me._Label2_10.Name = "_Label2_10"
        Me._Label2_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_10.Size = New System.Drawing.Size(153, 21)
        Me._Label2_10.TabIndex = 49
        Me._Label2_10.Text = "Columna Cantidad :"
        '
        '_Label2_9
        '
        Me._Label2_9.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_9.Location = New System.Drawing.Point(431, 122)
        Me._Label2_9.Name = "_Label2_9"
        Me._Label2_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_9.Size = New System.Drawing.Size(153, 21)
        Me._Label2_9.TabIndex = 50
        Me._Label2_9.Text = "Columna Código :"
        '
        '_Label2_159
        '
        Me._Label2_159.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_159.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_159.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_159.Location = New System.Drawing.Point(31, 266)
        Me._Label2_159.Name = "_Label2_159"
        Me._Label2_159.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_159.Size = New System.Drawing.Size(96, 21)
        Me._Label2_159.TabIndex = 15
        Me._Label2_159.Text = "Coord. Teléfono :"
        '
        '_lblEtiqueta_157
        '
        Me._lblEtiqueta_157.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_157.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_157.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_157.Location = New System.Drawing.Point(222, 74)
        Me._lblEtiqueta_157.Name = "_lblEtiqueta_157"
        Me._lblEtiqueta_157.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_157.Size = New System.Drawing.Size(107, 21)
        Me._lblEtiqueta_157.TabIndex = 27
        Me._lblEtiqueta_157.Text = "Coord. Estado :"
        '
        '_lblEtiqueta_156
        '
        Me._lblEtiqueta_156.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_156.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_156.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_156.Location = New System.Drawing.Point(222, 49)
        Me._lblEtiqueta_156.Name = "_lblEtiqueta_156"
        Me._lblEtiqueta_156.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_156.Size = New System.Drawing.Size(107, 21)
        Me._lblEtiqueta_156.TabIndex = 26
        Me._lblEtiqueta_156.Text = "Coord. Ciudad :"
        '
        '_lblEtiqueta_155
        '
        Me._lblEtiqueta_155.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_155.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_155.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_155.Location = New System.Drawing.Point(31, 217)
        Me._lblEtiqueta_155.Name = "_lblEtiqueta_155"
        Me._lblEtiqueta_155.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_155.Size = New System.Drawing.Size(107, 21)
        Me._lblEtiqueta_155.TabIndex = 13
        Me._lblEtiqueta_155.Text = "Coord. Colonia :"
        '
        '_lblEtiqueta_154
        '
        Me._lblEtiqueta_154.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_154.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_154.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_154.Location = New System.Drawing.Point(31, 194)
        Me._lblEtiqueta_154.Name = "_lblEtiqueta_154"
        Me._lblEtiqueta_154.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_154.Size = New System.Drawing.Size(107, 21)
        Me._lblEtiqueta_154.TabIndex = 12
        Me._lblEtiqueta_154.Text = "Coord. Dirección :"
        '
        '_lblEtiqueta_152
        '
        Me._lblEtiqueta_152.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_152.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_152.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_152.Location = New System.Drawing.Point(31, 122)
        Me._lblEtiqueta_152.Name = "_lblEtiqueta_152"
        Me._lblEtiqueta_152.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_152.Size = New System.Drawing.Size(107, 21)
        Me._lblEtiqueta_152.TabIndex = 9
        Me._lblEtiqueta_152.Text = "Coord. Fecha :"
        '
        '_lblEtiqueta_151
        '
        Me._lblEtiqueta_151.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_151.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_151.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_151.Location = New System.Drawing.Point(31, 170)
        Me._lblEtiqueta_151.Name = "_lblEtiqueta_151"
        Me._lblEtiqueta_151.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_151.Size = New System.Drawing.Size(107, 21)
        Me._lblEtiqueta_151.TabIndex = 11
        Me._lblEtiqueta_151.Text = "Coord. R.F.C. :"
        '
        '_lblEtiqueta_150
        '
        Me._lblEtiqueta_150.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_150.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_150.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_150.Location = New System.Drawing.Point(31, 146)
        Me._lblEtiqueta_150.Name = "_lblEtiqueta_150"
        Me._lblEtiqueta_150.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_150.Size = New System.Drawing.Size(107, 21)
        Me._lblEtiqueta_150.TabIndex = 10
        Me._lblEtiqueta_150.Text = "Coord. Cliente :"
        '
        '_lblEtiqueta_1
        '
        Me._lblEtiqueta_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblEtiqueta_1.Location = New System.Drawing.Point(31, 75)
        Me._lblEtiqueta_1.Name = "_lblEtiqueta_1"
        Me._lblEtiqueta_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_1.Size = New System.Drawing.Size(107, 21)
        Me._lblEtiqueta_1.TabIndex = 7
        Me._lblEtiqueta_1.Text = "Renglones Factura :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.mskFecha)
        Me.Frame1.Controls.Add(Me.mskHora)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(11, 11)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(704, 49)
        Me.Frame1.TabIndex = 71
        Me.Frame1.TabStop = False
        Me.Frame1.Visible = False
        '
        'mskFecha
        '
        Me.mskFecha.AllowPromptAsInput = False
        Me.mskFecha.Enabled = False
        Me.mskFecha.Location = New System.Drawing.Point(64, 16)
        Me.mskFecha.Name = "mskFecha"
        Me.mskFecha.Size = New System.Drawing.Size(97, 20)
        Me.mskFecha.TabIndex = 75
        '
        'mskHora
        '
        Me.mskHora.AllowPromptAsInput = False
        Me.mskHora.Enabled = False
        Me.mskHora.Location = New System.Drawing.Point(640, 16)
        Me.mskHora.Name = "mskHora"
        Me.mskHora.Size = New System.Drawing.Size(97, 20)
        Me.mskHora.TabIndex = 70
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(608, 20)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(49, 17)
        Me._Label1_1.TabIndex = 73
        Me._Label1_1.Text = "Hora   :"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(16, 20)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(49, 17)
        Me._Label1_0.TabIndex = 72
        Me._Label1_0.Text = "Fecha :"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(16, 412)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(377, 74)
        Me.Panel3.TabIndex = 72
        '
        'btnSalir
        '
        Me.btnSalir.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.salir
        Me.btnSalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSalir.Location = New System.Drawing.Point(208, 14)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(50, 42)
        Me.btnSalir.TabIndex = 70
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.buscar
        Me.btnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBuscar.Location = New System.Drawing.Point(160, 14)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(50, 42)
        Me.btnBuscar.TabIndex = 67
        Me.btnBuscar.Text = " "
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.grabar
        Me.btnGuardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnGuardar.Location = New System.Drawing.Point(11, 14)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(50, 42)
        Me.btnGuardar.TabIndex = 64
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.nuevo
        Me.btnLimpiar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnLimpiar.Location = New System.Drawing.Point(110, 14)
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
        Me.btnEliminar.Location = New System.Drawing.Point(61, 14)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(50, 42)
        Me.btnEliminar.TabIndex = 65
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'frmPVConfigFacturacion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(722, 525)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(212, 118)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPVConfigFacturacion"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Configuración de Facturas"
        Me.Frame2.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Marco.ResumeLayout(False)
        Me.Marco.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblEtiqueta, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optTicket, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click

    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub
End Class