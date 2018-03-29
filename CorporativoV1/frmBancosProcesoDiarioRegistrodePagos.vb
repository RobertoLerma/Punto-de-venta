Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms
Imports System.Data
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosProcesoDiarioRegistrodePagos
    Inherits System.Windows.Forms.Form

    Public bandera As Boolean = False

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdOrigenyAplicacion As System.Windows.Forms.Button
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents txtFolioEgreso As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblMoneda As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents txtFolioElectronico As System.Windows.Forms.TextBox
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Frame6 As System.Windows.Forms.GroupBox
    Public WithEvents _optFormaPago_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optFormaPago_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optFormaPago_2 As System.Windows.Forms.RadioButton
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents txtImporte As System.Windows.Forms.TextBox
    Public WithEvents txtConceptoCancelacion As System.Windows.Forms.TextBox
    Public WithEvents chkCancelado As System.Windows.Forms.CheckBox
    Public WithEvents txtNumeroCheque As System.Windows.Forms.TextBox
    Public WithEvents dtpFechaCheque As System.Windows.Forms.DateTimePicker
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents txtConcepto As System.Windows.Forms.TextBox
    Public WithEvents txtBeneficiario As System.Windows.Forms.TextBox
    Public WithEvents dbcBanco As System.Windows.Forms.ComboBox
    Public WithEvents _optTipoPago_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optTipoPago_0 As System.Windows.Forms.RadioButton
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dbcCuentaBancaria As System.Windows.Forms.ComboBox
    Public WithEvents lblCancelada As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents optFormaPago As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents Panel1 As Panel
    Public WithEvents optTipoPago As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public strControlActual As String 'Nombre del control actual


    'Variables
    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Dim mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Dim mblnSalir As Boolean 'Para Salir Con el Esc
    Dim FueraChange As Boolean
    Dim PierdeFoco As Boolean
    Dim intCodBanco As Integer
    Dim intUltFormaPago As Integer
    Dim tecla As Integer
    Dim LetraFolio As String
    Dim ConsecutivoCheque As Integer
    Dim MonedaProgramacion As String
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    Dim cMoneda As String
    Public intNumPartida As Integer
    Public ConsultaPagos As Boolean
    Dim Concepto As String

    'Variables y constantes utilizadas por Paimí para realizar los pagos elegidos en el formulario de Emision de Pagos
    Const C_PAGAR As String = "P"

    Const P_RENENCABEZADO As Integer = 0 : Const S_RENENCABEZADO As Integer = 1

    Const P_COLFACTURA As Integer = 0 : Const P_COLPAGO As Integer = 1 : Const P_COLFECHAFACTURA As Integer = 2
    Const P_COLFECHAVENCTO As Integer = 3 : Const P_COLFECHAPAGO As Integer = 4 : Const P_COLIMPORTE As Integer = 5
    Const P_COLPAGOS As Integer = 6 : Const P_COLSALDO As Integer = 7 : Const P_COLIMPORTEPAGO As Integer = 8
    Const P_COLMONEDA As Integer = 9 : Const P_COLDESCTOPORC As Integer = 10 : Const P_COLDESCTOFIN As Integer = 11
    Const P_COLSUBTOTALDF As Integer = 12 : Const P_COLIVADF As Integer = 13 : Const P_COLAPAGAR As Integer = 14
    Const P_COLNUMPARTIDA As Integer = 15 : Const P_COLESTATUS As Integer = 16

    Const A_COLFOLIO As Integer = 0 : Const A_COLFECHA As Integer = 1 : Const A_COLIMPORTE As Integer = 2
    Const A_COLESTATUS As Integer = 3
    Const N_COLFOLIO As Integer = 0 : Const N_COLFECHA As Integer = 1 : Const N_COLFACTURA As Integer = 2
    Const N_COLTOTAL As Integer = 3 : Const N_COLTIPO As Integer = 4 : Const N_COLESTATUS As Integer = 5

    Dim nEP_CodProveedor As Integer
    Dim cEP_DescProveedor As String
    Dim nEP_ImportePago As Decimal
    Dim cEP_Moneda As String
    Public blnEmisionPagos As Boolean 'Se utiliza para indicar que el formulario fue llamado desde Emisión de Pagos

    Dim fFolioProgramacionP As String
    Dim fCodProvAcreed As Integer
    Dim fTipoFacturaCxP As String
    Dim fTipoGasto As String
    Dim fFolioFactura As String
    Dim fFechaFactura As Date
    Dim fFechaPago As Date
    Dim fTotalPago As Decimal
    Dim fMoneda As String
    Dim fTipoCambio As Decimal
    Dim fTipoCambioE As Decimal
    Dim fDescuentoFinanciero As Decimal
    Dim fSubTotalDF As Decimal
    Dim fIvaDF As Decimal
    Dim fEstatus As String
    Dim fFechaCancel As Date
    Dim fTipoPagoProg As String
    Dim fEfectivo As Boolean

    Public frmPagos2 As frmBancosProcesoDiarioOrigenyAplicacion = New frmBancosProcesoDiarioOrigenyAplicacion()


    Public Sub RecuperaDatosProgramacionPagos(ByRef cFolioProg As Object, ByRef nNumPartida As Integer)
        Try 'On Error GoTo Merr
            Dim lStrSql As String
            Dim rsLocal As ADODB.Recordset
            lStrSql = "SELECT * FROM ProgramacionPagos WHERE FolioProgramacionP = '" & Trim(cFolioProg) & "' and NumPartida = " & nNumPartida
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, lStrSql))
            rsLocal = Cmd.Execute
            If rsLocal.RecordCount > 0 Then
                fFolioProgramacionP = rsLocal.Fields("FolioProgramacionP").Value
                fCodProvAcreed = rsLocal.Fields("CodProvAcreed").Value
                fTipoFacturaCxP = rsLocal.Fields("TipoFacturaCxP").Value
                fTipoGasto = rsLocal.Fields("TipoGasto").Value
                fFolioFactura = rsLocal.Fields("FolioFactura").Value
                fFechaFactura = rsLocal.Fields("FechaFactura").Value
                fFechaPago = rsLocal.Fields("FechaPago").Value
                fTotalPago = rsLocal.Fields("TotalPago").Value
                fMoneda = rsLocal.Fields("Moneda").Value
                fTipoCambio = rsLocal.Fields("TipoCambio").Value
                fTipoCambioE = rsLocal.Fields("TipoCambioE").Value
                fDescuentoFinanciero = rsLocal.Fields("DescuentoFinanciero").Value
                fSubTotalDF = rsLocal.Fields("SubTotalDF").Value
                fIvaDF = rsLocal.Fields("IvaDF").Value
                fEstatus = rsLocal.Fields("Estatus").Value
                fFechaCancel = rsLocal.Fields("FechaCancel").Value
                fTipoPagoProg = rsLocal.Fields("TipoPagoProg").Value
                fEfectivo = rsLocal.Fields("Efectivo").Value
            Else
                fFolioProgramacionP = ""
                fCodProvAcreed = 0
                fTipoFacturaCxP = ""
                fTipoGasto = ""
                fFolioFactura = ""
                fFechaFactura = #1/1/1900#
                fFechaPago = #1/1/1900#
                fTotalPago = 0
                fMoneda = ""
                fTipoCambio = 0
                fTipoCambioE = 0
                fDescuentoFinanciero = 0
                fSubTotalDF = 0
                fIvaDF = 0
                fEstatus = ""
                fFechaCancel = #1/1/1900#
                fTipoPagoProg = ""
                fEfectivo = False
            End If
        Catch 'Merr:
            If Err.Number <> 0 Then
                ModEstandar.MostrarError()
            End If
        End Try
    End Sub

    Public Sub NuevoPagos()
        Try 'On Error GoTo Merr
            FueraChange = True
            Me.dbcBanco.Text = ""
            Me.dbcBanco.Tag = Me.dbcBanco.Text
            intCodBanco = 0
            Me.dbcCuentaBancaria.Text = ""
            Me.dbcCuentaBancaria.Tag = Me.dbcCuentaBancaria.Text
            FueraChange = False
            If blnEmisionPagos Then
                'Me.chkPagoProgramado.Value = vbUnchecked
                'Me.chkPagoProgramado.Visible = False
                'Me.Label8.Visible = False
                'Me.txtFolioProgramacion.text = ""
                'Me.txtFolioProgramacion.Visible = False
                Me.txtBeneficiario.ReadOnly = True
                Me.txtImporte.ReadOnly = True
            Else
                'Me.chkPagoProgramado.Value = vbUnchecked
                'Me.chkPagoProgramado.Visible = True
                'Me.Label8.Visible = True
                'Me.txtFolioProgramacion.text = ""
                'Me.txtFolioProgramacion.Visible = True
                Me.txtBeneficiario.ReadOnly = False
                Me.txtImporte.ReadOnly = False
            End If
        Catch 'Merr:
            If Err.Number <> 0 Then
                ModEstandar.MostrarError()
            End If
        End Try
    End Sub

    'Procedimiento utilizado por Paimí para vaciar los datos de los pagos en este formulario
    Public Sub LlenaDatosPagos(ByRef nCodProveedor As Integer, ByRef cDescProveedor As String, ByRef nImportePago As Decimal, ByRef cTipoMoneda As String)

        If (bandera = True) Then
            Exit Sub
        End If

        Try 'On Error GoTo Merr
            Dim I As Integer
            Nuevo()
            nEP_CodProveedor = nCodProveedor
            cEP_DescProveedor = Trim(cDescProveedor)
            nEP_ImportePago = nImportePago
            cEP_Moneda = cTipoMoneda

            Me.NuevoPagos()

            If Trim(Me.txtFolioEgreso.Text) = "" Then
                Me.txtFolioEgreso.Text = C_TIPOMOVEGRESO & Format(Me.dtpFecha.Value.Year, "0000") & Format(Me.dtpFecha.Value.Month, "00") & Format(Me.dtpFecha.Value.Day, "00") & "0000"
            End If

            If frmCXPEmisionPagos.optOrigen(0).Checked Then
                Me._optTipoPago_0.Checked = True
                Me._optTipoPago_1.Checked = False
            Else
                Me._optTipoPago_0.Checked = True
                Me._optTipoPago_1.Checked = False
            End If
            Me._optTipoPago_0.Enabled = False
            Me._optTipoPago_1.Enabled = False

            Me.txtFolioEgreso.Focus()
            ModEstandar.SelTxt()

            Me.txtBeneficiario.Text = Trim(cDescProveedor)
            Me.txtImporte.Text = VB6.Format(nImportePago, "###,###,##0.00")

        Catch ''Merr:
            If Err.Number <> 0 Then
                ModEstandar.MostrarError()
            End If
        End Try
    End Sub

    Sub Buscar()
        Try 'On Error GoTo Merr
            Dim strSQL As String
            Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
            Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas 
            Dim I As Integer


            'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
            strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control

            Select Case strControlActual
                Case "TXTFOLIOEGRESO"
                    strCaptionForm = "Consulta de Registro de Pagos"
                    gStrSql = "SELECT FolioMovto AS FOLIO,Concepto AS CONCEPTO,Beneficiario AS BENEFICIARIO," & "FechaMovto AS FECHA,Importe AS IMPORTE FROM MovimientosBancarios " & "WHERE FolioMovto LIKE '" & txtFolioEgreso.Text & "%' AND Movimiento = '" & C_MOVPAGO & "' AND TipoMovto = '" & C_TIPOMOVEGRESO & "' ORDER BY FechaMovto DESC ,FolioMovto DESC"
                    'Case "TXTFOLIOPROGRAMACION"
                    '    strCaptionForm = "Consulta de Folio    s de Programación de Pagos"
                    '    gStrSql = "SELECT FolioProgramacionP AS 'FOLIO DE PROGRAMACION',NumPartida AS 'N° DE PARTIDA', FechaPago AS FECHA, TotalPago AS 'IMPORTE DEL PAGO' " & _
                    ''    "FROM ProgramacionPagos WHERE Estatus = 'V' AND PasoBancos = 0 ORDER BY FolioProgramacionP DESC, NumPartida ASC, FechaPago DESC"
                Case Else
                    Exit Sub
            End Select

            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute

            'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
            If RsGral.RecordCount = 0 Then
                MsjNoExiste(C_msgSINDATOS, gstrNombCortoEmpresa)
                Exit Sub
            End If

            'Carga el formulario de consulta 
            Dim FrmConsultas As FrmConsultas = New FrmConsultas()
            ConfiguraConsultas(FrmConsultas, 12700, RsGral, strTag, strCaptionForm)

            With FrmConsultas.Flexdet
                Select Case strControlActual
                    Case "TXTFOLIOEGRESO"
                        .set_ColWidth(0, 0, 1400) 'Columna del Folio
                        .set_ColWidth(1, 0, 4000) 'Columna del Concepto del Movimiento
                        .set_ColWidth(2, 0, 3900) 'Columna del Beneficiario del Movimiento
                        .set_ColWidth(3, 0, 1200) 'Columna de la Fecha del Movimiento
                        .set_ColWidth(4, 0, 1800) 'Columna del Importe del Movimiento
                        .set_ColAlignment(3, 4)

                        For I = 1 To FrmConsultas.Flexdet.Rows - 1
                            FrmConsultas.Flexdet.set_TextMatrix(I, 3, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 3), "dd/MMM/yyyy"))
                            FrmConsultas.Flexdet.set_TextMatrix(I, 4, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 4), "###,##0.00"))
                        Next I

                        FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
                        FrmConsultas.Left = VB6.TwipsToPixelsX(1200)

                End Select
            End With
            FrmConsultas.ShowDialog()
        Catch 'Merr:
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Function Guardar() As Boolean

        If (bandera = True) Then
            Exit Function
        End If

        On Error GoTo Err_Renamed
        Dim blnTransaccion As Boolean
        Dim Ejercicio As Integer
        Dim Periodo As String
        Dim I As Integer
        Dim J As Integer
        Dim cFolioPago As String
        Dim nNumPartida As Integer
        Dim strFolioCancelacion As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim cModulo As String

        'Do While (VB.Timer() - sglTiempoCambio) <= 2.1
        'Loop
        'System.Windows.Forms.Application.DoEvents()

        If Not mblnNuevo Then
            Exit Function
        End If

        If ValidaDatos() = False Then
            Exit Function
        End If
        If blnEmisionPagos Then 'Código de Paimí
            If cMoneda <> cEP_Moneda Then
                MsgBox("La moneda de la cuenta bancaria no coincide con el tipo de moneda que seleccionó en la emisión de pagos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Function
            End If
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        'Generar Folio del Movimiento
        Ejercicio = CInt(Format(Year(CDate(dtpFecha.Value)), "0000"))
        Periodo = Format(Month(CDate(dtpFecha.Value)), "00")
        BuscaEjercicio(dtpFecha.Value)
        ObtenerLimitedeFechas(CInt(Periodo), Ejercicio, FechaInicial, FechaFinal)
        FechaFinal = Format(FechaFinal, "dd/mmm/yyyy")
        gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVEGRESO & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFolioEgreso.Text = C_TIPOMOVEGRESO & Format(Year(CDate(dtpFecha.Value)), "0000") & Format(Month(CDate(dtpFecha.Value)), "00") & Format(VB.Day(CDate(dtpFecha.Value)), "00") & Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
            ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVEGRESO, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        'Obtener el Consecutivo de Cheque si es que se Genero Cheque y Actualizar el Consecutivo
        If _optFormaPago_1.Checked Then
            ConsecutivoCheque = ObtieneNumCheque(CInt(intCodBanco), dbcCuentaBancaria.Text, LetraFolio)
            txtNumeroCheque.Text = Trim(LetraFolio) & Format(CStr(ConsecutivoCheque), "000000")
            ModStoredProcedures.PR_IMECatCuentasBancarias(CStr(intCodBanco), dbcCuentaBancaria.Text, "", "", "", "", "0", Format(CStr(ConsecutivoCheque), "000000"), "", C_MODIFICACION, CStr(1))
            Cmd.Execute()
        End If
        'Guardar el Movimiento Bancario
        If blnEmisionPagos Then 'Este código de Paimí sirve para indicar el módulo de donde se originó el pago
            cModulo = C_MODULOCXP
        Else
            cModulo = C_MODULOBANCOS
        End If
        ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioEgreso.Text, Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_MOVPAGO, C_TIPOMOVEGRESO, IIf(_optFormaPago_0.Checked, C_NATURALEZAINTERNA, C_NATURALEZACOMERCIAL), IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(gcurCorpoTIPOCAMBIODOLAR), IIf(_optFormaPago_0.Checked, C_FORMAPAGOEFECTIVO, IIf(_optFormaPago_1.Checked, C_FORMAPAGOCHEQUE, C_FORMAPAGOELECTRONICO)), IIf(_optTipoPago_0.Checked, C_TIPOPAGOJOYERIA, C_TIPOPAGOPERSONAL), CStr(intCodBanco), dbcCuentaBancaria.Text, txtBeneficiario.Text, txtConcepto.Text, "0", "", CStr(intNumPartida), IIf(_optFormaPago_1.Checked, Format(dtpFechaCheque.Value, C_FORMATFECHAGUARDAR), "01/01/1900"), txtNumeroCheque.Text, txtImporte.Text, "V", "01/01/1900", "", CStr(0), "01/01/1900", cModulo, "", IIf(_optFormaPago_2.Checked = True, txtFolioElectronico.Text, ""), C_INSERCION, CStr(0))
        Cmd.Execute()
        'Actualizar en la Tabla de Programación de Pagos el Paso a Bancos
        '    If chkPagoProgramado.Value = 1 Then
        '        ModStoredProcedures.PR_IMEProgramacionPagos txtFolioProgramacion, CStr(intNumPartida), "0", "", "", "", "01/01/1900", "01/01/1900", _
        ''        "0", "", "0", "0", "0", "0", "0", "P", "01/01/1900", "", "0", "1", Format(dtpFecha, C_FORMATFECHAGUARDAR), C_MODIFICACION, 3
        '        Cmd.Execute
        '        'Guarda los Datos de Programación de Pagos en la Tabla de Pagos
        '        ModStoredProcedures.PR_IMEPagos txtFolioProgramacion, CStr(intNumPartida), CStr(fCodProvAcreed), fTipoFacturaCxP, fTipoGasto, Trim(fFolioFactura), _
        ''        Format(fFechaFactura, C_FORMATFECHAGUARDAR), Format(fFechaPago, C_FORMATFECHAGUARDAR), CStr(fTotalPago), fMoneda, CStr(fTipoCambio), CStr(fTipoCambioE), _
        ''        CStr(fSubTotalDF), CStr(fIvaDF), fEstatus, Format(fFechaCancel, C_FORMATFECHAGUARDAR), fTipoPagoProg, CStr(fEfectivo), "1", Format(dtpFecha, C_FORMATFECHAGUARDAR), _
        ''        txtFolioEgreso, "1", C_INSERCION, 0
        '        Cmd.Execute
        '    End If
        'Aquí guarda los datos de la Emisión de Pagos (Paimí)
        If blnEmisionPagos Then
            With frmCXPEmisionPagos.mshPagos
                For I = 2 To .Rows - 2
                    If Trim(.get_TextMatrix(I, P_COLPAGO)) = "" Then
                        Exit For
                    End If
                    If Trim(.get_TextMatrix(I, P_COLESTATUS)) = C_PAGAR Then
                        cFolioPago = Trim(.get_TextMatrix(I, P_COLPAGO))
                        nNumPartida = CInt(Numerico(.get_TextMatrix(I, P_COLNUMPARTIDA)))
                        Call Me.RecuperaDatosProgramacionPagos(cFolioPago, nNumPartida)
                        ModStoredProcedures.PR_IMEPagos(cFolioPago, CStr(nNumPartida), CStr(fCodProvAcreed), fTipoFacturaCxP, fTipoGasto, Trim(fFolioFactura), Format(fFechaFactura, C_FORMATFECHAGUARDAR), Format(Today, C_FORMATFECHAGUARDAR), Trim(.get_TextMatrix(I, P_COLIMPORTEPAGO)), VB.Left(Trim(.get_TextMatrix(I, P_COLMONEDA)), 1), Trim(frmCXPEmisionPagos.txtTipoCambio.Text), Trim(frmCXPEmisionPagos.txtTipoCambioEuro.Text), Trim(.get_TextMatrix(I, P_COLSUBTOTALDF)), Trim(.get_TextMatrix(I, P_COLIVADF)), C_STVIGENTE, Format(#1/1/1900#, C_FORMATFECHAGUARDAR), fTipoPagoProg, CStr(IIf(Me._optFormaPago_0.Checked, True, False)), CStr(False), Format(#1/1/1900#, C_FORMATFECHAGUARDAR), Trim(Me.txtFolioEgreso.Text), "0", C_INSERCION, CStr(0))
                        Cmd.Execute()
                        'Si el pago menos el descuento es mayor o igual que el saldo, cambia el estatus del pago a Pagado ("P")
                        If CDec(Numerico(.get_TextMatrix(I, P_COLIMPORTEPAGO))) >= CDec(Numerico(.get_TextMatrix(I, P_COLSALDO))) Then
                            '(CCur(Numerico(.TextMatrix(i, P_COLIMPORTEPAGO))) + CCur(Numerico(.TextMatrix(i, P_COLDESCTOFIN))))
                            ModStoredProcedures.PR_IMEProgramacionPagos(cFolioPago, CStr(nNumPartida), "0", "", "", "", "01/01/1900", "01/01/1900", "0", "", "0", "0", "0", "0", "0", C_STPAGADO, "01/01/1900", "", CStr(IIf(Me._optFormaPago_0.Checked, True, False)), CStr(False), Format(#1/1/1900#, C_FORMATFECHAGUARDAR), C_MODIFICACION, CStr(3))
                            Cmd.Execute()
                        End If
                    End If
                Next I
            End With
            With frmCXPEmisionPagos.mshNotasCredito
                For I = 1 To .get_Cols() - 1
                    If Trim(.get_TextMatrix(I, N_COLFOLIO)) = "" Then
                        Exit For
                    End If
                    If Trim(.get_TextMatrix(I, N_COLESTATUS)) = C_PAGAR Then
                        ModStoredProcedures.PR_IMENotasCreditoCab(Trim(.get_TextMatrix(I, N_COLFOLIO)), Format(#1/1/1900#, C_FORMATFECHAGUARDAR), "", "0", "", "", "", "0", "0", "0", "0", C_STAPLICADA, Format(#1/1/1900#, C_FORMATFECHAGUARDAR), "0", "0", Trim(frmCXPEmisionPagos.txtTipoCambio.Text), Format(Today, C_FORMATFECHAGUARDAR), Trim(Me.txtFolioEgreso.Text), "", Trim(frmCXPEmisionPagos.txtTipoCambioEuro.Text), C_MODIFICACION, CStr(3))
                        Cmd.Execute()
                    End If
                Next I
            End With
            With frmCXPEmisionPagos.mshAnticipos
                For I = 1 To .get_Cols() - 1
                    If Trim(.get_TextMatrix(I, A_COLFOLIO)) = "" Then
                        Exit For
                    End If
                    If Trim(.get_TextMatrix(I, A_COLESTATUS)) = C_PAGAR Then
                        ModStoredProcedures.PR_IME_Anticipos(Trim(.get_TextMatrix(I, A_COLFOLIO)), Format(#1/1/1900#, "mm/dd/yyyy"), "", CStr(nEP_CodProveedor), "", cEP_Moneda, "0", "0", "0", "0", C_STAPLICADA, Format(#1/1/1900#, "mm/dd/yyyy"), "0", Trim(frmCXPEmisionPagos.txtTipoCambio.Text), Format(Me.dtpFecha.Value, "mm/dd/yyyy"), Trim(Me.txtFolioEgreso.Text), C_MODIFICACION, CStr(2))
                        Cmd.Execute()
                    End If
                Next I
            End With
            frmCXPEmisionPagos.Limpiar()
        End If
        'Guardar los Movimientos de Origen y Aplicación
        If Not frmPagos2.GuardarMovimientosOrigenAplicacion("REGISTRO DE PAGOS") Then
            Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Exit Function
        End If
        'Cancelar el Movimiento si esta Activada la Opción de Cancelar Cheque
        If chkCancelado.CheckState = 1 Then
            gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVCANCELACION & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                strFolioCancelacion = C_TIPOMOVCANCELACION & Format(Year(CDate(dtpFecha.Value)), "0000") & Format(Month(CDate(dtpFecha.Value)), "00") & Format(VB.Day(CDate(dtpFecha.Value)), "00") & Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
                ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVCANCELACION, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
                Cmd.Execute()
            End If
            'Guardar el Movimiento Bancario de Cancelación
            ModStoredProcedures.PR_IMEMovimientosBancarios(strFolioCancelacion, Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_MOVCANCELACION, C_TIPOMOVEGRESO, IIf(_optFormaPago_0.Checked, C_NATURALEZAINTERNA, C_NATURALEZACOMERCIAL), IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(gcurCorpoTIPOCAMBIODOLAR), IIf(_optFormaPago_0.Checked, C_FORMAPAGOEFECTIVO, IIf(_optFormaPago_1.Checked, C_FORMAPAGOCHEQUE, C_FORMAPAGOELECTRONICO)), IIf(_optTipoPago_0.Checked, C_TIPOPAGOJOYERIA, C_TIPOPAGOPERSONAL), CStr(intCodBanco), dbcCuentaBancaria.Text, txtBeneficiario.Text, QuitaEnter(txtConceptoCancelacion.Text), "0", "", CStr(intNumPartida), IIf(_optFormaPago_1.Checked, Format(dtpFechaCheque.Value, C_FORMATFECHAGUARDAR), "01/01/1900"), txtNumeroCheque.Text, CStr(CDbl(Numerico(txtImporte.Text)) * -1), "V", "01/01/1900", "", "1", Format(FechaFinal, C_FORMATFECHAGUARDAR), C_MODULOBANCOS, txtFolioEgreso.Text, IIf(_optFormaPago_2.Checked = True, txtFolioElectronico.Text, ""), C_INSERCION, CStr(0))
            Cmd.Execute()
            'Restaurar el Folio de Programación
            '        If Trim(txtFolioProgramacion) <> "" Then
            '            ModStoredProcedures.PR_IMEProgramacionPagos txtFolioProgramacion, CStr(intNumPartida), "0", "", "", "", "01/01/1900", "01/01/1900", _
            ''            "0", "", "0", "0", "0", "0", "0", "V", "01/01/1900", "", "0", "0", "01/01/1900", C_MODIFICACION, 3
            '            Cmd.Execute
            '            'Cancelar en Pagos
            '            ModStoredProcedures.PR_IMEPagos txtFolioProgramacion, CStr(intNumPartida), "0", "", "", "", "01/01/1900", "01/01/1900", "0", "", "0", "0", _
            ''            "0", "0", "C", Format(dtpFecha, C_FORMATFECHAGUARDAR), "", "0", "0", "01/01/1900", "", "0", C_MODIFICACION, 0
            '            Cmd.Execute
            '        End If
            'Cancelar los Movimientos de Origen y Aplicación
            ModStoredProcedures.PR_IMEMovimientosOrigenAplic(txtFolioEgreso.Text, "0", "0", "0", "0", "", "0", "C", Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_MODIFICACION, CStr(0))
            Cmd.Execute()
            'Conciliar el Movimiento Cancelado
            ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioEgreso.Text, "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "1", Format(FechaFinal, C_FORMATFECHAGUARDAR), "", "", "", C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False

        If chkCancelado.CheckState = 0 Then
            MsgBox("Los Datos se Han Guardado con Exito" & Chr(13) & "Se ha Generado el Folio de Egreso " & txtFolioEgreso.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        ElseIf chkCancelado.CheckState = 1 Then
            MsgBox("Los Datos se Han Guardado con Exito" & Chr(13) & "Se ha Generado el Folio de Egreso " & txtFolioEgreso.Text & Chr(13) & "Y el Folio de Cancelacion " & strFolioCancelacion, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        End If
        Limpiar()
        If blnEmisionPagos Then
            frmCXPEmisionPagos.dbcProveedor.Text = ""
            frmCXPEmisionPagos.dbcProveedor.Tag = ""
            '''frmCXPEmisionPagos.dbcProveedor.SetFocus   OJO
            frmCXPEmisionPagos.Limpiar()
            Me.Close()
            frmCXPEmisionPagos.dbcProveedor.Focus()
            Exit Function
        End If
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    'Sub LlenaDatosProgramacion()
    '    On Local Error GoTo MErr
    ''    If Trim(txtFolioProgramacion) = "" Then
    ''        Exit Sub
    ''    End If
    '    gStrSql = "SELECT * " & _
    ''    "FROM ProgramacionPagos PP, CatProvAcreed CP WHERE PP.FolioProgramacionP = '" & txtFolioProgramacion & "' AND " & _
    ''    "CP.CodProvAcreed = PP.CodProvAcreed AND PP.Estatus = 'V' AND PP.PasoBancos = 0 AND PP.NumPartida = " & intNumPartida
    '    ModEstandar.BorraCmd
    '    Cmd.CommandText = "dbo.Up_Select_Datos"
    '    Cmd.CommandType = adCmdStoredProc
    '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
    '    Set RsGral = Cmd.Execute
    '    If RsGral.RecordCount > 0 Then
    '        txtBeneficiario = Trim(RsGral!DescProvACreed)
    '        If RsGral!TipoGasto = C_TIPOPAGOJOYERIA Then
    '            _optTipoPago_0.Value = True
    '        ElseIf RsGral!TipoGasto = C_TIPOPAGOPERSONAL Then
    '            _optTipoPago_1.Value = True
    '        End If
    '        dtpFechaCheque.Value = Format(RsGral!FechaPago, C_FORMATFECHAMOSTRAR)
    '        txtImporte = Format(RsGral!TotalPago, "###,##0.00")
    '        MonedaProgramacion = Trim(RsGral!Moneda)
    '        intNumPartida = RsGral!NumPartida
    '        If lblMoneda.Caption <> "" Then
    '            If MonedaProgramacion <> Left(lblMoneda.Caption, 1) Then
    '                MsgBox "Este Folio fue Programado para ser Pagado en " & IIf(MonedaProgramacion = C_PESO, C_DESCPESOS, C_DESCDOLARES) & Chr(13) & _
    ''                "La Cuenta Bancaria Seleccionada Maneja " & lblMoneda.Caption & Chr(13) & _
    ''                "Favor de Seleccionar Una Cuenta Valida ...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
    '                'dbcCuentaBancaria.SetFocus
    '                If PierdeFoco = False Then
    '                    ModEstandar.RetrocederTab Me
    '                End If
    '                Me.Refresh
    '            End If
    '        End If
    '        fFolioProgramacionP = RsGral!FolioProgramacionP
    '        fCodProvAcreed = RsGral!CodProvAcreed
    '        fTipoFacturaCxP = RsGral!TipoFacturaCxP
    '        fTipoGasto = RsGral!TipoGasto
    '        fFolioFactura = RsGral!FolioFactura
    '        fFechaFactura = RsGral!FechaFactura
    '        fFechaPago = RsGral!FechaPago
    '        fTotalPago = RsGral!TotalPago
    '        fMoneda = RsGral!Moneda
    '        fTipoCambio = RsGral!TipoCambio
    '        fTipoCambioE = RsGral!TipoCambioE
    '        fDescuentoFinanciero = RsGral!DescuentoFinanciero
    '        fSubTotalDF = RsGral!SubTotalDF
    '        fIvaDF = RsGral!IvaDF
    '        fEstatus = RsGral!Estatus
    '        fFechaCancel = RsGral!FechaCancel
    '        fTipoPagoProg = RsGral!TipoPagoProg
    '        fEfectivo = RsGral!Efectivo
    '    Else
    '        MsgBox "Folio de Programación no Existe, Favor de Verificar.", vbOKOnly + vbInformation, gstrNombCortoEmpresa
    '        txtFolioProgramacion = ""
    '        txtFolioProgramacion.SetFocus
    '    End If
    'MErr:
    '    If Err.Number <> 0 Then ModEstandar.MostrarError
    'End Sub

    Sub LlenaDatos()

        If (bandera = True) Then
            Exit Sub
        End If

        On Error GoTo Merr
        Dim RsAux As New ADODB.Recordset
        Dim I As Integer
        Dim Total As Decimal
        If Trim(txtFolioEgreso.Text) = "" Then
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT * FROM MovimientosBancarios MB,CatBancos CB WHERE MB.FolioMovto = '" & txtFolioEgreso.Text & "' AND MB.Movimiento = '" & C_MOVPAGO & "' AND " & "MB.TipoMovto = '" & C_TIPOMOVEGRESO & "' AND CB.CodBanco = MB.CodBanco"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            gStrSql = "SELECT * FROM MovimientosBancarios WHERE Referencia = '" & txtFolioEgreso.Text & "' AND Movimiento = '" & C_MOVCANCELACION & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsAux = Cmd.Execute
            If RsAux.RecordCount > 0 Then
                lblCancelada.Text = "Movimiento de Cancelación : " & RsAux.Fields("FolioMovto").Value
                If RsGral.Fields("FechaMovto").Value = RsAux.Fields("FechaMovto").Value And RsGral.Fields("FormaPago").Value = C_FORMAPAGOCHEQUE Then
                    txtConceptoCancelacion.Text = RsAux.Fields("Concepto").Value
                End If
            End If
            dtpFecha.Value = Format(RsGral.Fields("FechaMovto").Value, C_FORMATFECHAMOSTRAR)
            If RsGral.Fields("FormaPago").Value = C_FORMAPAGOEFECTIVO Then
                _optFormaPago_0.Checked = True
            ElseIf RsGral.Fields("FormaPago").Value = C_FORMAPAGOCHEQUE Then
                _optFormaPago_1.Checked = True
            ElseIf RsGral.Fields("FormaPago").Value = C_FORMAPAGOELECTRONICO Then
                _optFormaPago_2.Checked = True
            End If
            If RsGral.Fields("TipoPago").Value = C_TIPOPAGOJOYERIA Then
                _optTipoPago_0.Checked = True
            ElseIf RsGral.Fields("TipoPago").Value = C_TIPOPAGOPERSONAL Then
                _optTipoPago_1.Checked = True
            End If
            dbcBanco.Text = Trim(RsGral.Fields("DescBanco").Value)
            dbcCuentaBancaria.Text = Trim(RsGral.Fields("CtaBancaria").Value)
            txtBeneficiario.Text = Trim(RsGral.Fields("Beneficiario").Value)
            txtConcepto.Text = Trim(RsGral.Fields("Concepto").Value)
            '        If RsGral!PagoProgramado Then
            '            chkPagoProgramado.Value = 1
            '            txtFolioProgramacion = RsGral!FolioProgramacion
            '        Else
            '            chkPagoProgramado.Value = 0
            '            txtFolioProgramacion = ""
            '        End If
            If Trim(RsGral.Fields("NoDocto").Value) <> "" Then
                txtNumeroCheque.Text = RsGral.Fields("NoDocto").Value
                dtpFechaCheque.Value = Format(RsGral.Fields("FechaDocto").Value, C_FORMATFECHAMOSTRAR)
            End If
            txtImporte.Text = Format(RsGral.Fields("importe").Value, "###,##0.00")

            Frame1.Enabled = False

            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMoneda.Text = C_DESCPESOS
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMoneda.Text = C_DESCDOLARES
            End If
            If RsGral.Fields("FormaPago").Value = C_FORMAPAGOELECTRONICO Then
                txtFolioElectronico.Text = Trim(RsGral.Fields("FolioElectronico").Value)
            End If
            gStrSql = "SELECT * FROM MovimientosOrigenAplic MO,CatOrigenAplicRecursos CO,CatRubrosOrigenAplicRecursos CR " & "WHERE FolioMovto = '" & Trim(txtFolioEgreso.Text) & "' AND CO.CodOrigenAplicR = MO.CodOrigenAplicR AND CR.CodRubro = MO.CodRubro AND CO.CodOrigenAplicR = CR.CodOrigAplicR"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then

                With frmPagos2.flexDetalle
                    I = 1
                    .Row = 1
                    frmPagos2.lblTotal.Text = "0.00"
                    Do While Not RsGral.EOF
                        .set_TextMatrix(.Row, 0, VB6.Format(RsGral.Fields("CodOrigenAplicR").Value, "0000"))
                        .set_TextMatrix(.Row, 1, Trim(RsGral.Fields("DescOrigenAplicR").Value))
                        .set_TextMatrix(.Row, 2, VB6.Format(RsGral.Fields("CodRubro").Value, "000000"))
                        .set_TextMatrix(.Row, 3, Trim(RsGral.Fields("DescRubro").Value))
                        .set_TextMatrix(.Row, 4, VB6.Format(RsGral.Fields("importe").Value, "###,##0.00"))
                        With frmPagos2
                            .lblTotal.Text = CStr(CDec(Numerico(VB6.Format(.lblTotal.Text, "#####0.00"))) + CDbl(VB6.Format(RsGral.Fields("importe").Value, "###,##0.00")))
                        End With
                        If .Row = .Rows - 1 Then
                            .Rows = .Rows + 1
                        End If
                        .Row = .Row + 1
                        I = I + 1
                        RsGral.MoveNext()
                    Loop
                    frmPagos2.lblTotal.Text = VB6.Format(frmPagos2.lblTotal.Text, "###,##0.00")
                    frmPagos2.lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                    frmPagos2.Nuevo = True
                End With
            End If
            txtImporte.Enabled = False
            cmdOrigenyAplicacion.Enabled = True
            mblnNuevo = False
            dtpFecha.Enabled = False
            ConsultaPagos = True
        Else
            MsgBox("Folio de Movimiento de Egreso no Existe ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFolioEgreso.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Not BuscaUltimoCierre(dtpFecha.Value) Then
            Exit Function
        End If
        If Len(Trim(dbcBanco.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Nombre del Banco", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dbcBanco.Focus()
            Exit Function
        End If
        If Len(Trim(dbcCuentaBancaria.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Cuenta Bancaria", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dbcCuentaBancaria.Focus()
            Exit Function
        End If
        If Len(Trim(txtBeneficiario.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Beneficiario", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtBeneficiario.Focus()
            Exit Function
        End If
        If Len(Trim(txtConcepto.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Concepto", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtConcepto.Focus()
            Exit Function
        End If
        If Me.dtpFechaCheque.Value < Today Then
            MsgBox("La fecha de expedición del cheque no debe ser menor a la fecha actual", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Me.dtpFechaCheque.Focus()
            Exit Function
        End If
        If CDbl(Numerico(txtImporte.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Importe del Pago", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtImporte.Focus()
            Exit Function
        End If
        '    If chkPagoProgramado.Value = 1 Then
        '        If Len(Trim(txtFolioProgramacion)) < 15 Then
        '            MsgBox "Folio de Programación no Valido...", vbInformation + vbOKOnly, gstrNombCortoEmpresa
        '            txtFolioProgramacion.SetFocus
        '            Exit Function
        '        End If
        '        If Trim(MonedaProgramacion) <> Left(lblMoneda.Caption, 1) Then
        '            MsgBox "La Moneda de la Cuenta Bancaria no Coincide con la del Folio de Programación " & Chr(13) & _
        ''                   "                     Seleccione Una Cuenta Bancaria Valida", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '            dbcCuentaBancaria.SetFocus
        '            Exit Function
        '        End If
        '    End If
        If _optFormaPago_2.Checked = True Then
            If Trim(txtFolioElectronico.Text) = "" Then
                MsgBox("Proporcione el Folio Electrónico...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtFolioElectronico.Focus()
                Exit Function
            End If
        End If
        If chkCancelado.CheckState = 1 Then
            If Trim(txtConceptoCancelacion.Text) = "" Then
                MsgBox("Proporcione el Concepto de Cancelación del Cheque...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtConceptoCancelacion.Focus()
                Exit Function
            End If
        End If
        If Not ChecaGrid(frmPagos2) Then
            MsgBox("No se Han Capturado los Movimientos de Origen y Aplicación ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            cmdOrigenyAplicacion_Click(cmdOrigenyAplicacion, New System.EventArgs())
            Exit Function
        End If
        If Numerico((frmPagos2.lblImporte).Text) <> Numerico((frmPagos2.lblTotal).Text) Then
            MsgBox("El Total de los Movimientos de Origen y Aplicación no es Igual al Importe del Pago...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            cmdOrigenyAplicacion_Click(cmdOrigenyAplicacion, New System.EventArgs())
            Exit Function
        End If
        If Not ChecaSaldo(CInt(intCodBanco), Trim(dbcCuentaBancaria.Text), CDec(txtImporte.Text)) Then
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Sub Limpiar()
        Nuevo()
        txtFolioEgreso.Text = ""
        txtFolioEgreso.Focus()
    End Sub

    Sub Nuevo()

        If (bandera = True) Then
            Exit Sub
        End If

        txtFolioEgreso.Text = ""
        lblMoneda.Text = ""
        'dtpFecha.Value = Format(Date.Now(), "dd/mmm/yyyy")
        'dtpFecha.Value = Format(DateTime.Now, "dd/mmm/yyyy")
        lblCancelada.Text = ""
        _optTipoPago_0.Enabled = True
        _optTipoPago_1.Enabled = True
        _optFormaPago_0.Checked = True
        _optTipoPago_0.Checked = True
        dbcBanco.Text = ""
        'dbcBanco.RowSource = Nothing
        dbcCuentaBancaria.Text = ""
        'dbcCuentaBancaria.RowSource = Nothing
        txtBeneficiario.Text = ""
        txtBeneficiario.ReadOnly = False
        txtImporte.ReadOnly = False
        txtConcepto.Text = ""
        'chkPagoProgramado.Value = 0
        'txtFolioProgramacion = ""
        'txtFolioProgramacion.Enabled = False
        chkCancelado.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkCancelado.Enabled = False
        dtpFechaCheque.Value = Now
        dtpFechaCheque.Enabled = False
        txtConceptoCancelacion.Enabled = False
        txtConceptoCancelacion.Text = ""
        txtNumeroCheque.Text = ""
        txtNumeroCheque.Enabled = False
        txtFolioElectronico.Enabled = False
        txtFolioElectronico.Text = ""
        txtImporte.Enabled = True
        txtImporte.Text = "0.00"
        Frame1.Enabled = True
        InicializaVariables()
        gblnSalir = True
        'frmPagos2.Close()
        gblnSalir = False
        frmPagos2.Nuevo = False
        'ConsultaPagos = False
    End Sub

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnSalir = False
        FueraChange = False
        intCodBanco = 0
        MonedaProgramacion = ""
    End Sub


    Private Sub chkCancelado_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCancelado.CheckStateChanged
        If chkCancelado.CheckState = 1 Then
            txtConceptoCancelacion.Enabled = True
        ElseIf chkCancelado.CheckState = 0 Then
            txtConceptoCancelacion.Enabled = False
            txtConceptoCancelacion.Text = ""
        End If
    End Sub

    'Private Sub chkPagoProgramado_Click()
    '    If chkPagoProgramado.Value = 1 Then
    '        txtFolioProgramacion.Enabled = True
    '        txtBeneficiario.Locked = True
    '        txtImporte.Locked = True
    '        _optTipoPago_0.Enabled = False
    '        _optTipoPago_1.Enabled = False
    '    ElseIf chkPagoProgramado.Value = 0 Then
    '        txtFolioProgramacion.Enabled = False
    '        txtFolioProgramacion = ""
    '        txtBeneficiario.Locked = False
    '        txtImporte.Locked = False
    '        _optTipoPago_0.Enabled = True
    '        _optTipoPago_1.Enabled = True
    '    End If
    'End Sub

    Private Sub chkPagoProgramado_GotFocus()
        Pon_Tool()
    End Sub

    Private Sub cmdOrigenyAplicacion_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOrigenyAplicacion.Click
        'Dim frmPagos2 As New frmBancosProcesoDiarioOrigenyAplicacion() 
        If Trim(dbcBanco.Text) <> "" And Trim(dbcCuentaBancaria.Text) <> "" Then
            If CDbl(Numerico(txtImporte.Text)) > 0 Then
                If frmPagos2.Nuevo Then
                    frmPagos2.cmdAceptar.TabIndex = 0
                    frmPagos2.flexDetalle.TabIndex = 1
                    frmPagos2.flexDetalle.Enabled = False
                Else
                    frmPagos2.flexDetalle.TabIndex = 0
                    frmPagos2.cmdAceptar.TabIndex = 1
                    frmPagos2.cmdAceptar.Enabled = False
                End If
                frmPagos2.Tag = "frmPagos2"
                frmPagos2.Text = "Aplicación de Recursos (Registro de Pagos)"
                frmPagos2.lblMoneda.Text = lblMoneda.Text
                frmPagos2.lblFechaMovimiento.Text = VB.Format(dtpFecha.Value, "dd/MM/yyyy")
                frmPagos2.lblImporte.Text = txtImporte.Text
                frmPagos2.flexDetalle.Col = 0
                frmPagos2.flexDetalle.Row = 1
                frmPagos2.ShowDialog()
            Else
                MsgBox("El Importe del Pago debe ser Mayor que Cero, Favor de Teclear un Importe ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtImporte.Focus()
            End If
        Else
            MsgBox("Favor de Seleccionar Una Cuenta Bancaria Valida ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcCuentaBancaria.Focus()
        End If
    End Sub


    Private Sub dbcBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Enter
        If _optFormaPago_0.Checked Then
            gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE ControlInterno = 1 ORDER BY DescBanco"
        Else
            gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE ControlInterno = 0 ORDER BY DescBanco"
        End If
        DCGotFocus(gStrSql, dbcBanco)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcBanco_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcBanco.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            If Me._optTipoPago_0.Enabled And Me._optTipoPago_1.Enabled Then
                If _optTipoPago_0.Checked Then
                    _optTipoPago_0.Focus()
                ElseIf _optTipoPago_1.Checked Then
                    _optTipoPago_1.Focus()
                End If
            Else
                If Me._optFormaPago_0.Checked Then
                    Me._optFormaPago_0.Focus()
                ElseIf Me._optFormaPago_1.Checked Then
                    Me._optFormaPago_1.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub dbcBanco_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcBanco.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Leave
        If _optFormaPago_0.Checked Then
            gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' AND ControlInterno = 1 ORDER BY DescBanco"
        Else
            gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' AND ControlInterno = 0 ORDER BY DescBanco"
        End If
        DCLostFocus(dbcBanco, gStrSql, intCodBanco)
    End Sub

    Private Sub dbcCuentaBancaria_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.CursorChanged
        If FueraChange = True Then Exit Sub

        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcCuentaBancaria" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCChange(gStrSql, tecla)
        If Trim(dbcCuentaBancaria.Text) = "" Then
            lblMoneda.Text = ""
        End If
        'intCodBanco = 0
    End Sub

    Private Sub dbcCuentaBancaria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As EventArgs) Handles dbcCuentaBancaria.Enter
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCGotFocus(gStrSql, dbcCuentaBancaria)
        Pon_Tool()
        FueraChange = False
        PierdeFoco = False
    End Sub

    Private Sub dbcCuentaBancaria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCuentaBancaria.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            dbcBanco.Focus()
        End If
    End Sub

    Private Sub dbcCuentaBancaria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcCuentaBancaria.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcCuentaBancaria_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCuentaBancaria.KeyUp
        Dim Aux As String
        Aux = dbcCuentaBancaria.Text
        'If dbcCuentaBancaria.SelectedItem <> 0 Then
        '    PierdeFoco = True
        '    dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        '    PierdeFoco = False
        'End If
        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub dbcCuentaBancaria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Leave
        On Error GoTo Err_Renamed
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCLostFocus(dbcCuentaBancaria, gStrSql, intCodBanco)
        gStrSql = "SELECT Moneda FROM CatCuentasBancarias WHERE CtaBancaria = '" & Trim(dbcCuentaBancaria.Text) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMoneda.Visible = True
                lblMoneda.Text = C_DESCPESOS
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMoneda.Visible = True
                lblMoneda.Text = C_DESCDOLARES
            End If
            cMoneda = RsGral.Fields("Moneda").Value
        End If
        If _optFormaPago_1.Checked Then
            ConsecutivoCheque = ObtieneNumCheque(CInt(intCodBanco), dbcCuentaBancaria.Text, LetraFolio)
            txtNumeroCheque.Text = Trim(LetraFolio) & Format(CStr(ConsecutivoCheque), "000000")
        End If
        If PierdeFoco = True Then Exit Sub
        '    If Trim(txtFolioProgramacion) <> "" Then
        '        If lblMoneda.Caption = "" Then Exit Sub
        '        If Trim(MonedaProgramacion) <> Left(lblMoneda.Caption, 1) Then
        '            MsgBox "La Moneda de la Cuenta Bancaria no Coincide con la del Folio de Programación " & Chr(13) & _
        ''                   "                             Favor de Verificar ...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '        End If
        '    End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcCuentaBancaria_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcCuentaBancaria.MouseUp
        Dim Aux As String
        Aux = dbcCuentaBancaria.Text
        'If dbcCuentaBancaria.SelectedItem <> 0 Then
        '    PierdeFoco = True
        '    dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        '    PierdeFoco = False
        'End If
        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub dtpFecha_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFecha.CursorChanged
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFecha_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFecha.Click
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFecha_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFecha.KeyPress
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaCheque_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaCheque.Enter
        Pon_Tool()
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodePagos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodePagos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodePagos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape

                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtFolioEgreso" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodePagos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodePagos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        bandera = True
        frmPagos2.InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        'gstrMovimiento = "S"
        InicializaVariables()
        Nuevo()
        BuscaEjercicio(dtpFecha.Value)
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodePagos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not blnEmisionPagos Then 'Si el formulario no fue llamado desde Emisión de pagos, cierra normalmente
        '    'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        '    ModEstandar.RestaurarForma(Me, False)
        '    'Si se cierra el formulario y existio algun cambio en el registro se
        '    'informa al usuario del cabio y si desea guardar el registro, ya sea
        '    'que sea nuevo o un registro modificado
        '    If Not mblnSalir Then
        '        'If Cambios = True And mblnNuevo = False Then
        '        'Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
        '        'Case vbYes: 'Guardar el registro
        '        'If Guardar = False Then
        '        'Cancel = 1
        '        'End If
        '        'Case vbNo: 'No hace nada y permite el cierre del formulario
        '        'Case vbCancel: 'Cancela el cierre del formulario sin guardar
        '        'Cancel = 1
        '        'End Select
        '        'End If
        '    Else
        '        Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes
        '                Cancel = 0
        '            Case MsgBoxResult.No
        '                mblnSalir = False
        '                Cancel = 1
        '        End Select
        '    End If
        'Else
        '    Cancel = 0
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodePagos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        If blnEmisionPagos Then
            NuevoPagos()
            blnEmisionPagos = False
        End If
        If Me.Tag = "FRMCXPEMISIONPAGOS" Then
            frmCXPEmisionPagos.Enabled = True
        End If
        'Me = Nothing
        IsNothing(Me)

        'Me.Close()
        Me.Hide()
        gblnSalir = True
        frmPagos2.Close()
        frmPagos2 = Nothing
    End Sub

    Private Sub _optFormaPago_0_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optFormaPago_0.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = 0
            '= optFormaPago.GetIndex(eventSender)
            Select Case Index
                Case 0
                    dbcCuentaBancaria.Text = ""
                    lblMoneda.Text = ""
                    dbcBanco.Text = ""
                    intUltFormaPago = 0
                    chkCancelado.Enabled = False
                    chkCancelado.CheckState = System.Windows.Forms.CheckState.Unchecked
                    dtpFechaCheque.Enabled = False
                    txtNumeroCheque.Enabled = False
                    txtConceptoCancelacion.Enabled = False
                    txtConceptoCancelacion.Text = ""
                    txtFolioElectronico.Enabled = False
                    txtFolioElectronico.Text = ""
                    txtNumeroCheque.Text = ""
                Case 1
                    If intUltFormaPago = 0 Then
                        dbcBanco.Text = ""
                        dbcCuentaBancaria.Text = ""
                        lblMoneda.Text = ""
                    End If
                    intUltFormaPago = 1
                    dtpFechaCheque.Enabled = True
                    txtNumeroCheque.Enabled = True
                    txtFolioElectronico.Text = ""
                    txtFolioElectronico.Enabled = False

                    If dbcCuentaBancaria.Text <> "" Then
                        ConsecutivoCheque = ObtieneNumCheque(CInt(intCodBanco), dbcCuentaBancaria.Text, LetraFolio)
                        txtNumeroCheque.Text = Trim(LetraFolio) & Format(CStr(ConsecutivoCheque), "000000")
                    End If
                    chkCancelado.Enabled = True
                Case 2
                    If intUltFormaPago = 0 Then
                        dbcBanco.Text = ""
                        dbcCuentaBancaria.Text = ""
                        lblMoneda.Text = ""
                    End If
                    intUltFormaPago = 2
                    chkCancelado.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkCancelado.Enabled = False
                    dtpFechaCheque.Enabled = False
                    txtNumeroCheque.Enabled = False
                    txtFolioElectronico.Enabled = True
                    txtNumeroCheque.Text = ""
            End Select
        End If
    End Sub

    Private Sub _optFormaPago_0_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optFormaPago_0.Enter
        Dim Index As Integer = 0
        '= optFormaPago.GetIndex(eventSender)
        Select Case Index
            Case 0
                Pon_Tool()
            Case 1
                Pon_Tool()
            Case 2
                Pon_Tool()
        End Select
    End Sub

    Private Sub _optFormaPago_1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optFormaPago_1.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = 1
            '= optFormaPago.GetIndex(eventSender)
            Select Case Index
                Case 0
                    dbcCuentaBancaria.Text = ""
                    lblMoneda.Text = ""
                    dbcBanco.Text = ""
                    intUltFormaPago = 0
                    chkCancelado.Enabled = False
                    chkCancelado.CheckState = System.Windows.Forms.CheckState.Unchecked
                    dtpFechaCheque.Enabled = False
                    txtNumeroCheque.Enabled = False
                    txtConceptoCancelacion.Enabled = False
                    txtConceptoCancelacion.Text = ""
                    txtFolioElectronico.Enabled = False
                    txtFolioElectronico.Text = ""
                    txtNumeroCheque.Text = ""
                Case 1
                    If intUltFormaPago = 0 Then
                        dbcBanco.Text = ""
                        dbcCuentaBancaria.Text = ""
                        lblMoneda.Text = ""
                    End If
                    intUltFormaPago = 1
                    dtpFechaCheque.Enabled = True
                    txtNumeroCheque.Enabled = True
                    txtFolioElectronico.Text = ""
                    txtFolioElectronico.Enabled = False

                    If dbcCuentaBancaria.Text <> "" Then
                        ConsecutivoCheque = ObtieneNumCheque(CInt(intCodBanco), dbcCuentaBancaria.Text, LetraFolio)
                        txtNumeroCheque.Text = Trim(LetraFolio) & Format(CStr(ConsecutivoCheque), "000000")
                    End If
                    chkCancelado.Enabled = True
                Case 2
                    If intUltFormaPago = 0 Then
                        dbcBanco.Text = ""
                        dbcCuentaBancaria.Text = ""
                        lblMoneda.Text = ""
                    End If
                    intUltFormaPago = 2
                    chkCancelado.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkCancelado.Enabled = False
                    dtpFechaCheque.Enabled = False
                    txtNumeroCheque.Enabled = False
                    txtFolioElectronico.Enabled = True
                    txtNumeroCheque.Text = ""
            End Select
        End If
    End Sub

    Private Sub _optFormaPago_1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optFormaPago_1.Enter
        Dim Index As Integer = 1
        '= optFormaPago.GetIndex(eventSender)
        Select Case Index
            Case 0
                Pon_Tool()
            Case 1
                Pon_Tool()
            Case 2
                Pon_Tool()
        End Select
    End Sub

    Private Sub _optFormaPago_2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optFormaPago_2.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = 2
            '= optFormaPago.GetIndex(eventSender)
            Select Case Index
                Case 0
                    dbcCuentaBancaria.Text = ""
                    lblMoneda.Text = ""
                    dbcBanco.Text = ""
                    intUltFormaPago = 0
                    chkCancelado.Enabled = False
                    chkCancelado.CheckState = System.Windows.Forms.CheckState.Unchecked
                    dtpFechaCheque.Enabled = False
                    txtNumeroCheque.Enabled = False
                    txtConceptoCancelacion.Enabled = False
                    txtConceptoCancelacion.Text = ""
                    txtFolioElectronico.Enabled = False
                    txtFolioElectronico.Text = ""
                    txtNumeroCheque.Text = ""
                Case 1
                    If intUltFormaPago = 0 Then
                        dbcBanco.Text = ""
                        dbcCuentaBancaria.Text = ""
                        lblMoneda.Text = ""
                    End If
                    intUltFormaPago = 1
                    dtpFechaCheque.Enabled = True
                    txtNumeroCheque.Enabled = True
                    txtFolioElectronico.Text = ""
                    txtFolioElectronico.Enabled = False

                    If dbcCuentaBancaria.Text <> "" Then
                        ConsecutivoCheque = ObtieneNumCheque(CInt(intCodBanco), dbcCuentaBancaria.Text, LetraFolio)
                        txtNumeroCheque.Text = Trim(LetraFolio) & Format(CStr(ConsecutivoCheque), "000000")
                    End If
                    chkCancelado.Enabled = True
                Case 2
                    If intUltFormaPago = 0 Then
                        dbcBanco.Text = ""
                        dbcCuentaBancaria.Text = ""
                        lblMoneda.Text = ""
                    End If
                    intUltFormaPago = 2
                    chkCancelado.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkCancelado.Enabled = False
                    dtpFechaCheque.Enabled = False
                    txtNumeroCheque.Enabled = False
                    txtFolioElectronico.Enabled = True
                    txtNumeroCheque.Text = ""
            End Select
        End If
    End Sub

    Private Sub _optFormaPago_2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optFormaPago_2.Enter
        Dim Index As Integer = 2
        '= optFormaPago.GetIndex(eventSender)
        Select Case Index
            Case 0
                Pon_Tool()
            Case 1
                Pon_Tool()
            Case 2
                Pon_Tool()
        End Select
    End Sub

    Private Sub optTipoPago_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTipoPago.Enter
        Dim Index As Integer = optTipoPago.GetIndex(eventSender)
        Select Case Index
            Case 0
                Pon_Tool()
            Case 1
                Pon_Tool()
        End Select
    End Sub

    Private Sub txtBeneficiario_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBeneficiario.Enter
        Pon_Tool()
    End Sub

    Private Sub txtBeneficiario_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBeneficiario.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "!""#$%&/()=?'¡¿*,;.:<>@+-_")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtConcepto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.Enter
        Pon_Tool()
    End Sub

    Private Sub txtConcepto_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConcepto.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "!""#$%&/()=?'¡¿*,;.:<>@+-_")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtConceptoCancelacion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConceptoCancelacion.Enter
        SelTextoTxt(txtConceptoCancelacion)
        Pon_Tool()
    End Sub

    Private Sub txtConceptoCancelacion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConceptoCancelacion.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Concepto = txtConceptoCancelacion.Text
        End If
        'If KeyAscii = vbKeyEscape And Trim(txtConceptoCancelacion) = "" Then
        '    Concepto = ""
        'End If
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "!""#$%&/()=?'¡¿*,;.:<>@+-_")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtConceptoCancelacion_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConceptoCancelacion.Leave
        If Trim(Concepto) <> "" Then
            txtConceptoCancelacion.Text = Concepto
        End If
        Trim(QuitaEnter(txtConceptoCancelacion.Text))
        Concepto = ""
    End Sub

    Private Sub txtFolioEgreso_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEgreso.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtFolioEgreso_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEgreso.Enter
        strControlActual = UCase("txtFolioEgreso")
        SelTextoTxt(txtFolioEgreso)
        Pon_Tool()
    End Sub

    Private Sub txtFolioEgreso_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioEgreso.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, C_TIPOMOVEGRESO)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioEgreso_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEgreso.Leave

        If Me.ActiveControl.Name = "btnBuscar" Then
            Exit Sub
        End If

        If Trim(txtFolioEgreso.Text) = "" Then
            txtFolioEgreso.Text = C_TIPOMOVEGRESO & Format(Year(CDate(dtpFecha.Value)), "0000") & Format(Month(CDate(dtpFecha.Value)), "00") & Format(VB.Day(CDate(dtpFecha.Value)), "00") & "0000"
            bandera = True
            Exit Sub
        End If

        If mblnCambiosEnCodigo = True And txtFolioEgreso.Text <> "" And VB.Right(txtFolioEgreso.Text, 4) <> "0000" Then
            LlenaDatos()
            frmPagos2.Hide()
            Me.BringToFront()
        End If
    End Sub


    'Private Sub txtFolioProgramacion_Change()
    '    MonedaProgramacion = ""
    'End Sub

    'Private Sub txtFolioProgramacion_GotFocus()
    '    SelTextoTxt txtFolioProgramacion
    '    Pon_Tool
    'End Sub

    'Private Sub txtFolioProgramacion_KeyPress(KeyAscii As Integer)
    '    ModEstandar.gp_CampoNumerico KeyAscii, "X"
    'End Sub

    'Private Sub txtFolioProgramacion_LostFocus()
    '    If FueraChange Then
    '        FueraChange = False
    '        Exit Sub
    '    End If
    '    If Trim(txtFolioProgramacion) <> "" Then
    '        PierdeFoco = True
    '        LlenaDatosProgramacion
    '        PierdeFoco = False
    '    End If
    'End Sub

    Private Sub txtImporte_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.TextChanged
        If Trim(txtImporte.Text) = "" Then
            txtImporte.Text = "0.00"
        End If
    End Sub

    Private Sub txtImporte_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Enter
        SelTextoTxt(txtImporte)
        Pon_Tool()
    End Sub

    Private Sub txtImporte_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImporte.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.MskCantidad(txtImporte.Text, KeyAscii, 15, 2, (txtImporte.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtImporte_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Leave
        txtImporte.Text = VB6.Format(txtImporte.Text, "###,##0.00")
    End Sub

    Private Sub txtNumeroCheque_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumeroCheque.Enter
        SelTextoTxt(txtNumeroCheque)
        Pon_Tool()
    End Sub



    'Private Sub dbcBanco_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dbcBanco.SelectedIndexChanged
    '    If FueraChange = True Then Exit Sub

    '    If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcCuentaBancaria" Then
    '        Exit Sub
    '    End If
    '    gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
    '    DCChange(gStrSql, tecla)
    '    If Trim(dbcCuentaBancaria.Text) = "" Then
    '        lblMoneda.Text = ""
    '    End If
    '    'intCodBanco = 0
    'End Sub

    'Private Sub dbcCuentaBancaria_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dbcCuentaBancaria.SelectedIndexChanged
    '    If FueraChange = True Then Exit Sub

    '    If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcCuentaBancaria" Then
    '        Exit Sub
    '    End If
    '    gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
    '    DCChange(gStrSql, tecla)
    '    If Trim(dbcCuentaBancaria.Text) = "" Then
    '        lblMoneda.Text = ""
    '    End If
    '    'intCodBanco = 0
    'End Sub


    'Private Sub dbcBanco_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dbcBanco.SelectedIndexChanged

    Private Sub dbcBanco_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.CursorChanged
        If FueraChange = True Then Exit Sub

        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcBanco" Then
            Exit Sub
        End If
        dbcCuentaBancaria.Text = ""
        lblMoneda.Text = ""
        If _optFormaPago_0.Checked Then
            gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' AND ControlInterno = 1 ORDER BY DescBanco"
        Else
            gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' AND ControlInterno = 0 ORDER BY DescBanco"
        End If
        DCChange(gStrSql, tecla)
        intCodBanco = 0
    End Sub

    Private Sub dbcBanco_click(sender As Object, e As EventArgs) Handles dbcBanco.Click

        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCGotFocus(gStrSql, dbcCuentaBancaria)
        Pon_Tool()
        FueraChange = False
        PierdeFoco = False

        '        On Error GoTo Err
        '        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        '        DCLostFocus(dbcCuentaBancaria, gStrSql, intCodBanco)
        '        gStrSql = "SELECT Moneda FROM CatCuentasBancarias WHERE CtaBancaria = '" & Trim(dbcCuentaBancaria.Text) & "'"
        '        ModEstandar.BorraCmd()
        '        Cmd.CommandText = "dbo.Up_Select_Datos"
        '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        '        RsGral = Cmd.Execute
        '        If RsGral.RecordCount > 0 Then
        '            If RsGral.Fields("Moneda").Value = C_PESO Then
        '                lblMoneda.Visible = True
        '                lblMoneda.Text = C_DESCPESOS
        '            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
        '                lblMoneda.Visible = True
        '                lblMoneda.Text = C_DESCDOLARES
        '            End If
        '            cMoneda = RsGral.Fields("Moneda").Value
        '        End If


        '        If _optFormaPago_1.Checked Then
        '            ConsecutivoCheque = ObtieneNumCheque(CInt(intCodBanco), dbcCuentaBancaria.Text, LetraFolio)
        '            txtNumeroCheque.Text = Trim(LetraFolio) & Format(CStr(ConsecutivoCheque), "000000")
        '            ModStoredProcedures.PR_IMECatCuentasBancarias(CStr(intCodBanco), dbcCuentaBancaria.Text, "", "", "", "", "0", Format(CStr(ConsecutivoCheque), "000000"), "", C_MODIFICACION, CStr(1))
        '            Cmd.Execute()
        '        End If

        '        If PierdeFoco = True Then Exit Sub
        '        '    If Trim(txtFolioProgramacion) <> "" Then
        '        '        If lblMoneda.Caption = "" Then Exit Sub
        '        '        If Trim(MonedaProgramacion) <> Left(lblMoneda.Caption, 1) Then
        '        '            MsgBox "La Moneda de la Cuenta Bancaria no Coincide con la del Folio de Programación " & Chr(13) & _
        '        '                   "                             Favor de Verificar ...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '        '        End If
        '        '    End If
        'Err:
        '        If Err.Number <> 0 Then ModEstandar.MostrarError()


    End Sub

    Private Sub dbcCuentaBancaria_click(sender As Object, e As EventArgs) Handles dbcCuentaBancaria.Click
        'If FueraChange = True Then Exit Sub
        'gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) &
        '    "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        'DCChange(gStrSql, tecla)

        'If Trim(dbcCuentaBancaria.Text) = "" Then
        '    'lblMoneda.Caption = ""
        'End If
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtFolioEgreso = New System.Windows.Forms.TextBox()
        Me.txtFolioElectronico = New System.Windows.Forms.TextBox()
        Me._optFormaPago_0 = New System.Windows.Forms.RadioButton()
        Me._optFormaPago_1 = New System.Windows.Forms.RadioButton()
        Me._optFormaPago_2 = New System.Windows.Forms.RadioButton()
        Me.txtImporte = New System.Windows.Forms.TextBox()
        Me.txtConceptoCancelacion = New System.Windows.Forms.TextBox()
        Me.txtNumeroCheque = New System.Windows.Forms.TextBox()
        Me.txtConcepto = New System.Windows.Forms.TextBox()
        Me.txtBeneficiario = New System.Windows.Forms.TextBox()
        Me._optTipoPago_1 = New System.Windows.Forms.RadioButton()
        Me._optTipoPago_0 = New System.Windows.Forms.RadioButton()
        Me.cmdOrigenyAplicacion = New System.Windows.Forms.Button()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblMoneda = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Frame6 = New System.Windows.Forms.GroupBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.chkCancelado = New System.Windows.Forms.CheckBox()
        Me.dtpFechaCheque = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dbcCuentaBancaria = New System.Windows.Forms.ComboBox()
        Me.lblCancelada = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.optFormaPago = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optTipoPago = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Frame4.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.Frame6.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        CType(Me.optFormaPago, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optTipoPago, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtFolioEgreso
        '
        Me.txtFolioEgreso.AcceptsReturn = True
        Me.txtFolioEgreso.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioEgreso.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioEgreso.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioEgreso.Location = New System.Drawing.Point(99, 14)
        Me.txtFolioEgreso.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioEgreso.MaxLength = 13
        Me.txtFolioEgreso.Name = "txtFolioEgreso"
        Me.txtFolioEgreso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioEgreso.Size = New System.Drawing.Size(215, 20)
        Me.txtFolioEgreso.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtFolioEgreso, "Folio del Egreso.")
        '
        'txtFolioElectronico
        '
        Me.txtFolioElectronico.AcceptsReturn = True
        Me.txtFolioElectronico.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioElectronico.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioElectronico.Enabled = False
        Me.txtFolioElectronico.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioElectronico.Location = New System.Drawing.Point(37, 23)
        Me.txtFolioElectronico.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioElectronico.MaxLength = 20
        Me.txtFolioElectronico.Name = "txtFolioElectronico"
        Me.txtFolioElectronico.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioElectronico.Size = New System.Drawing.Size(171, 20)
        Me.txtFolioElectronico.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.txtFolioElectronico, "Folio Electrónico del Banco.")
        '
        '_optFormaPago_0
        '
        Me._optFormaPago_0.BackColor = System.Drawing.SystemColors.Control
        Me._optFormaPago_0.Checked = True
        Me._optFormaPago_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optFormaPago_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFormaPago.SetIndex(Me._optFormaPago_0, CType(0, Short))
        Me._optFormaPago_0.Location = New System.Drawing.Point(9, 15)
        Me._optFormaPago_0.Margin = New System.Windows.Forms.Padding(2)
        Me._optFormaPago_0.Name = "_optFormaPago_0"
        Me._optFormaPago_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optFormaPago_0.Size = New System.Drawing.Size(74, 21)
        Me._optFormaPago_0.TabIndex = 2
        Me._optFormaPago_0.TabStop = True
        Me._optFormaPago_0.Text = "Efectivo"
        Me.ToolTip1.SetToolTip(Me._optFormaPago_0, "Pago en Efectivo.")
        Me._optFormaPago_0.UseVisualStyleBackColor = False
        '
        '_optFormaPago_1
        '
        Me._optFormaPago_1.BackColor = System.Drawing.SystemColors.Control
        Me._optFormaPago_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optFormaPago_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFormaPago.SetIndex(Me._optFormaPago_1, CType(1, Short))
        Me._optFormaPago_1.Location = New System.Drawing.Point(164, 15)
        Me._optFormaPago_1.Margin = New System.Windows.Forms.Padding(2)
        Me._optFormaPago_1.Name = "_optFormaPago_1"
        Me._optFormaPago_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optFormaPago_1.Size = New System.Drawing.Size(62, 20)
        Me._optFormaPago_1.TabIndex = 3
        Me._optFormaPago_1.TabStop = True
        Me._optFormaPago_1.Text = "Cheque"
        Me.ToolTip1.SetToolTip(Me._optFormaPago_1, "Pago con Cheque.")
        Me._optFormaPago_1.UseVisualStyleBackColor = False
        '
        '_optFormaPago_2
        '
        Me._optFormaPago_2.BackColor = System.Drawing.SystemColors.Control
        Me._optFormaPago_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optFormaPago_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFormaPago.SetIndex(Me._optFormaPago_2, CType(2, Short))
        Me._optFormaPago_2.Location = New System.Drawing.Point(288, 18)
        Me._optFormaPago_2.Margin = New System.Windows.Forms.Padding(2)
        Me._optFormaPago_2.Name = "_optFormaPago_2"
        Me._optFormaPago_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optFormaPago_2.Size = New System.Drawing.Size(92, 15)
        Me._optFormaPago_2.TabIndex = 4
        Me._optFormaPago_2.TabStop = True
        Me._optFormaPago_2.Text = "Electrónico"
        Me.ToolTip1.SetToolTip(Me._optFormaPago_2, "Pago Electrónico.")
        Me._optFormaPago_2.UseVisualStyleBackColor = False
        '
        'txtImporte
        '
        Me.txtImporte.AcceptsReturn = True
        Me.txtImporte.BackColor = System.Drawing.SystemColors.Window
        Me.txtImporte.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImporte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImporte.Location = New System.Drawing.Point(565, 312)
        Me.txtImporte.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImporte.MaxLength = 18
        Me.txtImporte.Name = "txtImporte"
        Me.txtImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImporte.Size = New System.Drawing.Size(92, 20)
        Me.txtImporte.TabIndex = 16
        Me.txtImporte.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtImporte, "Importe del Pago.")
        '
        'txtConceptoCancelacion
        '
        Me.txtConceptoCancelacion.AcceptsReturn = True
        Me.txtConceptoCancelacion.BackColor = System.Drawing.SystemColors.Window
        Me.txtConceptoCancelacion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConceptoCancelacion.Enabled = False
        Me.txtConceptoCancelacion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConceptoCancelacion.Location = New System.Drawing.Point(15, 84)
        Me.txtConceptoCancelacion.Margin = New System.Windows.Forms.Padding(2)
        Me.txtConceptoCancelacion.MaxLength = 100
        Me.txtConceptoCancelacion.Multiline = True
        Me.txtConceptoCancelacion.Name = "txtConceptoCancelacion"
        Me.txtConceptoCancelacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConceptoCancelacion.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtConceptoCancelacion.Size = New System.Drawing.Size(344, 80)
        Me.txtConceptoCancelacion.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.txtConceptoCancelacion, "Concepto de Cancelación del Cheque.")
        '
        'txtNumeroCheque
        '
        Me.txtNumeroCheque.AcceptsReturn = True
        Me.txtNumeroCheque.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumeroCheque.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumeroCheque.Enabled = False
        Me.txtNumeroCheque.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumeroCheque.Location = New System.Drawing.Point(287, 42)
        Me.txtNumeroCheque.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNumeroCheque.MaxLength = 10
        Me.txtNumeroCheque.Name = "txtNumeroCheque"
        Me.txtNumeroCheque.ReadOnly = True
        Me.txtNumeroCheque.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumeroCheque.Size = New System.Drawing.Size(72, 20)
        Me.txtNumeroCheque.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.txtNumeroCheque, "Numero de Cheque.")
        '
        'txtConcepto
        '
        Me.txtConcepto.AcceptsReturn = True
        Me.txtConcepto.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtConcepto.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConcepto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtConcepto.Location = New System.Drawing.Point(84, 136)
        Me.txtConcepto.Margin = New System.Windows.Forms.Padding(2)
        Me.txtConcepto.MaxLength = 100
        Me.txtConcepto.Name = "txtConcepto"
        Me.txtConcepto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConcepto.Size = New System.Drawing.Size(422, 20)
        Me.txtConcepto.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txtConcepto, "Concepto de Pago.")
        '
        'txtBeneficiario
        '
        Me.txtBeneficiario.AcceptsReturn = True
        Me.txtBeneficiario.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtBeneficiario.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBeneficiario.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBeneficiario.Location = New System.Drawing.Point(84, 115)
        Me.txtBeneficiario.Margin = New System.Windows.Forms.Padding(2)
        Me.txtBeneficiario.MaxLength = 50
        Me.txtBeneficiario.Name = "txtBeneficiario"
        Me.txtBeneficiario.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBeneficiario.Size = New System.Drawing.Size(422, 20)
        Me.txtBeneficiario.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.txtBeneficiario, "Persona que Recibira el Pago.")
        '
        '_optTipoPago_1
        '
        Me._optTipoPago_1.BackColor = System.Drawing.SystemColors.Control
        Me._optTipoPago_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoPago_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTipoPago.SetIndex(Me._optTipoPago_1, CType(1, Short))
        Me._optTipoPago_1.Location = New System.Drawing.Point(18, 32)
        Me._optTipoPago_1.Margin = New System.Windows.Forms.Padding(2)
        Me._optTipoPago_1.Name = "_optTipoPago_1"
        Me._optTipoPago_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoPago_1.Size = New System.Drawing.Size(73, 24)
        Me._optTipoPago_1.TabIndex = 6
        Me._optTipoPago_1.TabStop = True
        Me._optTipoPago_1.Text = "Personal"
        Me.ToolTip1.SetToolTip(Me._optTipoPago_1, "Pago Personal.")
        Me._optTipoPago_1.UseVisualStyleBackColor = False
        '
        '_optTipoPago_0
        '
        Me._optTipoPago_0.BackColor = System.Drawing.SystemColors.Control
        Me._optTipoPago_0.Checked = True
        Me._optTipoPago_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoPago_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTipoPago.SetIndex(Me._optTipoPago_0, CType(0, Short))
        Me._optTipoPago_0.Location = New System.Drawing.Point(18, 13)
        Me._optTipoPago_0.Margin = New System.Windows.Forms.Padding(2)
        Me._optTipoPago_0.Name = "_optTipoPago_0"
        Me._optTipoPago_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoPago_0.Size = New System.Drawing.Size(73, 20)
        Me._optTipoPago_0.TabIndex = 5
        Me._optTipoPago_0.TabStop = True
        Me._optTipoPago_0.Text = "Joyería"
        Me.ToolTip1.SetToolTip(Me._optTipoPago_0, "Pago de la Joyería.")
        Me._optTipoPago_0.UseVisualStyleBackColor = False
        '
        'cmdOrigenyAplicacion
        '
        Me.cmdOrigenyAplicacion.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOrigenyAplicacion.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOrigenyAplicacion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOrigenyAplicacion.Location = New System.Drawing.Point(572, 341)
        Me.cmdOrigenyAplicacion.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdOrigenyAplicacion.Name = "cmdOrigenyAplicacion"
        Me.cmdOrigenyAplicacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOrigenyAplicacion.Size = New System.Drawing.Size(85, 35)
        Me.cmdOrigenyAplicacion.TabIndex = 32
        Me.cmdOrigenyAplicacion.Text = "A&plicación"
        Me.cmdOrigenyAplicacion.UseVisualStyleBackColor = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.dtpFecha)
        Me.Frame4.Controls.Add(Me.txtFolioEgreso)
        Me.Frame4.Controls.Add(Me.Label1)
        Me.Frame4.Controls.Add(Me.lblMoneda)
        Me.Frame4.Controls.Add(Me.Label3)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(12, 11)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(645, 40)
        Me.Frame4.TabIndex = 27
        Me.Frame4.TabStop = False
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(535, 12)
        Me.dtpFecha.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(97, 20)
        Me.dtpFecha.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(6, 15)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(92, 17)
        Me.Label1.TabIndex = 30
        Me.Label1.Text = "Folio de Egreso :"
        '
        'lblMoneda
        '
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMoneda.Location = New System.Drawing.Point(326, 12)
        Me.lblMoneda.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.Size = New System.Drawing.Size(121, 17)
        Me.lblMoneda.TabIndex = 29
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(494, 14)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(37, 17)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Fecha :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.Color.Gainsboro
        Me.Frame1.Controls.Add(Me.Frame5)
        Me.Frame1.Controls.Add(Me.txtConcepto)
        Me.Frame1.Controls.Add(Me.txtBeneficiario)
        Me.Frame1.Controls.Add(Me.dbcBanco)
        Me.Frame1.Controls.Add(Me.Frame2)
        Me.Frame1.Controls.Add(Me.dbcCuentaBancaria)
        Me.Frame1.Controls.Add(Me.Label7)
        Me.Frame1.Controls.Add(Me.Label6)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(21, 65)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(645, 169)
        Me.Frame1.TabIndex = 17
        Me.Frame1.TabStop = False
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.txtFolioElectronico)
        Me.Frame6.Controls.Add(Me.Label12)
        Me.Frame6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame6.Location = New System.Drawing.Point(431, 241)
        Me.Frame6.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(226, 61)
        Me.Frame6.TabIndex = 35
        Me.Frame6.TabStop = False
        Me.Frame6.Text = "Pago Electrónico"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(8, 24)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(33, 17)
        Me.Label12.TabIndex = 36
        Me.Label12.Text = "Folio Electrónico :"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(508, 314)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(56, 17)
        Me.Label11.TabIndex = 26
        Me.Label11.Text = "Importe :"
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me._optFormaPago_0)
        Me.Frame5.Controls.Add(Me._optFormaPago_1)
        Me.Frame5.Controls.Add(Me._optFormaPago_2)
        Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame5.Location = New System.Drawing.Point(12, 14)
        Me.Frame5.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(369, 46)
        Me.Frame5.TabIndex = 31
        Me.Frame5.TabStop = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.txtConceptoCancelacion)
        Me.Frame3.Controls.Add(Me.chkCancelado)
        Me.Frame3.Controls.Add(Me.txtNumeroCheque)
        Me.Frame3.Controls.Add(Me.dtpFechaCheque)
        Me.Frame3.Controls.Add(Me.Label2)
        Me.Frame3.Controls.Add(Me.Label10)
        Me.Frame3.Controls.Add(Me.Label9)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(12, 228)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(369, 174)
        Me.Frame3.TabIndex = 23
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Cheque"
        '
        'chkCancelado
        '
        Me.chkCancelado.BackColor = System.Drawing.SystemColors.Control
        Me.chkCancelado.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCancelado.Enabled = False
        Me.chkCancelado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCancelado.Location = New System.Drawing.Point(12, 13)
        Me.chkCancelado.Margin = New System.Windows.Forms.Padding(2)
        Me.chkCancelado.Name = "chkCancelado"
        Me.chkCancelado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCancelado.Size = New System.Drawing.Size(127, 20)
        Me.chkCancelado.TabIndex = 11
        Me.chkCancelado.Text = "Generar Cheque Cancelado"
        Me.chkCancelado.UseVisualStyleBackColor = False
        '
        'dtpFechaCheque
        '
        Me.dtpFechaCheque.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaCheque.Location = New System.Drawing.Point(49, 42)
        Me.dtpFechaCheque.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaCheque.Name = "dtpFechaCheque"
        Me.dtpFechaCheque.Size = New System.Drawing.Size(103, 20)
        Me.dtpFechaCheque.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(15, 69)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(103, 14)
        Me.Label2.TabIndex = 34
        Me.Label2.Text = "Concepto de Cancelación :"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(241, 44)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(46, 17)
        Me.Label10.TabIndex = 25
        Me.Label10.Text = "Numero :"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(15, 44)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(37, 17)
        Me.Label9.TabIndex = 24
        Me.Label9.Text = "Fecha :"
        '
        'dbcBanco
        '
        Me.dbcBanco.Location = New System.Drawing.Point(84, 72)
        Me.dbcBanco.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(298, 21)
        Me.dbcBanco.TabIndex = 7
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me._optTipoPago_1)
        Me.Frame2.Controls.Add(Me._optTipoPago_0)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(515, 11)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(116, 62)
        Me.Frame2.TabIndex = 18
        Me.Frame2.TabStop = False
        '
        'dbcCuentaBancaria
        '
        Me.dbcCuentaBancaria.Location = New System.Drawing.Point(84, 93)
        Me.dbcCuentaBancaria.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcCuentaBancaria.Name = "dbcCuentaBancaria"
        Me.dbcCuentaBancaria.Size = New System.Drawing.Size(298, 21)
        Me.dbcCuentaBancaria.TabIndex = 8
        '
        'lblCancelada
        '
        Me.lblCancelada.BackColor = System.Drawing.SystemColors.Control
        Me.lblCancelada.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCancelada.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblCancelada.Location = New System.Drawing.Point(11, 408)
        Me.lblCancelada.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCancelada.Name = "lblCancelada"
        Me.lblCancelada.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCancelada.Size = New System.Drawing.Size(370, 17)
        Me.lblCancelada.TabIndex = 33
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(12, 138)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(67, 17)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Concepto :"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(12, 117)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(67, 17)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Beneficiario :"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(12, 95)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(67, 17)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Cuenta Bancaria :"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(12, 73)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(49, 17)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Banco :"
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(11, 464)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 37
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'btnEliminar
        '
        Me.btnEliminar.BackColor = System.Drawing.SystemColors.Control
        Me.btnEliminar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEliminar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEliminar.Location = New System.Drawing.Point(126, 464)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEliminar.Size = New System.Drawing.Size(109, 36)
        Me.btnEliminar.TabIndex = 38
        Me.btnEliminar.Text = "&Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = False
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackColor = System.Drawing.SystemColors.Control
        Me.btnLimpiar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLimpiar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLimpiar.Location = New System.Drawing.Point(356, 464)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLimpiar.Size = New System.Drawing.Size(109, 36)
        Me.btnLimpiar.TabIndex = 40
        Me.btnLimpiar.Text = "&Nuevo"
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.BackColor = System.Drawing.SystemColors.Control
        Me.btnBuscar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBuscar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBuscar.Location = New System.Drawing.Point(241, 464)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 39
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'optTipoPago
        '
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Panel1.Controls.Add(Me.Frame6)
        Me.Panel1.Controls.Add(Me.Frame3)
        Me.Panel1.Controls.Add(Me.Frame4)
        Me.Panel1.Controls.Add(Me.btnLimpiar)
        Me.Panel1.Controls.Add(Me.txtImporte)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.btnBuscar)
        Me.Panel1.Controls.Add(Me.cmdOrigenyAplicacion)
        Me.Panel1.Controls.Add(Me.btnEliminar)
        Me.Panel1.Controls.Add(Me.btnGuardar)
        Me.Panel1.Controls.Add(Me.lblCancelada)
        Me.Panel1.Location = New System.Drawing.Point(9, 10)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(670, 528)
        Me.Panel1.TabIndex = 36
        '
        'frmBancosProcesoDiarioRegistrodePagos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(688, 547)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoDiarioRegistrodePagos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Registro de Pagos."
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame6.ResumeLayout(False)
        Me.Frame6.PerformLayout()
        Me.Frame5.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        CType(Me.optFormaPago, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optTipoPago, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

End Class


