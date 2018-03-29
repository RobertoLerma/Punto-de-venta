Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports VB = Microsoft.VisualBasic

Public Class frmCXPRegFactComprasCargaInicial
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtTipoCambioEuro As System.Windows.Forms.TextBox
    Public WithEvents txtTipoCambio As System.Windows.Forms.TextBox
    Public WithEvents lblEuro As System.Windows.Forms.Label
    Public WithEvents lblDolar As System.Windows.Forms.Label
    Public WithEvents _fraRegistro_4 As System.Windows.Forms.GroupBox
    Public WithEvents _txtIVA_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtSubTotal_2 As System.Windows.Forms.TextBox
    Public WithEvents txtDesctoFinanciero As System.Windows.Forms.TextBox
    Public WithEvents _lblRegistro_15 As System.Windows.Forms.Label
    Public WithEvents _lblRegistro_6 As System.Windows.Forms.Label
    Public WithEvents lblDesctoFinanciero As System.Windows.Forms.Label
    Public WithEvents fraDF As System.Windows.Forms.GroupBox
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_2 As System.Windows.Forms.RadioButton
    Public WithEvents fraMoneda As System.Windows.Forms.Panel
    Public WithEvents _fraRegistro_6 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblRegistro_3 As System.Windows.Forms.Label
    Public WithEvents fraFecha As System.Windows.Forms.Panel
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents txtFolioContrarecibo As System.Windows.Forms.TextBox
    Public WithEvents chkCheque As System.Windows.Forms.CheckBox
    Public WithEvents _txtSubTotal_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtDescuento_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtIVA_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtTotal_0 As System.Windows.Forms.TextBox
    Public WithEvents txtFolioFactura As System.Windows.Forms.TextBox
    Public WithEvents _fraRegistro_1 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaFactura As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaVence As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblRegistro_1 As System.Windows.Forms.Label
    Public WithEvents _lblRegistro_14 As System.Windows.Forms.Label
    Public WithEvents _lblRegistro_13 As System.Windows.Forms.Label
    Public WithEvents _lblRegistro_12 As System.Windows.Forms.Label
    Public WithEvents _lblRegistro_11 As System.Windows.Forms.Label
    Public WithEvents _lblRegistro_5 As System.Windows.Forms.Label
    Public WithEvents _lblRegistro_4 As System.Windows.Forms.Label
    Public WithEvents _lblRegistro_2 As System.Windows.Forms.Label
    Public WithEvents fraDatosFactura As System.Windows.Forms.GroupBox
    Public WithEvents lblEstatus As System.Windows.Forms.Label
    Public WithEvents lblProveedor As System.Windows.Forms.Label
    Public WithEvents fraRegistro As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRegistro As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents txtDescuento As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtIVA As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtSubTotal As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtTotal As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray

    Const nFACT As Integer = 0
    Const nOC As Integer = 1
    Const nDF As Integer = 2 'Descuento financiero

    Const C_RENENCABEZADO As Integer = 0

    Const C_COLCODIGO As Integer = 0
    Const C_ColDESCRIPCION As Integer = 1
    Const C_ColUNIDAD As Integer = 2
    Const C_ColCANTIDAD As Integer = 3
    Const C_COLPRECIOUNITARIO As Integer = 4
    Const C_COLDESCTO As Integer = 5
    Const C_COLIVA As Integer = 6
    Const C_ColIMPORTE As Integer = 7

    Dim cESTATUSFAC As String
    Public nNUMDOCTO As Integer
    Public lOperExt As Boolean 'Variable que sirve para indicar si el procedimiento LlenarDatos se invocó desde el formulario
    'o desde un formulario distinto a éste (p. ej. frmConsulta)

    Dim cMonedadeCompra As String
    Dim cMonedadeCompraTag As String

    'Variables tipo Currency
    'Dim (txtSubTotal(0)) As Currency
    'Dim numerico(txtDescuento(0)) As Currency
    'Dim numerico(txtIVA(0)) As Currency
    'Dim numerico(txtTotal(0)) As Currency

    Dim mcurSubTotalDF As Decimal
    Dim mcurImpIVADF As Decimal

    Dim rsLocal As ADODB.Recordset

    Dim mblnCambiosEnCodigoOC As Boolean
    Dim mblnCambiosEnCodigo As Boolean
    Dim mblnNuevo As Boolean

    Dim mblnSalir As Boolean

    Dim mblnLoad As Boolean

    'Variables para el Combo
    Dim mblnFueraChange As Boolean
    Dim tecla As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnGuardar As Button
    Friend WithEvents btnBuscar As Button
    Public WithEvents btnCancelar As Button
    Public mintCodProveedor As Integer
    Public strControlActual As String = ""

    Public Sub CalcularDescuentoFinanciero()
        Dim nSubTotal As Decimal
        Dim nDescuento As Decimal
        Dim nIMPIVA As Decimal
        Dim nSubTotalDF As Decimal
        Dim nPorcDF As Decimal
        'Debe ser con los datos de la Orden de Compra
        '   %DesctoF    = Porcentaje de descuento financiero / 100
        '   SubTotalDF  = (SubTotal - Descto) * (%DesctoF)
        '   IvaDF       = (Importe de IVA) * (%DesctoF)

        nSubTotal = CDec(Numerico(txtSubTotal(0).Text))
        nDescuento = CDec(Numerico(txtDescuento(0).Text))
        nIMPIVA = CDec(Numerico(txtIVA(0).Text))
        nPorcDF = System.Math.Round(CDec(Numerico((Me.txtDesctoFinanciero.Text))) / 100, 2)

        mcurSubTotalDF = (nSubTotal - nDescuento) * nPorcDF
        mcurImpIVADF = nIMPIVA * nPorcDF

        Me.txtSubTotal(nDF).Text = VB6.Format(System.Math.Round((nSubTotal - nDescuento) * nPorcDF, 2), "###,###,##0.00")
        Me.txtIVA(nDF).Text = VB6.Format(System.Math.Round(nIMPIVA * nPorcDF, 2), "###,###,##0.00")
    End Sub

    'Public Sub CalcularTotalUSD()
    '    Dim nTipoCambioDolar As Currency
    '    Dim nTipoCambioEuro As Currency
    '    Dim nTotal As Currency
    '    nTipoCambioDolar = CCur(numerico(Me.txtTipoCambio.text))
    '    nTipoCambioEuro = CCur(numerico(Me.txtTipoCambioEuro.text))
    '
    '    If nTipoCambioDolar = 0 Then
    '        Me.txtTotalUSD.text = "(Indefinido)"
    '        Exit Sub
    '    End If
    '
    '    If Me.optMoneda(0).Value Then
    '        Me.txtTotalUSD.text = Format(Me.txtTotal(nOC).text, "###,###,##0.00")
    '        Exit Sub
    '    End If
    '
    '    nTotal = CCur(numerico(Me.txtTotal(nOC).text))
    '
    '    If Me.optMoneda(1).Value Then 'Pesos a Dólares
    '        nTotal = Round(nTotal / nTipoCambioDolar, 2)
    '    ElseIf Me.optMoneda(2).Value Then 'Euros a Dólares
    '        nTotal = Round(nTotal * nTipoCambioEuro / nTipoCambioDolar, 2)
    '    End If
    '    Me.txtTotalUSD.text = Format(nTotal, "###,###,##0.00")
    'End Sub

    Public Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendrá el string del tag que se le mandara al fromulario de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim strControlActual As String 'Nombre del control actual
        Dim nColumnaActual As Integer
        Dim I As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        Select Case UCase(strControlActual)
            Case "TXTFOLIO"
                If Not mblnNuevo Then
                    Exit Sub
                End If
                If mintCodProveedor = 0 Then
                    MsgBox("Debe indicar el proveedor al que pertenece la Orden de Compra", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Me.dbcProveedor.Focus()
                    Exit Sub
                End If
                frmCXPConsultaOrden.dbcProveedor.Text = True
                frmCXPConsultaOrden.nPROV = mintCodProveedor
                frmCXPConsultaOrden.lDESC = True
                frmCXPConsultaOrden.cFORM = "frmCXPRegFactComprasCargaInicial"
                frmCXPConsultaOrden.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Unchecked
                frmCXPConsultaOrden.chkTipoConsulta(0).Enabled = False
                frmCXPConsultaOrden.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Checked
                frmCXPConsultaOrden.chkTipoConsulta(1).Enabled = True
                frmCXPConsultaOrden.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                frmCXPConsultaOrden.chkTipoConsulta(2).Enabled = False
                frmCXPConsultaOrden.chkTipoConsulta(3).CheckState = System.Windows.Forms.CheckState.Unchecked
                frmCXPConsultaOrden.chkTipoConsulta(3).Enabled = False
                frmCXPConsultaOrden.ShowDialog()
                Exit Sub
        End Select

        If UCase(Me.ActiveControl.Name) <> "TXTFOLIOFACTURA" Then
            Exit Sub
        End If

        If mintCodProveedor = 0 Then
            MsgBox("Selecione un Proveedor", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Me.dbcProveedor.Focus()
            ModEstandar.SelTxt()
            Exit Sub
        Else
            If Trim(Me.txtFolioFactura.Text) <> "" Then
                gStrSql = "SELECT LTrim(RTrim(FolioFactura)) as FACTURA, LTrim(RTrim(FolioContraRecibo)) AS CONTRARECIBO, dbo.FormatFecha(fechaFactura, 5) as FECHA, Total as IMPORTE, dbo.EstatusStr(Estatus) as ESTATUS, CodProvAcreed as PROVEEDOR, NumDocto as DOCUMENTO FROM CXPFacturas WHERE CodProvAcreed = " & mintCodProveedor & " and FolioFactura LIKE '" & Trim(Me.txtFolioFactura.Text) & "%' ORDER BY FolioFactura DESC"
            Else
                gStrSql = "SELECT LTrim(RTrim(FolioFactura)) as FACTURA, LTrim(RTrim(FolioContraRecibo)) AS CONTRARECIBO, dbo.FormatFecha(fechaFactura, 5) as FECHA, Total as IMPORTE, dbo.EstatusStr(Estatus) as ESTATUS, CodProvAcreed as PROVEEDOR, NumDocto as DOCUMENTO FROM CXPFacturas WHERE CodProvAcreed = " & mintCodProveedor & " ORDER BY FolioFactura DESC"
            End If
        End If

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsgBox(C_msgSINDATOS, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta
        'Load(FrmConsultas)

        'Call ConfiguraConsultas(FrmConsultas, 7900, RsGral, strTag, strCaptionForm)
        FrmConsultas.Width = VB6.TwipsToPixelsX(8500)
        FrmConsultas.Tag = strTag
        FrmConsultas.Text = "Consulta de Facturas de Compras ..."
        With FrmConsultas.Flexdet
            .set_Cols(0, 7)
            Select Case strControlActual
                Case "TXTFOLIOFACTURA"
                    .set_ColWidth(0, 0, 1700) 'Contrarecibo
                    .set_ColWidth(1, 0, 1700) 'Factura
                    .set_ColWidth(2, 0, 1500) 'Fecha
                    .set_ColWidth(3, 0, 1500) 'Importe
                    .set_ColWidth(4, 0, 1500) 'Estatus
                    .set_ColWidth(5, 0, 0) 'Proveedor
                    .set_ColWidth(6, 0, 0) 'Número de Documento
            End Select
            'Poner Nombre a las Columnas
            .set_TextMatrix(0, 0, "FACTURA")
            .set_TextMatrix(0, 1, "CONTRARECIBO")
            .set_TextMatrix(0, 2, "FECHA")
            .set_TextMatrix(0, 3, "IMPORTE")
            .set_TextMatrix(0, 4, "ESTATUS")
            .set_TextMatrix(0, 5, "PROVEEDOR")
            .set_TextMatrix(0, 6, "DOCUMENTO")

            'Colocar los textos de los encabezados centrados
            .Row = 0
            For I = 0 To (.get_Cols() - 1)
                .Col = I
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next I

            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(5, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            .set_ColAlignment(6, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)

            'Llenar el Grid
            .Rows = RsGral.RecordCount + 1
            .Width = VB6.TwipsToPixelsX(8200)

            If .Rows <= 7 Then
                .Height = VB6.TwipsToPixelsY(.Rows * 305)
            Else
                .Height = VB6.TwipsToPixelsY(7 * 305)
            End If

            FrmConsultas.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(.Height) + 310 + 400)

            .ScrollBars = MSHierarchicalFlexGridLib.ScrollBarsSettings.flexScrollBarVertical

            RsGral.MoveFirst()
            For I = 1 To RsGral.RecordCount
                .set_TextMatrix(I, 0, RsGral.Fields("Factura").Value)
                .set_TextMatrix(I, 1, RsGral.Fields("Contrarecibo").Value)
                .set_TextMatrix(I, 2, VB6.Format(RsGral.Fields("Fecha").Value, "dd/MMM/yyyy"))
                .set_TextMatrix(I, 3, VB6.Format(RsGral.Fields("importe").Value, "###,###,##0.00"))
                .set_TextMatrix(I, 4, RsGral.Fields("Estatus").Value)
                .set_TextMatrix(I, 5, RsGral.Fields("Proveedor").Value)
                .set_TextMatrix(I, 6, RsGral.Fields("Documento").Value)
                RsGral.MoveNext()
            Next I
        End With

        FrmConsultas.ShowDialog()

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Sub Limpiar()
        On Error Resume Next
        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                Case MsgBoxResult.Cancel 'Cancela la acción de limpiar pantalla
                    Exit Sub
            End Select
        End If
        Me.txtFolioFactura.Text = ""
        mblnFueraChange = True
        mintCodProveedor = 0
        Me.dbcProveedor.Text = ""
        Me.dbcProveedor.Tag = ""
        nNUMDOCTO = 0
        mblnFueraChange = False
        Nuevo()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        Me.dbcProveedor.Focus()
    End Sub

    Public Sub Nuevo()
        On Error Resume Next
        If mblnNuevo Then
            Me.txtFolioFactura.Text = ""
            Me.txtFolioFactura.Tag = ""
        End If

        Me.txtFolioContrarecibo.Text = ""
        Me.txtFolioContrarecibo.Tag = ""

        Me.dtpFecha.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.dtpFecha.Tag = VB6.Format(Today, "dd/MMM/yyyy")

        Me.chkCheque.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCheque.Tag = Me.chkCheque.CheckState

        Me.dtpFechaFactura.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.dtpFechaFactura.Tag = VB6.Format(Today, "dd/MMM/yyyy")

        Me.dtpFechaVence.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.dtpFechaVence.Tag = VB6.Format(Today, "dd/MMM/yyyy")

        Me.txtSubTotal(nFACT).Text = "0.00"
        Me.txtSubTotal(nFACT).Tag = "0.00"
        Me.txtDescuento(nFACT).Text = "0.00"
        Me.txtDescuento(nFACT).Tag = "0.00"
        Me.txtIVA(nFACT).Text = "0.00"
        Me.txtIVA(nFACT).Tag = "0.00"
        Me.txtTotal(nFACT).Text = "0.00"
        Me.txtTotal(nFACT).Tag = "0.00"

        Me.txtDesctoFinanciero.Text = "0.00"
        Me.txtDesctoFinanciero.Tag = "0.00"
        Me.txtSubTotal(nDF).Text = "0.00"
        Me.txtSubTotal(nDF).Tag = "0.00"
        Me.txtIVA(nDF).Text = "0.00"
        Me.txtIVA(nDF).Tag = "0.00"
        Me.txtTipoCambio.Text = "0.00"
        Me.txtTipoCambio.Tag = "0.00"
        Me.txtTipoCambioEuro.Text = "0.00"
        Me.txtTipoCambioEuro.Tag = "0.00"

        'Desactivar los controles para que no modifiquen la factura
        'Todos menos dbcProveedor y txtFolioFactura
        Me.txtFolioContrarecibo.ReadOnly = False
        Me.chkCheque.Enabled = True
        Me.dtpFechaFactura.Enabled = True
        Me.dtpFechaVence.Enabled = True
        Me.txtSubTotal(nFACT).ReadOnly = False
        Me.txtDescuento(nFACT).ReadOnly = False
        Me.txtIVA(nFACT).ReadOnly = False
        Me.txtDesctoFinanciero.ReadOnly = False
        Me.txtTipoCambio.ReadOnly = False
        Me.txtTipoCambioEuro.ReadOnly = False
        Me.fraMoneda.Enabled = True
        optMoneda(0).Checked = True
        cMonedadeCompra = "D"
        CalcularDescuentoFinanciero()

        cESTATUSFAC = ""
        Me.lblEstatus.Text = ""
        Me.lblEstatus.Visible = False
    End Sub

    Public Function Cambios() As Boolean
        On Error Resume Next
        Select Case True
            Case Trim(Me.txtFolioContrarecibo.Text) <> Trim(Me.txtFolioContrarecibo.Tag)
                Cambios = True
            Case CInt(Me.chkCheque.CheckState) <> CInt(Me.chkCheque.Tag)
                Cambios = True
            Case VB6.Format(Me.dtpFechaFactura.Value, "dd/MMM/yyyy") <> VB6.Format(Me.dtpFechaFactura.Tag, "dd/MMM/yyyy")
                Cambios = True
            Case VB6.Format(Me.dtpFechaVence.Value, "dd/MMM/yyyy") <> VB6.Format(Me.dtpFechaVence.Tag, "dd/MMM/yyyy")
                Cambios = True
            Case CDec(Numerico((Me.txtTipoCambio.Text))) <> CDec(Numerico((Me.txtTipoCambio.Tag)))
                Cambios = True
            Case CDec(Numerico((Me.txtTipoCambioEuro.Text))) <> CDec(Numerico((Me.txtTipoCambioEuro.Tag)))
                Cambios = True
            Case CDec(Numerico((Me.txtDesctoFinanciero.Text))) <> CDec(Numerico((Me.txtDesctoFinanciero.Tag)))
                Cambios = True
            Case CDec(Numerico(Me.txtSubTotal(nFACT).Text)) <> CDec(Numerico(Me.txtSubTotal(nFACT).Tag))
                Cambios = True
            Case CDec(Numerico(Me.txtDescuento(nFACT).Text)) <> CDec(Numerico(Me.txtDescuento(nFACT).Tag))
                Cambios = True
            Case CDec(Numerico(Me.txtIVA(nFACT).Text)) <> CDec(Numerico(Me.txtIVA(nFACT).Tag))
                Cambios = True
            Case CDec(Numerico(Me.txtTotal(nFACT).Text)) <> CDec(Numerico(Me.txtTotal(nFACT).Tag))
                Cambios = True
                '        Case Trim(Me.txtFolio.text) <> Trim(Me.txtFolio.Tag)
                '            Cambios = True
            Case Else
                Cambios = False
        End Select
    End Function

    Public Function ValidaDatos() As Boolean
        On Error Resume Next
        Select Case True
            Case mintCodProveedor = 0
                MsgBox("Necesita indicar un proveedor", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.dbcProveedor.Focus()
                ModEstandar.SelTxt()
                ValidaDatos = False
            Case Trim(Me.txtFolioFactura.Text) = ""
                MsgBox("Debe escribir el número de la factura del proveedor", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.txtFolioFactura.Focus()
                ModEstandar.SelTxt()
                ValidaDatos = False
                '        Case Trim(Me.txtFolio.text) = "" Or Len(Trim(Me.txtFolio.text)) < 19
                '            MsgBox "Debe especificar la orden de compra de la factura", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                '            Me.txtFolio.SetFocus
                '            ModEstandar.SelTextoTxt Me.txtFolio
                '            ValidaDatos = False
                '        Case Not BuscaOrden()
                '            MsgBox "La Orden de Compra no existe. Por favor, introduzca una orden válida", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                '            Me.txtFolio.SetFocus
                '            ModEstandar.SelTextoTxt Me.txtFolio
                '            ValidaDatos = False
                '        Case CCur(Numerico(Me.txtSubTotal(nOC).text)) <> CCur(Numerico(Me.txtSubTotal(nFACT).text))
                '            MsgBox "Los SubTotales de la Factura y de la Orden de compra deben coincidir", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                '            Me.txtSubTotal(nFACT).SetFocus
                '            ModEstandar.SelTextoTxt Me.txtSubTotal(nFACT)
                '            ValidaDatos = False
                '        Case CCur(Numerico(Me.txtDescuento(nOC).text)) <> CCur(Numerico(Me.txtDescuento(nFACT).text))
                '            MsgBox "Los Descuentos de la Factura y de la Orden de compra deben coincidir", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                '            Me.txtDescuento(nFACT).SetFocus
                '            ModEstandar.SelTextoTxt Me.txtDescuento(nFACT)
                '            ValidaDatos = False
                '        Case CCur(Numerico(Me.txtIVA(nOC).text)) <> CCur(Numerico(Me.txtIVA(nFACT).text))
                '            MsgBox "Los Impuestos (IVA) de la Factura y de la Orden de compra deben coincidir", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                '            Me.txtIVA(nFACT).SetFocus
                '            ModEstandar.SelTextoTxt Me.txtIVA(nFACT)
                '            ValidaDatos = False
            Case CDec(Numerico((Me.txtTipoCambio.Text))) = 0
                MsgBox("El tipo de Cambio para Dólar no debe estar en ceros", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.txtTipoCambio.Focus()
                ModEstandar.SelTextoTxt((Me.txtTipoCambio))
                ValidaDatos = False
            Case CDec(Numerico((Me.txtTipoCambioEuro.Text))) = 0
                MsgBox("El tipo de Cambio para Euro no debe estar en ceros", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.txtTipoCambioEuro.Focus()
                ModEstandar.SelTextoTxt((Me.txtTipoCambioEuro))
                ValidaDatos = False
            Case Me.dtpFechaFactura.Value > Me.dtpFecha.Value
                MsgBox("La Fecha de la Factura no puede ser mayor a la Fecha de Registro", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.dtpFechaFactura.Focus()
                ValidaDatos = False
            Case Me.dtpFechaVence.Value < Me.dtpFechaFactura.Value
                MsgBox("La Fecha de Vencimiento no puede ser menor a la Fecha de la Factura", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.dtpFechaVence.Focus()
                ValidaDatos = False
            Case Else
                ValidaDatos = True
        End Select
    End Function

    'Public Sub NuevaOrdenCompra()
    '    Me.txtTipoCambio.text = "0.00"
    '    Me.txtTipoCambio.Tag = Me.txtTipoCambio.text
    '    Me.txtTipoCambioEuro.text = "0.00"
    '    Me.txtTipoCambioEuro.Tag = Me.txtTipoCambioEuro.text
    '    Me.optMoneda(0).Value = True
    '    Me.optMoneda(1).Value = False
    '    Me.optMoneda(2).Value = False
    '
    '    cMonedadeCompra = C_DOLAR
    '    cMonedadeCompraTag = C_DOLAR
    '
    '
    '
    ''    Me.txtDescripcion.Caption = ""
    ''    Me.txtDescripcion.Tag = ""
    '
    '    numerico(txtSubTotal(0)) = 0
    '    numerico(txtdescuento(0)) = 0
    '    numerico(txtIva(0)) = 0
    '    numerico(txtTotal(0)) = 0
    '
    '    Me.txtSubTotal(nOC).text = "0.00"
    '    Me.txtSubTotal(nOC).Tag = "0.00"
    '    Me.txtDescuento(nOC).text = "0.00"
    '    Me.txtDescuento(nOC).Tag = "0.00"
    '    Me.txtIVA(nOC).text = "0.00"
    '    Me.txtIVA(nOC).Tag = "0.00"
    '    Me.txtTotal(nOC).text = "0.00"
    '    Me.txtTotal(nOC).Tag = "0.00"
    '
    ''    Me.txtTotalUSD.text = "0.00"
    ''    Me.txtTotalUSD.Tag = "0.00"
    '
    '    Call CalcularDescuentoFinanciero
    'End Sub
    '
    'Public Sub LlenaDatosOrdenCompra()
    '    On Local Error GoTo MErr
    '    Dim I As Integer
    '    'gStrSql = "select * from OrdenesCompra where FolioOrdenCompra ='" & Trim(Me.txtFolio.text) & "' and Estatus = '" & C_STGENERADA & "'"
    '    gStrSql = "select * from OrdenesCompra where FolioOrdenCompra ='" & Trim(Me.txtFolio.text) & "'"
    '    ModEstandar.BorraCmd
    '    Cmd.CommandText = "dbo.UP_Select_Datos"
    '    Cmd.CommandType = adCmdStoredProc
    '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
    '    Set RsGral = Cmd.Execute
    '    If RsGral.RecordCount > 0 Then
    '        'Si la Orden de Compra no existe todavía, toma el DescuentoFinanciero de la tabla OrdenesCompra
    '        'pero si la Orden ya existe, toma el valor de la tabla CXPFacturas
    '        If Not BuscaFactura(mintCodProveedor, Trim(Me.txtFolioFactura.text)) Then
    '            Me.txtTipoCambio.text = Format(RsGral!TipoCambioC, "###,###,##0.00")
    '            Me.txtTipoCambio.Tag = Me.txtTipoCambio.text
    '            Me.txtTipoCambioEuro.text = Format(RsGral!TipoCambioEuroC, "###,###,##0.00")
    '            Me.txtTipoCambioEuro.Tag = Me.txtTipoCambioEuro.text
    '            Me.txtDesctoFinanciero.text = Format(RsGral!PorcDesctoFinanciero, "##0.00")
    '            Me.txtDesctoFinanciero.Tag = Me.txtDesctoFinanciero.text
    '        End If
    '
    '        'mintCodProveedor = RsGral!CodProvAcreed
    '
    '        numerico(txtSubTotal(0)) = RsGral!SubTotal
    '        numerico(txtdescuento(0)) = RsGral!Descuento
    '        numerico(txtIva(0)) = RsGral!Iva
    '        numerico(txtTotal(0)) = RsGral!Total
    '
    '        Me.txtSubTotal(nOC).text = Format(RsGral!SubTotal, "###,###,##0.00")
    '        Me.txtSubTotal(nOC).Tag = Me.txtSubTotal(nOC).text
    '        Me.txtDescuento(nOC).text = Format(RsGral!Descuento, "###,###,##0.00")
    '        Me.txtDescuento(nOC).Tag = Me.txtDescuento(nOC).text
    '        Me.txtIVA(nOC).text = Format(RsGral!Iva, "###,###,##0.00")
    '        Me.txtIVA(nOC).Tag = Me.txtIVA(nOC).text
    '        Me.txtTotal(nOC).text = Format(RsGral!Total, "###,###,##0.00")
    '        Me.txtTotal(nOC).Tag = Me.txtTotal(nOC).text
    '
    '        If RsGral!Moneda = C_DOLAR Then
    '            Me.optMoneda(0).Value = True
    '            Me.optMoneda(1).Value = False
    '            Me.optMoneda(2).Value = False
    '        ElseIf RsGral!Moneda = "P" Then
    '            Me.optMoneda(0).Value = False
    '            Me.optMoneda(1).Value = True
    '            Me.optMoneda(2).Value = False
    '        Else
    '            Me.optMoneda(0).Value = False
    '            Me.optMoneda(1).Value = False
    '            Me.optMoneda(2).Value = True
    '        End If
    '
    '        cMonedadeCompra = Trim(RsGral!Moneda)
    '        cMonedadeCompraTag = cMonedadeCompra
    '
    '        'Llenar el Grid
    '        gStrSql = "select * from OrdenesCompraPreCat where FolioOrdenCompra ='" & Trim(Me.txtFolio.text) & "'"
    '        ModEstandar.BorraCmd
    '        Cmd.CommandText = "dbo.UP_Select_Datos"
    '        Cmd.CommandType = adCmdStoredProc
    '        Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '        Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
    '        Set RsGral = Cmd.Execute
    '        Encabezado
    '        If RsGral.RecordCount > 0 Then
    '            RsGral.MoveFirst
    '            With Me.mshFlex
    '                If RsGral.RecordCount < 11 Then
    '                    .Rows = 11
    '                Else
    '                    .Rows = RsGral.RecordCount + 2
    '                End If
    '                For I = 1 To RsGral.RecordCount
    '                    If RsGral!CodArticulo = 0 Then
    '                        .TextMatrix(I, C_ColCODIGO) = ""
    '                    Else
    '                        .TextMatrix(I, C_ColCODIGO) = RsGral!CodArticulo
    '                    End If
    '                    .TextMatrix(I, C_ColDESCRIPCION) = Trim(RsGral!DescArticulo)
    '                    .TextMatrix(I, C_ColUNIDAD) = Trim(BuscaUnidad(RsGral!CodUnidad))
    '                    .TextMatrix(I, C_ColCANTIDAD) = Trim(RsGral!CantidadRecepcion)
    '                    .TextMatrix(I, C_COLPRECIOUNITARIO) = Format(RsGral!CostoUnitario, "###,###,##0.00")
    '                    .TextMatrix(I, C_COLDESCTO) = Format(RsGral!Descuento, "###,###,##0.00")
    '                    .TextMatrix(I, C_COLIVA) = Format(RsGral!Iva, "###,###,##0.00")
    '                    .TextMatrix(I, C_ColIMPORTE) = Format(RsGral!CostoUnitario * RsGral!CantidadRecepcion, "###,###,##0.00")
    '                    RsGral.MoveNext
    '                Next I
    '            End With
    '        End If
    '        'Calcular el total en dólares
    '        Call CalcularTotalUSD
    '        Call CalcularDescuentoFinanciero
    '    Else
    '        Me.mshFlex.Rows = 11
    '        Call NuevaOrdenCompra
    '    End If
    '    mblnCambiosEnCodigoOC = False
    'MErr:
    '    If Err.Number <> 0 Then
    '        ModEstandar.MostrarError
    '    End If
    'End Sub
    '
    Public Sub LlenaDatos()
        On Error GoTo Merr
        Dim I As Integer
        If Len(Trim(Me.txtFolioFactura.Text)) = 0 Then
            Nuevo()
            Exit Sub
        End If

        'Determinar si la función fue llamada por un procedimiento externo, o fue desde este formulario
        If lOperExt Then
            gStrSql = "select * from CxPFacturas where FolioFactura = '" & Trim(Me.txtFolioFactura.Text) & "' and CodProvAcreed = " & mintCodProveedor & " and NumDocto = " & nNUMDOCTO
            lOperExt = False
        Else
            If mintCodProveedor <> 0 Then
                'Buscar el Número de documento cuyo estatus sea diferente de Cancelado
                gStrSql = "select * from CxPFacturas where FolioFactura = '" & Trim(Me.txtFolioFactura.Text) & "' and CodProvAcreed = " & mintCodProveedor & " and Estatus <> '" & Trim(C_STCANCELADA) & "'"
            Else
                MsgBox("Seleccione un proveedor para buscar la factura", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Sub
            End If
        End If

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then

            cESTATUSFAC = RsGral.Fields("Estatus").Value
            nNUMDOCTO = RsGral.Fields("NumDocto").Value

            Me.lblEstatus.Visible = True
            If cESTATUSFAC = C_STCANCELADA Then
                Me.lblEstatus.Text = "CANCELADA"
            ElseIf cESTATUSFAC = C_STVIGENTE Then
                Me.lblEstatus.Text = "VIGENTE"
            Else
                Me.lblEstatus.Visible = False
            End If

            mblnFueraChange = True
            mintCodProveedor = RsGral.Fields("CodProvAcreed").Value
            Me.dbcProveedor.Text = Trim(BuscaNombreProveedor(mintCodProveedor))
            Me.dbcProveedor.Tag = Trim(Me.dbcProveedor.Text)
            mblnFueraChange = False

            Me.txtFolioContrarecibo.Text = Trim(RsGral.Fields("FolioContraRecibo").Value)
            Me.txtFolioContrarecibo.Tag = Me.txtFolioContrarecibo.Text

            Select Case RsGral.Fields("PagoConChq").Value
                Case True
                    Me.chkCheque.CheckState = System.Windows.Forms.CheckState.Checked
                Case Else
                    Me.chkCheque.CheckState = System.Windows.Forms.CheckState.Unchecked
            End Select
            Me.chkCheque.Tag = Me.chkCheque.CheckState

            Dim fecha As String = AgregarHoraAFecha(RsGral.Fields("FechaRegistro").Value)
            Dim fechaFact As String = AgregarHoraAFecha(RsGral.Fields("FechaFactura").Value)
            Dim fechaVenc As String = AgregarHoraAFecha(RsGral.Fields("FechaVencto").Value)


            Me.dtpFecha.Value = fecha
            Me.dtpFecha.Tag = fecha
            Me.dtpFechaFactura.Value = fechaFact
            Me.dtpFechaFactura.Tag = fechaFact
            Me.dtpFechaVence.Value = fechaVenc
            Me.dtpFechaVence.Tag = fechaVenc

            'Me.dtpFecha.Value = VB6.Format(RsGral.Fields("FechaRegistro").Value, "dd/MMM/yyyy")
            'Me.dtpFecha.Tag = VB6.Format(RsGral.Fields("FechaRegistro").Value, "dd/MMM/yyyy")
            'Me.dtpFechaFactura.Value = VB6.Format(RsGral.Fields("FechaFactura").Value, "dd/MMM/yyyy")
            'Me.dtpFechaFactura.Tag = Me.dtpFechaFactura.Value
            'Me.dtpFechaVence.Value = VB6.Format(RsGral.Fields("FechaVencto").Value, "dd/MMM/yyyy")
            'Me.dtpFechaVence.Tag = Me.dtpFechaVence.Value

            Me.txtSubTotal(nFACT).Text = VB6.Format(RsGral.Fields("SubTotal").Value, "###,###,##0.00")
            Me.txtSubTotal(nFACT).Tag = Me.txtSubTotal(nFACT).Text
            Me.txtDescuento(nFACT).Text = VB6.Format(RsGral.Fields("Descuento").Value, "###,###,##0.00")
            Me.txtDescuento(nFACT).Tag = Me.txtDescuento(nFACT).Text
            Me.txtIVA(nFACT).Text = VB6.Format(RsGral.Fields("Iva").Value, "###,###,##0.00")
            Me.txtIVA(nFACT).Tag = Me.txtIVA(nFACT).Text
            Me.txtTotal(nFACT).Text = VB6.Format(RsGral.Fields("Total").Value, "###,###,##0.00")
            Me.txtTotal(nFACT).Tag = Me.txtTotal(nFACT).Text

            Me.txtSubTotal(nDF).Text = VB6.Format(RsGral.Fields("SubTotalDF").Value, "###,###,##0.00")
            Me.txtSubTotal(nDF).Tag = Me.txtSubTotal(nDF).Text
            Me.txtIVA(nDF).Text = VB6.Format(RsGral.Fields("IvaDF").Value, "###,###,##0.00")
            Me.txtIVA(nDF).Tag = Me.txtIVA(nDF).Text

            '        If Trim(RsGral!FolioOrdenCompra) <> "" Then
            Me.txtTipoCambio.Text = VB6.Format(RsGral.Fields("TipoCambio").Value, "###,###,##0.00")
            Me.txtTipoCambio.Tag = Me.txtTipoCambio.Text
            Me.txtTipoCambioEuro.Text = VB6.Format(RsGral.Fields("TipoCambioEuro").Value, "###,###,##0.00")
            Me.txtTipoCambioEuro.Tag = Me.txtTipoCambioEuro.Text
            Me.txtDesctoFinanciero.Text = VB6.Format(RsGral.Fields("DescuentoFinanciero").Value, "##0.00")
            Me.txtDesctoFinanciero.Tag = Me.txtDesctoFinanciero.Text
            Me.txtFolioContrarecibo.ReadOnly = True
            Me.chkCheque.Enabled = False
            Me.dtpFechaFactura.Enabled = False
            Me.dtpFechaVence.Enabled = False
            Me.txtSubTotal(nFACT).ReadOnly = True
            Me.txtDescuento(nFACT).ReadOnly = True
            Me.txtIVA(nFACT).ReadOnly = True
            Me.txtDesctoFinanciero.ReadOnly = True
            Me.txtTipoCambio.ReadOnly = True
            Me.txtTipoCambioEuro.ReadOnly = True
            Me.fraMoneda.Enabled = False
            If RsGral.Fields("Moneda").Value = C_DOLAR Then
                optMoneda(0).Checked = True
            ElseIf RsGral.Fields("Moneda").Value = C_PESO Then
                optMoneda(1).Checked = True
            ElseIf RsGral.Fields("Moneda").Value = C_EURO Then
                optMoneda(2).Checked = True
            End If

            Me.CalcularDescuentoFinanciero()
            mblnCambiosEnCodigo = False
            mblnNuevo = False
        Else

            Dim fechaHoy As String = AgregarHoraAFecha(Today)

            Me.dtpFecha.Value = fechaHoy
            Me.dtpFecha.Tag = fechaHoy
            mblnNuevo = False
            Nuevo()
            mblnNuevo = True
            mblnCambiosEnCodigo = False
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function BuscaNombreProveedor(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT DescProvAcreed FROM CatProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' and codProvAcreed = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaNombreProveedor = Trim(rsLocal.Fields("DescProvACreed").Value)
        Else
            BuscaNombreProveedor = ""
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub BuscaPorcDescto(ByRef Codigo As Integer)
        On Error GoTo Merr
        Dim cTaxID As String
        gStrSql = "SELECT codProvAcreed, DesctoFinanciero FROM CatProvAcreed WHERE codProvAcreed = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            Me.txtDesctoFinanciero.Text = VB6.Format(rsLocal.Fields("DesctoFinanciero").Value, "##0.00")
        Else
            Me.txtDesctoFinanciero.Text = "0.00"
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function BuscaFactura(ByRef nCodProveedor As Integer, ByRef cFolioFactura As String) As Boolean
        On Error GoTo Merr
        gStrSql = "select * from CXPFacturas where CodProvAcreed =" & nCodProveedor & " and FolioFactura = '" & Trim(cFolioFactura) & "' and NumDocto = " & nNUMDOCTO & " and Estatus <> '" & C_STCANCELADA & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaFactura = True
        Else
            BuscaFactura = False
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    'Public Function BuscaOrden() As Boolean
    '    On Local Error GoTo MErr
    '    gStrSql = "select * from OrdenesCompra where FolioOrdenCompra ='" & Trim(Me.txtFolio.text) & "' and (Estatus = '" & C_STGENERADA & "' or Estatus = '" & C_STREGISTRADA & "')"
    '    ModEstandar.BorraCmd
    '    Cmd.CommandText = "dbo.UP_Select_Datos"
    '    Cmd.CommandType = adCmdStoredProc
    '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, gStrSql)
    '    Set rsLocal = Cmd.Execute
    '    If rsLocal.RecordCount > 0 Then
    '        BuscaOrden = True
    '    Else
    '        BuscaOrden = False
    '    End If
    'MErr:
    '    If Err.Number <> 0 Then
    '        ModEstandar.MostrarError
    '    End If
    'End Function

    Public Function BuscaUnidad(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codUnidad, DescUnidad FROM CatUnidades WHERE CodUnidad = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaUnidad = Trim(rsLocal.Fields("DescUnidad").Value)
        Else
            BuscaUnidad = ""
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub Cancelar()
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        gStrSql = "SELECT * FROM CXPFacturas WHERE codProvAcreed = " & mintCodProveedor & " and FolioFactura = '" & Trim(Me.txtFolioFactura.Text) & "' and NumDocto = " & nNUMDOCTO
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount <= 0 Then
            MsgBox("Proporcione la información adecuada de Proveedor y Folio de Factura para Cancelar dicha Factura", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        Else
            If RsGral.Fields("Estatus").Value = "C" Then
                MsgBox("La factura que quiere afectar ya está cancelada", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                RsGral.Close()
                Exit Sub
            End If
        End If

        'Validar si la factura tiene, o no, pagos efectuados
        If ModCorporativo.Referencia("select * from pagos where CodProvAcreed = " & mintCodProveedor & " and FolioFactura = '" & Trim(Me.txtFolioFactura.Text) & "' and Estatus <> 'C'") Then
            MsgBox("No puede cancelar esta factura debido a que ya tiene pagos efectuados", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If

        'Preguntar si desea cancelar el registro
        If MsgBox("¿Está seguro de CANCELAR esta Factura?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        Cnn.BeginTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        blnTransaction = True
        ModStoredProcedures.PR_IMECXPFacturas(CStr(mintCodProveedor), C_TIPOFACTURAPROV, C_TIPOGASTOJOYERIA, "", VB6.Format(Me.dtpFecha.Value, C_FORMATFECHAGUARDAR), Trim(Me.txtFolioContrarecibo.Text), Trim(Me.txtFolioFactura.Text), CStr(nNUMDOCTO), VB6.Format(Me.dtpFechaFactura.Value, C_FORMATFECHAGUARDAR), VB6.Format(Me.dtpFechaVence.Value, C_FORMATFECHAGUARDAR), CStr(#1/1/1900#), "0", VB6.Format(Me.dtpFechaVence.Value, C_FORMATFECHAGUARDAR), CStr(Numerico(txtSubTotal(0).Text)), CStr(Numerico(txtDescuento(0).Text)), "0.00", CStr(Numerico(txtIVA(0).Text)), CStr(Numerico(txtTotal(0).Text)), "0.00", "0.00", cMonedadeCompra, Trim(Me.txtTipoCambio.Text), Trim(Me.txtTipoCambioEuro.Text), CStr(IIf(Me.chkCheque.CheckState = System.Windows.Forms.CheckState.Checked, True, False)), Trim(Me.txtDesctoFinanciero.Text), CStr(mcurSubTotalDF), CStr(mcurImpIVADF), C_STCANCELADA, VB6.Format(Today, C_FORMATFECHAGUARDAR), C_MODIFICACION, CStr(1))
        Cmd.Execute()
        '
        '    'Pone como generada la Orden de Compra
        '    ModStoredProcedures.PR_IMEOrdenesCompra Trim(Me.txtFolio.text), CStr(0), CStr(#1/1/1900#), CStr(#1/1/1900#), "", "", CStr(0), CStr(0), CStr(0), CStr(0), "", CStr(0), CStr(0), CStr(0), CStr(0), "", C_STGENERADA, CStr(#1/1/1900#), "0", "0", "0", "0", "0", "0", "0", CStr(#1/1/1900#), "", C_MODIFICACION, 1 'Pone la orden como Vigente
        '    Cmd.Execute

        Cnn.CommitTrans()
        blnTransaction = False
        Limpiar()
        mblnFueraChange = True
        mintCodProveedor = 0
        Me.dbcProveedor.Text = ""
        Me.dbcProveedor.Tag = ""
        mblnFueraChange = False
        Me.txtFolioFactura.Text = ""
        MsgBox("La operación finalizó satisfactoriamente.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        Me.dbcProveedor.Focus()
        ModEstandar.SelTxt()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function Guardar() As Boolean
        On Error GoTo Merr
        Dim blnTransaction As Boolean

        'Valida si todos los datos han sido llenados correctamnte para poder ser guardados
        If Not ValidaDatos() Then
            If Not BuscaFactura(mintCodProveedor, Trim(Me.txtFolioFactura.Text)) Then
                mblnNuevo = True
            End If
            Exit Function
        End If
        If Not Cambios() Then
            Limpiar()
            Exit Function
        End If

        Cnn.BeginTrans()
        blnTransaction = True
        Dim nPartida As Integer
        If mblnNuevo Then
            ModStoredProcedures.PR_IMECXPFacturas(CStr(mintCodProveedor), C_TIPOFACTURAPROV, C_TIPOGASTOJOYERIA, "", VB6.Format(Me.dtpFecha.Value, C_FORMATFECHAGUARDAR), Trim(Me.txtFolioContrarecibo.Text), Trim(Me.txtFolioFactura.Text), CStr(nNUMDOCTO), VB6.Format(Me.dtpFechaFactura.Value, C_FORMATFECHAGUARDAR), VB6.Format(Me.dtpFechaVence.Value, C_FORMATFECHAGUARDAR), CStr(#1/1/1900#), "0", VB6.Format(Me.dtpFechaVence.Value, C_FORMATFECHAGUARDAR), CStr(Numerico(txtSubTotal(0).Text)), CStr(Numerico(txtDescuento(0).Text)), "0.00", CStr(Numerico(txtIVA(0).Text)), CStr(Numerico(txtTotal(0).Text)), "0.00", "0.00", cMonedadeCompra, Trim(Me.txtTipoCambio.Text), Trim(Me.txtTipoCambioEuro.Text), CStr(IIf(Me.chkCheque.CheckState = System.Windows.Forms.CheckState.Checked, True, False)), Trim(Me.txtDesctoFinanciero.Text), CStr(mcurSubTotalDF), CStr(mcurImpIVADF), C_STVIGENTE, CStr(#1/1/1900#), C_INSERCION, CStr(0))
            Cmd.Execute()
            ModStoredProcedures.PR_IMEProgramacionPagos("", "0", CStr(mintCodProveedor), C_TIPOFACTURAPROV, C_TIPOGASTOJOYERIA, Trim(Me.txtFolioFactura.Text), VB6.Format(Me.dtpFechaFactura.Value, C_FORMATFECHAGUARDAR), VB6.Format(Me.dtpFechaVence.Value, C_FORMATFECHAGUARDAR), CStr(Numerico(txtTotal(0).Text)), cMonedadeCompra, Trim(Me.txtTipoCambio.Text), Trim(Me.txtTipoCambioEuro.Text), Trim(Me.txtDesctoFinanciero.Text), Trim(Me.txtSubTotal(2).Text), Trim(Me.txtIVA(2).Text), C_STVIGENTE, VB6.Format(#1/1/1900#, C_FORMATFECHAGUARDAR), "", CStr(False), CStr(False), VB6.Format(#1/1/1900#, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
            Cmd.Execute()
            nPartida = Cmd.Parameters("ID").Value
        Else
            ModStoredProcedures.PR_IMECXPFacturas(CStr(mintCodProveedor), C_TIPOFACTURAPROV, C_TIPOGASTOJOYERIA, "", VB6.Format(Me.dtpFecha.Value, C_FORMATFECHAGUARDAR), Trim(Me.txtFolioContrarecibo.Text), Trim(Me.txtFolioFactura.Text), CStr(nNUMDOCTO), VB6.Format(Me.dtpFechaFactura.Value, C_FORMATFECHAGUARDAR), VB6.Format(Me.dtpFechaVence.Value, C_FORMATFECHAGUARDAR), CStr(#1/1/1900#), "0", VB6.Format(Me.dtpFechaVence.Value, C_FORMATFECHAGUARDAR), CStr(Numerico(txtSubTotal(0).Text)), CStr(Numerico(txtDescuento(0).Text)), "0.00", CStr(Numerico(txtIVA(0).Text)), CStr(Numerico(txtTotal(0).Text)), "0.00", "0.00", cMonedadeCompra, Trim(Me.txtTipoCambio.Text), Trim(Me.txtTipoCambioEuro.Text), CStr(IIf(Me.chkCheque.CheckState = System.Windows.Forms.CheckState.Checked, True, False)), Trim(Me.txtDesctoFinanciero.Text), CStr(mcurSubTotalDF), CStr(mcurImpIVADF), C_STVIGENTE, CStr(#1/1/1900#), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        Cnn.CommitTrans()
        blnTransaction = False
        If mblnNuevo Then
            MsgBox("La Factura ha sido GUARDADA correctamente", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        End If
        mblnNuevo = True
        Nuevo()
        Guardar = True
        Limpiar()
        mblnFueraChange = True
        mintCodProveedor = 0
        Me.dbcProveedor.Text = ""
        Me.dbcProveedor.Tag = ""
        mblnFueraChange = False

Merr:
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Function

    'Public Sub Encabezado()
    '    Dim LnContador As Integer
    '
    '    With mshFlex
    '        If Not mblnLoad Then
    '            .Rows = 2
    '            .Rows = 12
    '            .COLS = 8
    '            .RemoveItem (1)
    '            Exit Sub
    '        End If
    '        .Height = 2050
    '        .COLS = 8
    '
    '        .Clear
    '
    '        .ColWidth(C_ColCODIGO) = 1140
    '        .ColWidth(C_ColDESCRIPCION) = 2815
    '        .ColWidth(C_ColUNIDAD) = 555
    '        .ColWidth(C_ColCANTIDAD) = 555
    '        .ColWidth(C_COLPRECIOUNITARIO) = 1350
    '        .ColWidth(C_COLDESCTO) = 1350
    '        .ColWidth(C_COLIVA) = 1350
    '        .ColWidth(C_ColIMPORTE) = 1350
    '
    '        .TextMatrix(C_RENENCABEZADO, C_ColCODIGO) = "Código"
    '        .TextMatrix(C_RENENCABEZADO, C_ColDESCRIPCION) = "Descripción"
    '        .TextMatrix(C_RENENCABEZADO, C_ColUNIDAD) = "Ud."
    '        .TextMatrix(C_RENENCABEZADO, C_ColCANTIDAD) = "Cant."
    '        .TextMatrix(C_RENENCABEZADO, C_COLPRECIOUNITARIO) = "Costo Unitario"
    '        .TextMatrix(C_RENENCABEZADO, C_COLDESCTO) = "Descuento"
    '        .TextMatrix(C_RENENCABEZADO, C_COLIVA) = "IVA"
    '        .TextMatrix(C_RENENCABEZADO, C_ColIMPORTE) = "Importe"
    '
    '        'Colocar los textos de los encabezados centrados
    '        .Row = C_RENENCABEZADO
    '        For LnContador = 0 To (.COLS - 1) Step 1
    '            .Col = LnContador
    '            .CellAlignment = flexAlignCenterCenter
    '            .CellFontBold = False
    '        Next LnContador
    '
    '        .TopRow = 1
    '        .Row = 1
    '        .Col = C_ColCODIGO
    '
    '    End With
    'End Sub
    '
    Private Sub chkCheque_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCheque.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcProveedor)

        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    If Guardar() Then
                    End If
                Case MsgBoxResult.No
                Case MsgBoxResult.Cancel
            End Select
        End If

        If Trim(Me.dbcProveedor.Text) = "" Then
            dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        Pon_Tool()
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' ORDER BY descProvAcreed"
        ModDCombo.DCGotFocus(gStrSql, dbcProveedor)
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcProveedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        Aux = mintCodProveedor
        mintCodProveedor = 0
        ModDCombo.DCLostFocus(dbcProveedor, gStrSql, mintCodProveedor)
        If mintCodProveedor = 0 Then
            mblnNuevo = True
            Nuevo()
        ElseIf Aux <> mintCodProveedor Then
            mblnNuevo = True
            Nuevo()
        End If
        If cESTATUSFAC = "" Then
            Call Me.BuscaPorcDescto(mintCodProveedor)
        End If
    End Sub

    Private Sub dbcProveedor_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcProveedor.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcProveedor.Text)
        'If Me.dbcProveedor.SelectedItem <> 0 Then
        '    dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        'End If
        Me.dbcProveedor.Text = Aux
    End Sub

    Private Sub dtpFechaFactura_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFactura.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaVence_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaVence.Enter
        Pon_Tool()
    End Sub
    Private Sub frmCXPRegFactComprasCargaInicial_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCXPRegFactComprasCargaInicial_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPRegFactComprasCargaInicial_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If UCase(Me.ActiveControl.Name) = "DBCPROVEEDOR" Then
                    Me.txtFolioFactura.Focus()
                Else
                    ModEstandar.AvanzarTab(Me)
                End If
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmCXPRegFactComprasCargaInicial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPRegFactComprasCargaInicial_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        mblnLoad = True
        'Me.mshFlex.Rows = 11
        Nuevo()
        mblnLoad = False
        mblnCambiosEnCodigo = False
        mblnNuevo = True
    End Sub

    Private Sub frmCXPRegFactComprasCargaInicial_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSalir Then
        '    'Si desea cerrar la forma y ésta se encuentra minimizada, se debe restaurar
        '    ModEstandar.RestaurarForma(Me, False)
        '    If Cambios() Then ' And Not mblnNuevo
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes
        '                If Not Guardar() Then 'Si falla el guardar, no cierra la forma
        '                    Cancel = 1
        '                Else
        '                    mblnNuevo = True
        '                    Cancel = 0
        '                End If
        '            Case MsgBoxResult.No 'No hace nada y permite que se cierre el formulario
        '                mblnNuevo = True
        '                Cancel = 0
        '            Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
        '                Cancel = 1
        '        End Select
        '    End If
        'Else 'Se quiere salir con escape
        '    mblnSalir = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Me.dbcProveedor.Focus()
        '            ModEstandar.SelTxt()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCXPRegFactComprasCargaInicial_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optMoneda_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoneda.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = optMoneda.GetIndex(eventSender)
            If Me.optMoneda(0).Checked Then
                cMonedadeCompra = "D"
            ElseIf Me.optMoneda(1).Checked Then
                cMonedadeCompra = "P"
            ElseIf Me.optMoneda(2).Checked Then
                cMonedadeCompra = "E"
            Else
                cMonedadeCompra = ""
            End If
        End If
    End Sub

    Private Sub optMoneda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoneda.Enter
        Dim Index As Integer = optMoneda.GetIndex(eventSender)
        Pon_Tool()
    End Sub
    Private Sub txtDesctoFinanciero_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesctoFinanciero.TextChanged
        Call CalcularDescuentoFinanciero()
    End Sub

    Private Sub txtDesctoFinanciero_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesctoFinanciero.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt((Me.txtDesctoFinanciero))
    End Sub

    Private Sub txtDesctoFinanciero_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDesctoFinanciero.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtDesctoFinanciero.Text = VB6.Format(Numerico((Me.txtDesctoFinanciero.Text)), "##0.00")
        End If
        KeyAscii = ModEstandar.MskCantidad((Me.txtDesctoFinanciero.Text), KeyAscii, 3, 2, (Me.txtDesctoFinanciero.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDesctoFinanciero_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesctoFinanciero.Leave
        Me.txtDesctoFinanciero.Text = VB6.Format(Numerico((Me.txtDesctoFinanciero.Text)), "##0.00")
    End Sub

    Private Sub txtDescuento_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescuento.TextChanged
        Dim Index As Integer = txtDescuento.GetIndex(eventSender)
        Select Case Index
            Case nFACT
                Me.txtTotal(Index).Text = CStr(System.Math.Round((CDec(Numerico(Me.txtSubTotal(Index).Text)) - CDec(Numerico(Me.txtDescuento(Index).Text))) + CDec(Numerico(Me.txtIVA(Index).Text)), 2))
        End Select
    End Sub

    Private Sub txtDescuento_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescuento.Enter
        Dim Index As Integer = txtDescuento.GetIndex(eventSender)
        Pon_Tool()
        ModEstandar.SelTextoTxt(Me.txtDescuento(Index))
    End Sub

    Private Sub txtDescuento_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDescuento.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = txtDescuento.GetIndex(eventSender)
        Select Case Index
            Case nFACT
                If KeyAscii = 13 Then
                    Me.txtDescuento(Index).Text = VB6.Format(Numerico(Me.txtDescuento(Index).Text), "###,###,##0.00")
                End If
                KeyAscii = ModEstandar.MskCantidad(Me.txtDescuento(Index).Text, KeyAscii, 9, 2, Me.txtDescuento(Index).SelectionStart)
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDescuento_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescuento.Leave
        Dim Index As Integer = txtDescuento.GetIndex(eventSender)
        Me.txtDescuento(Index).Text = VB6.Format(Numerico(Me.txtDescuento(Index).Text), "###,###,##0.00")
    End Sub

    Private Sub txtFolio_KeyPress(ByRef KeyAscii As Integer)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "/-")
    End Sub

    'Private Sub txtFolio_LostFocus()
    '    If Screen.ActiveForm.Caption = Me.Caption Then
    '        If mblnCambiosEnCodigoOC Then
    '            Call Me.LlenaDatosOrdenCompra
    '        End If
    '    End If
    'End Sub

    Private Sub txtFolioContrarecibo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioContrarecibo.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt((Me.txtFolioContrarecibo))
    End Sub

    Private Sub txtFolioContrarecibo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioContrarecibo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "-/")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioFactura_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioFactura.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
        If Trim(Me.txtFolioFactura.Text) = "" Then
            mblnFueraChange = True
            mintCodProveedor = 0
            Me.dbcProveedor.Text = ""
            Me.dbcProveedor.Tag = Me.dbcProveedor.Text
            mblnFueraChange = False
        End If
    End Sub

    Private Sub txtFolioFactura_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioFactura.Enter
        strControlActual = "TXTFOLIO"
        Pon_Tool()
        ModEstandar.SelTextoTxt((Me.txtFolioFactura))
    End Sub

    Private Sub txtFolioFactura_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFolioFactura.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Pregunta sólo en caso de que existan cambios en la clave (esto es, cuando se teclea una clave diferente a la actual)
        If Cambios() And KeyCode = System.Windows.Forms.Keys.Delete Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        KeyCode = 0
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se borre el contenido del text
                Case MsgBoxResult.Cancel
                    KeyCode = 0
                    Me.txtFolioFactura.Focus()
            End Select
        End If
    End Sub

    Private Sub txtFolioFactura_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioFactura.KeyPress
        'Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'ModEstandar.gp_CampoAlfanumerico(KeyAscii, "/-.")
        'If KeyAscii <> 0 Then
        '    'Pregunta sólo si ha habido cambios
        '    If Cambios() And Not mblnNuevo Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes
        '                If Not Guardar() Then
        '                    KeyAscii = 0
        '                End If
        '            Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
        '            Case MsgBoxResult.Cancel 'Cancela la captura
        '                KeyAscii = 0
        '                Me.txtFolioFactura.Focus()
        '        End Select
        '    End If
        'End If
        'eventArgs.KeyChar = Chr(KeyAscii)
        'If KeyAscii = 0 Then
        '    eventArgs.Handled = True
        'End If
    End Sub

    Private Sub txtFolioFactura_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioFactura.Leave
        'If System.Windows.Forms.Form.ActiveForm.Text = Me.Text Then
        'If mblnCambiosEnCodigo Then
        If txtFolioFactura.Text <> "" Then
            LlenaDatos()
        End If
        'End If
    End Sub

    Private Sub txtIVA_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtIVA.TextChanged
        Dim Index As Integer = txtIVA.GetIndex(eventSender)
        Select Case Index
            Case nFACT
                Me.txtTotal(Index).Text = CStr(System.Math.Round((CDec(Numerico(Me.txtSubTotal(Index).Text)) - CDec(Numerico(Me.txtDescuento(Index).Text))) + CDec(Numerico(Me.txtIVA(Index).Text)), 2))
        End Select
    End Sub

    Private Sub txtIVA_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtIVA.Enter
        Dim Index As Integer = txtIVA.GetIndex(eventSender)
        Pon_Tool()
        ModEstandar.SelTextoTxt(Me.txtIVA(Index))
    End Sub

    Private Sub txtIVA_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtIVA.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = txtIVA.GetIndex(eventSender)
        Select Case Index
            Case nFACT
                If KeyAscii = 13 Then
                    Me.txtIVA(Index).Text = VB6.Format(Numerico(Me.txtIVA(Index).Text), "###,###,##0.00")
                End If
                KeyAscii = ModEstandar.MskCantidad(Me.txtIVA(Index).Text, KeyAscii, 9, 2, Me.txtIVA(Index).SelectionStart)
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtIVA_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtIVA.Leave
        Dim Index As Integer = txtIVA.GetIndex(eventSender)
        Me.txtIVA(Index).Text = VB6.Format(Numerico(Me.txtIVA(Index).Text), "###,###,##0.00")
    End Sub

    Private Sub txtSubTotal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSubTotal.TextChanged
        Dim Index As Integer = txtSubTotal.GetIndex(eventSender)
        Select Case Index
            Case nFACT
                Me.txtTotal(Index).Text = CStr(System.Math.Round((CDec(Numerico(Me.txtSubTotal(Index).Text)) - CDec(Numerico(Me.txtDescuento(Index).Text))) + CDec(Numerico(Me.txtIVA(Index).Text)), 2))
        End Select
    End Sub

    Private Sub txtSubTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSubTotal.Enter
        Dim Index As Integer = txtSubTotal.GetIndex(eventSender)
        Pon_Tool()
        ModEstandar.SelTextoTxt(Me.txtSubTotal(Index))
    End Sub

    Private Sub txtSubTotal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSubTotal.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = txtSubTotal.GetIndex(eventSender)
        Select Case Index
            Case nFACT
                If KeyAscii = 13 Then
                    Me.txtSubTotal(Index).Text = VB6.Format(Numerico(Me.txtSubTotal(Index).Text), "###,###,##0.00")
                End If
                KeyAscii = ModEstandar.MskCantidad(Me.txtSubTotal(Index).Text, KeyAscii, 9, 2, Me.txtSubTotal(Index).SelectionStart)
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtSubTotal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSubTotal.Leave
        Dim Index As Integer = txtSubTotal.GetIndex(eventSender)
        Me.txtSubTotal(Index).Text = VB6.Format(Numerico(Me.txtSubTotal(Index).Text), "###,###,##0.00")
    End Sub

    Private Sub txtTipoCambio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtTipoCambio)
    End Sub

    Private Sub txtTipoCambio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtTipoCambio.Text = VB6.Format(Me.txtTipoCambio.Text, "###,###,##0.00")
        End If
        KeyAscii = ModEstandar.MskCantidad((Me.txtTipoCambio.Text), KeyAscii, 9, 2, (Me.txtTipoCambio.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Leave
        Me.txtTipoCambio.Text = VB6.Format(Numerico((Me.txtTipoCambio.Text)), "###,###,##0.00")
    End Sub

    Private Sub txtTipoCambioEuro_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.TextChanged
        '    Call CalcularTotalUSD
    End Sub

    Private Sub txtTipoCambioEuro_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtTipoCambioEuro)
    End Sub

    Private Sub txtTipoCambioEuro_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambioEuro.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtTipoCambioEuro.Text = VB6.Format(Me.txtTipoCambioEuro.Text, "###,###,##0.00")
        End If
        KeyAscii = ModEstandar.MskCantidad((Me.txtTipoCambioEuro.Text), KeyAscii, 9, 2, (Me.txtTipoCambioEuro.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambioEuro_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.Leave
        Me.txtTipoCambioEuro.Text = VB6.Format(Numerico((Me.txtTipoCambioEuro.Text)), "###,###,##0.00")
    End Sub

    Private Sub txtTotal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTotal.TextChanged
        Dim Index As Integer = txtTotal.GetIndex(eventSender)
        Me.txtTotal(Index).Text = VB6.Format(Numerico(Me.txtTotal(Index).Text), "###,###,##0.00")
    End Sub

    Private Sub txtTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTotal.Enter
        Dim Index As Integer = txtTotal.GetIndex(eventSender)
        Pon_Tool()
        ModEstandar.SelTextoTxt(Me.txtTotal(Index))
    End Sub

    Private Sub txtTotal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTotal.Leave
        Dim Index As Integer = txtTotal.GetIndex(eventSender)
        Me.txtTotal(Index).Text = VB6.Format(Numerico(Me.txtTotal(Index).Text), "###,###,##0.00")
    End Sub

    'Private Sub txtTotalUSD_GotFocus()
    '    Pon_Tool
    '    ModEstandar.SelTextoTxt Me.txtTotalUSD
    'End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtTipoCambioEuro = New System.Windows.Forms.TextBox()
        Me.txtTipoCambio = New System.Windows.Forms.TextBox()
        Me._txtIVA_2 = New System.Windows.Forms.TextBox()
        Me._txtSubTotal_2 = New System.Windows.Forms.TextBox()
        Me.txtDesctoFinanciero = New System.Windows.Forms.TextBox()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_2 = New System.Windows.Forms.RadioButton()
        Me.txtFolioContrarecibo = New System.Windows.Forms.TextBox()
        Me.chkCheque = New System.Windows.Forms.CheckBox()
        Me._txtSubTotal_0 = New System.Windows.Forms.TextBox()
        Me._txtDescuento_0 = New System.Windows.Forms.TextBox()
        Me._txtIVA_0 = New System.Windows.Forms.TextBox()
        Me._txtTotal_0 = New System.Windows.Forms.TextBox()
        Me.txtFolioFactura = New System.Windows.Forms.TextBox()
        Me._fraRegistro_4 = New System.Windows.Forms.GroupBox()
        Me.lblEuro = New System.Windows.Forms.Label()
        Me.lblDolar = New System.Windows.Forms.Label()
        Me.fraDF = New System.Windows.Forms.GroupBox()
        Me._lblRegistro_15 = New System.Windows.Forms.Label()
        Me._lblRegistro_6 = New System.Windows.Forms.Label()
        Me.lblDesctoFinanciero = New System.Windows.Forms.Label()
        Me._fraRegistro_6 = New System.Windows.Forms.GroupBox()
        Me.fraMoneda = New System.Windows.Forms.Panel()
        Me.fraFecha = New System.Windows.Forms.Panel()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me._lblRegistro_3 = New System.Windows.Forms.Label()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me.fraDatosFactura = New System.Windows.Forms.GroupBox()
        Me._fraRegistro_1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaFactura = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaVence = New System.Windows.Forms.DateTimePicker()
        Me._lblRegistro_1 = New System.Windows.Forms.Label()
        Me._lblRegistro_14 = New System.Windows.Forms.Label()
        Me._lblRegistro_13 = New System.Windows.Forms.Label()
        Me._lblRegistro_12 = New System.Windows.Forms.Label()
        Me._lblRegistro_11 = New System.Windows.Forms.Label()
        Me._lblRegistro_5 = New System.Windows.Forms.Label()
        Me._lblRegistro_4 = New System.Windows.Forms.Label()
        Me._lblRegistro_2 = New System.Windows.Forms.Label()
        Me.lblEstatus = New System.Windows.Forms.Label()
        Me.lblProveedor = New System.Windows.Forms.Label()
        Me.fraRegistro = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRegistro = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.txtDescuento = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtIVA = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtSubTotal = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtTotal = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnCancelar = New System.Windows.Forms.Button()
        Me._fraRegistro_4.SuspendLayout()
        Me.fraDF.SuspendLayout()
        Me._fraRegistro_6.SuspendLayout()
        Me.fraMoneda.SuspendLayout()
        Me.fraFecha.SuspendLayout()
        Me.fraDatosFactura.SuspendLayout()
        CType(Me.fraRegistro, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRegistro, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDescuento, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtIVA, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSubTotal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTotal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtTipoCambioEuro
        '
        Me.txtTipoCambioEuro.AcceptsReturn = True
        Me.txtTipoCambioEuro.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtTipoCambioEuro.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambioEuro.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambioEuro.Location = New System.Drawing.Point(230, 22)
        Me.txtTipoCambioEuro.MaxLength = 0
        Me.txtTipoCambioEuro.Name = "txtTipoCambioEuro"
        Me.txtTipoCambioEuro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambioEuro.Size = New System.Drawing.Size(81, 20)
        Me.txtTipoCambioEuro.TabIndex = 25
        Me.txtTipoCambioEuro.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambioEuro, "Tipo de Cambio (de Euros a Pesos)")
        '
        'txtTipoCambio
        '
        Me.txtTipoCambio.AcceptsReturn = True
        Me.txtTipoCambio.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtTipoCambio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambio.Location = New System.Drawing.Point(98, 22)
        Me.txtTipoCambio.MaxLength = 0
        Me.txtTipoCambio.Name = "txtTipoCambio"
        Me.txtTipoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambio.Size = New System.Drawing.Size(81, 20)
        Me.txtTipoCambio.TabIndex = 23
        Me.txtTipoCambio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambio, "Tipo de Cambio (de Dólares a Pesos)")
        '
        '_txtIVA_2
        '
        Me._txtIVA_2.AcceptsReturn = True
        Me._txtIVA_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._txtIVA_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtIVA_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtIVA.SetIndex(Me._txtIVA_2, CType(2, Short))
        Me._txtIVA_2.Location = New System.Drawing.Point(112, 72)
        Me._txtIVA_2.MaxLength = 0
        Me._txtIVA_2.Name = "_txtIVA_2"
        Me._txtIVA_2.ReadOnly = True
        Me._txtIVA_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtIVA_2.Size = New System.Drawing.Size(113, 20)
        Me._txtIVA_2.TabIndex = 32
        Me._txtIVA_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtIVA_2, "Total de los Impuestos")
        '
        '_txtSubTotal_2
        '
        Me._txtSubTotal_2.AcceptsReturn = True
        Me._txtSubTotal_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._txtSubTotal_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtSubTotal_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubTotal.SetIndex(Me._txtSubTotal_2, CType(2, Short))
        Me._txtSubTotal_2.Location = New System.Drawing.Point(112, 48)
        Me._txtSubTotal_2.MaxLength = 0
        Me._txtSubTotal_2.Name = "_txtSubTotal_2"
        Me._txtSubTotal_2.ReadOnly = True
        Me._txtSubTotal_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtSubTotal_2.Size = New System.Drawing.Size(113, 20)
        Me._txtSubTotal_2.TabIndex = 30
        Me._txtSubTotal_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtSubTotal_2, "Total del Importe sin Impuestos")
        '
        'txtDesctoFinanciero
        '
        Me.txtDesctoFinanciero.AcceptsReturn = True
        Me.txtDesctoFinanciero.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtDesctoFinanciero.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDesctoFinanciero.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDesctoFinanciero.Location = New System.Drawing.Point(170, 16)
        Me.txtDesctoFinanciero.MaxLength = 0
        Me.txtDesctoFinanciero.Name = "txtDesctoFinanciero"
        Me.txtDesctoFinanciero.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDesctoFinanciero.Size = New System.Drawing.Size(57, 20)
        Me.txtDesctoFinanciero.TabIndex = 28
        Me.txtDesctoFinanciero.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDesctoFinanciero, "Porcentaje Adicional a la Factura")
        '
        '_optMoneda_0
        '
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_0, CType(0, Short))
        Me._optMoneda_0.Location = New System.Drawing.Point(27, 6)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_0.TabIndex = 35
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me._optMoneda_0, "Moneda de Compra (Dólares)")
        Me._optMoneda_0.UseVisualStyleBackColor = False
        '
        '_optMoneda_1
        '
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_1, CType(1, Short))
        Me._optMoneda_1.Location = New System.Drawing.Point(108, 6)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(59, 17)
        Me._optMoneda_1.TabIndex = 36
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me._optMoneda_1, "Modeda de Compra (Pesos)")
        Me._optMoneda_1.UseVisualStyleBackColor = False
        '
        '_optMoneda_2
        '
        Me._optMoneda_2.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_2, CType(2, Short))
        Me._optMoneda_2.Location = New System.Drawing.Point(190, 4)
        Me._optMoneda_2.Name = "_optMoneda_2"
        Me._optMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_2.Size = New System.Drawing.Size(58, 17)
        Me._optMoneda_2.TabIndex = 37
        Me._optMoneda_2.TabStop = True
        Me._optMoneda_2.Text = "Euros"
        Me.ToolTip1.SetToolTip(Me._optMoneda_2, "Modeda de Compra (Euros)")
        Me._optMoneda_2.UseVisualStyleBackColor = False
        '
        'txtFolioContrarecibo
        '
        Me.txtFolioContrarecibo.AcceptsReturn = True
        Me.txtFolioContrarecibo.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioContrarecibo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioContrarecibo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioContrarecibo.Location = New System.Drawing.Point(264, 24)
        Me.txtFolioContrarecibo.MaxLength = 15
        Me.txtFolioContrarecibo.Name = "txtFolioContrarecibo"
        Me.txtFolioContrarecibo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioContrarecibo.Size = New System.Drawing.Size(97, 20)
        Me.txtFolioContrarecibo.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtFolioContrarecibo, "Folio del Contrarecibo de la factura")
        '
        'chkCheque
        '
        Me.chkCheque.BackColor = System.Drawing.SystemColors.Control
        Me.chkCheque.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCheque.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCheque.Location = New System.Drawing.Point(56, 58)
        Me.chkCheque.Name = "chkCheque"
        Me.chkCheque.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCheque.Size = New System.Drawing.Size(113, 17)
        Me.chkCheque.TabIndex = 5
        Me.chkCheque.Text = "Pago con Cheque"
        Me.ToolTip1.SetToolTip(Me.chkCheque, "Pago con Cheque")
        Me.chkCheque.UseVisualStyleBackColor = False
        '
        '_txtSubTotal_0
        '
        Me._txtSubTotal_0.AcceptsReturn = True
        Me._txtSubTotal_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtSubTotal_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtSubTotal_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubTotal.SetIndex(Me._txtSubTotal_0, CType(0, Short))
        Me._txtSubTotal_0.Location = New System.Drawing.Point(432, 56)
        Me._txtSubTotal_0.MaxLength = 0
        Me._txtSubTotal_0.Name = "_txtSubTotal_0"
        Me._txtSubTotal_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtSubTotal_0.Size = New System.Drawing.Size(113, 20)
        Me._txtSubTotal_0.TabIndex = 14
        Me._txtSubTotal_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtSubTotal_0, "Total del Importe sin Impuestos")
        '
        '_txtDescuento_0
        '
        Me._txtDescuento_0.AcceptsReturn = True
        Me._txtDescuento_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtDescuento_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtDescuento_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescuento.SetIndex(Me._txtDescuento_0, CType(0, Short))
        Me._txtDescuento_0.Location = New System.Drawing.Point(432, 88)
        Me._txtDescuento_0.MaxLength = 0
        Me._txtDescuento_0.Name = "_txtDescuento_0"
        Me._txtDescuento_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtDescuento_0.Size = New System.Drawing.Size(113, 20)
        Me._txtDescuento_0.TabIndex = 16
        Me._txtDescuento_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtDescuento_0, "Total del Descuento")
        '
        '_txtIVA_0
        '
        Me._txtIVA_0.AcceptsReturn = True
        Me._txtIVA_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtIVA_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtIVA_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtIVA.SetIndex(Me._txtIVA_0, CType(0, Short))
        Me._txtIVA_0.Location = New System.Drawing.Point(592, 56)
        Me._txtIVA_0.MaxLength = 0
        Me._txtIVA_0.Name = "_txtIVA_0"
        Me._txtIVA_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtIVA_0.Size = New System.Drawing.Size(113, 20)
        Me._txtIVA_0.TabIndex = 18
        Me._txtIVA_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtIVA_0, "Total de los Impuestos")
        '
        '_txtTotal_0
        '
        Me._txtTotal_0.AcceptsReturn = True
        Me._txtTotal_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._txtTotal_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtTotal_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTotal.SetIndex(Me._txtTotal_0, CType(0, Short))
        Me._txtTotal_0.Location = New System.Drawing.Point(592, 88)
        Me._txtTotal_0.MaxLength = 0
        Me._txtTotal_0.Name = "_txtTotal_0"
        Me._txtTotal_0.ReadOnly = True
        Me._txtTotal_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtTotal_0.Size = New System.Drawing.Size(113, 20)
        Me._txtTotal_0.TabIndex = 20
        Me._txtTotal_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtTotal_0, "Total Neto")
        '
        'txtFolioFactura
        '
        Me.txtFolioFactura.AcceptsReturn = True
        Me.txtFolioFactura.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioFactura.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioFactura.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioFactura.Location = New System.Drawing.Point(56, 24)
        Me.txtFolioFactura.MaxLength = 15
        Me.txtFolioFactura.Name = "txtFolioFactura"
        Me.txtFolioFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioFactura.Size = New System.Drawing.Size(121, 20)
        Me.txtFolioFactura.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtFolioFactura, "Folio de Factura [F3] - Consulta de facturas")
        '
        '_fraRegistro_4
        '
        Me._fraRegistro_4.BackColor = System.Drawing.SystemColors.Control
        Me._fraRegistro_4.Controls.Add(Me.txtTipoCambioEuro)
        Me._fraRegistro_4.Controls.Add(Me.txtTipoCambio)
        Me._fraRegistro_4.Controls.Add(Me.lblEuro)
        Me._fraRegistro_4.Controls.Add(Me.lblDolar)
        Me._fraRegistro_4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRegistro.SetIndex(Me._fraRegistro_4, CType(4, Short))
        Me._fraRegistro_4.Location = New System.Drawing.Point(377, 169)
        Me._fraRegistro_4.Name = "_fraRegistro_4"
        Me._fraRegistro_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRegistro_4.Size = New System.Drawing.Size(354, 53)
        Me._fraRegistro_4.TabIndex = 21
        Me._fraRegistro_4.TabStop = False
        Me._fraRegistro_4.Text = "Tipo de Cambio"
        '
        'lblEuro
        '
        Me.lblEuro.AutoSize = True
        Me.lblEuro.BackColor = System.Drawing.SystemColors.Control
        Me.lblEuro.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEuro.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEuro.Location = New System.Drawing.Point(200, 26)
        Me.lblEuro.Name = "lblEuro"
        Me.lblEuro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEuro.Size = New System.Drawing.Size(29, 13)
        Me.lblEuro.TabIndex = 24
        Me.lblEuro.Text = "Euro"
        '
        'lblDolar
        '
        Me.lblDolar.AutoSize = True
        Me.lblDolar.BackColor = System.Drawing.SystemColors.Control
        Me.lblDolar.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDolar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDolar.Location = New System.Drawing.Point(64, 26)
        Me.lblDolar.Name = "lblDolar"
        Me.lblDolar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDolar.Size = New System.Drawing.Size(32, 13)
        Me.lblDolar.TabIndex = 22
        Me.lblDolar.Text = "Dólar"
        '
        'fraDF
        '
        Me.fraDF.BackColor = System.Drawing.SystemColors.Control
        Me.fraDF.Controls.Add(Me._txtIVA_2)
        Me.fraDF.Controls.Add(Me._txtSubTotal_2)
        Me.fraDF.Controls.Add(Me.txtDesctoFinanciero)
        Me.fraDF.Controls.Add(Me._lblRegistro_15)
        Me.fraDF.Controls.Add(Me._lblRegistro_6)
        Me.fraDF.Controls.Add(Me.lblDesctoFinanciero)
        Me.fraDF.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraDF.Location = New System.Drawing.Point(7, 169)
        Me.fraDF.Name = "fraDF"
        Me.fraDF.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraDF.Size = New System.Drawing.Size(332, 105)
        Me.fraDF.TabIndex = 26
        Me.fraDF.TabStop = False
        Me.fraDF.Text = "Descuento Financiero"
        '
        '_lblRegistro_15
        '
        Me._lblRegistro_15.AutoSize = True
        Me._lblRegistro_15.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_15, CType(15, Short))
        Me._lblRegistro_15.Location = New System.Drawing.Point(56, 76)
        Me._lblRegistro_15.Name = "_lblRegistro_15"
        Me._lblRegistro_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_15.Size = New System.Drawing.Size(24, 13)
        Me._lblRegistro_15.TabIndex = 31
        Me._lblRegistro_15.Text = "IVA"
        '
        '_lblRegistro_6
        '
        Me._lblRegistro_6.AutoSize = True
        Me._lblRegistro_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_6, CType(6, Short))
        Me._lblRegistro_6.Location = New System.Drawing.Point(56, 52)
        Me._lblRegistro_6.Name = "_lblRegistro_6"
        Me._lblRegistro_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_6.Size = New System.Drawing.Size(50, 13)
        Me._lblRegistro_6.TabIndex = 29
        Me._lblRegistro_6.Text = "SubTotal"
        '
        'lblDesctoFinanciero
        '
        Me.lblDesctoFinanciero.AutoSize = True
        Me.lblDesctoFinanciero.BackColor = System.Drawing.SystemColors.Control
        Me.lblDesctoFinanciero.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDesctoFinanciero.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDesctoFinanciero.Location = New System.Drawing.Point(112, 20)
        Me.lblDesctoFinanciero.Name = "lblDesctoFinanciero"
        Me.lblDesctoFinanciero.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDesctoFinanciero.Size = New System.Drawing.Size(58, 13)
        Me.lblDesctoFinanciero.TabIndex = 27
        Me.lblDesctoFinanciero.Text = "Porcentaje"
        '
        '_fraRegistro_6
        '
        Me._fraRegistro_6.BackColor = System.Drawing.SystemColors.Control
        Me._fraRegistro_6.Controls.Add(Me.fraMoneda)
        Me._fraRegistro_6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRegistro.SetIndex(Me._fraRegistro_6, CType(6, Short))
        Me._fraRegistro_6.Location = New System.Drawing.Point(378, 229)
        Me._fraRegistro_6.Name = "_fraRegistro_6"
        Me._fraRegistro_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRegistro_6.Size = New System.Drawing.Size(353, 46)
        Me._fraRegistro_6.TabIndex = 33
        Me._fraRegistro_6.TabStop = False
        Me._fraRegistro_6.Text = "Moneda de la Compra"
        '
        'fraMoneda
        '
        Me.fraMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.fraMoneda.Controls.Add(Me._optMoneda_0)
        Me.fraMoneda.Controls.Add(Me._optMoneda_1)
        Me.fraMoneda.Controls.Add(Me._optMoneda_2)
        Me.fraMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraMoneda.Location = New System.Drawing.Point(25, 16)
        Me.fraMoneda.Name = "fraMoneda"
        Me.fraMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMoneda.Size = New System.Drawing.Size(317, 27)
        Me.fraMoneda.TabIndex = 34
        '
        'fraFecha
        '
        Me.fraFecha.BackColor = System.Drawing.SystemColors.Control
        Me.fraFecha.Controls.Add(Me.dtpFecha)
        Me.fraFecha.Controls.Add(Me._lblRegistro_3)
        Me.fraFecha.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraFecha.Enabled = False
        Me.fraFecha.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraFecha.Location = New System.Drawing.Point(531, 6)
        Me.fraFecha.Name = "fraFecha"
        Me.fraFecha.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraFecha.Size = New System.Drawing.Size(199, 25)
        Me.fraFecha.TabIndex = 38
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(80, 3)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(113, 20)
        Me.dtpFecha.TabIndex = 40
        '
        '_lblRegistro_3
        '
        Me._lblRegistro_3.AutoSize = True
        Me._lblRegistro_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_3, CType(3, Short))
        Me._lblRegistro_3.Location = New System.Drawing.Point(32, 7)
        Me._lblRegistro_3.Name = "_lblRegistro_3"
        Me._lblRegistro_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_3.Size = New System.Drawing.Size(37, 13)
        Me._lblRegistro_3.TabIndex = 39
        Me._lblRegistro_3.Text = "Fecha"
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(80, 8)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(241, 21)
        Me.dbcProveedor.TabIndex = 1
        '
        'fraDatosFactura
        '
        Me.fraDatosFactura.BackColor = System.Drawing.SystemColors.Control
        Me.fraDatosFactura.Controls.Add(Me.txtFolioContrarecibo)
        Me.fraDatosFactura.Controls.Add(Me.chkCheque)
        Me.fraDatosFactura.Controls.Add(Me._txtSubTotal_0)
        Me.fraDatosFactura.Controls.Add(Me._txtDescuento_0)
        Me.fraDatosFactura.Controls.Add(Me._txtIVA_0)
        Me.fraDatosFactura.Controls.Add(Me._txtTotal_0)
        Me.fraDatosFactura.Controls.Add(Me.txtFolioFactura)
        Me.fraDatosFactura.Controls.Add(Me._fraRegistro_1)
        Me.fraDatosFactura.Controls.Add(Me.dtpFechaFactura)
        Me.fraDatosFactura.Controls.Add(Me.dtpFechaVence)
        Me.fraDatosFactura.Controls.Add(Me._lblRegistro_1)
        Me.fraDatosFactura.Controls.Add(Me._lblRegistro_14)
        Me.fraDatosFactura.Controls.Add(Me._lblRegistro_13)
        Me.fraDatosFactura.Controls.Add(Me._lblRegistro_12)
        Me.fraDatosFactura.Controls.Add(Me._lblRegistro_11)
        Me.fraDatosFactura.Controls.Add(Me._lblRegistro_5)
        Me.fraDatosFactura.Controls.Add(Me._lblRegistro_4)
        Me.fraDatosFactura.Controls.Add(Me._lblRegistro_2)
        Me.fraDatosFactura.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraDatosFactura.Location = New System.Drawing.Point(8, 40)
        Me.fraDatosFactura.Name = "fraDatosFactura"
        Me.fraDatosFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraDatosFactura.Size = New System.Drawing.Size(721, 121)
        Me.fraDatosFactura.TabIndex = 2
        Me.fraDatosFactura.TabStop = False
        Me.fraDatosFactura.Text = "Datos de la Factura"
        '
        '_fraRegistro_1
        '
        Me._fraRegistro_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraRegistro_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRegistro.SetIndex(Me._fraRegistro_1, CType(1, Short))
        Me._fraRegistro_1.Location = New System.Drawing.Point(368, 8)
        Me._fraRegistro_1.Name = "_fraRegistro_1"
        Me._fraRegistro_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRegistro_1.Size = New System.Drawing.Size(2, 105)
        Me._fraRegistro_1.TabIndex = 12
        Me._fraRegistro_1.TabStop = False
        '
        'dtpFechaFactura
        '
        Me.dtpFechaFactura.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFactura.Location = New System.Drawing.Point(264, 56)
        Me.dtpFechaFactura.Name = "dtpFechaFactura"
        Me.dtpFechaFactura.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaFactura.TabIndex = 9
        '
        'dtpFechaVence
        '
        Me.dtpFechaVence.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaVence.Location = New System.Drawing.Point(264, 88)
        Me.dtpFechaVence.Name = "dtpFechaVence"
        Me.dtpFechaVence.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaVence.TabIndex = 11
        '
        '_lblRegistro_1
        '
        Me._lblRegistro_1.AutoSize = True
        Me._lblRegistro_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_1, CType(1, Short))
        Me._lblRegistro_1.Location = New System.Drawing.Point(196, 28)
        Me._lblRegistro_1.Name = "_lblRegistro_1"
        Me._lblRegistro_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_1.Size = New System.Drawing.Size(67, 13)
        Me._lblRegistro_1.TabIndex = 6
        Me._lblRegistro_1.Text = "Contrarecibo"
        '
        '_lblRegistro_14
        '
        Me._lblRegistro_14.AutoSize = True
        Me._lblRegistro_14.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_14, CType(14, Short))
        Me._lblRegistro_14.Location = New System.Drawing.Point(376, 60)
        Me._lblRegistro_14.Name = "_lblRegistro_14"
        Me._lblRegistro_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_14.Size = New System.Drawing.Size(50, 13)
        Me._lblRegistro_14.TabIndex = 13
        Me._lblRegistro_14.Text = "SubTotal"
        '
        '_lblRegistro_13
        '
        Me._lblRegistro_13.AutoSize = True
        Me._lblRegistro_13.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_13, CType(13, Short))
        Me._lblRegistro_13.Location = New System.Drawing.Point(376, 92)
        Me._lblRegistro_13.Name = "_lblRegistro_13"
        Me._lblRegistro_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_13.Size = New System.Drawing.Size(59, 13)
        Me._lblRegistro_13.TabIndex = 15
        Me._lblRegistro_13.Text = "Descuento"
        '
        '_lblRegistro_12
        '
        Me._lblRegistro_12.AutoSize = True
        Me._lblRegistro_12.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_12, CType(12, Short))
        Me._lblRegistro_12.Location = New System.Drawing.Point(560, 60)
        Me._lblRegistro_12.Name = "_lblRegistro_12"
        Me._lblRegistro_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_12.Size = New System.Drawing.Size(24, 13)
        Me._lblRegistro_12.TabIndex = 17
        Me._lblRegistro_12.Text = "IVA"
        '
        '_lblRegistro_11
        '
        Me._lblRegistro_11.AutoSize = True
        Me._lblRegistro_11.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_11, CType(11, Short))
        Me._lblRegistro_11.Location = New System.Drawing.Point(560, 92)
        Me._lblRegistro_11.Name = "_lblRegistro_11"
        Me._lblRegistro_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_11.Size = New System.Drawing.Size(31, 13)
        Me._lblRegistro_11.TabIndex = 19
        Me._lblRegistro_11.Text = "Total"
        '
        '_lblRegistro_5
        '
        Me._lblRegistro_5.AutoSize = True
        Me._lblRegistro_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_5, CType(5, Short))
        Me._lblRegistro_5.Location = New System.Drawing.Point(192, 92)
        Me._lblRegistro_5.Name = "_lblRegistro_5"
        Me._lblRegistro_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_5.Size = New System.Drawing.Size(71, 13)
        Me._lblRegistro_5.TabIndex = 10
        Me._lblRegistro_5.Text = "Fecha Vence"
        '
        '_lblRegistro_4
        '
        Me._lblRegistro_4.AutoSize = True
        Me._lblRegistro_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_4, CType(4, Short))
        Me._lblRegistro_4.Location = New System.Drawing.Point(187, 62)
        Me._lblRegistro_4.Name = "_lblRegistro_4"
        Me._lblRegistro_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_4.Size = New System.Drawing.Size(76, 13)
        Me._lblRegistro_4.TabIndex = 8
        Me._lblRegistro_4.Text = "Fecha Factura"
        '
        '_lblRegistro_2
        '
        Me._lblRegistro_2.AutoSize = True
        Me._lblRegistro_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRegistro_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRegistro_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRegistro.SetIndex(Me._lblRegistro_2, CType(2, Short))
        Me._lblRegistro_2.Location = New System.Drawing.Point(26, 28)
        Me._lblRegistro_2.Name = "_lblRegistro_2"
        Me._lblRegistro_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRegistro_2.Size = New System.Drawing.Size(29, 13)
        Me._lblRegistro_2.TabIndex = 3
        Me._lblRegistro_2.Text = "Folio"
        '
        'lblEstatus
        '
        Me.lblEstatus.BackColor = System.Drawing.SystemColors.Info
        Me.lblEstatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEstatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEstatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblEstatus.Location = New System.Drawing.Point(344, 8)
        Me.lblEstatus.Name = "lblEstatus"
        Me.lblEstatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEstatus.Size = New System.Drawing.Size(181, 21)
        Me.lblEstatus.TabIndex = 41
        Me.lblEstatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblEstatus.Visible = False
        '
        'lblProveedor
        '
        Me.lblProveedor.AutoSize = True
        Me.lblProveedor.BackColor = System.Drawing.SystemColors.Control
        Me.lblProveedor.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblProveedor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblProveedor.Location = New System.Drawing.Point(16, 12)
        Me.lblProveedor.Name = "lblProveedor"
        Me.lblProveedor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblProveedor.Size = New System.Drawing.Size(56, 13)
        Me.lblProveedor.TabIndex = 0
        Me.lblProveedor.Text = "Proveedor"
        '
        'optMoneda
        '
        '
        'txtDescuento
        '
        '
        'txtIVA
        '
        '
        'txtSubTotal
        '
        '
        'txtTotal
        '
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(123, 308)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 136
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(8, 308)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 135
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(353, 308)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 134
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'btnCancelar
        '
        Me.btnCancelar.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancelar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancelar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancelar.Location = New System.Drawing.Point(238, 308)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancelar.Size = New System.Drawing.Size(109, 36)
        Me.btnCancelar.TabIndex = 137
        Me.btnCancelar.Text = "&Cancelar"
        Me.btnCancelar.UseVisualStyleBackColor = False
        '
        'frmCXPRegFactComprasCargaInicial
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(737, 360)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me._fraRegistro_4)
        Me.Controls.Add(Me.fraDF)
        Me.Controls.Add(Me._fraRegistro_6)
        Me.Controls.Add(Me.fraFecha)
        Me.Controls.Add(Me.dbcProveedor)
        Me.Controls.Add(Me.fraDatosFactura)
        Me.Controls.Add(Me.lblEstatus)
        Me.Controls.Add(Me.lblProveedor)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 36)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCXPRegFactComprasCargaInicial"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Registro de facturas de compras inicial"
        Me._fraRegistro_4.ResumeLayout(False)
        Me._fraRegistro_4.PerformLayout()
        Me.fraDF.ResumeLayout(False)
        Me.fraDF.PerformLayout()
        Me._fraRegistro_6.ResumeLayout(False)
        Me.fraMoneda.ResumeLayout(False)
        Me.fraFecha.ResumeLayout(False)
        Me.fraFecha.PerformLayout()
        Me.fraDatosFactura.ResumeLayout(False)
        Me.fraDatosFactura.PerformLayout()
        CType(Me.fraRegistro, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRegistro, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDescuento, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtIVA, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSubTotal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTotal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Cancelar()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub
End Class