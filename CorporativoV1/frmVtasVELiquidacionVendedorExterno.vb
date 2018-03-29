Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasVELiquidacionVendedorExterno
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtDescSucursal As System.Windows.Forms.TextBox
    Public WithEvents txtCodVendExterno As System.Windows.Forms.TextBox
    Public WithEvents cmdProcesarPago As System.Windows.Forms.Button
    Public WithEvents optDolaresAnt As System.Windows.Forms.RadioButton
    Public WithEvents optPesosAnt As System.Windows.Forms.RadioButton
    Public WithEvents txtSaldo As System.Windows.Forms.TextBox
    Public WithEvents txtAnticipo As System.Windows.Forms.TextBox
    Public WithEvents Label18 As System.Windows.Forms.Label
    Public WithEvents Label17 As System.Windows.Forms.Label
    Public WithEvents Frame6 As System.Windows.Forms.GroupBox
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents txtDescSucMatriz As System.Windows.Forms.TextBox
    Public WithEvents txtCodSucMatriz As System.Windows.Forms.TextBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents cmdABCClientes As System.Windows.Forms.Button
    Public WithEvents txtTotalPesos As System.Windows.Forms.TextBox
    Public WithEvents txtTotalDolares As System.Windows.Forms.TextBox
    Public WithEvents txtRedondeo As System.Windows.Forms.TextBox
    Public WithEvents txtTotal As System.Windows.Forms.TextBox
    Public WithEvents txtIVA As System.Windows.Forms.TextBox
    Public WithEvents txtDescuento As System.Windows.Forms.TextBox
    Public WithEvents txtSubtotal As System.Windows.Forms.TextBox
    Public WithEvents Label15 As System.Windows.Forms.Label
    Public WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents dbcVendedor As System.Windows.Forms.ComboBox
    Public WithEvents optCredito As System.Windows.Forms.RadioButton
    Public WithEvents optContado As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents txtTipoCambio As System.Windows.Forms.TextBox
    Public WithEvents TxtTelefono As System.Windows.Forms.TextBox
    Public WithEvents txtRFC As System.Windows.Forms.TextBox
    Public WithEvents txtDomicilio As System.Windows.Forms.TextBox
    Public WithEvents txtNombre As System.Windows.Forms.TextBox
    Public WithEvents Label16 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents txtFolioEntrega As System.Windows.Forms.TextBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents txtFolio As System.Windows.Forms.TextBox
    Public WithEvents Label19 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents lblArticulo As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label


    'Variables
    Dim mblnSalir As Boolean
    Dim mblnNuevo As Boolean
    Dim mblnCambios As Boolean
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim intCodVendedor As Integer
    Dim intCodCliente As Integer
    Dim SubTotal As Double
    Dim Descuento As Double
    Dim Iva As Double
    Dim Total As Double
    Dim Redondeo As Double
    Dim RedondeoPesos As Double
    Dim RedondeoDolares As Double
    Dim TotalPesos As Double
    Dim TotalDolares As Double
    Dim TipoCambio As Double
    Dim TipoMovto As String
    Dim FolioSalida As String
    Dim intCodCaja As Integer
    Dim EmitePago As Boolean

    'Constantes para el Grid
    Const C_COLCODARTICULO As Integer = 0
    Const C_COLDESCARTICULO As Integer = 1
    Const C_COLDESCUNIDAD As Integer = 2
    Const C_ColEXISTENCIA As Integer = 3
    Const C_COLPRECIOPUBDOLAR As Integer = 4
    Const C_COLDESCUENTOCONIVA As Integer = 5
    Const C_COLIVAREALCON2DECIMALES As Integer = 6
    Const C_ColIMPORTE As Integer = 7
    Const C_COLIMPORTECONDESCTO As Integer = 8
    Const C_COLPRECIOSINIVA As Integer = 9
    Const C_COLPRECIOREAL As Integer = 10
    Const C_COLPORCENTAJEDEDESCUENTO As Integer = 11
    Const C_COLDESCUENTOSINIVA As Integer = 12
    Const C_COLIVAREALCON4DECIMALES As Integer = 13
    Const C_COLCOSTOREAL As Integer = 14
    Const C_COLORIGEN As Integer = 15
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents btnGuardar As Button
    Public bandera As Boolean = False
    Public strControlActual As String 'Nombre del control actual

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas

        Dim I As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "TXTFOLIO"
                '''SE AGREGO REDONDEO AL TOTAL DE LA VENTA - ROUND1
                strCaptionForm = "Busqueda de Folios de Liquidación a Vendedores Externos"
                gStrSql = "SELECT VTACAB.FOLIOVENTA AS FOLIO,VTACAB.FECHAVENTA AS FECHA,CATCLI.DESCCLIENTE AS CLIENTE," & "CASE VTACAB.CONDICION WHEN 'CO' THEN 'CONTADO' ELSE 'CREDITO' END AS 'TIPO DE VENTA'," & "ROUND(VTACAB.TOTAL+VTACAB.REDONDEO,1) AS TOTAL,CASE VTACAB.ESTATUS WHEN 'V' THEN 'VIGENTE' ELSE 'CANCELADO' END AS ESTATUS " & "FROM MOVIMIENTOSVENTASCAB VTACAB INNER JOIN CATCLIENTES CATCLI ON VTACAB.CODCLIENTE = CATCLI.CODCLIENTE " & "WHERE VTACAB.VTAVEXT = 1 AND TIPOMOVTO = 'V' " & "ORDER BY VTACAB.FECHAVENTA DESC,VTACAB.FOLIOVENTA DESC"
            Case "TXTFOLIOENTREGA"
                strCaptionForm = "Busqueda de Folios de Entrega de Mercancia"
                gStrSql = "SELECT FolioAlmacen AS FOLIO,FechaAlmacen AS FECHA, Concepto AS CONCEPTO FROM " & "MovtosAlmacenCab WHERE CodAlmacen = " & txtCodSucMatriz.Text & " AND CodMovtoAlm = " & C_SalidaAVendedoresExternos & " AND Estatus = 'V' ORDER BY FolioAlmacen Desc,FechaAlmacen Desc"
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
        ConfiguraConsultas(frmconsultas, 13000, RsGral, strTag, strCaptionForm)

        With frmconsultas.Flexdet
            Select Case strControlActual
                Case "TXTFOLIO"
                    .set_ColAlignment(3, 4)
                    .set_ColAlignment(1, 4)
                    .set_ColAlignment(5, 4)
                    .set_ColWidth(0, 0, 1800)
                    .set_ColWidth(1, 0, 1400)
                    .set_ColWidth(2, 0, 4600)
                    .set_ColWidth(3, 0, 1800)
                    .set_ColWidth(4, 0, 1800)
                    .set_ColWidth(5, 0, 1200)
                    For I = 1 To .Rows - 1
                        .set_TextMatrix(I, 1, VB6.Format(.get_TextMatrix(I, 1), "dd/mmm/yyyy"))
                        .set_TextMatrix(I, 4, VB6.Format(.get_TextMatrix(I, 4), "###,##0.00"))
                    Next
                    frmconsultas.Left = VB6.TwipsToPixelsX(975)
                Case "TXTFOLIOENTREGA"
                    'ConfiguraConsultas(frmconsultas, 11000, RsGral, strTag, strCaptionForm)
                    .set_ColAlignment(0, 0)
                    .set_ColAlignment(1, 3)
                    .set_ColAlignment(2, 0)
                    .set_ColWidth(0, 0, 1600)
                    .set_ColWidth(1, 0, 1400)
                    .set_ColWidth(2, 0, 7500)
                    For I = 1 To .Rows - 1
                        .set_TextMatrix(I, 1, VB6.Format(.get_TextMatrix(I, 1), "dd/mmm/yyyy"))
                        .set_TextMatrix(I, 2, Trim(QuitaEnter(.get_TextMatrix(I, 2))))
                    Next
                    frmconsultas.Left = VB6.TwipsToPixelsX(2000)
            End Select
        End With
        frmconsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVtasVELiquidacionVendedorExterno))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCodVendExterno = New System.Windows.Forms.TextBox()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.txtTotalPesos = New System.Windows.Forms.TextBox()
        Me.txtTotalDolares = New System.Windows.Forms.TextBox()
        Me.txtRedondeo = New System.Windows.Forms.TextBox()
        Me.txtTotal = New System.Windows.Forms.TextBox()
        Me.txtIVA = New System.Windows.Forms.TextBox()
        Me.txtDescuento = New System.Windows.Forms.TextBox()
        Me.txtSubtotal = New System.Windows.Forms.TextBox()
        Me.optCredito = New System.Windows.Forms.RadioButton()
        Me.optContado = New System.Windows.Forms.RadioButton()
        Me.txtTipoCambio = New System.Windows.Forms.TextBox()
        Me.TxtTelefono = New System.Windows.Forms.TextBox()
        Me.txtRFC = New System.Windows.Forms.TextBox()
        Me.txtDomicilio = New System.Windows.Forms.TextBox()
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me.txtDescSucursal = New System.Windows.Forms.TextBox()
        Me.cmdProcesarPago = New System.Windows.Forms.Button()
        Me.Frame6 = New System.Windows.Forms.GroupBox()
        Me.optDolaresAnt = New System.Windows.Forms.RadioButton()
        Me.optPesosAnt = New System.Windows.Forms.RadioButton()
        Me.txtSaldo = New System.Windows.Forms.TextBox()
        Me.txtAnticipo = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.txtDescSucMatriz = New System.Windows.Forms.TextBox()
        Me.txtCodSucMatriz = New System.Windows.Forms.TextBox()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.cmdABCClientes = New System.Windows.Forms.Button()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.dbcVendedor = New System.Windows.Forms.ComboBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtFolioEntrega = New System.Windows.Forms.TextBox()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.txtFolio = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblArticulo = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Frame6.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frame4.SuspendLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtCodVendExterno
        '
        Me.txtCodVendExterno.AcceptsReturn = True
        Me.txtCodVendExterno.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodVendExterno.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodVendExterno.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodVendExterno.Location = New System.Drawing.Point(103, 54)
        Me.txtCodVendExterno.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodVendExterno.MaxLength = 3
        Me.txtCodVendExterno.Name = "txtCodVendExterno"
        Me.txtCodVendExterno.ReadOnly = True
        Me.txtCodVendExterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodVendExterno.Size = New System.Drawing.Size(26, 20)
        Me.txtCodVendExterno.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtCodVendExterno, "Codigo del Vendedor Externo")
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(12, 36)
        Me.optDolares.Margin = New System.Windows.Forms.Padding(2)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(62, 18)
        Me.optDolares.TabIndex = 12
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me.optDolares, "Venta de Contado")
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Checked = True
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(12, 15)
        Me.optPesos.Margin = New System.Windows.Forms.Padding(2)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(62, 19)
        Me.optPesos.TabIndex = 13
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me.optPesos, "Venta de Credito")
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'txtTotalPesos
        '
        Me.txtTotalPesos.AcceptsReturn = True
        Me.txtTotalPesos.BackColor = System.Drawing.SystemColors.Window
        Me.txtTotalPesos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTotalPesos.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtTotalPesos.Location = New System.Drawing.Point(252, 81)
        Me.txtTotalPesos.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTotalPesos.MaxLength = 0
        Me.txtTotalPesos.Name = "txtTotalPesos"
        Me.txtTotalPesos.ReadOnly = True
        Me.txtTotalPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotalPesos.Size = New System.Drawing.Size(84, 20)
        Me.txtTotalPesos.TabIndex = 29
        Me.txtTotalPesos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTotalPesos, "Total en Pesos")
        '
        'txtTotalDolares
        '
        Me.txtTotalDolares.AcceptsReturn = True
        Me.txtTotalDolares.BackColor = System.Drawing.SystemColors.Window
        Me.txtTotalDolares.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTotalDolares.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtTotalDolares.Location = New System.Drawing.Point(252, 58)
        Me.txtTotalDolares.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTotalDolares.MaxLength = 0
        Me.txtTotalDolares.Name = "txtTotalDolares"
        Me.txtTotalDolares.ReadOnly = True
        Me.txtTotalDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotalDolares.Size = New System.Drawing.Size(84, 20)
        Me.txtTotalDolares.TabIndex = 28
        Me.txtTotalDolares.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTotalDolares, "Total en Dolares")
        '
        'txtRedondeo
        '
        Me.txtRedondeo.AcceptsReturn = True
        Me.txtRedondeo.BackColor = System.Drawing.SystemColors.Window
        Me.txtRedondeo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRedondeo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtRedondeo.Location = New System.Drawing.Point(252, 35)
        Me.txtRedondeo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRedondeo.MaxLength = 0
        Me.txtRedondeo.Name = "txtRedondeo"
        Me.txtRedondeo.ReadOnly = True
        Me.txtRedondeo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRedondeo.Size = New System.Drawing.Size(84, 20)
        Me.txtRedondeo.TabIndex = 27
        Me.txtRedondeo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtRedondeo, "Redondeo")
        '
        'txtTotal
        '
        Me.txtTotal.AcceptsReturn = True
        Me.txtTotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTotal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtTotal.Location = New System.Drawing.Point(252, 11)
        Me.txtTotal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTotal.MaxLength = 0
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.ReadOnly = True
        Me.txtTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotal.Size = New System.Drawing.Size(84, 20)
        Me.txtTotal.TabIndex = 26
        Me.txtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTotal, "Total sin Redondeo")
        '
        'txtIVA
        '
        Me.txtIVA.AcceptsReturn = True
        Me.txtIVA.BackColor = System.Drawing.SystemColors.Window
        Me.txtIVA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtIVA.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtIVA.Location = New System.Drawing.Point(78, 61)
        Me.txtIVA.Margin = New System.Windows.Forms.Padding(2)
        Me.txtIVA.MaxLength = 0
        Me.txtIVA.Name = "txtIVA"
        Me.txtIVA.ReadOnly = True
        Me.txtIVA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtIVA.Size = New System.Drawing.Size(84, 20)
        Me.txtIVA.TabIndex = 25
        Me.txtIVA.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtIVA, "IVA Total")
        '
        'txtDescuento
        '
        Me.txtDescuento.AcceptsReturn = True
        Me.txtDescuento.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescuento.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescuento.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtDescuento.Location = New System.Drawing.Point(78, 38)
        Me.txtDescuento.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescuento.MaxLength = 0
        Me.txtDescuento.Name = "txtDescuento"
        Me.txtDescuento.ReadOnly = True
        Me.txtDescuento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescuento.Size = New System.Drawing.Size(84, 20)
        Me.txtDescuento.TabIndex = 24
        Me.txtDescuento.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDescuento, "Descuento Total")
        '
        'txtSubtotal
        '
        Me.txtSubtotal.AcceptsReturn = True
        Me.txtSubtotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtSubtotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSubtotal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtSubtotal.Location = New System.Drawing.Point(78, 14)
        Me.txtSubtotal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSubtotal.MaxLength = 0
        Me.txtSubtotal.Name = "txtSubtotal"
        Me.txtSubtotal.ReadOnly = True
        Me.txtSubtotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSubtotal.Size = New System.Drawing.Size(84, 20)
        Me.txtSubtotal.TabIndex = 23
        Me.txtSubtotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtSubtotal, "Subtotal de la Venta")
        '
        'optCredito
        '
        Me.optCredito.BackColor = System.Drawing.SystemColors.Control
        Me.optCredito.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCredito.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCredito.Location = New System.Drawing.Point(18, 34)
        Me.optCredito.Margin = New System.Windows.Forms.Padding(2)
        Me.optCredito.Name = "optCredito"
        Me.optCredito.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCredito.Size = New System.Drawing.Size(67, 20)
        Me.optCredito.TabIndex = 11
        Me.optCredito.TabStop = True
        Me.optCredito.Text = "Crédito"
        Me.ToolTip1.SetToolTip(Me.optCredito, "Venta de Credito")
        Me.optCredito.UseVisualStyleBackColor = False
        '
        'optContado
        '
        Me.optContado.BackColor = System.Drawing.SystemColors.Control
        Me.optContado.Checked = True
        Me.optContado.Cursor = System.Windows.Forms.Cursors.Default
        Me.optContado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optContado.Location = New System.Drawing.Point(18, 15)
        Me.optContado.Margin = New System.Windows.Forms.Padding(2)
        Me.optContado.Name = "optContado"
        Me.optContado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optContado.Size = New System.Drawing.Size(67, 21)
        Me.optContado.TabIndex = 10
        Me.optContado.TabStop = True
        Me.optContado.Text = "Contado"
        Me.ToolTip1.SetToolTip(Me.optContado, "Venta de Contado")
        Me.optContado.UseVisualStyleBackColor = False
        '
        'txtTipoCambio
        '
        Me.txtTipoCambio.AcceptsReturn = True
        Me.txtTipoCambio.BackColor = System.Drawing.SystemColors.Info
        Me.txtTipoCambio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambio.Location = New System.Drawing.Point(601, 119)
        Me.txtTipoCambio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTipoCambio.MaxLength = 0
        Me.txtTipoCambio.Name = "txtTipoCambio"
        Me.txtTipoCambio.ReadOnly = True
        Me.txtTipoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambio.Size = New System.Drawing.Size(56, 20)
        Me.txtTipoCambio.TabIndex = 14
        Me.txtTipoCambio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambio, "Tipo de Cambio")
        '
        'TxtTelefono
        '
        Me.TxtTelefono.AcceptsReturn = True
        Me.TxtTelefono.BackColor = System.Drawing.SystemColors.Info
        Me.TxtTelefono.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtTelefono.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtTelefono.Location = New System.Drawing.Point(69, 62)
        Me.TxtTelefono.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtTelefono.MaxLength = 0
        Me.TxtTelefono.Name = "TxtTelefono"
        Me.TxtTelefono.ReadOnly = True
        Me.TxtTelefono.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtTelefono.Size = New System.Drawing.Size(110, 20)
        Me.TxtTelefono.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.TxtTelefono, "R.F.C. del Cliente")
        '
        'txtRFC
        '
        Me.txtRFC.AcceptsReturn = True
        Me.txtRFC.BackColor = System.Drawing.SystemColors.Info
        Me.txtRFC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRFC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRFC.Location = New System.Drawing.Point(206, 62)
        Me.txtRFC.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRFC.MaxLength = 0
        Me.txtRFC.Name = "txtRFC"
        Me.txtRFC.ReadOnly = True
        Me.txtRFC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRFC.Size = New System.Drawing.Size(141, 20)
        Me.txtRFC.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.txtRFC, "R.F.C. del Cliente")
        '
        'txtDomicilio
        '
        Me.txtDomicilio.AcceptsReturn = True
        Me.txtDomicilio.BackColor = System.Drawing.SystemColors.Info
        Me.txtDomicilio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDomicilio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDomicilio.Location = New System.Drawing.Point(69, 40)
        Me.txtDomicilio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDomicilio.MaxLength = 0
        Me.txtDomicilio.Name = "txtDomicilio"
        Me.txtDomicilio.ReadOnly = True
        Me.txtDomicilio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDomicilio.Size = New System.Drawing.Size(278, 20)
        Me.txtDomicilio.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtDomicilio, "Domicilio del Cliente")
        '
        'txtNombre
        '
        Me.txtNombre.AcceptsReturn = True
        Me.txtNombre.BackColor = System.Drawing.SystemColors.Info
        Me.txtNombre.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombre.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombre.Location = New System.Drawing.Point(69, 18)
        Me.txtNombre.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNombre.MaxLength = 0
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.ReadOnly = True
        Me.txtNombre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombre.Size = New System.Drawing.Size(278, 20)
        Me.txtNombre.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtNombre, "Nombre del Cliente")
        '
        'txtDescSucursal
        '
        Me.txtDescSucursal.AcceptsReturn = True
        Me.txtDescSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescSucursal.Location = New System.Drawing.Point(132, 55)
        Me.txtDescSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescSucursal.MaxLength = 0
        Me.txtDescSucursal.Name = "txtDescSucursal"
        Me.txtDescSucursal.ReadOnly = True
        Me.txtDescSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescSucursal.Size = New System.Drawing.Size(226, 20)
        Me.txtDescSucursal.TabIndex = 4
        '
        'cmdProcesarPago
        '
        Me.cmdProcesarPago.BackColor = System.Drawing.SystemColors.Control
        Me.cmdProcesarPago.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdProcesarPago.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdProcesarPago.Location = New System.Drawing.Point(6, 436)
        Me.cmdProcesarPago.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdProcesarPago.Name = "cmdProcesarPago"
        Me.cmdProcesarPago.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdProcesarPago.Size = New System.Drawing.Size(115, 33)
        Me.cmdProcesarPago.TabIndex = 21
        Me.cmdProcesarPago.Text = "P&rocesar Pago"
        Me.cmdProcesarPago.UseVisualStyleBackColor = False
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.optDolaresAnt)
        Me.Frame6.Controls.Add(Me.optPesosAnt)
        Me.Frame6.Controls.Add(Me.txtSaldo)
        Me.Frame6.Controls.Add(Me.txtAnticipo)
        Me.Frame6.Controls.Add(Me.Label18)
        Me.Frame6.Controls.Add(Me.Label17)
        Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame6.Location = New System.Drawing.Point(7, 367)
        Me.Frame6.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(301, 53)
        Me.Frame6.TabIndex = 53
        Me.Frame6.TabStop = False
        '
        'optDolaresAnt
        '
        Me.optDolaresAnt.BackColor = System.Drawing.SystemColors.Control
        Me.optDolaresAnt.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolaresAnt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolaresAnt.Location = New System.Drawing.Point(73, -1)
        Me.optDolaresAnt.Margin = New System.Windows.Forms.Padding(2)
        Me.optDolaresAnt.Name = "optDolaresAnt"
        Me.optDolaresAnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolaresAnt.Size = New System.Drawing.Size(66, 19)
        Me.optDolaresAnt.TabIndex = 18
        Me.optDolaresAnt.TabStop = True
        Me.optDolaresAnt.Text = "Dolares"
        Me.optDolaresAnt.UseVisualStyleBackColor = False
        '
        'optPesosAnt
        '
        Me.optPesosAnt.BackColor = System.Drawing.SystemColors.Control
        Me.optPesosAnt.Checked = True
        Me.optPesosAnt.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesosAnt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesosAnt.Location = New System.Drawing.Point(10, -1)
        Me.optPesosAnt.Margin = New System.Windows.Forms.Padding(2)
        Me.optPesosAnt.Name = "optPesosAnt"
        Me.optPesosAnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesosAnt.Size = New System.Drawing.Size(62, 19)
        Me.optPesosAnt.TabIndex = 17
        Me.optPesosAnt.TabStop = True
        Me.optPesosAnt.Text = "Pesos"
        Me.optPesosAnt.UseVisualStyleBackColor = False
        '
        'txtSaldo
        '
        Me.txtSaldo.AcceptsReturn = True
        Me.txtSaldo.BackColor = System.Drawing.SystemColors.Window
        Me.txtSaldo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSaldo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSaldo.Location = New System.Drawing.Point(206, 22)
        Me.txtSaldo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSaldo.MaxLength = 15
        Me.txtSaldo.Name = "txtSaldo"
        Me.txtSaldo.ReadOnly = True
        Me.txtSaldo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSaldo.Size = New System.Drawing.Size(72, 20)
        Me.txtSaldo.TabIndex = 20
        Me.txtSaldo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAnticipo
        '
        Me.txtAnticipo.AcceptsReturn = True
        Me.txtAnticipo.BackColor = System.Drawing.SystemColors.Window
        Me.txtAnticipo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAnticipo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnticipo.Location = New System.Drawing.Point(67, 22)
        Me.txtAnticipo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtAnticipo.MaxLength = 15
        Me.txtAnticipo.Name = "txtAnticipo"
        Me.txtAnticipo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAnticipo.Size = New System.Drawing.Size(72, 20)
        Me.txtAnticipo.TabIndex = 19
        Me.txtAnticipo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Location = New System.Drawing.Point(155, 25)
        Me.Label18.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label18.Size = New System.Drawing.Size(47, 17)
        Me.Label18.TabIndex = 55
        Me.Label18.Text = "Saldo :"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label17.Location = New System.Drawing.Point(12, 27)
        Me.Label17.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label17.Size = New System.Drawing.Size(52, 17)
        Me.Label17.TabIndex = 54
        Me.Label17.Text = "Anticipo :"
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.optDolares)
        Me.Frame5.Controls.Add(Me.optPesos)
        Me.Frame5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame5.Location = New System.Drawing.Point(488, 82)
        Me.Frame5.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(80, 58)
        Me.Frame5.TabIndex = 52
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Moneda"
        '
        'txtDescSucMatriz
        '
        Me.txtDescSucMatriz.AcceptsReturn = True
        Me.txtDescSucMatriz.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescSucMatriz.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescSucMatriz.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescSucMatriz.Location = New System.Drawing.Point(126, 542)
        Me.txtDescSucMatriz.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescSucMatriz.MaxLength = 0
        Me.txtDescSucMatriz.Name = "txtDescSucMatriz"
        Me.txtDescSucMatriz.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescSucMatriz.Size = New System.Drawing.Size(86, 20)
        Me.txtDescSucMatriz.TabIndex = 51
        '
        'txtCodSucMatriz
        '
        Me.txtCodSucMatriz.AcceptsReturn = True
        Me.txtCodSucMatriz.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucMatriz.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucMatriz.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucMatriz.Location = New System.Drawing.Point(36, 542)
        Me.txtCodSucMatriz.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucMatriz.MaxLength = 0
        Me.txtCodSucMatriz.Name = "txtCodSucMatriz"
        Me.txtCodSucMatriz.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucMatriz.Size = New System.Drawing.Size(56, 20)
        Me.txtCodSucMatriz.TabIndex = 50
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(15, 228)
        Me.txtFlex.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.ReadOnly = True
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(64, 20)
        Me.txtFlex.TabIndex = 49
        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtFlex.Visible = False
        '
        'cmdABCClientes
        '
        Me.cmdABCClientes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdABCClientes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdABCClientes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdABCClientes.Location = New System.Drawing.Point(7, 473)
        Me.cmdABCClientes.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdABCClientes.Name = "cmdABCClientes"
        Me.cmdABCClientes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdABCClientes.Size = New System.Drawing.Size(115, 33)
        Me.cmdABCClientes.TabIndex = 22
        Me.cmdABCClientes.Text = "ABC Clien&tes"
        Me.cmdABCClientes.UseVisualStyleBackColor = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.txtTotalPesos)
        Me.Frame4.Controls.Add(Me.txtTotalDolares)
        Me.Frame4.Controls.Add(Me.txtRedondeo)
        Me.Frame4.Controls.Add(Me.txtTotal)
        Me.Frame4.Controls.Add(Me.txtIVA)
        Me.Frame4.Controls.Add(Me.txtDescuento)
        Me.Frame4.Controls.Add(Me.txtSubtotal)
        Me.Frame4.Controls.Add(Me.Label15)
        Me.Frame4.Controls.Add(Me.Label14)
        Me.Frame4.Controls.Add(Me.Label13)
        Me.Frame4.Controls.Add(Me.Label12)
        Me.Frame4.Controls.Add(Me.Label11)
        Me.Frame4.Controls.Add(Me.Label10)
        Me.Frame4.Controls.Add(Me.Label9)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(319, 367)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(358, 111)
        Me.Frame4.TabIndex = 40
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Totales"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label15.Location = New System.Drawing.Point(171, 84)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label15.Size = New System.Drawing.Size(72, 13)
        Me.Label15.TabIndex = 47
        Me.Label15.Text = "Total Pesos :"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Location = New System.Drawing.Point(171, 63)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(87, 17)
        Me.Label14.TabIndex = 46
        Me.Label14.Text = "Total Dólares :"
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(171, 41)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(66, 17)
        Me.Label13.TabIndex = 45
        Me.Label13.Text = "Redondeo :"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(171, 18)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(56, 17)
        Me.Label12.TabIndex = 44
        Me.Label12.Text = "Total :"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(9, 62)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(49, 17)
        Me.Label11.TabIndex = 43
        Me.Label11.Text = "IVA :"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(9, 41)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(65, 17)
        Me.Label10.TabIndex = 42
        Me.Label10.Text = "Descuento :"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(9, 20)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(65, 17)
        Me.Label9.TabIndex = 41
        Me.Label9.Text = "SubTotal :"
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(11, 185)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(666, 151)
        Me.flexDetalle.TabIndex = 16
        '
        'dbcVendedor
        '
        Me.dbcVendedor.Location = New System.Drawing.Point(398, 158)
        Me.dbcVendedor.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcVendedor.Name = "dbcVendedor"
        Me.dbcVendedor.Size = New System.Drawing.Size(164, 21)
        Me.dbcVendedor.TabIndex = 15
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.optCredito)
        Me.Frame3.Controls.Add(Me.optContado)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(398, 82)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(85, 58)
        Me.Frame3.TabIndex = 38
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Tipo Venta"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.TxtTelefono)
        Me.Frame2.Controls.Add(Me.txtRFC)
        Me.Frame2.Controls.Add(Me.txtDomicilio)
        Me.Frame2.Controls.Add(Me.txtNombre)
        Me.Frame2.Controls.Add(Me.Label16)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Controls.Add(Me.Label5)
        Me.Frame2.Controls.Add(Me.Label4)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(7, 82)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(370, 92)
        Me.Frame2.TabIndex = 33
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Cliente"
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.Control
        Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label16.Location = New System.Drawing.Point(12, 63)
        Me.Label16.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label16.Size = New System.Drawing.Size(61, 17)
        Me.Label16.TabIndex = 56
        Me.Label16.Text = "Telefono :"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(182, 63)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(31, 17)
        Me.Label6.TabIndex = 36
        Me.Label6.Text = "R.F.C."
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(12, 41)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(61, 17)
        Me.Label5.TabIndex = 35
        Me.Label5.Text = "Domicilio :"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(12, 20)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(52, 17)
        Me.Label4.TabIndex = 34
        Me.Label4.Text = "Nombre :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtFolioEntrega)
        Me.Frame1.Controls.Add(Me.dtpFecha)
        Me.Frame1.Controls.Add(Me.txtFolio)
        Me.Frame1.Controls.Add(Me.Label19)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(7, 6)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(668, 40)
        Me.Frame1.TabIndex = 30
        Me.Frame1.TabStop = False
        '
        'txtFolioEntrega
        '
        Me.txtFolioEntrega.AcceptsReturn = True
        Me.txtFolioEntrega.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioEntrega.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioEntrega.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioEntrega.Location = New System.Drawing.Point(356, 11)
        Me.txtFolioEntrega.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioEntrega.MaxLength = 17
        Me.txtFolioEntrega.Name = "txtFolioEntrega"
        Me.txtFolioEntrega.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioEntrega.Size = New System.Drawing.Size(98, 20)
        Me.txtFolioEntrega.TabIndex = 1
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(543, 11)
        Me.dtpFecha.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(95, 20)
        Me.dtpFecha.TabIndex = 2
        '
        'txtFolio
        '
        Me.txtFolio.AcceptsReturn = True
        Me.txtFolio.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolio.Location = New System.Drawing.Point(81, 11)
        Me.txtFolio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolio.MaxLength = 17
        Me.txtFolio.Name = "txtFolio"
        Me.txtFolio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolio.Size = New System.Drawing.Size(121, 20)
        Me.txtFolio.TabIndex = 0
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.SystemColors.Control
        Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label19.Location = New System.Drawing.Point(258, 15)
        Me.Label19.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label19.Size = New System.Drawing.Size(112, 17)
        Me.Label19.TabIndex = 58
        Me.Label19.Text = "Folio de Entrega :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(500, 14)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(45, 17)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Fecha :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(86, 17)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Folio Venta :"
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(132, 54)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(158, 21)
        Me.dbcSucursal.TabIndex = 5
        Me.dbcSucursal.Visible = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(6, 56)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(115, 17)
        Me.Label2.TabIndex = 57
        Me.Label2.Text = "Vendedor Externo :"
        '
        'lblArticulo
        '
        Me.lblArticulo.BackColor = System.Drawing.SystemColors.Info
        Me.lblArticulo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblArticulo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblArticulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblArticulo.Location = New System.Drawing.Point(7, 347)
        Me.lblArticulo.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblArticulo.Name = "lblArticulo"
        Me.lblArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblArticulo.Size = New System.Drawing.Size(670, 17)
        Me.lblArticulo.TabIndex = 48
        Me.lblArticulo.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(399, 142)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(70, 17)
        Me.Label8.TabIndex = 39
        Me.Label8.Text = "Vendedor :"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(588, 100)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(87, 16)
        Me.Label7.TabIndex = 37
        Me.Label7.Text = "Tipo de Cambio"
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Location = New System.Drawing.Point(470, 502)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(93, 35)
        Me.btnLimpiar.TabIndex = 75
        Me.btnLimpiar.Text = "Nuevo"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(371, 502)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(93, 35)
        Me.btnBuscar.TabIndex = 74
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(272, 502)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(93, 35)
        Me.btnEliminar.TabIndex = 73
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.Location = New System.Drawing.Point(165, 502)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(93, 35)
        Me.btnGuardar.TabIndex = 72
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'frmVtasVELiquidacionVendedorExterno
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(689, 542)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.txtDescSucursal)
        Me.Controls.Add(Me.txtCodVendExterno)
        Me.Controls.Add(Me.cmdProcesarPago)
        Me.Controls.Add(Me.Frame6)
        Me.Controls.Add(Me.Frame5)
        Me.Controls.Add(Me.txtDescSucMatriz)
        Me.Controls.Add(Me.txtCodSucMatriz)
        Me.Controls.Add(Me.txtFlex)
        Me.Controls.Add(Me.cmdABCClientes)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.flexDetalle)
        Me.Controls.Add(Me.dbcVendedor)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.txtTipoCambio)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblArticulo)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(76, 126)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasVELiquidacionVendedorExterno"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Liquidación de Vendedor Externo"
        Me.Frame6.ResumeLayout(False)
        Me.Frame6.PerformLayout()
        Me.Frame5.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Sub ObtenerCaja()
        On Error GoTo Merr
        gStrSql = "SELECT CodCaja FROM CatCajas WHERE CodAlmacen = " & txtCodSucMatriz.Text & " ORDER BY CodCaja"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            intCodCaja = RsGral.Fields("CodCaja").Value
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function BuscaMovimientos() As Boolean
        On Error GoTo Merr
        Dim RsAux As ADODB.Recordset
        gStrSql = "SELECT FOLIOALMACEN,FECHAALMACEN,CODMOVTOALM,REFERENCIADEORIGEN FROM MOVTOSALMACENCAB " & "Where CODALMACEN = " & txtCodVendExterno.Text & " AND (CodMovtoAlm = " & C_EntradaaAlmacendeVendedorExterno & " Or CodMovtoAlm = " & C_SalidadeAlmacendeVendedorExterno & " Or " & "CodMovtoAlm = " & C_SalidaPorVentadeVendedoresExternos & ") " & "ORDER BY FOLIOALMACEN DESC,FECHAALMACEN DESC"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsAux = Cmd.Execute
        If RsAux.RecordCount > 0 Then
            RsAux.MoveFirst()
            If RsAux.Fields("CodMovtoAlm").Value = C_SalidaPorVentadeVendedoresExternos Then
                MsgBox("Este vendedor externo ya liquido su última salida, no es posible hacerle otra liquidación ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                BuscaMovimientos = False
            Else
                FolioSalida = RsAux.Fields("ReferenciaDeOrigen").Value
                BuscaMovimientos = True
            End If
        Else
            MsgBox("Este vendedor externo no ha tenido movimientos, no es posible hacerle una liquidación...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            BuscaMovimientos = False
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function GeneraFolioAdicional(ByRef FolioVenta As String) As String
        On Error GoTo Err_Renamed
        Dim TipoMovto As String
        Dim CajaSucursal As String
        Dim Fecha As String
        Dim Consecutivo As Integer
        TipoMovto = (FolioVenta)
        CajaSucursal = Mid(FolioVenta, 2, 4)
        Fecha = Mid(FolioVenta, 6, 8)
        gStrSql = "SELECT ISNULL(MAX(RIGHT(FOLIOADICIONAL,4))+1,0) AS Consecutivo FROM MOVIMIENTOSVENTASDET " & "WHERE LEFT(FOLIOADICIONAL,1) = '" & TipoMovto & "' AND SUBSTRING(FOLIOADICIONAL,2,4) = '" & CajaSucursal & "' AND SUBSTRING(FOLIOADICIONAL,6,8) = '" & Fecha & "' "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Consecutivo").Value = 0 Then
                Consecutivo = 1
            Else
                Consecutivo = RsGral.Fields("Consecutivo").Value
            End If
            GeneraFolioAdicional = TipoMovto & CajaSucursal & Fecha & VB6.Format(Consecutivo, "0000")
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function Guardar() As Boolean
        On Error GoTo Err_Renamed
        Dim blnTransaccion As Boolean
        Dim Prefijo As String
        Dim Consecutivo As Integer
        Dim FolioVenta As String
        Dim Condicion As String
        Dim TipoIngreso As String
        Dim MonedaAnticipo As Object
        Dim I As Integer
        Dim FolioInventario As String
        Dim FolioAdicional As String
        If Not mblnNuevo Then Exit Function
        If Not ValidaDatos() Then Exit Function
        Dim NumPartida As Integer
        Dim DescFamilia As String

        'Verificar si se introdujo todo el Importe que se debe Pagar por la venta. De no ser así, no podrá grabarse la venta.
        'Si el Cambio es Mayor o Igual a Cero, sí podrá grabarse.
        If optContado.Checked = True Or (optCredito.Checked = True And CDbl(Numerico(txtAnticipo.Text)) > 0) Then
            'If CDbl(Numerico((frmPagosSalMercancia.txtmnCambio).Text)) < 0 Then
            '    MsgBox("El importe pagado por el cliente, es menor que el importe total de la venta." & vbNewLine & "Verifique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            '    frmPagosSalMercancia.msgFormasPago.Focus()
            '    Exit Function
            'End If
        End If

        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        'Obtener el Folio con el cual se almacenará la Venta:
        'La estructura es: Prefijo - Sucursal - CodCaja - Fecha(Año-Mes-Dia) - Consecutivo
        gStrSql = "Select PrefijoSalidasMcia as Prefijo, ConsecSalidasMcia + 1 as Consecutivo From CatCajas Where CodCaja = " & intCodCaja & " AND CodAlmacen = " & txtCodSucMatriz.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Prefijo = Trim(RsGral.Fields("Prefijo").Value)
            Consecutivo = RsGral.Fields("Consecutivo").Value
            txtFolio.Text = Prefijo & VB6.Format(CStr(txtCodSucMatriz.Text), "00") & VB6.Format(CStr(intCodCaja), "00") & CStr(Year(dtpFecha.Value)) & VB6.Format(CStr(Month(dtpFecha.Value)), "00") & VB6.Format(CStr((dtpFecha.Value)), "00") & VB6.Format(Consecutivo, "0000")
        End If
        If optContado.Checked Then
            Condicion = "CO"
            TipoIngreso = "P"
        ElseIf optCredito.Checked Then
            Condicion = "CR"
            TipoIngreso = "A"
        End If
        If optPesosAnt.Checked = True Then
            MonedaAnticipo = "P"
        ElseIf optDolaresAnt.Checked = True Then
            MonedaAnticipo = "D"
        End If

        'Guardar el Movimiento de Cabecero de las Ventas
        ModStoredProcedures.PR_IMEMovimientosVentasCab(txtFolio.Text, VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), txtCodSucMatriz.Text, CStr(intCodCaja), CStr(intCodVendedor), CStr(intCodCliente), txtNombre.Text, txtRFC.Text, Condicion, IIf(optDolares.Checked = True, "D", "P"), txtTipoCambio.Text, CStr(SubTotal), CStr(Descuento), CStr(Iva), CStr(Total), CStr(RedondeoDolares), IIf(optPesosAnt.Checked = True, VB6.Format(CDbl(txtAnticipo.Text) / CDbl(txtTipoCambio.Text), "#####0.0000"), VB6.Format(txtAnticipo.Text, "#####0.00")), CStr(gcurCorpoTASAIVA), "", "V", "01/01/1900", CStr(SubTotal), CStr(Descuento), CStr(Iva), CStr(Total), CStr(RedondeoDolares), VB6.Format(txtAnticipo.Text, "#####0.00"), "", "", "01/01/1900", TipoMovto, "1", CStr(MonedaAnticipo), 0, 0, C_INSERCION, CStr(0))
        Cmd.Execute()
        NumPartida = 1
        'Guardar el Detalle de la Venta
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODARTICULO)) <> "" Then
                    ModStoredProcedures.PR_IE_MovimientosVentasDet(txtFolio.Text, CStr(NumPartida), .get_TextMatrix(I, C_COLCODARTICULO), .get_TextMatrix(I, C_COLDESCARTICULO), .get_TextMatrix(I, C_ColEXISTENCIA), "0", VB6.Format(.get_TextMatrix(I, C_COLPORCENTAJEDEDESCUENTO), "#####0.0000"), "0", VB6.Format(.get_TextMatrix(I, C_COLDESCUENTOSINIVA), "#####0.0000"), VB6.Format(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR), "#####0.0000"), VB6.Format(.get_TextMatrix(I, C_COLPRECIOSINIVA), "#####0.0000"), VB6.Format(.get_TextMatrix(I, C_COLPRECIOREAL), "#####0.0000"), VB6.Format(.get_TextMatrix(I, C_COLIVAREALCON4DECIMALES), "#####0.0000"), VB6.Format(.get_TextMatrix(I, C_COLCOSTOREAL), "#####0.0000"), "0", VB6.Format(.get_TextMatrix(I, C_COLDESCUENTOSINIVA), "#####0.0000"), VB6.Format(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR), "#####0.0000"), VB6.Format(.get_TextMatrix(I, C_COLPRECIOSINIVA), "#####0.0000"), VB6.Format(.get_TextMatrix(I, C_COLPRECIOREAL), "#####0.0000"), VB6.Format(.get_TextMatrix(I, C_COLIVAREALCON4DECIMALES), "#####0.0000"), TipoMovto, "", "", CStr(0), "", "", "", CStr(0), CStr(0), "", "", CStr(0), "0", "0", "01/01/1900", "0", "0", "0", "0", "", C_INSERCION, CStr(0))
                    Cmd.Execute()
                    NumPartida = NumPartida + 1
                End If
            Next
        End With
        'Guardar los Ingresos
        If (Condicion = "CR" And CDbl(Numerico(txtAnticipo.Text)) <> 0) Or Condicion = "CO" Then
            'Si es una venta a crédito, y el anticipo es mayor de cero, se puede guardar ingresos, o por otro lado si la venta es de contado, tambien se guardan los ingresos.
            'Guardar los Importes de Pagos.
            'If Not frmPagosSalMercancia.GuardarIngresos(txtFolio.Text, CDate(dtpFecha.Value), intCodCliente, intCodVendedor, TipoIngreso, IIf(optDolares.Checked = True, "D", "P"), CDec(txtTipoCambio.Text), "V", intCodCaja) Then
            '    Cnn.RollbackTrans()
            '    Me.Cursor = System.Windows.Forms.Cursors.Default
            '    Exit Function
            'End If
        End If
        'Generar el Folio de Salida por Venta a Vendedores Externos
        ModStoredProcedures.PR_I_FoliosAlmacen(txtCodVendExterno.Text, CStr(Consecutivo), "", CStr(0))
        Cmd.Execute()
        Consecutivo = Cmd.Parameters("Consecutivo").Value
        FolioInventario = C_PrefijoFoliosAlmacen & VB6.Format(txtCodVendExterno.Text, "00") & Year(dtpFecha.Value) & VB6.Format(Month(dtpFecha.Value), "00") & VB6.Format((dtpFecha.Value), "00") & VB6.Format(Consecutivo, "000000")
        'Guardar el Cabecero de Movimiento de Salida por Venta a Vendedor Externo
        ModStoredProcedures.PR_IE_MovtosAlmacenCab(FolioInventario, VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), txtCodVendExterno.Text, "0", "", "0", "", "0", CStr(C_SalidaPorVentadeVendedoresExternos), C_SALIDA, "", "", "", "", "V", gStrNomUsuario, "01/01/1900", "", Trim(txtFolioEntrega.Text), VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), "0", "", "01/01/1900", CStr(gcurCorpoTIPOCAMBIODOLAR), txtFolio.Text, C_INSERCION, CStr(0))
        Cmd.Execute()
        'Guardar el Detalle de Movimiento de Salida por Venta a Vendedor Externo
        NumPartida = 1

        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODARTICULO)) <> "" Then
                    '''se modifico el preciovta por precio pub dolar
                    '''                'Guarda el Detalle de Salida por Venta a Vendedor Externo
                    '''                ModStoredProcedures.PR_IE_MovtosAlmacenDet FolioInventario, CStr(NumPartida), VB6.Format(dtpFecha, C_FormatFECHAGUARDAR), _
                    ''''                .TextMatrix(I, C_COLCODARTICULO), "0", .TextMatrix(I, C_COLEXISTENCIA), CCur(Numerico(.TextMatrix(I, C_COLCOSTOREAL))), CStr(CCur(Numerico(.TextMatrix(I, C_COLPRECIOPUBDOLAR))) / (1 + Round(gcurCorpoTASAIVA / 100, 2))), _
                    ''''                "0", "V", "01/01/1900", "0", C_INSERCION, 0

                    '''20OCT2004
                    '''correccion del origen del articulo - ponia 0 por default
                    ''' txtCodSucMatriz
                    'Guarda el Detalle de Salida por Venta a Vendedor Externo
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(FolioInventario, CStr(NumPartida), VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), .get_TextMatrix(I, C_COLCODARTICULO), Trim(.get_TextMatrix(I, C_COLORIGEN)), .get_TextMatrix(I, C_ColEXISTENCIA), CStr(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOREAL)))), CStr(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR)))), "0", "V", "01/01/1900", "0", C_INSERCION, CStr(0))
                    Cmd.Execute()

                    'Guardar el Detalle de Inventario de Salida
                    ModStoredProcedures.PR_IE_Inventario(txtCodVendExterno.Text, "0", .get_TextMatrix(I, C_COLCODARTICULO), Trim(.get_TextMatrix(I, C_COLORIGEN)), "0", "0", "0", CStr(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOREAL))) * gcurCorpoTIPOCAMBIODOLAR), CStr(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOREAL)))), "0", .get_TextMatrix(I, C_ColEXISTENCIA), "0", CStr(C_SalidaPorVentadeVendedoresExternos), VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    NumPartida = NumPartida + 1
                End If
            Next
        End With

        'Si se pago con tarjeta se genera folio adicional y se actualizan los datos adicionales de la tabla movimientosventasdet
        If gblnPagoVentasconTarjeta Then
            'Genero folio adicional
            FolioAdicional = GeneraFolioAdicional(txtFolio.Text)
            'Guardo los Datos Adicionales en la tabla de MovimientosVentasDet
            With flexDetalle
                For I = 1 To .Rows - 1
                    If CDbl(Numerico(.get_TextMatrix(I, C_COLCODARTICULO))) > 0 Then
                        gStrSql = "SELECT DescFamilia FROM CatArticulos A INNER JOIN CatFamilias F ON A.CodFamilia = F.CodFamilia AND A.CodGrupo = F.CodGrupo " & "WHERE A.CodArticulo = " & Numerico(.get_TextMatrix(I, C_COLCODARTICULO))
                        ModEstandar.BorraCmd()
                        Cmd.CommandText = "dbo.Up_Select_Datos"
                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                        RsGral = Cmd.Execute
                        If RsGral.RecordCount > 0 Then
                            DescFamilia = Trim(RsGral.Fields("DescFamilia").Value)
                        Else
                            DescFamilia = ""
                        End If
                        ModStoredProcedures.PR_IE_MovimientosVentasDet(txtFolio.Text, CStr(I), .get_TextMatrix(I, C_COLCODARTICULO), "", CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), "", DescFamilia, "", CStr(0), FolioAdicional, "", "V", .get_TextMatrix(I, C_ColEXISTENCIA), Trim(txtTipoCambio.Text), IIf(optDolares.Checked = True, "D", "P"), Condicion, CStr(gcurCorpoTASAIVA), CStr(RedondeoDolares), VB6.Format(txtAnticipo.Text, "#####0.00"), VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), txtCodSucMatriz.Text, CStr(intCodCaja), CStr(intCodVendedor), CStr(intCodCliente), "", C_MODIFICACION, CStr(8))
                        Cmd.Execute()
                    End If
                Next
            End With
        End If

        Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("La Liquidación ha sido grabada correctamente con el código: " & txtFolio.Text, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        ModCorporativo.CargarRutaImpresoras()

        If optPesos.Checked = True Then
            ', txtCodSucMatriz
            ModCorporativo.TicketVentaPesos(txtFolio.Text, intCodCaja, Condicion)
        ElseIf optDolares.Checked = True Then
            ', txtCodSucMatriz
            ModCorporativo.TicketVenta(txtFolio.Text, intCodCaja, Condicion)
        End If

        'Descargar el Formulario de Pagos
        'frmPagosSalMercancia.Close()
        'frmPVRegCheque.Close()
        'frmPVRegNotasCred.Close()
        'frmPVRegTarjeta_PV.Close()
        Limpiar()

Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Sub CalculaImporte()

        If (bandera = False) Then
            Exit Sub
        End If

        Dim I As Integer
        txtSubtotal.Text = CStr(0)
        txtIVA.Text = CStr(0)
        txtDescuento.Text = CStr(0)
        txtTotal.Text = CStr(0)
        txtRedondeo.Text = CStr(0)
        txtTotalPesos.Text = CStr(0)
        txtTotalDolares.Text = CStr(0)
        flexDetalle.Rows = 11
        With Me.flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODARTICULO)) = "" Then Exit For
                .set_TextMatrix(I, C_COLDESCUENTOCONIVA, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR))) * System.Math.Round(CDbl(Numerico(.get_TextMatrix(I, C_COLPORCENTAJEDEDESCUENTO))) / 100, 2), "###,##0.00"))
                .set_TextMatrix(I, C_ColIMPORTE, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_ColEXISTENCIA))) * CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR))), "###,##0.00"))
                .set_TextMatrix(I, C_COLPRECIOSINIVA, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR))) / (1 + System.Math.Round(gcurCorpoTASAIVA / 100, 2)), "###,##0.0000"))
                .set_TextMatrix(I, C_COLDESCUENTOSINIVA, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOSINIVA))) * System.Math.Round(CDbl(Numerico(.get_TextMatrix(I, C_COLPORCENTAJEDEDESCUENTO))) / 100, 2), "###,##0.0000"))
                .set_TextMatrix(I, C_COLPRECIOREAL, VB6.Format(((CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR))) / (1 + System.Math.Round(gcurCorpoTASAIVA / 100, 2))) - (CDbl(Numerico(.get_TextMatrix(I, C_COLDESCUENTOSINIVA))))) * (1 + System.Math.Round(gcurCorpoTASAIVA / 100, 2)), "###,##0.0000"))
                .set_TextMatrix(I, C_COLIVAREALCON4DECIMALES, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOREAL))) / (1 + System.Math.Round(gcurCorpoTASAIVA / 100, 2))) * System.Math.Round(gcurCorpoTASAIVA / 100, 2), "###,##0.0000"))
                .set_TextMatrix(I, C_COLIVAREALCON2DECIMALES, VB6.Format(.get_TextMatrix(I, C_COLIVAREALCON4DECIMALES), "###,##0.00"))
                .set_TextMatrix(I, C_COLIMPORTECONDESCTO, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOREAL))) * CDbl(Numerico(.get_TextMatrix(I, C_ColEXISTENCIA))), "###,##0.00"))
            Next
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODARTICULO)) = "" Then Exit For
                txtSubtotal.Text = VB6.Format(CDbl(Numerico(txtSubtotal.Text)) + CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOSINIVA))) * CDbl(Numerico(.get_TextMatrix(I, C_ColEXISTENCIA))), "#####0.0000")
                txtDescuento.Text = VB6.Format(CDbl(Numerico(txtDescuento.Text)) + (CDbl(Numerico(.get_TextMatrix(I, C_COLDESCUENTOSINIVA))) * CDbl(Numerico(.get_TextMatrix(I, C_ColEXISTENCIA)))), "#####0.0000")
                txtIVA.Text = VB6.Format(CDbl(Numerico(txtIVA.Text)) + (CDbl(Numerico(.get_TextMatrix(I, C_COLIVAREALCON4DECIMALES))) * CDbl(Numerico(.get_TextMatrix(I, C_ColEXISTENCIA)))), "#####0.0000")
                txtTotal.Text = VB6.Format((CDbl(Numerico(txtSubtotal.Text)) - CDbl(Numerico(txtDescuento.Text))) + CDbl(Numerico(txtIVA.Text)), "#####0.0000")
            Next

            SubTotal = CDbl(txtSubtotal.Text)
            txtSubtotal.Text = VB6.Format(txtSubtotal.Text, "###,##0.00")
            Descuento = CDbl(txtDescuento.Text)
            txtDescuento.Text = VB6.Format(txtDescuento.Text, "###,##0.00")
            Iva = CDbl(txtIVA.Text)
            txtIVA.Text = VB6.Format(txtIVA.Text, "###,##0.00")


            TotalDolares = 0
            TotalPesos = 0
            Total = CDbl(txtTotal.Text)
            txtTotal.Text = VB6.Format((String.Concat(txtTotal.Text, "###,##0.00")))
            TipoCambio = CDbl(VB6.Format(String.Concat(txtTipoCambio.Text, "#####0.00")))
            TotalPesos = CDbl(VB6.Format(CDbl(String.Concat(Numerico(txtTotal.Text) * TipoCambio)), "#####0.000000"))
            TotalPesos = CDbl(VB6.Format((String.Concat(TotalPesos, "#####0.00"))))
            TotalDolares = CDbl(Numerico((String.Concat(VB6.Format(txtTotal.Text, "#####0.0000")))))
            RedondeoPesos = 0
            RedondeoDolares = 0

            '''RedondeoPesos = ModCorporativo.RedondeoUnidadFinal(CCur(TotalPesos), CDbl(gcurRedondeo))
            '''RedondeoDolares = VB6.Format(RedondeoPesos / TipoCambio, "#####0.0000")
            '''txtRedondeo = VB6.Format(RedondeoDolares, "###,##0.00")
            ''OJO
            If optPesos.Checked Then
                RedondeoPesos = ModCorporativo.RedondeoUnidadFinal(CDbl(TotalPesos), CDbl(gcurRedondeo))
                If RedondeoPesos = 0 And TipoCambio = 0 Then RedondeoDolares = 0 Else RedondeoDolares = System.Math.Round(RedondeoPesos / TipoCambio, 4)
            Else
                RedondeoDolares = ModCorporativo.RedondeoUnidadFinal(CDbl(Total), 1)
                If RedondeoDolares = 0 And TipoCambio = 0 Then RedondeoPesos = 0 Else RedondeoPesos = System.Math.Round(RedondeoDolares * TipoCambio, 4)
            End If
            txtRedondeo.Text = VB6.Format(RedondeoDolares, "###,##0.00")
            TotalPesos = System.Math.Round(RedondeoPesos + TotalPesos, 1)
            TotalDolares = System.Math.Round(TotalDolares + RedondeoDolares, 2)
            '''RedondeoGralVenta = RedondeoD
            '''OJO
            TotalDolares = CDbl(Numerico(VB6.Format(txtTotal.Text, "#####0.0000"))) + RedondeoDolares
            txtTotalDolares.Text = VB6.Format(TotalDolares, "###,##0.00")
            txtTotalPesos.Text = VB6.Format(System.Math.Round(TotalDolares * TipoCambio, 1), "###,##0.00")
            '''txtTotalPesos = VB6.Format(txtTotalPesos, "###,##0.00")
            '''        TotalDolares = CDbl(Numerico(VB6.Format(txtTotal, "#####0.0000"))) + RedondeoDolares
            '''        txtTotalDolares = VB6.Format(TotalDolares, "###,##0.00")
            '''        txtTotalPesos = VB6.Format(TotalPesos + RedondeoPesos, "#####0.0")
            '''        txtTotalPesos = VB6.Format(txtTotalPesos, "###,##0.00")
        End With
    End Sub

    Function ObtenerDescuentoJoyeria(ByRef PrecioPubDol As Double) As Double
        Dim RsAux As ADODB.Recordset
        On Error GoTo Err_Renamed
        If PrecioPubDol = 0 Then
            ObtenerDescuentoJoyeria = 0
            Exit Function
        End If
        gStrSql = "SELECT * FROM CatDesctosVExternos WHERE CodGrupo = " & gCODJOYERIA
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsAux = Cmd.Execute
        If RsAux.RecordCount > 0 Then
            Do While Not RsAux.EOF
                If PrecioPubDol >= RsAux.Fields("ImporteIni").Value And PrecioPubDol <= RsAux.Fields("ImporteFin").Value Then
                    ObtenerDescuentoJoyeria = RsAux.Fields("PorcDescto").Value
                    Exit Function
                Else
                    If PrecioPubDol >= RsAux.Fields("ImporteIni").Value And RsAux.Fields("ImporteFin").Value = 0 Then
                        ObtenerDescuentoJoyeria = RsAux.Fields("PorcDescto").Value
                        Exit Function
                    End If
                End If
                RsAux.MoveNext()
            Loop
        End If
        ObtenerDescuentoJoyeria = 0
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ObtenerDescuentoRelojeria(ByRef CodMArca As Integer) As Double
        Dim RsAux As ADODB.Recordset
        On Error GoTo Err_Renamed
        gStrSql = "SELECT * FROM CatDesctosVExternos WHERE CodGrupo = " & gCODRELOJERIA
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsAux = Cmd.Execute
        If RsAux.RecordCount > 0 Then
            Do While Not RsAux.EOF
                If RsAux.Fields("CodMArca").Value = CodMArca Then
                    ObtenerDescuentoRelojeria = RsAux.Fields("PorcDescto").Value
                    Exit Function
                End If
                RsAux.MoveNext()
            Loop
        End If
        ObtenerDescuentoRelojeria = 0
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ObtenerDescuentoVarios(ByRef CodFamilia As Integer) As Double
        Dim RsAux As ADODB.Recordset
        On Error GoTo Err_Renamed
        gStrSql = "SELECT * FROM CatDesctosVExternos WHERE CodGrupo = " & gCODVARIOS
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsAux = Cmd.Execute
        If RsAux.RecordCount > 0 Then
            Do While Not RsAux.EOF
                If RsAux.Fields("CodFamilia").Value = CodFamilia Then
                    ObtenerDescuentoVarios = RsAux.Fields("PorcDescto").Value
                    Exit Function
                End If
                RsAux.MoveNext()
            Loop
        End If
        ObtenerDescuentoVarios = 0
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub BuscaExistencias()
        On Error GoTo Merr
        Dim I As Integer
        Dim RsAux As ADODB.Recordset
        Dim NumArt As Integer

        NumArt = 0
        If Trim(txtFolioEntrega.Text) = "" Then Exit Sub
        '    gStrSql = "SELECT CA.CODARTICULO,CA.DESCARTICULO,UNI.DESCUNIDAD," & _
        ''    "SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS)) AS EXISTENCIA,CA.CODGRUPO,CA.CODMARCA,CA.CODFAMILIA," & _
        ''    "CASE WHEN CA.PESOSFIJOS = 0 THEN CA.PRECIOPUBDOLAR WHEN CA.PESOSFIJOS = 1 THEN CA.PRECIOPUBDOLAR / " & Numerico(txtTipoCambio) & " END AS PRECIOPUBDOLAR,CA.COSTOREAL,CLI.CODCLIENTE,CLI.DESCCLIENTE,CLI.DOMICILIO,CLI.RFC,ltrim(rtrim(CLI.TelCasa)) + '     ' + ltrim(rtrim(CLI.TelOficina)) + '     ' + ltrim(rtrim(CLI.Fax)) as Telefono " & _
        ''    "FROM CATALMACEN SUC INNER JOIN INVENTARIO I ON SUC.CODALMACEN = I.CODALMACEN " & _
        ''    "INNER JOIN CATARTICULOS CA ON CA.CODARTICULO = I.CODARTICULO " & _
        ''    "INNER JOIN CATUNIDADES UNI ON CA.CODUNIDAD = UNI.CODUNIDAD " & _
        ''    "LEFT OUTER JOIN CATCLIENTES CLI ON SUC.CODALMACEN = CLI.ALMACENVEXT " & _
        ''    "WHERE I.CODALMACEN = " & txtCodVendExterno & " " & _
        ''    "GROUP BY CA.CODARTICULO,CA.DESCARTICULO,UNI.DESCUNIDAD,CA.CODGRUPO,CA.CODMARCA,CA.CODFAMILIA,CA.PESOSFIJOS,CA.PRECIOPUBDOLAR," & _
        ''    "CA.COSTOREAL,CLI.CODCLIENTE,CLI.DESCCLIENTE,CLI.DOMICILIO,CLI.RFC,CLI.TELCASA,CLI.TELOFICINA,CLI.FAX " & _
        ''    "HAVING SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS)) > 0 " & _
        ''    "ORDER BY CA.CODARTICULO,CA.DESCARTICULO"

        gStrSql = "SELECT CAB.FOLIOALMACEN,CAB.CODALMACENREF,SUC.DESCALMACEN,DET.CODARTICULO,CA.DESCARTICULO,UNI.DESCUNIDAD," & "SUM(DET.CANTIDAD) AS CANTENTRADA,ISNULL(VEXT.CANTSALXDEV,0) AS CANTDEVOL,ISNULL(VEXT.CANTSALXVENTA,0) AS CANTVENTA," & "CA.CODGRUPO,CA.CODMARCA,CA.CODFAMILIA,CASE WHEN CA.PESOSFIJOS = 0 THEN CA.PRECIOPUBDOLAR WHEN CA.PESOSFIJOS = 1 THEN CA.PRECIOPUBDOLAR / 11.6 END AS PRECIOPUBDOLAR," & "CA.COSTOREAL,CLI.CODCLIENTE,CLI.DESCCLIENTE,CLI.DOMICILIO,CLI.RFC,ltrim(rtrim(CLI.TelCasa)) + '     ' + ltrim(rtrim(CLI.TelOficina)) + '     ' + ltrim(rtrim(CLI.Fax)) as Telefono, CA.CodAlmacenOrigen  " & "FROM MOVTOSALMACENCAB CAB INNER JOIN MOVTOSALMACENDET DET ON CAB.FOLIOALMACEN = DET.FOLIOALMACEN " & "LEFT OUTER JOIN (SELECT CAB.REFERENCIADEORIGEN,DET.CODARTICULO,SUM(CASE WHEN CODMOVTOALM = " & C_SalidadeAlmacendeVendedorExterno & " THEN DET.CANTIDAD ELSE 0 END) AS CANTSALXDEV," & "SUM(CASE WHEN CODMOVTOALM = " & C_SalidaPorVentadeVendedoresExternos & " THEN DET.CANTIDAD ELSE 0 END) AS CANTSALXVENTA FROM MOVTOSALMACENCAB CAB " & "INNER JOIN MOVTOSALMACENDET DET ON CAB.FOLIOALMACEN = DET.FOLIOALMACEN WHERE CAB.REFERENCIADEORIGEN = '" & Trim(txtFolioEntrega.Text) & "' AND CAB.ESTATUS <> 'C' GROUP BY CAB.REFERENCIADEORIGEN,DET.CODARTICULO) VEXT " & "ON CAB.FOLIOALMACEN = VEXT.REFERENCIADEORIGEN AND DET.CODARTICULO = VEXT.CODARTICULO INNER JOIN CATALMACEN SUC ON CAB.CODALMACENREF = SUC.CODALMACEN " & "INNER JOIN CATARTICULOS CA ON CA.CODARTICULO = DET.CODARTICULO INNER JOIN CATUNIDADES UNI ON CA.CODUNIDAD = UNI.CODUNIDAD " & "LEFT OUTER JOIN CATCLIENTES CLI ON SUC.CODALMACEN = CLI.ALMACENVEXT WHERE CAB.FOLIOALMACEN = '" & Trim(txtFolioEntrega.Text) & "' AND CAB.CODMOVTOALM = " & C_SalidaAVendedoresExternos & " AND CAB.ESTATUS <> 'C' " & "GROUP BY CAB.FOLIOALMACEN,CAB.CODALMACENREF,SUC.DESCALMACEN,DET.CODARTICULO,CA.DESCARTICULO,UNI.DESCUNIDAD,VEXT.CANTSALXDEV," & "VEXT.CANTSALXVENTA,CA.CODGRUPO,CA.CODMARCA,CA.CODFAMILIA,CA.PESOSFIJOS,CA.PRECIOPUBDOLAR,CA.COSTOREAL,CLI.CODCLIENTE,CLI.DESCCLIENTE," & "CLI.Domicilio , CLI.Rfc, CLI.TelCasa, CLI.TelOficina, CLI.Fax, CA.CodAlmacenOrigen "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            If IsDBNull(RsGral.Fields("CodCliente").Value) Then
                MsgBox("Este vendedor externo no esta registrado en el catalogo de clientes. " & Chr(13) & "   Favor de registrar a este vendedor en el catalogo de clientes.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                FueraChange = True
                txtCodVendExterno.Text = ""
                dbcSucursal.Text = ""
                FueraChange = False
                Me.Activate()
                Exit Sub
            End If
            flexDetalle.Clear()
            Encabezado()
            RsGral.MoveFirst()
            With flexDetalle
                I = 1
                Do While Not RsGral.EOF
                    If RsGral.Fields("CantEntrada").Value - (RsGral.Fields("CantDevol").Value + RsGral.Fields("CantVenta").Value) > 0 Then
                        .set_TextMatrix(I, C_COLCODARTICULO, RsGral.Fields("CodArticulo").Value)
                        .set_TextMatrix(I, C_COLDESCARTICULO, RsGral.Fields("DescArticulo").Value)
                        .set_TextMatrix(I, C_COLDESCUNIDAD, RsGral.Fields("DescUnidad").Value)
                        .set_TextMatrix(I, C_ColEXISTENCIA, RsGral.Fields("CantEntrada").Value - (RsGral.Fields("CantDevol").Value + RsGral.Fields("CantVenta").Value))
                        .set_TextMatrix(I, C_COLPRECIOPUBDOLAR, VB6.Format(RsGral.Fields("PrecioPubDolar").Value, "###,##0.00"))
                        .set_TextMatrix(I, C_COLORIGEN, RsGral.Fields("CodAlmacenOrigen").Value)
                        '''If RsGral!PrecioPubDolar = 0 Then
                        '''    EmitePago = False
                        '''End If
                        If RsGral.Fields("CodGrupo").Value = gCODJOYERIA Then
                            .set_TextMatrix(I, C_COLPORCENTAJEDEDESCUENTO, VB6.Format(ObtenerDescuentoJoyeria(RsGral.Fields("PrecioPubDolar").Value), "###,##0.00"))
                        ElseIf RsGral.Fields("CodGrupo").Value = gCODRELOJERIA Then
                            .set_TextMatrix(I, C_COLPORCENTAJEDEDESCUENTO, VB6.Format(ObtenerDescuentoRelojeria(RsGral.Fields("CodMArca").Value), "###,##0.00"))
                        ElseIf RsGral.Fields("CodGrupo").Value = gCODVARIOS Then
                            .set_TextMatrix(I, C_COLPORCENTAJEDEDESCUENTO, VB6.Format(ObtenerDescuentoVarios(RsGral.Fields("CodFamilia").Value), "###,##0.00"))
                        End If
                        .set_TextMatrix(I, C_COLCOSTOREAL, RsGral.Fields("CostoReal").Value)
                        NumArt = NumArt + 1
                        RsGral.MoveNext()
                        If I = .Rows - 1 Then
                            If Not RsGral.EOF Then
                                .Rows = .Rows + 1
                            End If
                        End If
                        I = I + 1
                    Else
                        RsGral.MoveNext()
                    End If
                Loop
                If NumArt = 0 Then
                    MsgBox("Los articulos de este folio de entrega ya no existen en el inventario del vendedor externo" & vbNewLine & "        No es posible hacer una liquidación de este folio, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Limpiar()
                    Exit Sub
                End If
                CalculaImporte()
            End With
            RsGral.MoveFirst()
            intCodCliente = RsGral.Fields("CodCliente").Value
            txtNombre.Text = Trim(RsGral.Fields("DescCliente").Value)
            txtDomicilio.Text = Trim(RsGral.Fields("Domicilio").Value)
            TxtTelefono.Text = Trim(RsGral.Fields("Telefono").Value)
            txtRFC.Text = RsGral.Fields("Rfc").Value
            FueraChange = True
            txtCodVendExterno.Text = RsGral.Fields("CodALmacenREf").Value
            txtDescSucursal.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            FueraChange = False
        Else
            MsgBox("Este folio no existe o no es un folio de entrega de mercancia, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtFolioEntrega.Text = ""
            Me.Activate()
            Exit Sub
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub BuscaVendedorExterno()
        On Error GoTo Merr
        gStrSql = "SELECT DescAlmacen,TipoAlmacen FROM CatAlmacen WHERE CodAlmacen = " & txtCodVendExterno.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("TipoAlmacen").Value = "P" Then
                MsgBox("Este almacen no es un vendedor externo, Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodVendExterno.Text = ""
                txtCodVendExterno.Focus()
                Exit Sub
            Else
                'If BuscaMovimientos Then
                txtCodVendExterno.Text = txtCodVendExterno.Text
                dbcSucursal.Text = Trim(RsGral.Fields("DescAlmacen").Value)
                BuscaExistencias()
                txtNombre.Focus()
                'Else
                '    txtCodVendExterno = ""
                '    dbcSucursal.text = ""
                '    txtCodVendExterno.SetFocus
                'End If
            End If
        Else
            MsgBox("Codigo de almacen no existe, Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodVendExterno.Text = ""
            txtCodVendExterno.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Encabezado()
        With flexDetalle
            .set_Cols(0, 17)
            .Row = 0
            .Col = C_COLCODARTICULO
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(C_COLCODARTICULO, 0, 1000)
            .Text = "Código"
            .Col = C_COLDESCARTICULO
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(C_COLDESCARTICULO, 0, 3500)
            .Text = "Descripción"
            .Col = C_COLDESCUNIDAD
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(C_COLDESCUNIDAD, 0, 700)
            .Text = "Unidad"
            .Col = C_ColEXISTENCIA
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(C_ColEXISTENCIA, 0, 900)
            .Text = "Cantidad"
            .Col = C_COLPRECIOPUBDOLAR
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(C_COLPRECIOPUBDOLAR, 0, 1100)
            .Text = "Precio"
            .Col = C_COLDESCUENTOCONIVA
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(C_COLDESCUENTOCONIVA, 0, 1100)
            .Text = "Descto."
            .Col = C_COLIVAREALCON2DECIMALES
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(C_COLIVAREALCON2DECIMALES, 0, 1100)
            .Text = "IVA"
            .Col = C_ColIMPORTE
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(C_ColIMPORTE, 0, 1400)
            .Text = "Importe"
            .Col = C_COLIMPORTECONDESCTO
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(C_COLIMPORTECONDESCTO, 0, 1700)
            .Text = "Impte. con Descto."
            .Col = C_COLPRECIOSINIVA
            .set_ColWidth(C_COLPRECIOSINIVA, 0, 0)
            .Col = C_COLPRECIOREAL
            .set_ColWidth(C_COLPRECIOREAL, 0, 0)
            .Col = C_COLPORCENTAJEDEDESCUENTO
            .set_ColWidth(C_COLPORCENTAJEDEDESCUENTO, 0, 0)
            .Col = C_COLDESCUENTOSINIVA
            .set_ColWidth(C_COLDESCUENTOSINIVA, 0, 0)
            .Col = C_COLIVAREALCON4DECIMALES
            .set_ColWidth(C_COLIVAREALCON4DECIMALES, 0, 0)
            .set_ColAlignment(0, 7)
            .Col = C_COLCOSTOREAL
            .set_ColWidth(C_COLCOSTOREAL, 0, 0)
            .Col = C_COLORIGEN
            .set_ColWidth(C_COLORIGEN, 0, 0)

            .Rows = 11
            .Col = C_COLCODARTICULO
            .Row = 1
        End With
    End Sub

    Sub LlenaDatos()
        On Error GoTo Merr
        Dim I As Integer
        If Trim(txtFolio.Text) = "" Then
            Exit Sub
        End If
        Nuevo()
        gStrSql = "SELECT VTACAB.FOLIOVENTA,VTACAB.FECHAVENTA,ALM.CODALMACEN,ALM.DESCALMACEN,VTACAB.CODCLIENTE," & "CATCLI.DESCCLIENTE,CATCLI.RFC,CATCLI.DOMICILIO,VTACAB.CONDICION,VTACAB.TIPOCAMBIO,VTACAB.MONEDA," & "CATVEND.DESCVENDEDOR,VTADET.CODARTICULO,ART.DESCARTICULO,UNI.DESCUNIDAD,VTADET.CANTIDAD," & "VTADET.PRECIOLISTA,VTADET.IMPTEDESCUENTOS,VTADET.IVAREAL,VTADET.PRECIOREAL," & "VTACAB.SubTotal , VTACAB.Descuento, VTACAB.Iva, VTACAB.Total, VTACAB.Redondeo,VTACAB.PORCIVA, " & "VTACAB.MonedaAnticipo,VTACAB.Anticipo,ltrim(rtrim(CATCLI.TelCasa)) + '     ' + ltrim(rtrim(CATCLI.TelOficina)) + '     ' + ltrim(rtrim(CATCLI.Fax)) as Telefono,CAB.ReferenciadeOrigen " & "FROM MOVIMIENTOSVENTASCAB VTACAB INNER JOIN MOVIMIENTOSVENTASDET VTADET ON VTACAB.FOLIOVENTA = VTADET.FOLIOVENTA " & "INNER JOIN MOVTOSALMACENCAB CAB ON VTACAB.FOLIOVENTA = CAB.FOLIOVENTA " & "INNER JOIN CATCLIENTES CATCLI ON VTACAB.CODCLIENTE = CATCLI.CODCLIENTE " & "INNER JOIN CATALMACEN ALM ON CATCLI.ALMACENVEXT = ALM.CODALMACEN " & "INNER JOIN CATARTICULOS ART ON VTADET.CODARTICULO = ART.CODARTICULO " & "INNER JOIN CATUNIDADES UNI ON ART.CODUNIDAD = UNI.CODUNIDAD " & "INNER JOIN CATVENDEDORES CATVEND ON VTACAB.CODVENDEDOR = CATVEND.CODVENDEDOR " & "WHERE VTACAB.FOLIOVENTA = '" & txtFolio.Text & "' AND VTACAB.VTAVEXT = 1 "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            FueraChange = True
            txtCodVendExterno.Text = RsGral.Fields("CodAlmacen").Value
            txtDescSucursal.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            dtpFecha.Value = VB6.Format(RsGral.Fields("FechaVenta").Value, "dd/mmm/yyyy")
            txtNombre.Text = Trim(RsGral.Fields("DescCliente").Value)
            txtDomicilio.Text = Trim(RsGral.Fields("Domicilio").Value)
            TxtTelefono.Text = Trim(RsGral.Fields("Telefono").Value)
            txtRFC.Text = Trim(RsGral.Fields("Rfc").Value)
            txtFolioEntrega.Text = RsGral.Fields("ReferenciaDeOrigen").Value
            FueraChange = False
            If RsGral.Fields("Condicion").Value = "CO" Then
                optContado.Checked = True
            ElseIf RsGral.Fields("Condicion").Value = "CR" Then
                optCredito.Checked = True
            End If
            If RsGral.Fields("Moneda").Value = "D" Then
                optDolares.Checked = True
            ElseIf RsGral.Fields("Moneda").Value = "P" Then
                optPesos.Checked = True
            End If
            txtTipoCambio.Text = VB6.Format(RsGral.Fields("TipoCambio").Value, "###,##0.00")
            dbcVendedor.Text = Trim(RsGral.Fields("DescVendedor").Value)
            txtSubtotal.Text = VB6.Format(RsGral.Fields("SubTotal").Value, "###,##0.00")
            txtDescuento.Text = VB6.Format(RsGral.Fields("Descuento").Value, "###,##0.00")
            txtIVA.Text = VB6.Format(RsGral.Fields("Iva").Value, "###,##0.00")
            txtTotal.Text = VB6.Format(RsGral.Fields("Total").Value, "###,##0.00")
            txtRedondeo.Text = VB6.Format(RsGral.Fields("Redondeo").Value, "###,##0.00")
            txtTotalDolares.Text = VB6.Format(CDec(Numerico(txtTotal.Text)) + CDec(Numerico(txtRedondeo.Text)), "###,##0.00")
            txtTotalPesos.Text = VB6.Format(CDec(Numerico(txtTotalDolares.Text)) * CDec(Numerico(txtTipoCambio.Text)), "###,##0.0")
            txtTotalPesos.Text = VB6.Format(txtTotalPesos.Text, "###,##0.00")
            If RsGral.Fields("MonedaAnticipo").Value = "P" And RsGral.Fields("Condicion").Value = "CR" Then
                txtAnticipo.Text = VB6.Format(System.Math.Round(CDbl(Numerico(RsGral.Fields("Anticipo").Value)) * CDbl(Numerico(txtTipoCambio.Text)), 1), "###,##0.00")
                optPesosAnt.Checked = True
                txtSaldo.Text = VB6.Format(System.Math.Round(CDbl(txtTotalPesos.Text) - CDbl(txtAnticipo.Text)), "###,##0.00")
            ElseIf RsGral.Fields("MonedaAnticipo").Value = "D" And RsGral.Fields("Condicion").Value = "CR" Then
                optDolaresAnt.Checked = True
                txtAnticipo.Text = VB6.Format(System.Math.Round(CDbl(Numerico(RsGral.Fields("Anticipo").Value)), 2), "###,##0.00")
                txtSaldo.Text = VB6.Format(System.Math.Round(CDbl(txtTotalDolares.Text) - CDbl(txtAnticipo.Text)), gstrFormatoCantidad)
            End If
            With flexDetalle
                I = 1
                .Row = 1
                Do While Not RsGral.EOF
                    .set_TextMatrix(I, C_COLCODARTICULO, RsGral.Fields("CodArticulo").Value)
                    .set_TextMatrix(I, C_COLDESCARTICULO, RsGral.Fields("DescArticulo").Value)
                    .set_TextMatrix(I, C_COLDESCUNIDAD, RsGral.Fields("DescUnidad").Value)
                    .set_TextMatrix(I, C_ColEXISTENCIA, RsGral.Fields("Cantidad").Value)
                    .set_TextMatrix(I, C_COLPRECIOPUBDOLAR, VB6.Format(RsGral.Fields("PrecioLista").Value, "###,##0.00"))
                    .set_TextMatrix(I, C_COLDESCUENTOCONIVA, System.Math.Round(RsGral.Fields("ImpteDescuentos").Value * (1 + System.Math.Round(RsGral.Fields("PorcIva").Value / 100, 2)), 1))
                    .set_TextMatrix(I, C_COLDESCUENTOCONIVA, VB6.Format(.get_TextMatrix(I, C_COLDESCUENTOCONIVA), "###,##0.00"))
                    .set_TextMatrix(I, C_COLIVAREALCON2DECIMALES, VB6.Format(RsGral.Fields("IvaReal").Value, "###,##0.00"))
                    .set_TextMatrix(I, C_ColIMPORTE, VB6.Format(RsGral.Fields("Cantidad").Value * (RsGral.Fields("PrecioLista")).Value, "###,##0.00"))
                    .set_TextMatrix(I, C_COLIMPORTECONDESCTO, VB6.Format(RsGral.Fields("Cantidad").Value * (RsGral.Fields("PrecioReal")).Value, "###,##0.00"))
                    RsGral.MoveNext()
                    If Not RsGral.EOF Then
                        If I = .Rows - 1 Then
                            .Rows = .Rows + 1
                            .Row = .Row + 1
                        Else
                            .Row = .Row + 1
                        End If
                    End If
                    I = I + 1
                Loop
            End With
            mblnNuevo = False
            Frame2.Enabled = False
            Frame3.Enabled = False
            Frame5.Enabled = False
            Frame6.Enabled = False
            'txtCodVendExterno.Enabled = False
            dbcSucursal.Enabled = False
            dbcVendedor.Enabled = False
            cmdABCClientes.Enabled = False
            cmdProcesarPago.Enabled = False
            txtFolioEntrega.Enabled = False
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Limpiar()
        InicializaVariables()
        Nuevo()
        txtFolioEntrega.Text = ""
        txtFolio.Text = ""
        txtFolio.Focus()
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
        mblnNuevo = True
        mblnCambios = False
        gblnPagoVentasconTarjeta = False
        intCodSucursal = 0
        intCodCliente = 0
        intCodVendedor = 0
        SubTotal = 0
        Descuento = 0
        Iva = 0
        Total = 0
        Redondeo = 0
        RedondeoDolares = 0
        RedondeoPesos = 0
        TotalPesos = 0
        TotalDolares = 0
        TipoMovto = "V"
    End Sub

    Sub Nuevo()
        FueraChange = True
        dtpFecha.Value = Today
        txtCodVendExterno.Text = ""
        'dbcSucursal.text = ""
        txtNombre.Text = ""
        txtDomicilio.Text = ""
        txtRFC.Text = ""
        TxtTelefono.Text = ""
        txtFolioEntrega.Text = ""
        optContado.Checked = True
        txtTipoCambio.Text = VB6.Format(gcurCorpoTIPOCAMBIODOLAR, "###,##0.00")
        dbcVendedor.Text = ""
        txtDescSucursal.Text = ""
        flexDetalle.Clear()
        Encabezado()
        txtSubtotal.Text = "0.00"
        txtDescuento.Text = "0.00"
        txtIVA.Text = "0.00"
        txtTotal.Text = "0.00"
        txtRedondeo.Text = "0.00"
        txtTotalDolares.Text = "0.00"
        txtTotalPesos.Text = "0.00"
        Frame2.Enabled = True
        Frame3.Enabled = True
        Frame5.Enabled = True
        Frame6.Enabled = True
        txtCodVendExterno.Enabled = True
        'dbcSucursal.Enabled = True
        dbcVendedor.Enabled = True
        cmdABCClientes.Enabled = True
        cmdProcesarPago.Enabled = True
        optDolares.Checked = False
        optPesos.Checked = True
        optPesosAnt.Enabled = False
        optDolaresAnt.Enabled = False
        txtAnticipo.Enabled = False
        txtSaldo.Enabled = False
        txtFolioEntrega.Enabled = True
        txtAnticipo.Text = "0.00"
        txtSaldo.Text = "0.00"
        EmitePago = True
        FueraChange = False
    End Sub

    Sub PonerTotalesVentaenFrmPagos()
        'Este Proc. Transporta los datos existentes en el Form de Ventas Sal. de Mercancia, al Formulario de PAgos.
        'Los Datos que se pasan son: Subtotal, IVa, IEPS, Total,  Tpo Dólar.
        On Error GoTo Merr
        Dim SubtotalD As Decimal
        Dim SubtotalD4Decimales As Decimal
        Dim ImpIva As Decimal
        Dim ImpIva4Decimales As Decimal
        Dim Descuento As Decimal
        Dim TipoCambioDolar As Decimal
        Dim TotalP As Decimal
        Dim TotalP4decimales As Decimal
        Dim TotalD4Decimales As Decimal
        Dim TotalD As Decimal
        Dim PrecioListaSinIva As Decimal
        Dim IvaReal As Decimal

        ' En el Formulario de Pagos, existe un  Textbox apra alamacenar el Nombre del FOrmulario que ha invocado al FOrmulario de Pagos.
        ' Para posteriormente saber que función de Guardar se jecuta. (En este momento el FOrmulario de Pagos, se usa en en VEntas y Apartados.)
        'frmPagosSalMercancia.txtFormaOrigen.Text = Me.Name

        PrecioListaSinIva = ObtenerTotalPrecioListaSinIva()
        Descuento = ObtenerTotalPromDescuento()
        IvaReal = ObtenerTotalIvaReal()
        'Si es una venta a crédito y existe anticipo, el total el pago será el imoprte de anticipo el iva, se saca de el
        'Se está considerando que el ANticipo es en Pesos
        TipoCambioDolar = CDec(Numerico(txtTipoCambio.Text))
        If CDbl(Numerico(txtAnticipo.Text)) > 0 Then
            'Validar si el Importe es en Pesos o Dólares
            If optPesosAnt.Checked = True Then
                TotalP = FormateoDecimales(txtAnticipo)
                SubtotalD = FormateoDecimales(TotalP / TipoCambioDolar)
                ImpIva = FormateoDecimales(SubtotalD * (gcurCorpoTASAIVA / 100) / (1 + gcurCorpoTASAIVA / 100))
                SubtotalD = FormateoDecimales(SubtotalD - ImpIva)
                TotalD = FormateoDecimales(SubtotalD + ImpIva)
                Descuento = FormateoDecimales(0)

                TotalP4decimales = CDec(Numerico(txtAnticipo.Text))
                SubtotalD4Decimales = TotalP4decimales / TipoCambioDolar
                ImpIva4Decimales = SubtotalD4Decimales * (gcurCorpoTASAIVA / 100) / (1 + gcurCorpoTASAIVA / 100)
                SubtotalD4Decimales = SubtotalD4Decimales - ImpIva4Decimales
                TotalD4Decimales = SubtotalD4Decimales + ImpIva4Decimales
                Descuento = FormateoDecimales(0)
            ElseIf optDolaresAnt.Checked = True Then
                SubtotalD = FormateoDecimales(txtAnticipo)
                ImpIva = FormateoDecimales(SubtotalD * (gcurCorpoTASAIVA / 100) / (1 + gcurCorpoTASAIVA / 100))
                SubtotalD = FormateoDecimales(SubtotalD - ImpIva)
                TotalD = FormateoDecimales(SubtotalD + ImpIva)
                TotalP = FormateoDecimales(TotalD * TipoCambioDolar)
                Descuento = FormateoDecimales(0)
                SubtotalD4Decimales = CDec(Numerico(txtAnticipo.Text))
                ImpIva4Decimales = SubtotalD4Decimales * (gcurCorpoTASAIVA / 100) / (1 + gcurCorpoTASAIVA / 100)
                SubtotalD4Decimales = SubtotalD4Decimales - ImpIva4Decimales
                TotalD4Decimales = SubtotalD4Decimales + ImpIva4Decimales
                TotalP4decimales = TotalD4Decimales * TipoCambioDolar
                Descuento = FormateoDecimales(0)
            End If
        Else

            '        SubtotalD4Decimales = PrecioListaSinIva + CDbl(Numerico(txtRedondeo))
            '        ImpIva4Decimales = IvaReal
            '        TotalD4Decimales = SubtotalD4Decimales - Descuento + ImpIva4Decimales
            '        TotalP4decimales = TotalD4Decimales * TipoCambioDolar

            SubtotalD4Decimales = PrecioListaSinIva + CDec(Numerico(txtRedondeo.Text))
            ImpIva4Decimales = IvaReal
            TotalD4Decimales = SubtotalD4Decimales - Descuento + ImpIva4Decimales
            TotalP4decimales = TotalD4Decimales * TipoCambioDolar
            ImpIva = System.Math.Round(CDbl(Numerico(txtTotal.Text)) * (gcurCorpoTASAIVA / 100) / (1 + gcurCorpoTASAIVA / 100), 2)
            TotalD = CDec(Numerico(txtTotalDolares.Text))
            SubtotalD = System.Math.Round(TotalD + Descuento - ImpIva, 2)
            TotalP = FormateoDecimales(Numerico(txtTotalPesos.Text))
        End If
        'With frmPagosSalMercancia
        '    .txtSubTotal.Text = VB6.Format(FormateoDecimales(SubtotalD), gstrVB6.FormatoCantidad)
        '    .txtIVA.Text = VB6.Format(FormateoDecimales(ImpIva), gstrVB6.FormatoCantidad)
        '    .txtTotal.Text = VB6.Format(FormateoDecimales(TotalD), gstrVB6.FormatoCantidad)
        '    .txtDescuento.Text = VB6.Format(Descuento, gstrVB6.FormatoCantidad)
        '    .txtmnAPagar.Text = VB6.Format(FormateoDecimales(TotalP), gstrVB6.FormatoCantidad)
        '    .txtmnTotalPago.Text = VB6.Format(FormateoDecimales(0), gstrVB6.FormatoCantidad)
        '    .txtmnCambio.Text = VB6.Format(FormateoDecimales(CDbl(.txtmnTotalPago.Text) - CDbl(.txtmnAPagar.Text)), gstrVB6.FormatoCantidad)
        '    .txtdoAPagar.Text = VB6.Format(FormateoDecimales(TotalD), gstrVB6.FormatoCantidad)
        '    .txtdoTotalPago.Text = VB6.Format(FormateoDecimales(0), gstrVB6.FormatoCantidad)
        '    .txtdoCambio.Text = VB6.Format(FormateoDecimales(CDbl(.txtdoTotalPago.Text) - CDbl(.txtdoAPagar.Text)), gstrVB6.FormatoCantidad)
        '    .txtDolar.Text = VB6.Format(FormateoDecimales(TipoCambioDolar), gstrVB6.FormatoCantidad)
        '    .txtSubtotal4Decimales.Text = CStr(SubtotalD4Decimales)
        '    .txtIVA4Decimales.Text = CStr(ImpIva4Decimales)
        '    .txtTotal4Decimales.Text = CStr(TotalD4Decimales)
        '    .txtDescuento4Decimales.Text = CStr(Descuento)
        '    .txtmnAPagar4Decimales.Text = CStr(TotalP4decimales)
        '    .txtmnTotalPago4Decimales.Text = CStr(0)
        '    .txtmnCambio4Decimales.Text = CStr(CDbl(Numerico(.txtmnTotalPago4Decimales.Text)) - CDbl(Numerico(.txtmnAPagar4Decimales.Text)))
        '    .txtdoAPagar4Decimales.Text = CStr(TotalD4Decimales)
        '    .txtdoTotalPago4Decimales.Text = CStr(0)
        '    .txtdoCambio4Decimales.Text = CStr(CDbl(Numerico(.txtdoTotalPago4Decimales.Text)) - CDbl(Numerico(.txtdoAPagar4Decimales.Text)))
        'End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function ObtenerTotalPrecioListaSinIva() As Decimal
        Dim I As Integer
        'Esta Función obtiene el Total de la Columna de Precio Lista sin Iva, del total del detalle de la venta.
        ObtenerTotalPrecioListaSinIva = 0
        With flexDetalle
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, C_COLCODARTICULO) = "" Then Exit For
                ObtenerTotalPrecioListaSinIva = ObtenerTotalPrecioListaSinIva + (CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOSINIVA))) * CDbl(Numerico(.get_TextMatrix(I, C_ColEXISTENCIA))))
            Next
        End With
        ObtenerTotalPrecioListaSinIva = ObtenerTotalPrecioListaSinIva
    End Function

    Function ObtenerTotalPromDescuento() As Decimal
        Dim I As Integer
        'Esta Función obtiene el Total de la Columna de Promoción + Descuento de todo el Detalle de la Venta.
        'Multiplicado por  la cantidad en cada una de las filas. Es decir, Obtiene el Total de Promocion y Descuento que se dio al cliente (Importe)
        ObtenerTotalPromDescuento = 0
        Dim ImpteDescuento As Decimal
        Dim ImptePromocion As Decimal
        Dim Cantidad As Integer
        With flexDetalle
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, C_COLCODARTICULO) = "" Then Exit For
                ImpteDescuento = CDec(Numerico(.get_TextMatrix(I, C_COLDESCUENTOSINIVA)))
                'ImptePromocion = Numerico(.TextMatrix(i, C_ColPROMOCION))
                Cantidad = CShort(Numerico(.get_TextMatrix(I, C_ColEXISTENCIA)))
                ObtenerTotalPromDescuento = ObtenerTotalPromDescuento + ((ImpteDescuento + ImptePromocion) * Cantidad)
            Next
        End With
        ObtenerTotalPromDescuento = ObtenerTotalPromDescuento
    End Function

    Function ObtenerTotalIvaReal() As Decimal
        Dim I As Integer
        'Esta Función obtiene el Total de la Columna de Iva Real, del total del detalle de la venta.
        'Multiplicado por las cantidades de artículos, y saber el total de iva que se está cobrando en la Venta
        ObtenerTotalIvaReal = 0
        With flexDetalle
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, C_COLCODARTICULO) = "" Then Exit For
                ObtenerTotalIvaReal = ObtenerTotalIvaReal + (CDbl(Numerico(.get_TextMatrix(I, C_COLIVAREALCON4DECIMALES))) * CDbl(Numerico(.get_TextMatrix(I, C_ColEXISTENCIA))))
            Next
        End With
        ObtenerTotalIvaReal = ObtenerTotalIvaReal
    End Function

    Function ValidaDatos() As Boolean
        Dim I As Integer

        ValidaDatos = False
        If CDbl(Numerico(txtCodVendExterno.Text)) = 0 And Trim(dbcSucursal.Text) = "" Then
            MsgBox("Proporcione el código o el nombre del vendedor externo...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodVendExterno.Focus()
            Exit Function
        End If
        If Trim(txtNombre.Text) = "" Then
            MsgBox("Falta el nombre del cliente...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Function
        End If
        If Trim(TxtTelefono.Text) = "" Then
            MsgBox("Falta el Telefono del cliente...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Function
        End If
        If Trim(dbcVendedor.Text) = "" Then
            MsgBox("Proporcione el nombre del vendedor...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcVendedor.Focus()
            Exit Function
        End If
        If CDbl(Numerico(txtTotal.Text)) = 0 Then
            MsgBox("No hay detalle de venta...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodVendExterno.Focus()
            Exit Function
        End If
        If optCredito.Checked Then
            If CDbl(Numerico(txtSaldo.Text)) < 0 Then
                MsgBox("El saldo no puede ser negativo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtAnticipo.Focus()
                Exit Function
            End If
        End If
        EmitePago = True
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODARTICULO)) = "" Then Exit For
                If CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR))) = 0 Then
                    EmitePago = False
                    Exit Function
                End If
                If Trim(.get_TextMatrix(I, C_COLORIGEN)) = "" Then
                    EmitePago = False
                    Exit Function
                End If
            Next I
        End With
        ValidaDatos = True
    End Function

    Private Sub cmdABCClientes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdABCClientes.Click
        frmCorpoABCClientes.Show()
        frmCorpoABCClientes.BringToFront()
    End Sub

    Private Sub cmdABCClientes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdABCClientes.Enter
        Pon_Tool()
    End Sub

    Private Sub cmdProcesarPago_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdProcesarPago.Click
        If Not mblnNuevo Then
            Exit Sub
        End If

        If mblnNuevo And ValidaDatos() And EmitePago Then
            If CDbl(Numerico(txtAnticipo.Text)) = 0 And optCredito.Checked = True Then
                Guardar()
                Exit Sub
            End If
            'If frmPagosSalMercancia.ExistenFP = True Then
            '    Me.Enabled = False
            '    frmPagosSalMercancia.Show()
            '    PonerTotalesVentaenFrmPagos()
            'Else
            '    MsgBox("No existen formas de pago disponibles, Favor de verificar....", MsgBoxStyle.OkOnly + MsgBoxStyle.InFormation, gstrNombCortoEmpresa)
            'End If
        End If

        If Not EmitePago Then
            MsgBox("Existen articulo(s) con precio (0.00)" & vbNewLine & "o no tienen código de origen, esta liquidación no procede." & vbNewLine & "Favor de asignar un precio valido o código de origen a esto(s) articulo(s) en el catalogo de articulos.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            flexDetalle.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        If Trim(dbcSucursal.Text) = "" Then
            txtCodVendExterno.Text = ""
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'V' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursal)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            txtCodVendExterno.Focus()
        End If
    End Sub

    Private Sub dbcSucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcSucursal.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcSucursal_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        FueraChange = True
        If dbcSucursal.SelectedItem <> 0 Then
            gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
            DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
            If intCodSucursal <> 0 Then
                txtCodVendExterno.Text = CStr(intCodSucursal)
            Else
                txtCodVendExterno.Text = ""
            End If
        End If
        dbcSucursal.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        FueraChange = True
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
        If intCodSucursal <> 0 Then
            txtCodVendExterno.Text = CStr(intCodSucursal)
        Else
            txtCodVendExterno.Text = ""
        End If
        If CDbl(Numerico(txtCodVendExterno.Text)) <> 0 Then
            'If Not BuscaMovimientos Then
            '    FueraChange = False
            '    txtCodVendExterno = ""
            '    dbcSucursal.text = ""
            '    FueraChange = True
            '    ModEstandar.RetrocederTab Me
            'End If
        End If
        FueraChange = False
        BuscaExistencias()
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        FueraChange = True
        If dbcSucursal.SelectedItem <> 0 Then
            gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
            DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
            txtCodVendExterno.Text = CStr(intCodSucursal)
        End If
        dbcSucursal.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcvendedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVendedor.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcVendedor.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodVendedor,rtrim(ltrim(DescVendedor)) as DescVendedor FROM CatVendedores WHERE DescVendedor LIKE '" & Trim(dbcVendedor.Text) & "%' ORDER BY DescVendedor"
        DCChange(gStrSql, tecla)
        intCodVendedor = 0
    End Sub

    Private Sub dbcvendedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVendedor.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcVendedor.Name Then Exit Sub
        gStrSql = "SELECT CodVendedor,rtrim(ltrim(DescVendedor)) as DescVendedor FROM CatVendedores ORDER BY DescVendedor"
        DCGotFocus(gStrSql, dbcVendedor)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcvendedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcVendedor.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            txtTipoCambio.Focus()
        End If
    End Sub

    Private Sub dbcVendedor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcVendedor.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcVendedor_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcVendedor.KeyUp
        Dim Aux As String
        Aux = dbcVendedor.Text
        If dbcVendedor.SelectedItem <> 0 Then
            dbcvendedor_Leave(dbcVendedor, New System.EventArgs())
        End If
        FueraChange = True
        dbcVendedor.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcvendedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVendedor.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        FueraChange = True
        gStrSql = "SELECT CodVendedor,rtrim(ltrim(DescVendedor)) as DescVendedor FROM CatVendedores WHERE DescVendedor LIKE '" & Trim(dbcVendedor.Text) & "%' ORDER BY DescVendedor"
        DCLostFocus(dbcVendedor, gStrSql, intCodVendedor)
        FueraChange = False
    End Sub

    Private Sub dbcVendedor_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcVendedor.MouseUp
        'Dim Aux As String
        'Aux = dbcVendedor.Text
        'If dbcVendedor.SelectedItem <> 0 Then
        '    dbcvendedor_Leave(dbcVendedor, New System.EventArgs())
        'End If
        'FueraChange = True
        'dbcVendedor.Text = Aux
        'FueraChange = False
    End Sub

    Private Sub flexDetalle_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.ClickEvent
        txtFlex.Visible = False
    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.DblClick
        FlexDetalle_KeyPressEvent(flexDetalle, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub FlexDetalle_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.EnterCell
        lblArticulo.Text = Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, C_COLDESCARTICULO))
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        Pon_Tool()
        lblArticulo.Text = Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, C_COLDESCARTICULO))
    End Sub

    Private Sub FlexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexDetalle.KeyDownEvent
        Dim Ren As Integer
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete And mblnNuevo And Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, C_COLCODARTICULO)) <> "" Then
            Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Ren = flexDetalle.Rows
                    flexDetalle.RemoveItem(flexDetalle.Row)
                    flexDetalle.Rows = Ren
                    CalculaImporte()
                    If (CDec(Numerico(txtTotal.Text)) + CDec(Numerico(txtRedondeo.Text))) = 0 Then
                        Limpiar()
                        Exit Sub
                    End If
            End Select
            flexDetalle.Focus()
        End If
    End Sub

    Private Sub FlexDetalle_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexDetalle.KeyPressEvent
        If eventArgs.keyAscii = System.Windows.Forms.Keys.Return And mblnNuevo Then
            'Edita el campo sólo si es Editable
            If flexDetalle.Col = 5 And Trim(flexDetalle.Text) <> "" Then
                MSHFlexGridEdit(flexDetalle, txtFlex, eventArgs.keyAscii)
                txtFlex.Text = flexDetalle.get_TextMatrix(flexDetalle.Row, C_COLPORCENTAJEDEDESCUENTO) & " %"
            End If
        End If
    End Sub

    Private Sub FlexDetalle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Leave
        lblArticulo.Text = ""
    End Sub

    Private Sub FlexDetalle_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Scroll
        txtFlex.Visible = False
    End Sub

    Private Sub frmVtasVELiquidacionVendedorExterno_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasVELiquidacionVendedorExterno_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasVELiquidacionVendedorExterno_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If System.Windows.Forms.Form.ActiveForm.Name <> "frmVtasVELiquidacionVendedorExterno" Then
            Exit Sub
        End If
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "txtFolio" And Trim(txtFolio.Text) <> "" And (txtFolio.Text) <> "0000" Then
                    txtCodVendExterno.Enabled = False
                    dbcSucursal.Enabled = False
                End If
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtFolio" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
                '        Case vbKeyF8
                '            If Not mblnNuevo Then
                '                Exit Sub
                '            End If
                '            If Not EmitePago Then
                '                MsgBox "Existen articulo(s) con precio (0.00), esta liquidación no procede." & vbNewLine & _
                ''                "Favor de asignar un precio valido a esto(s) articulo(s) en el catalogo de articulos.", vbOKOnly + vbInVB6.Formation, gstrNombCortoEmpresa
                '                Exit Sub
                '            End If
                '            If mblnNuevo And ValidaDatos And EmitePago Then
                '                If Numerico(txtAnticipo) = 0 And optCredito.Value = True Then
                '
                '                    Guardar
                '                    Exit Sub
                '                End If
                '                If frmPagosSalMercancia.ExistenFP = True Then
                '                    Me.Enabled = False
                '                    frmPagosSalMercancia.Show
                '                    PonerTotalesVentaenFrmPagos
                '                Else
                '                    MsgBox "No existen formas de pago disponibles, Favor de verificar....", vbOKOnly + vbInVB6.Formation, gstrNombCortoEmpresa
                '                End If
                '            End If
        End Select
    End Sub

    Private Sub frmVtasVELiquidacionVendedorExterno_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasVELiquidacionVendedorExterno_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        bandera = True
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) - 350)
        InicializaVariables()
        Nuevo()
        ModCorporativo.ObtenerDatosSucursalMatriz(txtCodSucMatriz, txtDescSucMatriz)
        ObtenerCaja()
    End Sub

    Private Sub frmVtasVELiquidacionVendedorExterno_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
        '    'If Cambios = True And mblnNuevo = False Then
        '    'Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
        '    'Case vbYes: 'Guardar el registro
        '    'If Guardar = False Then
        '    'Cancel = 1
        '    'End If
        '    'Case vbNo: 'No hace nada y permite el cierre del formulario
        '    'Case vbCancel: 'Cancela el cierre del formulario sin guardar
        '    'Cancel = 1
        '    'End Select
        '    'End If
        'Else
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0
        '        Case MsgBoxResult.No
        '            mblnSalir = False
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasVELiquidacionVendedorExterno_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optContado_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optContado.CheckedChanged
        If eventSender.Checked Then
            optPesosAnt.Enabled = False
            optDolaresAnt.Enabled = False
            txtAnticipo.Enabled = False
            txtSaldo.Enabled = False
            txtAnticipo.Text = "0.00"
            txtSaldo.Text = "0.00"
        End If
    End Sub

    Private Sub optContado_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optContado.Enter
        Pon_Tool()
    End Sub

    Private Sub optCredito_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCredito.CheckedChanged
        If eventSender.Checked Then
            optPesosAnt.Enabled = True
            optDolaresAnt.Enabled = True
            txtAnticipo.Enabled = True
            txtSaldo.Enabled = True
            If optPesosAnt.Checked = True Then
                txtSaldo.Text = VB6.Format(CDbl(txtTotalPesos.Text) - CDbl(Numerico(txtAnticipo.Text)), "###,##0.00")
            ElseIf optDolaresAnt.Checked = True Then
                txtSaldo.Text = VB6.Format(CDbl(txtTotalDolares.Text) - CDbl(Numerico(txtAnticipo.Text)), "###,##0.00")
            End If
        End If
    End Sub

    Private Sub optCredito_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCredito.Enter
        Pon_Tool()
    End Sub

    Private Sub optDolares_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDolares.CheckedChanged
        If eventSender.Checked Then
            CalculaImporte()
        End If
    End Sub

    Private Sub optDolaresAnt_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDolaresAnt.CheckedChanged
        If eventSender.Checked Then
            txtAnticipo.Text = "0.00"
            txtSaldo.Text = VB6.Format(CDbl(txtTotalDolares.Text) - CDbl(Numerico(txtAnticipo.Text)), "###,##0.00")
        End If
    End Sub

    Private Sub optPesos_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPesos.CheckedChanged
        If eventSender.Checked Then
            CalculaImporte()
        End If
    End Sub
    Private Sub optPesosAnt_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPesosAnt.CheckedChanged
        If eventSender.Checked Then
            txtAnticipo.Text = "0.00"
            'txtSaldo.Text = (Convert.ToDouble(txtTotalPesos.Text)) - (Convert.ToDouble(Numerico(String.Concat(txtAnticipo.Text, "###,##0.00"))))
        End If
    End Sub

    Private Sub txtAnticipo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnticipo.TextChanged
        If optPesosAnt.Checked = True Then
            txtSaldo.Text = VB6.Format(VB6.Format(CDbl(txtTotalPesos.Text) - CDbl(Numerico(txtAnticipo.Text))), "###,##0.00")
        ElseIf optDolaresAnt.Checked = True Then
            txtSaldo.Text = VB6.Format(VB6.Format(CDbl(txtTotalDolares.Text) - CDbl(Numerico(txtAnticipo.Text))), "###,##0.00")
        End If
    End Sub

    Private Sub txtAnticipo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnticipo.Enter
        SelTextoTxt(txtAnticipo)
        Pon_Tool()
    End Sub

    Private Sub txtAnticipo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAnticipo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtAnticipo.Text, KeyAscii, 12, 2, (txtAnticipo.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtAnticipo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnticipo.Leave
        If CDbl(Numerico(txtAnticipo.Text)) = 0 Then
            txtAnticipo.Text = VB6.Format(0, "###,##0.00")
        Else
            txtAnticipo.Text = VB6.Format(txtAnticipo.Text, "###,##0.00")
        End If
        If optPesosAnt.Checked = True Then
            txtSaldo.Text = VB6.Format(VB6.Format(CDbl(txtTotalPesos.Text) - CDbl(Numerico(txtAnticipo.Text))), "###,##0.00")
        ElseIf optDolaresAnt.Checked = True Then
            txtSaldo.Text = VB6.Format(VB6.Format(CDbl(txtTotalDolares.Text) - CDbl(Numerico(txtAnticipo.Text))), "###,##0.00")
        End If
    End Sub

    Private Sub txtCodVendExterno_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodVendExterno.TextChanged
        If FueraChange Then Exit Sub
        If CDbl(Numerico(txtCodVendExterno.Text)) = 0 Then
            dbcSucursal.Text = ""
            Nuevo()
        End If
    End Sub

    Private Sub txtCodVendExterno_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodVendExterno.Enter
        Pon_Tool()
        SelTextoTxt(txtCodVendExterno)
    End Sub

    Private Sub txtCodVendExterno_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodVendExterno.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodVendExterno_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodVendExterno.Leave
        '    If Numerico(txtCodVendExterno) = 0 Then
        '        txtCodVendExterno = ""
        '    Else
        '        BuscaVendedorExterno
        '    End If
    End Sub

    Private Sub txtDescuento_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescuento.Enter
        Pon_Tool()
    End Sub

    Private Sub txtDomicilio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.Enter
        Pon_Tool()
    End Sub

    Private Sub txtFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Enter
        SelTextoTxt(txtFlex)
        Pon_Tool()
        lblArticulo.Text = "PORCENTAJE DE DESCUENTO AL VENDEDOR EXTERNO"
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            flexDetalle.Focus()
            txtFlex.Visible = False
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
            flexDetalle.Focus()
            txtFlex.Visible = False
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Private Sub txtFolio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambios = True
    End Sub

    Private Sub txtFolio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.Enter
        strControlActual = UCase("txtFolio")
        Pon_Tool()
        SelTextoTxt(txtFolio)
    End Sub

    Private Sub txtFolio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.Leave
        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If Trim(txtFolio.Text) = "" Then
            txtFolio.Text = "S" & VB6.Format(txtCodSucMatriz.Text, "00") & VB6.Format(intCodCaja, "00") & Year(dtpFecha.Value) & VB6.Format(Month(dtpFecha.Value), "00") & VB6.Format((dtpFecha.Value), "00") & "0000"
        End If
        If mblnCambios = True And txtFolio.Text <> "" And (txtFolio.Text) <> "0000" Then
            LlenaDatos()
        End If
    End Sub

    Private Sub txtFolioEntrega_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEntrega.TextChanged
        If FueraChange Then Exit Sub
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
    End Sub

    Private Sub txtFolioEntrega_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEntrega.Enter
        strControlActual = UCase("txtFolioEntrega")
        Pon_Tool()
        SelTextoTxt(txtFolioEntrega)
    End Sub

    Private Sub txtFolioEntrega_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioEntrega.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, C_PrefijoFoliosAlmacen)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioEntrega_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEntrega.Leave
        If Trim(txtFolioEntrega.Text) <> "" Then
            BuscaExistencias()
        End If
    End Sub

    Private Sub txtIVA_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtIVA.Enter
        Pon_Tool()
    End Sub

    Private Sub txtNombre_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.Enter
        Pon_Tool()
    End Sub

    Private Sub txtRedondeo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRedondeo.Enter
        Pon_Tool()
    End Sub

    Private Sub txtRFC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRFC.Enter
        Pon_Tool()
    End Sub

    Private Sub txtSaldo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSaldo.Enter
        SelTextoTxt(txtSaldo)
        Pon_Tool()
    End Sub

    Private Sub txtSaldo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSaldo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtSaldo.Text, KeyAscii, 12, 2, (txtSaldo.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtSaldo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSaldo.Leave
        If CDbl(Numerico(txtSaldo.Text)) = 0 Then
            txtSaldo.Text = VB6.Format(0, "###,##0.00")
        Else
            txtSaldo.Text = VB6.Format(txtSaldo.Text, "###,##0.00")
        End If
    End Sub

    Private Sub txtSubTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSubtotal.Enter
        Pon_Tool()
    End Sub

    Private Sub txtTipoCambio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Enter
        Pon_Tool()
    End Sub

    Private Sub txtTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTotal.Enter
        Pon_Tool()
    End Sub

    Private Sub txtTotalDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTotalDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub txtTotalPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTotalPesos.Enter
        Pon_Tool()
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
End Class