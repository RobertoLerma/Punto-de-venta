Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility


Public Class frmFactAnalisisVentas
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents DtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents lblFacturar As System.Windows.Forms.Label
    Public WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents lblExcluido As System.Windows.Forms.Label
    Public WithEvents Label18 As System.Windows.Forms.Label
    Public WithEvents Frame7 As System.Windows.Forms.GroupBox
    Public WithEvents cmdDatosFiscales As System.Windows.Forms.Button
    Public WithEvents chkDoctoCliente As System.Windows.Forms.CheckBox
    Public WithEvents chkDesglosarIva As System.Windows.Forms.CheckBox
    Public WithEvents cmdGenerarFactura As System.Windows.Forms.Button
    Public WithEvents cmdImpresionTickets As System.Windows.Forms.Button
    Public WithEvents cmdImprimirFactura As System.Windows.Forms.Button
    Public WithEvents Frame6 As System.Windows.Forms.GroupBox
    Public WithEvents txtFacturaAdicional As System.Windows.Forms.TextBox
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents lblSubTotal As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents lblRedondeo As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents lblTotal As System.Windows.Forms.Label
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents lblTotalPesos As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents txtDescripcion As System.Windows.Forms.TextBox
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents lblImporteRedondeo As System.Windows.Forms.Label
    Public WithEvents lblImporteTotal As System.Windows.Forms.Label
    Public WithEvents Label15 As System.Windows.Forms.Label
    Public WithEvents lblImporteSubTotal As System.Windows.Forms.Label
    Public WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents lblFactura As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents flexVentasPendientes As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents flexDetalleVenta As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents flexVentas As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents txtFolioFactura As System.Windows.Forms.TextBox
    Public WithEvents dtpFechaRegistro As System.Windows.Forms.DateTimePicker
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents txtPorcentaje As System.Windows.Forms.TextBox
    Public WithEvents optPorcentual As System.Windows.Forms.RadioButton
    Public WithEvents optManual As System.Windows.Forms.RadioButton
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents dtpFechaVenta As System.Windows.Forms.DateTimePicker
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents DtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents lblMoneda As System.Windows.Forms.Label
    Public WithEvents Label21 As System.Windows.Forms.Label
    Public WithEvents Label20 As System.Windows.Forms.Label
    Public WithEvents Label17 As System.Windows.Forms.Label
    Public WithEvents lblSubTot As System.Windows.Forms.Label
    Public WithEvents lblIva As System.Windows.Forms.Label
    Public WithEvents lblDescuento As System.Windows.Forms.Label
    Public WithEvents lblCantidad As System.Windows.Forms.Label
    Public WithEvents Label19 As System.Windows.Forms.Label
    Public WithEvents Label16 As System.Windows.Forms.Label
    Public WithEvents lblDesc As System.Windows.Forms.Label
    Public WithEvents lblDescripcion As System.Windows.Forms.Label
    Public WithEvents lblEstadoFolio As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label


    Dim mblnSalir As Boolean
    Dim mblnNuevo As Boolean
    Dim FueraChange As Boolean
    Dim intCodSucursal As Integer
    Dim tecla As Integer
    Dim Fecha As String
    Dim CodSucursal As Integer
    Dim CodCliente As Integer
    Dim DescEspecial As String
    Dim DescSucursal As String
    Dim ColorAnte As String
    Dim ColorAnteFolio As String
    Dim FolioAdicional As String
    Dim RenAnterior As Integer
    Dim RedondeoFactura As Decimal
    Dim TipoCambio As Decimal
    Dim mblnCambiarMetodo As Boolean
    Dim numPartidas As Integer
    Dim DesgloseIva As Byte
    Dim Caja As Integer
    Public gblnCambiosAnalisis As Boolean
    Public mblnFactura As Boolean
    Public GenerarFactPtoVenta As Boolean

    Dim blnFueraCell As Boolean
    Dim RenAct As Integer

    'Constantes para el Grid de Ventas
    Const C_COLFOLIOVENTA As Integer = 0
    Const C_ColIMPORTE As Integer = 1
    Const C_ColMONEDA As Integer = 2
    Const C_COLESTADOFOLIO As Integer = 3
    Const C_COLPORCIVA As Integer = 4
    Const C_ColTIPOCAMBIO As Integer = 5
    Const C_COLANTICIPO As Integer = 6
    Const C_COLFOLIOADICIONAL As Integer = 7
    Const C_COLMETODO As Integer = 8
    Const C_COLSUBTOTALADICIONAL As Integer = 9
    Const C_COLDESCUENTOADICIONAL As Integer = 10
    Const C_COLIVAADICIONAL As Integer = 11
    Const C_COLTOTALADICIONAL As Integer = 12
    Const C_COLREDONDEOADICIONAL As Integer = 13
    Const C_COLANTICIPOADICIONAL As Integer = 14
    Const C_COLESTATUSADICIONAL As Integer = 15
    Const C_COLFOLIOFACTURA As Integer = 16
    Const C_COLGRABADO As Integer = 17
    Const C_COLFOLIOEXCLUIDO As Integer = 18
    Const C_COLINCLUIRFACTURA As Integer = 19
    Const C_ColCONDICION As Integer = 20
    Const C_COLFECHADEVENTA As Integer = 21
    Const C_ColCODSUCURSAL As Integer = 22
    Const C_COLCODCAJA As Integer = 23
    Const C_COLCODVENDEDOR As Integer = 24
    Const C_COLCODCLIENTE As Integer = 25
    Const C_COLFACTURAPTOVTA As Integer = 26
    Const C_COLINCFACTURA As Integer = 27
    Const C_COLTIENEDEVOLUCION As Integer = 28
    Const C_COLCAMBIOS As Integer = 29
    Const C_COLTIPOMOVTO As Integer = 30

    'Constantes para el Grid del Detalle de la Venta
    Const C_COLCODIGOARTICULO As Integer = 0
    Const C_COLCANTIDADDEVOL As Integer = 1
    Const C_COLCANTIDAD As Integer = 2
    Const C_COLPORCDESCUENTO As Integer = 3
    Const C_COLPORCPROMOCION As Integer = 4
    Const C_ColPRECIOPUBLICO As Integer = 5
    Const C_COLIMPORTESINDESCUENTO As Integer = 6
    Const C_COLIMPORTECONDESCUENTO As Integer = 7
    Const C_COLNUEVADESCRIPCION As Integer = 8
    Const C_COLNUEVOPRECIOPUBLICO As Integer = 9
    Const C_COLNUEVOIMPORTESINDESCUENTO As Integer = 10
    Const C_COLNUEVOIMPORTECONDESCUENTO As Integer = 11
    Const C_COLDESCRIPCION As Integer = 12
    Const C_ColPRECIOLISTASINIVA As Integer = 13
    Const C_COLPRECIOREAL As Integer = 14
    Const C_COLIVAREAL As Integer = 15
    Const C_COLMODIFICADO As Integer = 16
    Const C_COLPORCDESCTO As Integer = 17
    Const C_COLPORCPROM As Integer = 18
    Const C_ColDESCUENTO As Integer = 19
    Const C_ColPROMOCION As Integer = 20
    Const C_COLFOLIOAGREGADO As Integer = 21
    Const C_COLFECHAFOLIOAGREGADO As Integer = 22
    Const C_COLEXCLUIDO As Integer = 23
    Const C_COLDESCRIPCIONFAMILIA As Integer = 24
    Const C_COLNUMPARTIDA As Integer = 25
    Const C_COLALMACEN As Integer = 26
    Const C_COLPRECIOLISTAADICIONAL As Integer = 27
    Const C_COLPORCENTAJEADICIONAL As Integer = 28
    Const C_COLGRAB As Integer = 29
    Const C_COLMODIFICADOTAG As Integer = 30

    'Constantes para el Grid de Detalles Pendientes
    Const C_ColFOLIO As Integer = 0
    Const C_COLFECHAVENTA As Integer = 1
    Const C_COLCODARTICULO As Integer = 2
    Const C_COLDESCARTICULO As Integer = 3
    Const C_COLDEVOL As Integer = 4
    Const C_COLCANT As Integer = 5
    Const C_COLDESCTO As Integer = 6
    Const C_COLPROM As Integer = 7
    Const C_COLPRECIOPUB As Integer = 8
    Const C_COLIMPTE As Integer = 9
    Const C_COLIMPTEDESCTO As Integer = 10
    Const C_COLDESCFAMILIA As Integer = 11
    Const C_COLPORDESCTO As Integer = 12
    Const C_COLPORPROM As Integer = 13
    Const C_COLNPARTIDA As Integer = 14
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnBuscar As Button
    Public strControlActual As String 'Nombre del control actual

    '''Const TipoMovto = 1
    Const TipoMovto As Integer = 5 '''No. de folio en Catfolios -->  Facturación



    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFactAnalisisVentas))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtFolioFactura = New System.Windows.Forms.TextBox()
        Me.txtPorcentaje = New System.Windows.Forms.TextBox()
        Me.optPorcentual = New System.Windows.Forms.RadioButton()
        Me.optManual = New System.Windows.Forms.RadioButton()
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.DtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.Frame7 = New System.Windows.Forms.GroupBox()
        Me.lblFacturar = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.lblExcluido = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Frame6 = New System.Windows.Forms.GroupBox()
        Me.cmdDatosFiscales = New System.Windows.Forms.Button()
        Me.chkDoctoCliente = New System.Windows.Forms.CheckBox()
        Me.chkDesglosarIva = New System.Windows.Forms.CheckBox()
        Me.cmdGenerarFactura = New System.Windows.Forms.Button()
        Me.cmdImpresionTickets = New System.Windows.Forms.Button()
        Me.cmdImprimirFactura = New System.Windows.Forms.Button()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.txtFacturaAdicional = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblSubTotal = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblRedondeo = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.lblTotalPesos = New System.Windows.Forms.Label()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.lblImporteRedondeo = New System.Windows.Forms.Label()
        Me.lblImporteTotal = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.lblImporteSubTotal = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.lblFactura = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.flexVentasPendientes = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.flexDetalleVenta = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.flexVentas = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaRegistro = New System.Windows.Forms.DateTimePicker()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpFechaVenta = New System.Windows.Forms.DateTimePicker()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.DtpHasta = New System.Windows.Forms.DateTimePicker()
        Me.lblMoneda = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.lblSubTot = New System.Windows.Forms.Label()
        Me.lblIva = New System.Windows.Forms.Label()
        Me.lblDescuento = New System.Windows.Forms.Label()
        Me.lblCantidad = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.lblDescripcion = New System.Windows.Forms.Label()
        Me.lblEstadoFolio = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Frame7.SuspendLayout()
        Me.Frame6.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        CType(Me.flexVentasPendientes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.flexDetalleVenta, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.flexVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtFolioFactura
        '
        Me.txtFolioFactura.AcceptsReturn = True
        Me.txtFolioFactura.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioFactura.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioFactura.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioFactura.Location = New System.Drawing.Point(83, 13)
        Me.txtFolioFactura.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioFactura.MaxLength = 17
        Me.txtFolioFactura.Name = "txtFolioFactura"
        Me.txtFolioFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioFactura.Size = New System.Drawing.Size(153, 20)
        Me.txtFolioFactura.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtFolioFactura, "Folio de la Factura.")
        '
        'txtPorcentaje
        '
        Me.txtPorcentaje.AcceptsReturn = True
        Me.txtPorcentaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtPorcentaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPorcentaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPorcentaje.Location = New System.Drawing.Point(84, 32)
        Me.txtPorcentaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPorcentaje.MaxLength = 3
        Me.txtPorcentaje.Name = "txtPorcentaje"
        Me.txtPorcentaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPorcentaje.Size = New System.Drawing.Size(36, 20)
        Me.txtPorcentaje.TabIndex = 5
        Me.txtPorcentaje.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtPorcentaje, "Porcentaje a Aplicar.")
        '
        'optPorcentual
        '
        Me.optPorcentual.BackColor = System.Drawing.SystemColors.Control
        Me.optPorcentual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPorcentual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPorcentual.Location = New System.Drawing.Point(12, 34)
        Me.optPorcentual.Margin = New System.Windows.Forms.Padding(2)
        Me.optPorcentual.Name = "optPorcentual"
        Me.optPorcentual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPorcentual.Size = New System.Drawing.Size(80, 17)
        Me.optPorcentual.TabIndex = 4
        Me.optPorcentual.TabStop = True
        Me.optPorcentual.Text = "Porcentual"
        Me.ToolTip1.SetToolTip(Me.optPorcentual, "Mediante este método, se Proporciona un porcentaje y se hacen los calculos automa" &
        "ticamente.")
        Me.optPorcentual.UseVisualStyleBackColor = False
        '
        'optManual
        '
        Me.optManual.BackColor = System.Drawing.SystemColors.Control
        Me.optManual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optManual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optManual.Location = New System.Drawing.Point(12, 15)
        Me.optManual.Margin = New System.Windows.Forms.Padding(2)
        Me.optManual.Name = "optManual"
        Me.optManual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optManual.Size = New System.Drawing.Size(61, 15)
        Me.optManual.TabIndex = 3
        Me.optManual.TabStop = True
        Me.optManual.Text = "Manual"
        Me.ToolTip1.SetToolTip(Me.optManual, "Mediante este método, los calculos se hacen manualmente.")
        Me.optManual.UseVisualStyleBackColor = False
        '
        'txtCodSucursal
        '
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.Enabled = False
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.Location = New System.Drawing.Point(214, 59)
        Me.txtCodSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucursal.MaxLength = 3
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(57, 20)
        Me.txtCodSucursal.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCodSucursal, "Codigo de la Sucursal.")
        Me.txtCodSucursal.Visible = False
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackColor = System.Drawing.SystemColors.Control
        Me.btnLimpiar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLimpiar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLimpiar.Location = New System.Drawing.Point(117, 652)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLimpiar.Size = New System.Drawing.Size(97, 43)
        Me.btnLimpiar.TabIndex = 91
        Me.btnLimpiar.Text = "Nuevo"
        Me.ToolTip1.SetToolTip(Me.btnLimpiar, "Registro de Clientes")
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.BackColor = System.Drawing.SystemColors.Control
        Me.btnBuscar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBuscar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBuscar.Location = New System.Drawing.Point(14, 652)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBuscar.Size = New System.Drawing.Size(97, 43)
        Me.btnBuscar.TabIndex = 89
        Me.btnBuscar.Text = "Buscar"
        Me.ToolTip1.SetToolTip(Me.btnBuscar, "Registro de Clientes")
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'DtpDesde
        '
        Me.DtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DtpDesde.Location = New System.Drawing.Point(26, 359)
        Me.DtpDesde.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpDesde.Name = "DtpDesde"
        Me.DtpDesde.Size = New System.Drawing.Size(95, 20)
        Me.DtpDesde.TabIndex = 20
        '
        'Frame7
        '
        Me.Frame7.BackColor = System.Drawing.SystemColors.Control
        Me.Frame7.Controls.Add(Me.lblFacturar)
        Me.Frame7.Controls.Add(Me.Label14)
        Me.Frame7.Controls.Add(Me.lblExcluido)
        Me.Frame7.Controls.Add(Me.Label18)
        Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame7.Location = New System.Drawing.Point(14, 81)
        Me.Frame7.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame7.Name = "Frame7"
        Me.Frame7.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame7.Size = New System.Drawing.Size(168, 54)
        Me.Frame7.TabIndex = 59
        Me.Frame7.TabStop = False
        '
        'lblFacturar
        '
        Me.lblFacturar.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblFacturar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFacturar.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFacturar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblFacturar.Location = New System.Drawing.Point(8, 30)
        Me.lblFacturar.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFacturar.Name = "lblFacturar"
        Me.lblFacturar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFacturar.Size = New System.Drawing.Size(13, 14)
        Me.lblFacturar.TabIndex = 63
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Location = New System.Drawing.Point(27, 26)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(134, 24)
        Me.Label14.TabIndex = 62
        Me.Label14.Text = "Folios Seleccionados Para Facturación del Pto. de Venta"
        '
        'lblExcluido
        '
        Me.lblExcluido.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblExcluido.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblExcluido.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblExcluido.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblExcluido.Location = New System.Drawing.Point(8, 10)
        Me.lblExcluido.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblExcluido.Name = "lblExcluido"
        Me.lblExcluido.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblExcluido.Size = New System.Drawing.Size(13, 14)
        Me.lblExcluido.TabIndex = 61
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label18.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Location = New System.Drawing.Point(27, 12)
        Me.Label18.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label18.Size = New System.Drawing.Size(134, 14)
        Me.Label18.TabIndex = 60
        Me.Label18.Text = "Partidas Excluidas"
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.cmdDatosFiscales)
        Me.Frame6.Controls.Add(Me.chkDoctoCliente)
        Me.Frame6.Controls.Add(Me.chkDesglosarIva)
        Me.Frame6.Controls.Add(Me.cmdGenerarFactura)
        Me.Frame6.Controls.Add(Me.cmdImpresionTickets)
        Me.Frame6.Controls.Add(Me.cmdImprimirFactura)
        Me.Frame6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame6.Location = New System.Drawing.Point(632, 493)
        Me.Frame6.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(259, 146)
        Me.Frame6.TabIndex = 52
        Me.Frame6.TabStop = False
        Me.Frame6.Text = "Documento..."
        '
        'cmdDatosFiscales
        '
        Me.cmdDatosFiscales.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDatosFiscales.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDatosFiscales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDatosFiscales.Location = New System.Drawing.Point(142, 55)
        Me.cmdDatosFiscales.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdDatosFiscales.Name = "cmdDatosFiscales"
        Me.cmdDatosFiscales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDatosFiscales.Size = New System.Drawing.Size(99, 39)
        Me.cmdDatosFiscales.TabIndex = 57
        Me.cmdDatosFiscales.Text = "Datos Fiscales"
        Me.cmdDatosFiscales.UseVisualStyleBackColor = False
        '
        'chkDoctoCliente
        '
        Me.chkDoctoCliente.BackColor = System.Drawing.SystemColors.Control
        Me.chkDoctoCliente.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDoctoCliente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDoctoCliente.Location = New System.Drawing.Point(142, 31)
        Me.chkDoctoCliente.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDoctoCliente.Name = "chkDoctoCliente"
        Me.chkDoctoCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDoctoCliente.Size = New System.Drawing.Size(98, 17)
        Me.chkDoctoCliente.TabIndex = 56
        Me.chkDoctoCliente.Text = "Docto. Cliente"
        Me.chkDoctoCliente.UseVisualStyleBackColor = False
        '
        'chkDesglosarIva
        '
        Me.chkDesglosarIva.BackColor = System.Drawing.SystemColors.Control
        Me.chkDesglosarIva.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDesglosarIva.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDesglosarIva.Location = New System.Drawing.Point(146, 97)
        Me.chkDesglosarIva.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDesglosarIva.Name = "chkDesglosarIva"
        Me.chkDesglosarIva.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDesglosarIva.Size = New System.Drawing.Size(94, 20)
        Me.chkDesglosarIva.TabIndex = 58
        Me.chkDesglosarIva.Text = "Desglosar Iva"
        Me.chkDesglosarIva.UseVisualStyleBackColor = False
        '
        'cmdGenerarFactura
        '
        Me.cmdGenerarFactura.BackColor = System.Drawing.SystemColors.Control
        Me.cmdGenerarFactura.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdGenerarFactura.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdGenerarFactura.Location = New System.Drawing.Point(9, 16)
        Me.cmdGenerarFactura.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdGenerarFactura.Name = "cmdGenerarFactura"
        Me.cmdGenerarFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdGenerarFactura.Size = New System.Drawing.Size(110, 36)
        Me.cmdGenerarFactura.TabIndex = 53
        Me.cmdGenerarFactura.Text = "Generar Factura"
        Me.cmdGenerarFactura.UseVisualStyleBackColor = False
        '
        'cmdImpresionTickets
        '
        Me.cmdImpresionTickets.BackColor = System.Drawing.SystemColors.Control
        Me.cmdImpresionTickets.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdImpresionTickets.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdImpresionTickets.Location = New System.Drawing.Point(9, 56)
        Me.cmdImpresionTickets.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdImpresionTickets.Name = "cmdImpresionTickets"
        Me.cmdImpresionTickets.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdImpresionTickets.Size = New System.Drawing.Size(110, 37)
        Me.cmdImpresionTickets.TabIndex = 54
        Me.cmdImpresionTickets.Text = "Impresión de Tickets"
        Me.cmdImpresionTickets.UseVisualStyleBackColor = False
        '
        'cmdImprimirFactura
        '
        Me.cmdImprimirFactura.BackColor = System.Drawing.SystemColors.Control
        Me.cmdImprimirFactura.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdImprimirFactura.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdImprimirFactura.Location = New System.Drawing.Point(9, 97)
        Me.cmdImprimirFactura.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdImprimirFactura.Name = "cmdImprimirFactura"
        Me.cmdImprimirFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdImprimirFactura.Size = New System.Drawing.Size(110, 37)
        Me.cmdImprimirFactura.TabIndex = 55
        Me.cmdImprimirFactura.Text = "Imprimir Factura"
        Me.cmdImprimirFactura.UseVisualStyleBackColor = False
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.txtFacturaAdicional)
        Me.Frame5.Controls.Add(Me.Label7)
        Me.Frame5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame5.Location = New System.Drawing.Point(455, 11)
        Me.Frame5.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(182, 58)
        Me.Frame5.TabIndex = 50
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Dato Adicional"
        '
        'txtFacturaAdicional
        '
        Me.txtFacturaAdicional.AcceptsReturn = True
        Me.txtFacturaAdicional.BackColor = System.Drawing.SystemColors.Window
        Me.txtFacturaAdicional.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFacturaAdicional.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFacturaAdicional.Location = New System.Drawing.Point(54, 24)
        Me.txtFacturaAdicional.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFacturaAdicional.MaxLength = 17
        Me.txtFacturaAdicional.Name = "txtFacturaAdicional"
        Me.txtFacturaAdicional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFacturaAdicional.Size = New System.Drawing.Size(111, 20)
        Me.txtFacturaAdicional.TabIndex = 6
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(8, 27)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(51, 15)
        Me.Label7.TabIndex = 51
        Me.Label7.Text = "Factura:"
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(254, 150)
        Me.txtFlex.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(70, 20)
        Me.txtFlex.TabIndex = 43
        Me.txtFlex.Visible = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.Label6)
        Me.Frame4.Controls.Add(Me.lblSubTotal)
        Me.Frame4.Controls.Add(Me.Label8)
        Me.Frame4.Controls.Add(Me.lblRedondeo)
        Me.Frame4.Controls.Add(Me.Label10)
        Me.Frame4.Controls.Add(Me.lblTotal)
        Me.Frame4.Controls.Add(Me.Label12)
        Me.Frame4.Controls.Add(Me.lblTotalPesos)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(674, 373)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(213, 116)
        Me.Frame4.TabIndex = 32
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Totales por Folio de Venta"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(24, 15)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(58, 17)
        Me.Label6.TabIndex = 40
        Me.Label6.Text = "SubTotal :"
        '
        'lblSubTotal
        '
        Me.lblSubTotal.BackColor = System.Drawing.SystemColors.Window
        Me.lblSubTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSubTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSubTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSubTotal.Location = New System.Drawing.Point(89, 15)
        Me.lblSubTotal.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSubTotal.Name = "lblSubTotal"
        Me.lblSubTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSubTotal.Size = New System.Drawing.Size(91, 17)
        Me.lblSubTotal.TabIndex = 39
        Me.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(24, 37)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(63, 17)
        Me.Label8.TabIndex = 38
        Me.Label8.Text = "Redondeo :"
        '
        'lblRedondeo
        '
        Me.lblRedondeo.BackColor = System.Drawing.SystemColors.Window
        Me.lblRedondeo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRedondeo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRedondeo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRedondeo.Location = New System.Drawing.Point(89, 37)
        Me.lblRedondeo.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRedondeo.Name = "lblRedondeo"
        Me.lblRedondeo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRedondeo.Size = New System.Drawing.Size(91, 17)
        Me.lblRedondeo.TabIndex = 37
        Me.lblRedondeo.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(24, 58)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(43, 17)
        Me.Label10.TabIndex = 36
        Me.Label10.Text = "Total :"
        '
        'lblTotal
        '
        Me.lblTotal.BackColor = System.Drawing.SystemColors.Window
        Me.lblTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotal.Location = New System.Drawing.Point(89, 57)
        Me.lblTotal.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotal.Size = New System.Drawing.Size(91, 17)
        Me.lblTotal.TabIndex = 35
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(5, 80)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(82, 17)
        Me.Label12.TabIndex = 34
        Me.Label12.Text = "Total Pesos :"
        '
        'lblTotalPesos
        '
        Me.lblTotalPesos.BackColor = System.Drawing.SystemColors.Window
        Me.lblTotalPesos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalPesos.Location = New System.Drawing.Point(89, 79)
        Me.lblTotalPesos.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTotalPesos.Name = "lblTotalPesos"
        Me.lblTotalPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalPesos.Size = New System.Drawing.Size(91, 17)
        Me.lblTotalPesos.TabIndex = 33
        Me.lblTotalPesos.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtDescripcion
        '
        Me.txtDescripcion.AcceptsReturn = True
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescripcion.Location = New System.Drawing.Point(662, 308)
        Me.txtDescripcion.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescripcion.MaxLength = 50
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(229, 20)
        Me.txtDescripcion.TabIndex = 19
        Me.txtDescripcion.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.Label11)
        Me.Frame3.Controls.Add(Me.lblImporteRedondeo)
        Me.Frame3.Controls.Add(Me.lblImporteTotal)
        Me.Frame3.Controls.Add(Me.Label15)
        Me.Frame3.Controls.Add(Me.lblImporteSubTotal)
        Me.Frame3.Controls.Add(Me.Label13)
        Me.Frame3.Controls.Add(Me.lblFactura)
        Me.Frame3.Controls.Add(Me.Label9)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(11, 599)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(612, 40)
        Me.Frame3.TabIndex = 25
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Totales de la Factura"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(333, 17)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(70, 14)
        Me.Label11.TabIndex = 46
        Me.Label11.Text = "Redondeo :"
        '
        'lblImporteRedondeo
        '
        Me.lblImporteRedondeo.BackColor = System.Drawing.SystemColors.Window
        Me.lblImporteRedondeo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImporteRedondeo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblImporteRedondeo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblImporteRedondeo.Location = New System.Drawing.Point(398, 16)
        Me.lblImporteRedondeo.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblImporteRedondeo.Name = "lblImporteRedondeo"
        Me.lblImporteRedondeo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblImporteRedondeo.Size = New System.Drawing.Size(64, 17)
        Me.lblImporteRedondeo.TabIndex = 45
        Me.lblImporteRedondeo.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblImporteTotal
        '
        Me.lblImporteTotal.BackColor = System.Drawing.SystemColors.Window
        Me.lblImporteTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImporteTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblImporteTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblImporteTotal.Location = New System.Drawing.Point(508, 16)
        Me.lblImporteTotal.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblImporteTotal.Name = "lblImporteTotal"
        Me.lblImporteTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblImporteTotal.Size = New System.Drawing.Size(97, 17)
        Me.lblImporteTotal.TabIndex = 31
        Me.lblImporteTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label15.Location = New System.Drawing.Point(466, 16)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label15.Size = New System.Drawing.Size(50, 14)
        Me.Label15.TabIndex = 30
        Me.Label15.Text = "Total :"
        '
        'lblImporteSubTotal
        '
        Me.lblImporteSubTotal.BackColor = System.Drawing.SystemColors.Window
        Me.lblImporteSubTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImporteSubTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblImporteSubTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblImporteSubTotal.Location = New System.Drawing.Point(238, 17)
        Me.lblImporteSubTotal.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblImporteSubTotal.Name = "lblImporteSubTotal"
        Me.lblImporteSubTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblImporteSubTotal.Size = New System.Drawing.Size(77, 17)
        Me.lblImporteSubTotal.TabIndex = 29
        Me.lblImporteSubTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(183, 17)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(58, 14)
        Me.Label13.TabIndex = 28
        Me.Label13.Text = "SubTotal :"
        '
        'lblFactura
        '
        Me.lblFactura.BackColor = System.Drawing.SystemColors.Window
        Me.lblFactura.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFactura.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFactura.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFactura.Location = New System.Drawing.Point(58, 15)
        Me.lblFactura.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFactura.Name = "lblFactura"
        Me.lblFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFactura.Size = New System.Drawing.Size(121, 17)
        Me.lblFactura.TabIndex = 27
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(9, 16)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(52, 14)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "Factura :"
        '
        'flexVentasPendientes
        '
        Me.flexVentasPendientes.DataSource = Nothing
        Me.flexVentasPendientes.Location = New System.Drawing.Point(19, 409)
        Me.flexVentasPendientes.Margin = New System.Windows.Forms.Padding(2)
        Me.flexVentasPendientes.Name = "flexVentasPendientes"
        Me.flexVentasPendientes.OcxState = CType(resources.GetObject("flexVentasPendientes.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexVentasPendientes.Size = New System.Drawing.Size(604, 151)
        Me.flexVentasPendientes.TabIndex = 22
        '
        'flexDetalleVenta
        '
        Me.flexDetalleVenta.DataSource = Nothing
        Me.flexDetalleVenta.Location = New System.Drawing.Point(256, 127)
        Me.flexDetalleVenta.Margin = New System.Windows.Forms.Padding(2)
        Me.flexDetalleVenta.Name = "flexDetalleVenta"
        Me.flexDetalleVenta.OcxState = CType(resources.GetObject("flexDetalleVenta.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalleVenta.Size = New System.Drawing.Size(628, 177)
        Me.flexDetalleVenta.TabIndex = 18
        '
        'flexVentas
        '
        Me.flexVentas.DataSource = Nothing
        Me.flexVentas.Location = New System.Drawing.Point(11, 161)
        Me.flexVentas.Margin = New System.Windows.Forms.Padding(2)
        Me.flexVentas.Name = "flexVentas"
        Me.flexVentas.OcxState = CType(resources.GetObject("flexVentas.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexVentas.Size = New System.Drawing.Size(228, 167)
        Me.flexVentas.TabIndex = 9
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtFolioFactura)
        Me.Frame2.Controls.Add(Me.dtpFechaRegistro)
        Me.Frame2.Controls.Add(Me.Label5)
        Me.Frame2.Controls.Add(Me.Label4)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(651, 11)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(240, 58)
        Me.Frame2.TabIndex = 14
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Facturas Registradas"
        '
        'dtpFechaRegistro
        '
        Me.dtpFechaRegistro.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaRegistro.Location = New System.Drawing.Point(112, 34)
        Me.dtpFechaRegistro.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaRegistro.Name = "dtpFechaRegistro"
        Me.dtpFechaRegistro.Size = New System.Drawing.Size(121, 20)
        Me.dtpFechaRegistro.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(6, 37)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(102, 13)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Fecha de Registro :"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(6, 17)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(80, 12)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Folio Factura :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtPorcentaje)
        Me.Frame1.Controls.Add(Me.optPorcentual)
        Me.Frame1.Controls.Add(Me.optManual)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(289, 12)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(152, 57)
        Me.Frame1.TabIndex = 12
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Método"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(124, 35)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(15, 13)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "%"
        '
        'dtpFechaVenta
        '
        Me.dtpFechaVenta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaVenta.Location = New System.Drawing.Point(81, 20)
        Me.dtpFechaVenta.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaVenta.Name = "dtpFechaVenta"
        Me.dtpFechaVenta.Size = New System.Drawing.Size(129, 20)
        Me.dtpFechaVenta.TabIndex = 0
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(81, 42)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(129, 21)
        Me.dbcSucursal.TabIndex = 2
        '
        'DtpHasta
        '
        Me.DtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DtpHasta.Location = New System.Drawing.Point(137, 359)
        Me.DtpHasta.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpHasta.Name = "DtpHasta"
        Me.DtpHasta.Size = New System.Drawing.Size(102, 20)
        Me.DtpHasta.TabIndex = 21
        '
        'lblMoneda
        '
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Window
        Me.lblMoneda.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMoneda.Location = New System.Drawing.Point(137, 339)
        Me.lblMoneda.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.Size = New System.Drawing.Size(102, 18)
        Me.lblMoneda.TabIndex = 67
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.BackColor = System.Drawing.SystemColors.Control
        Me.Label21.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label21.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label21.Location = New System.Drawing.Point(165, 381)
        Me.Label21.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label21.Size = New System.Drawing.Size(35, 13)
        Me.Label21.TabIndex = 66
        Me.Label21.Text = "Hasta"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.BackColor = System.Drawing.SystemColors.Control
        Me.Label20.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label20.Location = New System.Drawing.Point(55, 381)
        Me.Label20.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label20.Size = New System.Drawing.Size(38, 13)
        Me.Label20.TabIndex = 65
        Me.Label20.Text = "Desde"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label17.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label17.Location = New System.Drawing.Point(14, 141)
        Me.Label17.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label17.Size = New System.Drawing.Size(225, 18)
        Me.Label17.TabIndex = 64
        Me.Label17.Text = "Ventas"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblSubTot
        '
        Me.lblSubTot.BackColor = System.Drawing.SystemColors.Control
        Me.lblSubTot.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSubTot.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSubTot.Location = New System.Drawing.Point(600, 396)
        Me.lblSubTot.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSubTot.Name = "lblSubTot"
        Me.lblSubTot.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSubTot.Size = New System.Drawing.Size(49, 14)
        Me.lblSubTot.TabIndex = 49
        Me.lblSubTot.Visible = False
        '
        'lblIva
        '
        Me.lblIva.BackColor = System.Drawing.SystemColors.Control
        Me.lblIva.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblIva.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblIva.Location = New System.Drawing.Point(630, 356)
        Me.lblIva.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblIva.Name = "lblIva"
        Me.lblIva.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblIva.Size = New System.Drawing.Size(25, 20)
        Me.lblIva.TabIndex = 48
        Me.lblIva.Visible = False
        '
        'lblDescuento
        '
        Me.lblDescuento.BackColor = System.Drawing.SystemColors.Control
        Me.lblDescuento.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDescuento.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDescuento.Location = New System.Drawing.Point(564, 377)
        Me.lblDescuento.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDescuento.Name = "lblDescuento"
        Me.lblDescuento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDescuento.Size = New System.Drawing.Size(49, 14)
        Me.lblDescuento.TabIndex = 47
        Me.lblDescuento.Visible = False
        '
        'lblCantidad
        '
        Me.lblCantidad.BackColor = System.Drawing.SystemColors.Window
        Me.lblCantidad.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCantidad.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCantidad.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCantidad.Location = New System.Drawing.Point(613, 310)
        Me.lblCantidad.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCantidad.Name = "lblCantidad"
        Me.lblCantidad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCantidad.Size = New System.Drawing.Size(45, 17)
        Me.lblCantidad.TabIndex = 44
        Me.lblCantidad.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label19.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label19.Location = New System.Drawing.Point(192, 81)
        Me.Label19.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label19.Size = New System.Drawing.Size(699, 44)
        Me.Label19.TabIndex = 42
        Me.Label19.Text = resources.GetString("Label19.Text")
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label16.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label16.Location = New System.Drawing.Point(266, 339)
        Me.Label16.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label16.Size = New System.Drawing.Size(628, 17)
        Me.Label16.TabIndex = 41
        Me.Label16.Text = "Presione <F5> o Doble Click Para Incluir una Partida en el Grid de Detalle de Ven" &
    "tas"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblDesc
        '
        Me.lblDesc.BackColor = System.Drawing.SystemColors.Window
        Me.lblDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDesc.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDesc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblDesc.Location = New System.Drawing.Point(19, 569)
        Me.lblDesc.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDesc.Size = New System.Drawing.Size(604, 17)
        Me.lblDesc.TabIndex = 24
        Me.lblDesc.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblDescripcion
        '
        Me.lblDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.lblDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDescripcion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDescripcion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDescripcion.Location = New System.Drawing.Point(263, 311)
        Me.lblDescripcion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDescripcion.Name = "lblDescripcion"
        Me.lblDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDescripcion.Size = New System.Drawing.Size(346, 17)
        Me.lblDescripcion.TabIndex = 23
        Me.lblDescripcion.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblEstadoFolio
        '
        Me.lblEstadoFolio.BackColor = System.Drawing.SystemColors.Window
        Me.lblEstadoFolio.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEstadoFolio.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEstadoFolio.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblEstadoFolio.Location = New System.Drawing.Point(26, 340)
        Me.lblEstadoFolio.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblEstadoFolio.Name = "lblEstadoFolio"
        Me.lblEstadoFolio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEstadoFolio.Size = New System.Drawing.Size(95, 18)
        Me.lblEstadoFolio.TabIndex = 17
        Me.lblEstadoFolio.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 44)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(62, 12)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Sucursal :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 23)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(94, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Fecha Venta :"
        '
        'frmFactAnalisisVentas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(902, 707)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.DtpDesde)
        Me.Controls.Add(Me.Frame7)
        Me.Controls.Add(Me.Frame6)
        Me.Controls.Add(Me.Frame5)
        Me.Controls.Add(Me.txtFlex)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.txtDescripcion)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.flexVentasPendientes)
        Me.Controls.Add(Me.flexDetalleVenta)
        Me.Controls.Add(Me.flexVentas)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.txtCodSucursal)
        Me.Controls.Add(Me.dtpFechaVenta)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.DtpHasta)
        Me.Controls.Add(Me.lblMoneda)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.lblSubTot)
        Me.Controls.Add(Me.lblIva)
        Me.Controls.Add(Me.lblDescuento)
        Me.Controls.Add(Me.lblCantidad)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.lblDesc)
        Me.Controls.Add(Me.lblDescripcion)
        Me.Controls.Add(Me.lblEstadoFolio)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(8, 113)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmFactAnalisisVentas"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Análisis de las Ventas"
        Me.Frame7.ResumeLayout(False)
        Me.Frame6.ResumeLayout(False)
        Me.Frame5.ResumeLayout(False)
        Me.Frame5.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        CType(Me.flexVentasPendientes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.flexDetalleVenta, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.flexVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Public Function FoliosPendientesPtoVenta() As Boolean
        Dim I As Integer
        With flexVentas
            For I = 1 To .Rows - 1
                If (Trim(.get_TextMatrix(I, C_COLCAMBIOS)) = "S" And Trim(.get_TextMatrix(I, C_COLFACTURAPTOVTA)) = "S") Or (Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) = "TC" And Trim(.get_TextMatrix(I, C_COLFACTURAPTOVTA)) = "S") Then
                    FoliosPendientesPtoVenta = True
                    Exit Function
                End If
            Next
            FoliosPendientesPtoVenta = False
        End With
    End Function

    Public Function FoliosPendientesCorporativo() As Boolean
        Dim I As Integer
        With flexVentas
            For I = 1 To .Rows - 1
                If (Trim(.get_TextMatrix(I, C_COLCAMBIOS)) = "S" And Trim(.get_TextMatrix(I, C_COLFACTURAPTOVTA)) = "") Or (Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) = "TC" And Trim(.get_TextMatrix(I, C_COLFACTURAPTOVTA)) = "") Then
                    FoliosPendientesCorporativo = True
                    Exit Function
                End If
            Next
            FoliosPendientesCorporativo = False
        End With
    End Function

    Public Function FoliosPendientes() As Boolean
        Dim I As Integer
        With flexVentas
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCAMBIOS)) = "S" Then
                    FoliosPendientes = True
                    Exit Function
                End If
            Next
            FoliosPendientes = False
        End With
    End Function

    Sub Buscar()
        'On Error GoTo Err_Renamed
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas 
        Dim I As Integer

        'strControlActual = UCase(txtFolioFactura.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control


        Select Case strControlActual
            Case "TXTFOLIOFACTURA"
                If Me.FoliosPendientes() Then
                    MsgBox("No es posible consultar facturas, ya ha generado algun(s) folios adicionales" & vbNewLine & "  Para poder consultar otras facturas, debera generar la factura correspondiente", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Sub
                End If
                strCaptionForm = "Consulta de Facturas"
                gStrSql = "SELECT FolioFactura,FacturaAdicional,RTRIM(DBO.FormatFecha(FechaFactura,10)) AS Fecha,DBO.FormatCantidad(Total) AS Importe FROM Facturas WHERE Condicion = '' AND /*LEFT(DescEspecial,14) = 'Ventas del Dia' AND*/ TipoFactura = 'N' AND Origen = 'C' ORDER BY FechaFactura Desc, FolioFactura desc "
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

        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 7000, RsGral, strTag, strCaptionForm)
        'Carga el formulario de consulta  
        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTFOLIOFACTURA"
                    '.set_ColAlignment(2, 4) 
                    .set_ColWidth(0, 0, 1800) 'Columna de la Factura
                    .set_ColWidth(1, 0, 1800) 'Columna de la Factura Adicional
                    .set_ColWidth(2, 0, 1500) 'Columna de la Fecha
                    .set_ColWidth(3, 0, 1700) 'Columna del Importe
            End Select
        End With
        FueraChange = True
        FrmConsultas.ShowDialog()
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatos()
        On Error GoTo Err_Renamed
        Dim I As Integer
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        FueraChange = False
        gStrSql = "SELECT VtaDet.FolioAdicional,SUM(VtaDet.PrecioRealAdicional * VtaDet.CantidadAdicional) AS Importe,VtaDet.CodCajaAdicional,VtaDet.CondicionAdicional,VtaDet.CodSucursalAdicional,VtaDet.MonedaAdicional,VtaDet.CondicionAdicional,VtaDet.RedondeoAdicional,CASE WHEN LEFT(FolioAdicional,1) = 'S' THEN 'V' WHEN LEFT(FolioAdicional,1) = 'A' THEN 'A' END AS TipoMovto,Fact.FolioFactura,Fact.Total,Fact.Descuento,Fact.SubTotal,Fact.Iva,Fact.Redondeo,Fact.FechaFactura,Fact.TipoCambio, Fact.FacturaAdicional,Fact.CodCliente,Fact.DescEspecial,Fact.DesgloseIva,Fact.CodCaja,CA.DescAlmacen " & "FROM Facturas Fact INNER JOIN MovimientosVentasDet VtaDet ON Fact.FolioFactura = VtaDet.FolioFactura " & "INNER JOIN CatAlmacen CA ON Fact.CodSucursal = CA.CodAlmacen " & "WHERE VtaDet.FolioFactura = '" & txtFolioFactura.Text & "' AND Condicion = '' AND /*LEFT(DescEspecial,14) = 'Ventas del Dia' AND*/ Fact.TipoFactura = 'N' AND Fact.Origen = 'C' " & "GROUP BY VtaDet.FolioAdicional,VtaDet.CodCajaAdicional,VtaDet.CondicionAdicional,VtaDet.CodSucursalAdicional,VtaDet.MonedaAdicional,VtaDet.RedondeoAdicional,Fact.Total,Fact.FolioFactura,Fact.Descuento,Fact.SubTotal,Fact.Iva,Fact.Redondeo,Fact.FechaFactura,Fact.TipoCambio,Fact.FacturaAdicional,Fact.CodCliente,Fact.DescEspecial,Fact.DesgloseIva,Fact.CodCaja,CA.DescAlmacen ORDER BY FolioAdicional"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            dtpFechaVenta.Enabled = False
            dbcSucursal.Enabled = False
            FueraChange = True
            dbcSucursal.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            dtpFechaVenta.Value = VB6.Format(RsGral.Fields("FechaFactura").Value, "dd/MMM/yyyy")
            FueraChange = False
            txtFacturaAdicional.Text = Trim(RsGral.Fields("FacturaAdicional").Value)
            txtFacturaAdicional.Tag = txtFacturaAdicional.Text
            CodCliente = RsGral.Fields("CodCliente").Value
            DescEspecial = Trim(RsGral.Fields("DescEspecial").Value)
            chkDesglosarIva.Enabled = True
            If RsGral.Fields("DesgloseIva").Value = True Then
                DesgloseIva = 1
                chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                DesgloseIva = 0
                chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
            Caja = RsGral.Fields("CodCaja").Value
            chkDesglosarIva.Enabled = False
            LimpiarGridVentas()
            LimpiarGridDetalle()
            LimpiarGridPendientes()
            With flexVentas
                I = 1
                lblSubTot.Text = RsGral.Fields("SubTotal").Value
                lblImporteSubTotal.Text = VB6.Format(RsGral.Fields("Total").Value, "###,##0.00")
                lblDescuento.Text = RsGral.Fields("Descuento").Value
                lblIva.Text = RsGral.Fields("Iva").Value
                lblImporteRedondeo.Text = VB6.Format(RsGral.Fields("Redondeo").Value, "###,##0.00")
                lblImporteTotal.Text = VB6.Format(RsGral.Fields("Total").Value + RsGral.Fields("Redondeo").Value, "###,##0.00")
                dtpFechaRegistro.Value = VB6.Format(RsGral.Fields("FechaFactura").Value, "dd/mmm/yyyy")
                TipoCambio = CDec(VB6.Format(RsGral.Fields("TipoCambio").Value, "###,##0.00"))
                lblFactura.Text = RsGral.Fields("FolioFactura").Value
                Do While Not RsGral.EOF
                    .set_TextMatrix(I, C_COLFOLIOVENTA, Trim(RsGral.Fields("FolioAdicional").Value))
                    .set_TextMatrix(I, C_ColIMPORTE, VB6.Format(RsGral.Fields("importe").Value + RsGral.Fields("RedondeoAdicional").Value, "###,##0.00"))
                    .set_TextMatrix(I, C_ColMONEDA, RsGral.Fields("MonedaAdicional").Value)
                    .set_TextMatrix(I, C_COLCODCAJA, RsGral.Fields("CodCajaAdicional").Value)
                    .set_TextMatrix(I, C_ColCODSUCURSAL, RsGral.Fields("CodSucursalAdicional").Value)
                    .set_TextMatrix(I, C_ColCONDICION, RsGral.Fields("CondicionAdicional").Value)
                    .set_TextMatrix(I, C_COLGRABADO, "S")
                    .set_TextMatrix(I, C_COLTIPOMOVTO, Trim(RsGral.Fields("TipoMovto").Value))
                    If I = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    I = I + 1
                    RsGral.MoveNext()
                Loop
                If .Rows < 12 Then
                    .Rows = 12
                End If
            End With
            mblnNuevo = True
            '        FueraChange = True
            '        txtCodSucursal = ""
            '        dbcSucursal.text = ""
            '        FueraChange = False
            cmdGenerarFactura.Enabled = False
            cmdImpresionTickets.Enabled = True
            cmdImprimirFactura.Enabled = True
            chkDoctoCliente.Enabled = False
            flexVentas_EnterCell(flexVentas, New System.EventArgs())
        Else
            MsgBox("Este folio de factura no existe, o es un folio de factura generado desde el punto de venta" & vbNewLine & "                                                 Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            FueraChange = True
            txtFolioFactura.Text = ""
            FueraChange = False
            txtFolioFactura.Focus()
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
Err_Renamed:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub GenerarFacturaPuntoVenta()
        On Error GoTo Err_Renamed
        Dim blnTransaccion As Boolean
        Dim I As Integer
        Dim J As Integer
        Dim Fecha As String
        Dim ConsFactura As String
        Dim FolioFactura As String
        Dim CodCliente As Integer
        Dim NomCliente As String
        Dim RfcCliente As String
        Dim SubTotal As Decimal
        Dim TotalDescuento As Decimal
        Dim Total As Decimal
        Dim TotalPesos As Decimal
        Dim Iva As Decimal
        Dim RedondeoDolares As Decimal
        Dim RedondeoPesos As Decimal
        Dim RsAux As ADODB.Recordset
        Dim Sql As String
        Dim CodCaja As Integer
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        gStrSql = "SELECT TOP 1 CodCaja FROM CatCajas WHERE CodAlmacen = " & txtCodSucursal.Text & " ORDER BY CodCaja"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            CodCaja = RsGral.Fields("CodCaja").Value
        End If
        With flexVentas
            SubTotal = 0
            TotalDescuento = 0
            Total = 0
            Iva = 0
            RedondeoDolares = 0
            TipoCambio = gcurCorpoTIPOCAMBIODOLAR
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLFACTURAPTOVTA)) = "S" Then
                    SubTotal = SubTotal + System.Math.Round((CDec(Numerico(.get_TextMatrix(I, C_COLSUBTOTALADICIONAL))) + CDec(Numerico(.get_TextMatrix(I, C_COLREDONDEOADICIONAL)))) * CDec(Numerico(.get_TextMatrix(I, C_ColTIPOCAMBIO))), 1)
                    TotalDescuento = TotalDescuento + System.Math.Round(CDec(Numerico(.get_TextMatrix(I, C_COLDESCUENTOADICIONAL))) * CDec(Numerico(.get_TextMatrix(I, C_ColTIPOCAMBIO))), 1)
                    Iva = Iva + System.Math.Round(CDec(Numerico(.get_TextMatrix(I, C_COLIVAADICIONAL))) * CDec(Numerico(.get_TextMatrix(I, C_ColTIPOCAMBIO))), 1)
                    Total = ((SubTotal - TotalDescuento) + Iva)
                End If
            Next
            '        TotalPesos = Total * gcurCorpoTIPOCAMBIODOLAR
            '        RedondeoPesos = ModCorporativo.RedondeoUnidadFinal(CDbl(TotalPesos), CDbl(gcurRedondeo))
            '        RedondeoDolares = RedondeoPesos / gcurCorpoTIPOCAMBIODOLAR
            'Generar Folio de la Factura
            Fecha = ""
            Fecha = Fecha & (txtCodSucursal.Text + "00") & (Year(dtpFechaVenta.Value) + "0000") & (Month(dtpFechaVenta.Value) + "00") & (dtpFechaVenta.Value + "00")
            '''gStrSql = "SELECT Prefijo,Consecutivo FROM FoliosCorporativo WHERE CodFolio =" & TipoMovto
            gStrSql = "Select CodFolio, Prefijo,  Consecutivo + 1  as Consecutivo From CatFolios Where CodFolio = " & TipoMovto & " And CodAlmacen = " & CShort(Numerico((txtCodSucursal.Text))) & " "
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                ConsFactura = Str(RsGral.Fields("Consecutivo").Value)
            End If
            FolioFactura = Trim(RsGral.Fields("Prefijo").Value) & Fecha & VB6.Format(ConsFactura, "000000")
            ''''Actualizar el Consecutivo en Folios de Factura
            ''''ModStoredProcedures.PR_IMEFoliosCorporativo TipoMovto, "", "", ConsFactura, C_MODIFICACION, 0
            '''Actualiza el consecutivo de facturas para la sucursal seleccionada
            ModStoredProcedures.PR_IMECatFolios(CStr(TipoMovto), Trim(Str(CShort(Numerico((txtCodSucursal.Text))))), "", "", ConsFactura, C_MODIFICACION, CStr(1))
            Cmd.Execute()

            'Guardar el Folio de la Factura en el Detalle de las Ventas
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLFOLIOADICIONAL)) <> "" And Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) <> "F" And Trim(.get_TextMatrix(I, C_COLFACTURAPTOVTA)) = "S" Then
                    Sql = "SELECT * FROM MovimientosVentasDet WHERE FolioAdicional = '" & .get_TextMatrix(I, C_COLFOLIOADICIONAL) & "' AND EstatusAdicional = 'V'"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                    RsAux = Cmd.Execute
                    If RsAux.RecordCount > 0 Then
                        Do While Not RsAux.EOF
                            ModStoredProcedures.PR_IMEMovimientosVentasCab(RsAux.Fields("FolioVenta").Value, "01/01/1900", CStr(0), CStr(0), CStr(0), CStr(0), "", "", "", "", CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), "", "01/01/1900", CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), "V", FolioFactura, "01/01/1900", "", CStr(0), "", 0, 0, C_MODIFICACION, CStr(6))
                            Cmd.Execute()
                            ModStoredProcedures.PR_IE_MovimientosVentasDet(RsAux.Fields("FolioVenta").Value, CStr(RsAux.Fields("NumPartida").Value), "0", "", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "", "", "", "0", "", FolioFactura, "", "0", "0", "", "", "0", "0", "0", "01/01/1900", "0", "0", "0", "0", "", C_MODIFICACION, CStr(5))
                            Cmd.Execute()
                            RsAux.MoveNext()
                            .set_TextMatrix(I, C_COLCAMBIOS, "")
                            If Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) = "TC" Then
                                .set_TextMatrix(I, C_COLGRABADO, "S")
                                '.TextMatrix(I, C_COLESTADOFOLIO) = "N"
                            End If
                        Loop
                    End If
                End If
            Next
            If chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Checked Then
                DesgloseIva = 1
            ElseIf chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                DesgloseIva = 0
            End If

            'obtener importes en dolares
            SubTotal = System.Math.Round(SubTotal / TipoCambio, 4)
            TotalDescuento = System.Math.Round(TotalDescuento / TipoCambio, 4)
            Iva = System.Math.Round(Iva / TipoCambio, 4)
            TotalPesos = Total
            Total = System.Math.Round(Total / TipoCambio, 4)
            'obtener redondeo de la factura
            'RedondeoPesos = ModCorporativo.RedondeoUnidadFinal(CCur(TotalPesos), CDbl(gcurRedondeo))
            'RedondeoDolares = Round(RedondeoPesos / gcurCorpoTIPOCAMBIODOLAR, 4)
            RedondeoDolares = 0

            '        SubTotal = Format(SubTotal, "#####0.0000")
            '        TotalDescuento = Format(TotalDescuento, "#####0.0000")
            '        Iva = Format(Iva, "#####0.0000")
            '        Total = Format(Total, "#####0.0000")
            '        RedondeoDolares = Format(RedondeoDolares, "#####0.0000")
            'Guardar la Factura
            ModStoredProcedures.PR_IME_Facturas(FolioFactura, "1", txtCodSucursal.Text, CStr(CodCaja), VB6.Format(dtpFechaVenta.Value, C_FORMATFECHAGUARDAR), "N", "", CStr(gintCodRFC), gstrNombreCliente, gstrRFCCliente, C_DOLAR, CStr(TipoCambio), CStr(SubTotal), CStr(TotalDescuento), CStr(Iva), CStr(Total), CStr(RedondeoDolares), CStr(gcurCorpoTASAIVA), "V", "01/01/1900", "0", "COMPRAS JOYERIA Y RELOJERIA", "0", "0", "0", "", "C", CStr(DesgloseIva), C_INSERCION, CStr(0))
            Cmd.Execute()
            txtFolioFactura.Text = FolioFactura

        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("Se ha generado la factura " & FolioFactura, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        'Generar Notas de Credito
        With flexVentas
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) = "TC" And Trim(.get_TextMatrix(I, C_COLTIENEDEVOLUCION)) = "S" And Trim(.get_TextMatrix(I, C_COLFACTURAPTOVTA)) = "S" Then
                    gStrSql = "SELECT FolioDevolucion FROM DevolucionesCab WHERE FolioVenta = '" & .get_TextMatrix(I, C_COLFOLIOVENTA) & "'"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                    RsGral = Cmd.Execute
                    If RsGral.RecordCount > 0 Then
                        Do While Not RsGral.EOF
                            ModCorporativo.TicketNotaCredito(RsGral.Fields("FolioDevolucion").Value, .get_TextMatrix(I, C_COLFOLIOADICIONAL), CShort(.get_TextMatrix(I, C_COLCODCAJA)))
                            RsGral.MoveNext()
                        Loop
                    End If
                End If
            Next
        End With
        'Actualizamos el estatus del folio de venta a facturado
        With flexVentas
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLFOLIOADICIONAL)) <> "" And Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) <> "F" And Trim(.get_TextMatrix(I, C_COLFACTURAPTOVTA)) = "S" Then
                    .set_TextMatrix(I, C_COLESTADOFOLIO, "F")
                End If
            Next
        End With
        GenerarFactPtoVenta = False
        If Not Me.FoliosPendientes() Then
            Limpiar()
        End If
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Sub

    Function ExistenFoliosMarcados() As Boolean
        Dim I As Integer
        ExistenFoliosMarcados = False
        With flexVentas
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLFACTURAPTOVTA)) = "S" Then
                    ExistenFoliosMarcados = True
                    Exit Function
                End If
            Next
        End With
    End Function

    Sub GuardarFactura()
        On Error GoTo Err_Renamed
        Dim blnTransaccion As Boolean
        Dim I As Integer
        Dim J As Integer
        Dim Fecha As String
        Dim ConsFactura As String
        Dim FolioFactura As String
        Dim NomCliente As String
        Dim RfcCliente As String
        Dim SubTotal As Decimal
        Dim TotalDescuento As Decimal
        Dim Total As Decimal
        Dim TotalPesos As Decimal
        Dim Iva As Decimal
        Dim RedondeoDolares As Decimal
        Dim RedondeoPesos As Decimal
        Dim RsAux As ADODB.Recordset
        Dim Sql As String
        Dim CodCaja As Integer
        gStrSql = "SELECT TOP 1 CodCaja FROM CatCajas WHERE CodAlmacen = " & txtCodSucursal.Text & " ORDER BY CodCaja"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            CodCaja = RsGral.Fields("CodCaja").Value
        End If
        'Calcular Importes Factura
        'And GenerarFactPtoVenta
        '    If chkDoctoCliente.Value = vbChecked Then
        '        If Not ExistenFoliosMarcados() Then
        '            MsgBox "No ha seleccionado ningun folio para la factura del punto de venta, Favor de verificar...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '            Exit Sub
        '        End If
        '        If gintCodRFC = 0 And Trim(gstrNombreCliente) = "" And Trim(gstrRFCCliente) = "" Then
        '            MsgBox "No se han proporcionado los datos fiscales del cliente, Favor de verificar...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '            Exit Sub
        '        Else
        '            GenerarFacturaPuntoVenta
        '            Exit Sub
        '        End If
        'ElseIf chkDoctoCliente.Value = vbChecked And Not GenerarFactPtoVenta Then
        '    MsgBox "Ya se genero factura para el punto de venta, Favor de verificar...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '    Exit Sub
        'ElseIf gblnCambiosAnalisis Or mblnFactura Then
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        With flexVentas
            'Generar los Folios Adicionales Pendientes para folios facturados o pagados con tarjeta de credito
            '            For I = 1 To .Rows - 1
            '                If (.TextMatrix(I, C_COLESTADOFOLIO) = "TC" Or Trim(.TextMatrix(I, C_COLESTADOFOLIO)) = "F") And Trim(.TextMatrix(I, C_COLINCLUIRFACTURA)) = "S" And Trim(.TextMatrix(I, C_COLGRABADO)) = "" Then
            '                    sql = "SELECT VtaCab.FolioVenta,VtaDet.CodArticulo,CatArt.CodAlmacenOrigen,CatArt.DescArticulo," & _
            ''                    "SUM(ISNULL(DevDet.CantidadDevol,0)) AS CantidadDevuelta,VtaDet.Cantidad,VtaDet.PorcDescuentos," & _
            ''                    "VtaDet.PorcPromociones,VtaDet.PrecioLista,VtaDet.NumPartida," & _
            ''                    "VtaDet.PrecioLista * VtaDet.Cantidad AS Importe,VtaDet.PrecioReal * VtaDet.Cantidad AS ImporteDescto," & _
            ''                    "CatArt.CodGrupo,ISNULL(Cf.DescFamilia,'') AS DescFamilia," & _
            ''                    "VtaDet.ImptePromociones,VtaDet.ImpteDescuentos,VtaDet.PrecioLista,VtaDet.PrecioListaSinIva," & _
            ''                    "VtaDet.PrecioReal,VtaDet.IvaReal,VtaDet.PorcAdicional,VtaDet.DescArticuloAdicional,VtaDet.FolioAdicional,VtaDet.CantidadAdicional,VtaDet.EstatusAdicional " & _
            ''                    "FROM MovimientosVentasCab VtaCab INNER JOIN MovimientosVentasDet VtaDet " & _
            ''                    "ON VtaCab.FolioVenta = VtaDet.FolioVenta " & _
            ''                    "INNER JOIN CatArticulos CatArt ON VtaDet.CodArticulo = CatArt.CodArticulo " & _
            ''                    "LEFT OUTER JOIN CatFamilias Cf ON CatArt.CodGrupo = Cf.CodGrupo AND CatArt.CodFamilia = Cf.CodFamilia " & _
            ''                    "LEFT OUTER JOIN DevolucionesCab DevCab ON VtaDet.FolioVenta = DevCab.FolioVenta " & _
            ''                    "LEFT OUTER JOIN DevolucionesDet DevDet ON DevCab.FolioDevolucion = DevDet.FolioDevolucion AND VtaDet.CodArticulo = DevDet.CodArticulo " & _
            ''                    "WHERE VtaCab.FolioVenta ='" & .TextMatrix(I, C_COLFOLIOVENTA) & "' AND VtaCab.Estatus <> 'C' AND ISNULL(DevCab.Estatus,'') <> 'C' " & _
            ''                    "GROUP BY VtaCab.FolioVenta,VtaDet.CodArticulo,CatArt.CodAlmacenOrigen,CatArt.DescArticulo," & _
            ''                    "VtaDet.Cantidad,VtaDet.PorcDescuentos,VtaDet.PorcPromociones,VtaDet.PrecioLista,VtaDet.NumPartida," & _
            ''                    "VtaDet.PrecioLista * VtaDet.Cantidad,VtaDet.PrecioReal * VtaDet.Cantidad," & _
            ''                    "CatArt.CodGrupo,ISNULL(Cf.DescFamilia,''),VtaDet.ImptePromociones,VtaDet.ImpteDescuentos,VtaDet.PrecioLista,VtaDet.PrecioListaSinIva," & _
            ''                    "VtaDet.PrecioReal,VtaDet.IvaReal,VtaDet.PorcAdicional,VtaDet.DescArticuloAdicional,VtaDet.FolioAdicional,VtaDet.CantidadAdicional,VtaDet.EstatusAdicional"
            '                    ModEstandar.BorraCmd
            '                    Cmd.CommandText = "dbo.Up_Select_Datos"
            '                    Cmd.CommandType = adCmdStoredProc
            '                    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
            '                    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, sql)
            '                    Set RsAux = Cmd.Execute
            '                    If RsAux.RecordCount > 0 Then
            '                        GeneraFolioAdicional
            '                        Do While Not RsAux.EOF
            '                            'RsAux!CodAlmacenOrigen <> 0 And
            '                            If (RsAux!Cantidad - RsAux!CantidadDevuelta) > 0 Then
            '                                ModStoredProcedures.PR_IE_MovimientosVentasDet RsAux!FolioVenta, CStr(RsAux!NumPartida), "0", "", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", _
            ''                                CStr(RsAux!ImptePromociones), CStr(RsAux!ImpteDescuentos), CStr(RsAux!PrecioLista), CStr(RsAux!PrecioListaSinIva), CStr(RsAux!PrecioReal), CStr(RsAux!IvaReal), _
            ''                                "", IIf(RsAux!CodGrupo = gCODRELOJERIA, "RELOJ", RsAux!DescFamilia), "", "0", FolioAdicional, "", "V", CStr(RsAux!Cantidad - RsAux!CantidadDevuelta), _
            ''                                flexVentas.TextMatrix(flexVentas.Row, C_COLTIPOCAMBIO), flexVentas.TextMatrix(flexVentas.Row, C_ColMONEDA), flexVentas.TextMatrix(flexVentas.Row, C_COLCONDICION), _
            ''                                flexVentas.TextMatrix(flexVentas.Row, C_ColPORCIVA), flexVentas.TextMatrix(flexVentas.Row, C_COLREDONDEOADICIONAL), flexVentas.TextMatrix(flexVentas.Row, C_COLANTICIPOADICIONAL), _
            ''                                Format(Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLFECHADEVENTA)), C_FORMATFECHAGUARDAR), flexVentas.TextMatrix(flexVentas.Row, C_COLCODSUCURSAL), _
            ''                                flexVentas.TextMatrix(flexVentas.Row, C_COLCODCAJA), flexVentas.TextMatrix(flexVentas.Row, C_COLCODVENDEDOR), flexVentas.TextMatrix(flexVentas.Row, C_COLCODCLIENTE), C_MODIFICACION, 1
            '                                Cmd.Execute
            '                                .TextMatrix(I, C_COLFOLIOADICIONAL) = FolioAdicional
            '                            End If
            '                            RsAux.MoveNext
            '                        Loop
            '                        ModStoredProcedures.PR_IMEMovimientosVentasCab .TextMatrix(I, C_COLFOLIOVENTA), "01/01/1900", "0", "0", "0", "0", "", "", "", "", "0", "0", "0", "0", "0", "0", "0", "0", "", "", "01/01/1900", _
            ''                        "0", "0", "0", "0", "0", "0", "V", "", "01/01/1900", "", "0", "", "0", C_MODIFICACION, 3
            '                        Cmd.Execute
            '                    End If
            '                End If
            '            Next

            'Generar Folio de la Factura
            Fecha = ""
            Fecha = Fecha & (txtCodSucursal.Text + "00") & (Year(dtpFechaVenta.Value) + "0000") & (Month(dtpFechaVenta.Value) + "00") & ((dtpFechaVenta.Value) + "00")
            '''gStrSql = "SELECT Prefijo,Consecutivo FROM FoliosCorporativo WHERE CodFolio =" & TipoMovto
            gStrSql = "Select CodFolio, Prefijo,  Consecutivo + 1  as Consecutivo From CatFolios Where CodFolio = " & TipoMovto & " And CodAlmacen = " & CShort(Numerico((txtCodSucursal.Text))) & " "
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                ConsFactura = Str(RsGral.Fields("Consecutivo").Value)
            End If
            FolioFactura = Trim(RsGral.Fields("Prefijo").Value) & Fecha & VB6.Format(ConsFactura, "000000")

            ''''Actualizar el Consecutivo en Folios de Factura
            ''''ModStoredProcedures.PR_IMEFoliosCorporativo TipoMovto, "", "", ConsFactura, C_MODIFICACION, 0
            '''Actualiza el consecutivo de facturas para la sucursal seleccionada
            ModStoredProcedures.PR_IMECatFolios(CStr(TipoMovto), Trim(Str(CShort(Numerico((txtCodSucursal.Text))))), "", "", ConsFactura, C_MODIFICACION, CStr(1))
            Cmd.Execute()

            'Guardar el Folio de la Factura en el Detalle de las Ventas
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLFOLIOADICIONAL)) <> "" And Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) <> "F" And Trim(.get_TextMatrix(I, C_COLINCFACTURA)) = "S" Then
                    Sql = "SELECT * FROM MovimientosVentasDet WHERE FolioAdicional = '" & .get_TextMatrix(I, C_COLFOLIOADICIONAL) & "' AND EstatusAdicional = 'V'"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                    RsAux = Cmd.Execute
                    If RsAux.RecordCount > 0 Then
                        Do While Not RsAux.EOF
                            ModStoredProcedures.PR_IMEMovimientosVentasCab(RsAux.Fields("FolioVenta").Value, "01/01/1900", CStr(0), CStr(0), CStr(0), CStr(0), "", "", "", "", CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), "", "01/01/1900", CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), CStr(0), "V", FolioFactura, "01/01/1900", "", CStr(0), "", 0, 0, C_MODIFICACION, CStr(6))
                            Cmd.Execute()
                            ModStoredProcedures.PR_IE_MovimientosVentasDet(RsAux.Fields("FolioVenta").Value, CStr(RsAux.Fields("NumPartida").Value), "0", "", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "", "", "", "0", "", FolioFactura, "", "0", "0", "", "", "0", "0", "0", "01/01/1900", "0", "0", "0", "0", "", C_MODIFICACION, CStr(5))
                            Cmd.Execute()
                            RsAux.MoveNext()
                            .set_TextMatrix(I, C_COLCAMBIOS, "")
                            If Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) = "TC" Then
                                .set_TextMatrix(I, C_COLGRABADO, "S")
                                '.TextMatrix(I, C_COLESTADOFOLIO) = "N"
                            End If
                        Loop
                    End If
                End If
            Next
        End With
        'Obtener los Datos del Cliente Publico en General
        gStrSql = "SELECT * FROM CatClientes WHERE CodCliente = 1"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            CodCliente = RsGral.Fields("CodCliente").Value
            NomCliente = Trim(RsGral.Fields("DescCliente").Value)
            RfcCliente = Trim(RsGral.Fields("Rfc").Value)
        End If
        If chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Checked Then
            DesgloseIva = 1
        ElseIf chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            DesgloseIva = 0
        End If
        TipoCambio = gcurCorpoTIPOCAMBIODOLAR
        'Calcular Importes de la Factura en pesos
        SubTotal = 0
        TotalDescuento = 0
        Iva = 0
        Total = 0
        With flexVentas
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLFOLIOADICIONAL)) <> "" And Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) <> "F" And Trim(.get_TextMatrix(I, C_COLINCFACTURA)) = "S" Then
                    SubTotal = SubTotal + System.Math.Round((CDec(Numerico(.get_TextMatrix(I, C_COLSUBTOTALADICIONAL))) + CDec(Numerico(.get_TextMatrix(I, C_COLREDONDEOADICIONAL)))) * CDec(Numerico(.get_TextMatrix(I, C_ColTIPOCAMBIO))), 2)
                    TotalDescuento = TotalDescuento + System.Math.Round(CDec(Numerico(.get_TextMatrix(I, C_COLDESCUENTOADICIONAL))) * CDec(Numerico(.get_TextMatrix(I, C_ColTIPOCAMBIO))), 2)
                    Iva = Iva + System.Math.Round(CDec(Numerico(.get_TextMatrix(I, C_COLIVAADICIONAL))) * CDec(Numerico(.get_TextMatrix(I, C_ColTIPOCAMBIO))), 2)
                    Total = ((SubTotal - TotalDescuento) + Iva)
                End If
            Next
        End With
        'obtener importes en dolares
        SubTotal = System.Math.Round(SubTotal / TipoCambio, 4)
        TotalDescuento = System.Math.Round(TotalDescuento / TipoCambio, 4)
        Iva = System.Math.Round(Iva / TipoCambio, 4)
        TotalPesos = Total
        Total = System.Math.Round(Total / TipoCambio, 4)
        'obtener redondeo de la factura
        'RedondeoPesos = ModCorporativo.RedondeoUnidadFinal(CCur(TotalPesos), CDbl(gcurRedondeo))
        'RedondeoDolares = Round(RedondeoPesos / gcurCorpoTIPOCAMBIODOLAR, 4)
        RedondeoDolares = 0

        lblSubTot.Text = CStr(SubTotal)
        lblImporteSubTotal.Text = VB6.Format(Total, "###,##0.00")
        lblDescuento.Text = CStr(TotalDescuento)
        lblIva.Text = CStr(Iva)
        lblImporteRedondeo.Text = VB6.Format(RedondeoDolares, "###,##0.00")
        lblImporteTotal.Text = VB6.Format(Total + RedondeoDolares, "###,##0.00")
        '        SubTotal = Format(CCur(Numerico(lblSubTot)), "#####0.0000")
        '        TotalDescuento = Format(CCur(Numerico(lblDescuento)), "#####0.0000")
        '        Iva = Format(CCur(Numerico(lblIva)), "#####0.0000")
        '        Total = Format(CCur(Numerico(lblImporteSubTotal)), "#####0.0000")
        '        RedondeoDolares = Format(CCur(Numerico(lblImporteRedondeo)), "#####0.0000")
        'Guardar la Factura
        ModStoredProcedures.PR_IME_Facturas(FolioFactura, "1", txtCodSucursal.Text, CStr(CodCaja), VB6.Format(dtpFechaVenta.Value, C_FORMATFECHAGUARDAR), "N", "", CStr(CodCliente), NomCliente, RfcCliente, C_DOLAR, CStr(TipoCambio), CStr(SubTotal), CStr(TotalDescuento), CStr(Iva), CStr(Total), CStr(RedondeoDolares), CStr(gcurCorpoTASAIVA), "V", "01/01/1900", "0", "Ventas del Dia: " & VB6.Format(dtpFechaVenta.Value, "dd/mmm/yyyy"), "0", "0", "0", Trim(txtFacturaAdicional.Text), "C", CStr(DesgloseIva), C_INSERCION, CStr(0))
        Cmd.Execute()
        txtFolioFactura.Text = FolioFactura

        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("Se ha generado la factura " & FolioFactura, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        'Else
        '    Exit Sub
        'End If
        Select Case MsgBox("¿Desea imprimir los tickets?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
            Case MsgBoxResult.Yes
                With flexVentas
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, C_COLFOLIOADICIONAL)) <> "" And .get_TextMatrix(I, C_COLINCFACTURA) = "S" Then
                            If .get_TextMatrix(I, C_ColMONEDA) = "P" Then
                                ModCorporativo.TicketVentaReducidoPesos(Trim(.get_TextMatrix(I, C_COLFOLIOADICIONAL)), CShort(.get_TextMatrix(I, C_COLCODCAJA)), IIf(Trim(.get_TextMatrix(I, C_COLTIPOMOVTO)) = "V", Trim(.get_TextMatrix(I, C_ColCONDICION)), "CO"), CShort(.get_TextMatrix(I, C_ColCODSUCURSAL)), Trim(.get_TextMatrix(I, C_COLTIPOMOVTO)))
                            ElseIf .get_TextMatrix(I, C_ColMONEDA) = "D" Then
                                ModCorporativo.TicketVentaReducidoDolares(Trim(.get_TextMatrix(I, C_COLFOLIOADICIONAL)), CShort(.get_TextMatrix(I, C_COLCODCAJA)), IIf(Trim(.get_TextMatrix(I, C_COLTIPOMOVTO)) = "V", Trim(.get_TextMatrix(I, C_ColCONDICION)), "CO"), CShort(.get_TextMatrix(I, C_ColCODSUCURSAL)), Trim(.get_TextMatrix(I, C_COLTIPOMOVTO)))
                            End If
                        End If
                    Next
                End With
            Case MsgBoxResult.No
        End Select
        Select Case MsgBox("¿Desea imprimir la factura?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
            Case MsgBoxResult.Yes
                TipoCambio = gcurCorpoTIPOCAMBIODOLAR
                dtpFechaRegistro.Value = dtpFechaVenta.Value
                CodCliente = 1
                DescEspecial = "Ventas del Dia: " & VB6.Format(dtpFechaVenta.Value, "dd/mmm/yyyy")
                ImprimirFactura()
            Case MsgBoxResult.No
        End Select
        'Generar Notas de Credito
        With flexVentas
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) = "TC" And Trim(.get_TextMatrix(I, C_COLTIENEDEVOLUCION)) = "S" And Trim(.get_TextMatrix(I, C_COLINCFACTURA)) = "S" Then
                    gStrSql = "SELECT FolioDevolucion FROM DevolucionesCab WHERE FolioVenta = '" & .get_TextMatrix(I, C_COLFOLIOVENTA) & "'"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                    RsGral = Cmd.Execute
                    If RsGral.RecordCount > 0 Then
                        Do While Not RsGral.EOF
                            ModCorporativo.TicketNotaCredito(RsGral.Fields("FolioDevolucion").Value, .get_TextMatrix(I, C_COLFOLIOADICIONAL), CShort(.get_TextMatrix(I, C_COLCODCAJA)))
                            RsGral.MoveNext()
                        Loop
                    End If
                End If
            Next
        End With
        'Actualizamos el folio de la factura
        With flexVentas
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLFOLIOADICIONAL)) <> "" And Trim(.get_TextMatrix(I, C_COLESTADOFOLIO)) <> "F" And Trim(.get_TextMatrix(I, C_COLINCFACTURA)) = "S" Then
                    .set_TextMatrix(I, C_COLESTADOFOLIO, "F")
                End If
            Next
        End With
        If Not Me.FoliosPendientes() Then
            Limpiar()
        End If
        Exit Sub
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Sub

    Function GuardarFolio() As Boolean
        On Error GoTo Err_Renamed
        Dim blnTransaccion As Boolean
        Dim I As Integer
        Dim RsAux As ADODB.Recordset
        Dim Sql As String
        Dim GuardoFolioAdicional As Boolean

        'Validar que el prec. pub no sea cero
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODIGOARTICULO)) = "" Then Exit For
                If Trim(.get_TextMatrix(I, C_COLCODIGOARTICULO)) <> "" And Trim(.get_TextMatrix(I, C_COLEXCLUIDO)) <> "EXCLUIDO" And (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) > 0 Then
                    If CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) = 0 Then
                        MsgBox("No se puede generar folio adicional con precio publico en cero." & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                        flexVentas.Focus()
                        GuardarFolio = False
                        Exit Function
                    End If
                End If
            Next
        End With

        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        GuardoFolioAdicional = False
        If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOADICIONAL)) = "" Then
            GeneraFolioAdicional()
        Else
            FolioAdicional = Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOADICIONAL))
        End If
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) = "TC" Then
                    If Trim(.get_TextMatrix(I, C_COLCODIGOARTICULO)) <> "" Then
                        Sql = "SELECT * FROM MovimientosVentasDet WHERE FolioVenta = '" & flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA) & "' AND NumPartida = " & flexDetalleVenta.get_TextMatrix(I, C_COLNUMPARTIDA)
                        ModEstandar.BorraCmd()
                        Cmd.CommandText = "dbo.Up_Select_Datos"
                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                        RsAux = Cmd.Execute
                        If RsAux.RecordCount > 0 Then
                            If (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) > 0 Then
                                ModStoredProcedures.PR_IE_MovimientosVentasDet(RsAux.Fields("FolioVenta").Value, CStr(RsAux.Fields("NumPartida").Value), "0", "", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "", .get_TextMatrix(I, C_COLNUEVADESCRIPCION), IIf(optManual.Checked = True, "M", "P"), "0", "", "", "", "0", "0", "", "", "0", "0", "0", "01/01/1900", "0", "0", "0", "0", "", C_MODIFICACION, CStr(9))
                                Cmd.Execute()

                                '''nuevo 19AGO2005
                                ModStoredProcedures.PR_IE_MovimientosVentasDet(RsAux.Fields("FolioVenta").Value, CStr(RsAux.Fields("NumPartida").Value), "0", "", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "", "", "", "0", FolioAdicional, "", "V", "0", "0", "", "", "0", "0", "0", "01/01/1900", "0", "0", "0", "0", "", C_MODIFICACION, CStr(11))
                                Cmd.Execute()

                                ModStoredProcedures.PR_IE_MovimientosVentasDet(RsAux.Fields("FolioVenta").Value, .get_TextMatrix(I, C_COLNUMPARTIDA), "0", "", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", .get_TextMatrix(I, C_ColPROMOCION), .get_TextMatrix(I, C_ColDESCUENTO), VB6.Format(.get_TextMatrix(I, C_COLPRECIOLISTAADICIONAL), "#####0.00"), .get_TextMatrix(I, C_ColPRECIOLISTASINIVA), .get_TextMatrix(I, C_COLPRECIOREAL), .get_TextMatrix(I, C_COLIVAREAL), "", .get_TextMatrix(I, C_COLNUEVADESCRIPCION), IIf(optManual.Checked, "M", "P"), .get_TextMatrix(I, C_COLPORCENTAJEADICIONAL), FolioAdicional, "", "V", CStr(CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))), flexVentas.get_TextMatrix(flexVentas.Row, C_ColTIPOCAMBIO), flexVentas.get_TextMatrix(flexVentas.Row, C_ColMONEDA), flexVentas.get_TextMatrix(flexVentas.Row, C_ColCONDICION), flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA), flexVentas.get_TextMatrix(flexVentas.Row, C_COLREDONDEOADICIONAL), flexVentas.get_TextMatrix(flexVentas.Row, C_COLANTICIPOADICIONAL), VB6.Format(Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFECHADEVENTA)), C_FORMATFECHAGUARDAR), flexVentas.get_TextMatrix(flexVentas.Row, C_ColCODSUCURSAL), flexVentas.get_TextMatrix(flexVentas.Row, C_COLCODCAJA), flexVentas.get_TextMatrix(flexVentas.Row, C_COLCODVENDEDOR), flexVentas.get_TextMatrix(flexVentas.Row, C_COLCODCLIENTE), "", C_MODIFICACION, CStr(1))
                                Cmd.Execute()

                                '''nuevo 19AGO2005
                            End If
                        End If
                    Else
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLCAMBIOS, "S")
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLFOLIOADICIONAL, FolioAdicional)
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLINCFACTURA, "S")
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLGRABADO, "S")

                        With flexVentas
                            If .get_TextMatrix(.Row, C_COLFOLIOEXCLUIDO) = "EXCLUIDO" Then
                                ModStoredProcedures.PR_IMEMovimientosVentasCab(.get_TextMatrix(.Row, C_COLFOLIOVENTA), "01/01/1900", "0", "0", "0", "0", "", "", "", "", "0", "0", "0", "0", "0", "0", "0", "0", "", "", "01/01/1900", "0", "0", "0", "0", "0", "0", "O", "", "01/01/1900", "", "0", "", CBool("0"), CBool("0"), C_MODIFICACION, CStr(3))
                                Cmd.Execute()
                                .set_TextMatrix(.Row, C_COLGRABADO, "S")
                                Me.Cursor = System.Windows.Forms.Cursors.Default
                                Cnn.CommitTrans()
                                blnTransaccion = False
                                MsgBox("El folio se ha excluido con éxito...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                Exit Function
                            End If
                            If CDbl(Numerico(.get_TextMatrix(.Row, C_COLSUBTOTALADICIONAL))) <> 0 And CDbl(Numerico(.get_TextMatrix(.Row, C_COLIVAADICIONAL))) <> 0 And CDbl(Numerico(.get_TextMatrix(.Row, C_COLTOTALADICIONAL))) <> 0 Then

                                ModStoredProcedures.PR_IMEMovimientosVentasCab(.get_TextMatrix(.Row, C_COLFOLIOVENTA), "01/01/1900", "0", "0", "0", "0", "", "", "", "", "0", "0", "0", "0", "0", "0", "0", "0", "", "", "01/01/1900", .get_TextMatrix(.Row, C_COLSUBTOTALADICIONAL), .get_TextMatrix(.Row, C_COLDESCUENTOADICIONAL), .get_TextMatrix(.Row, C_COLIVAADICIONAL), .get_TextMatrix(.Row, C_COLTOTALADICIONAL), .get_TextMatrix(.Row, C_COLREDONDEOADICIONAL), .get_TextMatrix(.Row, C_COLANTICIPOADICIONAL), "V", "", "01/01/1900", "", "0", "", CBool("0"), CBool("0"), C_MODIFICACION, CStr(1))
                                Cmd.Execute()
                                .set_TextMatrix(.Row, C_COLGRABADO, "S")
                                .set_TextMatrix(.Row, C_COLFOLIOADICIONAL, FolioAdicional)

                            End If
                        End With

                        Cnn.CommitTrans()
                        Me.Cursor = System.Windows.Forms.Cursors.Default
                        blnTransaccion = False
                        gblnCambiosAnalisis = True
                        GuardarFolio = True
                        Exit Function

                    End If
                ElseIf .get_TextMatrix(I, C_COLMODIFICADO) = "M" Then
                    ModStoredProcedures.PR_IE_MovimientosVentasDet(IIf(Trim(.get_TextMatrix(I, C_COLFOLIOAGREGADO)) = "", Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA)), .get_TextMatrix(I, C_COLFOLIOAGREGADO)), .get_TextMatrix(I, C_COLNUMPARTIDA), "0", "", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", .get_TextMatrix(I, C_ColPROMOCION), .get_TextMatrix(I, C_ColDESCUENTO), VB6.Format(.get_TextMatrix(I, C_COLPRECIOLISTAADICIONAL), "#####0.00"), .get_TextMatrix(I, C_ColPRECIOLISTASINIVA), .get_TextMatrix(I, C_COLPRECIOREAL), .get_TextMatrix(I, C_COLIVAREAL), "", .get_TextMatrix(I, C_COLNUEVADESCRIPCION), IIf(optManual.Checked, "M", "P"), .get_TextMatrix(I, C_COLPORCENTAJEADICIONAL), FolioAdicional, "", "V", CStr(CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))), flexVentas.get_TextMatrix(flexVentas.Row, C_ColTIPOCAMBIO), flexVentas.get_TextMatrix(flexVentas.Row, C_ColMONEDA), flexVentas.get_TextMatrix(flexVentas.Row, C_ColCONDICION), flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA), flexVentas.get_TextMatrix(flexVentas.Row, C_COLREDONDEOADICIONAL), flexVentas.get_TextMatrix(flexVentas.Row, C_COLANTICIPOADICIONAL), VB6.Format(Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFECHADEVENTA)), C_FORMATFECHAGUARDAR), flexVentas.get_TextMatrix(flexVentas.Row, C_ColCODSUCURSAL), flexVentas.get_TextMatrix(flexVentas.Row, C_COLCODCAJA), flexVentas.get_TextMatrix(flexVentas.Row, C_COLCODVENDEDOR), flexVentas.get_TextMatrix(flexVentas.Row, C_COLCODCLIENTE), "", C_MODIFICACION, CStr(1))
                    Cmd.Execute()
                    If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFACTURAPTOVTA)) = "" Then
                        gblnCambiosAnalisis = True
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLINCFACTURA, "S")
                    Else

                        GenerarFactPtoVenta = True
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLINCFACTURA, "N")
                    End If
                    GuardoFolioAdicional = True
                ElseIf .get_TextMatrix(I, C_COLEXCLUIDO) = "EXCLUIDO" Then
                    ModStoredProcedures.PR_IE_MovimientosVentasDet(Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA)), .get_TextMatrix(I, C_COLNUMPARTIDA), "0", "", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "", "", "", "0", "", "", "O", "0", "0", "", "", "0", "0", "0", "01/01/1900", "0", "0", "0", "0", "", C_MODIFICACION, CStr(6))
                    Cmd.Execute()
                End If
            Next
        End With
        With flexVentas
            If .get_TextMatrix(.Row, C_COLFOLIOEXCLUIDO) = "EXCLUIDO" Then
                ModStoredProcedures.PR_IMEMovimientosVentasCab(.get_TextMatrix(.Row, C_COLFOLIOVENTA), "01/01/1900", "0", "0", "0", "0", "", "", "", "", "0", "0", "0", "0", "0", "0", "0", "0", "", "", "01/01/1900", "0", "0", "0", "0", "0", "0", "O", "", "01/01/1900", "", "0", "", CBool("0"), CBool("0"), C_MODIFICACION, CStr(3))
                Cmd.Execute()
                .set_TextMatrix(.Row, C_COLGRABADO, "S")
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Cnn.CommitTrans()
                blnTransaccion = False
                MsgBox("El folio se ha excluido con exito...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Function
            End If
            If CDbl(Numerico(.get_TextMatrix(.Row, C_COLSUBTOTALADICIONAL))) <> 0 And CDbl(Numerico(.get_TextMatrix(.Row, C_COLIVAADICIONAL))) <> 0 And CDbl(Numerico(.get_TextMatrix(.Row, C_COLTOTALADICIONAL))) <> 0 Then
                If Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) <> "F" Or Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) <> "TC" Then
                    ModStoredProcedures.PR_IMEMovimientosVentasCab(.get_TextMatrix(.Row, C_COLFOLIOVENTA), "01/01/1900", "0", "0", "0", "0", "", "", "", "", "0", "0", "0", "0", "0", "0", "0", "0", "", "", "01/01/1900", .get_TextMatrix(.Row, C_COLSUBTOTALADICIONAL), .get_TextMatrix(.Row, C_COLDESCUENTOADICIONAL), .get_TextMatrix(.Row, C_COLIVAADICIONAL), .get_TextMatrix(.Row, C_COLTOTALADICIONAL), .get_TextMatrix(.Row, C_COLREDONDEOADICIONAL), .get_TextMatrix(.Row, C_COLANTICIPOADICIONAL), "V", "", "01/01/1900", "", "0", "", CBool("0"), CBool("0"), C_MODIFICACION, CStr(1))
                    Cmd.Execute()
                    .set_TextMatrix(.Row, C_COLGRABADO, "S")
                    .set_TextMatrix(.Row, C_COLFOLIOADICIONAL, FolioAdicional)
                End If
            End If
        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        GuardarFolio = True
        If GuardoFolioAdicional Then
            flexVentas.set_TextMatrix(flexVentas.Row, C_COLCAMBIOS, "S")
            '        MsgBox "Se ha Generado el folio de venta adicional " & FolioAdicional, vbOKOnly + vbInformation, gstrNombCortoEmpresa
            '        If flexVentas.TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) <> "F" And flexVentas.TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) <> "TC" And flexVentas.TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) <> "" And Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLFACTURAPTOVTA)) = "" Then
            'CalculaImportesFactura
            '        End If
        End If
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Sub BuscaSucursal()
        On Error GoTo Merr
        gStrSql = "SELECT DescAlmacen,TipoAlmacen FROM CatAlmacen WHERE CodAlmacen = " & txtCodSucursal.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("TipoAlmacen").Value = "V" Then
                MsgBox("Este Almacen No Es Un Almacen Propio, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Nuevo()
                txtCodSucursal.Focus()
                Exit Sub
            Else
                txtCodSucursal.Text = VB6.Format(txtCodSucursal.Text, "000")
                dbcSucursal.Text = RsGral.Fields("DescAlmacen").Value
                CargaVentas()
            End If
        Else
            MsgBox("Codigo de Almacen no Existe, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Nuevo()
            txtCodSucursal.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub CambiarFormatoTxtenCaptura()
        With txtFlex
            .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
            .MaxLength = 15
        End With
    End Sub

    'Sub CalculaImportesFactura()
    '    Dim Total As Double
    '    Dim TotalPesos As Double
    '    Dim RedondeoPesos As Double
    '    Dim RedondeoDolares As Double
    '    If Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLFACTURAPTOVTA)) = "S" Then Exit Sub
    '
    '    lblDescuento = CCur(Numerico(lblDescuento)) + Round(CCur(Numerico(flexVentas.TextMatrix(flexVentas.Row, C_COLDESCUENTOADICIONAL))) * CCur(Numerico(flexVentas.TextMatrix(flexVentas.Row, C_COLTIPOCAMBIO))), 1)
    '    lblIva = CCur(Numerico(lblIva)) + Round(CCur(Numerico(flexVentas.TextMatrix(flexVentas.Row, C_COLIVAADICIONAL))) * CCur(Numerico(flexVentas.TextMatrix(flexVentas.Row, C_COLTIPOCAMBIO))), 1)
    '    lblSubTot = CCur(Numerico(lblSubTot)) + Round((CCur(Numerico(flexVentas.TextMatrix(flexVentas.Row, C_COLSUBTOTALADICIONAL))) + CCur(Numerico(flexVentas.TextMatrix(flexVentas.Row, C_COLREDONDEOADICIONAL)))) * CCur(Numerico(flexVentas.TextMatrix(flexVentas.Row, C_COLTIPOCAMBIO))), 1)
    '    lblImporteSubTotal = Round((CDbl(Numerico(lblSubTot)) - CDbl(Numerico(lblDescuento))) + CDbl(Numerico(lblIva)), 2)
    '    lblDescuento = Round(CDbl(Numerico(lblDescuento)), 2)
    '    lblIva = Round(CDbl(Numerico(lblIva)), 2)
    '    lblSubTot = Round(CDbl(Numerico(lblSubTot)), 2)
    '    Total = lblImporteSubTotal
    '    TotalPesos = Round(Total * gcurCorpoTIPOCAMBIODOLAR, 4)
    '    RedondeoPesos = ModCorporativo.RedondeoUnidadFinal(CCur(TotalPesos), CDbl(gcurRedondeo))
    '    RedondeoDolares = Round(RedondeoPesos / gcurCorpoTIPOCAMBIODOLAR, 4)
    '    RedondeoFactura = RedondeoDolares
    '
    '    lblImporteRedondeo = Round(RedondeoFactura, 2)
    '    lblImporteTotal = CDbl(Numerico(lblImporteSubTotal)) + CDbl(Numerico(lblImporteRedondeo))
    '    lblImporteSubTotal = Format(lblImporteSubTotal, "###,##0.00")
    '    lblImporteRedondeo = Format(lblImporteRedondeo, "###,##0.00")
    '    lblImporteTotal = Format(lblImporteTotal, "###,##0.00")
    'End Sub

    'Sub RecalculaImportesFactura()
    '    Dim Total As Double
    '    Dim TotalPesos As Double
    '    Dim RedondeoPesos As Double
    '    Dim RedondeoDolares As Double
    '    Dim I As Integer
    '    lblSubTot = 0
    '    lblDescuento = 0
    '    lblIva = 0
    '    With flexVentas
    '        For I = 1 To .Rows - 1
    '            If Trim(.TextMatrix(I, C_COLFOLIOVENTA)) <> "" Then
    '                If Trim(.TextMatrix(I, C_COLFACTURAPTOVTA)) = "" And Trim(.TextMatrix(I, C_COLINCFACTURA)) = "S" Then
    '                    lblSubTot = CDbl(Numerico(lblSubTot)) + CDbl(Numerico(Format(.TextMatrix(I, C_COLSUBTOTALADICIONAL), "#####0.000")))
    '                    lblDescuento = CDbl(Numerico(lblDescuento)) + CDbl(Numerico(Format(.TextMatrix(I, C_COLDESCUENTOADICIONAL), "#####0.000")))
    '                    lblIva = CDbl(Numerico(lblIva)) + CDbl(Numerico(Format(.TextMatrix(I, C_COLIVAADICIONAL), "#####0.000")))
    '                End If
    '            End If
    '        Next
    '        lblImporteSubTotal = Round((CDbl(Numerico(lblSubTot)) - CDbl(Numerico(lblDescuento))) + CDbl(Numerico(lblIva)), 2)
    '        Total = lblImporteSubTotal
    '        TotalPesos = Round(Total * gcurCorpoTIPOCAMBIODOLAR, 4)
    '        RedondeoPesos = ModCorporativo.RedondeoUnidadFinal(CCur(TotalPesos), CDbl(gcurRedondeo))
    '        RedondeoDolares = Round(RedondeoPesos / gcurCorpoTIPOCAMBIODOLAR, 4)
    '        RedondeoFactura = RedondeoDolares
    '        lblImporteRedondeo = Round(RedondeoFactura, 2)
    '        lblImporteTotal = CDbl(Numerico(lblImporteSubTotal)) + CDbl(Numerico(lblImporteRedondeo))
    '        lblImporteSubTotal = Format(lblImporteSubTotal, "###,##0.00")
    '        lblImporteRedondeo = Format(lblImporteRedondeo, "###,##0.00")
    '        lblImporteTotal = Format(lblImporteTotal, "###,##0.00")
    '    End With
    'End Sub

    Sub CalculaImportes()
        Dim I As Integer
        Dim Total As Double
        Dim TotalPesos As Double
        Dim RedondeoDolares As Double
        Dim RedondeoPesos As Double
        Dim PorcentajeAdicional As Double
        lblSubTotal.Text = CStr(0)
        lblRedondeo.Text = CStr(0)
        lblTotal.Text = CStr(0)
        lblTotalPesos.Text = CStr(0)
        flexVentas.set_TextMatrix(flexVentas.Row, C_COLSUBTOTALADICIONAL, "")
        flexVentas.set_TextMatrix(flexVentas.Row, C_COLDESCUENTOADICIONAL, "")
        flexVentas.set_TextMatrix(flexVentas.Row, C_COLIVAADICIONAL, "")
        flexVentas.set_TextMatrix(flexVentas.Row, C_COLTOTALADICIONAL, "")
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO)) <> "" Then
                    .set_TextMatrix(I, C_COLNUEVOIMPORTESINDESCUENTO, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) * CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))), "###,##0.00"))
                    If CDbl(Numerico(.get_TextMatrix(I, C_COLPORCDESCTO))) <> 0 Then
                        .set_TextMatrix(I, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) * ((CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) * CDbl(1 - CDbl(VB6.Format(Numerico(CStr(CDbl(.get_TextMatrix(I, C_COLPORCDESCTO)) / 100)), "#####0.0000"))))), "###,##0.00"))
                    ElseIf CDbl(Numerico(.get_TextMatrix(I, C_COLPORCPROM))) <> 0 Then
                        .set_TextMatrix(I, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) * ((CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) * CDbl(1 - CDbl(VB6.Format(Numerico(CStr(CDbl(.get_TextMatrix(I, C_COLPORCPROM)) / 100)), "#####0.0000"))))), "###,##0.00"))
                    Else
                        .set_TextMatrix(I, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) * ((CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) * CDbl(1 - CDbl(VB6.Format(Numerico(CStr(CDbl(.get_TextMatrix(I, C_COLPORCDESCTO)) / 100)), "#####0.0000"))))), "###,##0.00"))
                    End If
                    .set_TextMatrix(I, C_ColPRECIOLISTASINIVA, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) / (1 + CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000"))), "#####0.0000"))
                    If CDbl(Numerico(.get_TextMatrix(I, C_COLPORCDESCTO))) <> 0 Then
                        .set_TextMatrix(I, C_ColDESCUENTO, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) * CDbl(VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPORCDESCTO))) / 100, "#####0.0000")) / (1 + CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000"))), "#####0.0000"))
                        .set_TextMatrix(I, C_COLIVAREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColDESCUENTO)))) * CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000")), "#####0.0000"))
                        .set_TextMatrix(I, C_COLPRECIOREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColDESCUENTO)))) + CDbl(Numerico(.get_TextMatrix(I, C_COLIVAREAL))), "#####0.0000"))
                    ElseIf CDbl(Numerico(.get_TextMatrix(I, C_COLPORCPROM))) <> 0 Then
                        .set_TextMatrix(I, C_ColPROMOCION, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) * CDbl(VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPORCPROM))) / 100, "#####0.0000")) / (1 + CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000"))), "#####0.0000"))
                        .set_TextMatrix(I, C_COLIVAREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColPROMOCION)))) * CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000")), "#####0.0000"))
                        .set_TextMatrix(I, C_COLPRECIOREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColPROMOCION)))) + CDbl(Numerico(.get_TextMatrix(I, C_COLIVAREAL))), "#####0.0000"))
                    Else
                        .set_TextMatrix(I, C_ColDESCUENTO, 0)
                        .set_TextMatrix(I, C_ColPROMOCION, 0)
                        .set_TextMatrix(I, C_COLIVAREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColDESCUENTO)))) * CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000")), "#####0.0000"))
                        .set_TextMatrix(I, C_COLPRECIOREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColDESCUENTO)))) + CDbl(Numerico(.get_TextMatrix(I, C_COLIVAREAL))), "#####0.0000"))
                    End If
                    If CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) <> 0 Then
                        .set_TextMatrix(I, C_COLPORCENTAJEADICIONAL, VB6.Format(CDbl(VB6.Format(1 - CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) / CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOPUBLICO))), "#####0.0000")) * 100, "#####0.00"))
                    Else
                        .set_TextMatrix(I, C_COLPORCENTAJEADICIONAL, "")
                    End If
                End If
                'If Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLFACTURAPTOVTA)) = "" Then
                lblSubTotal.Text = CStr(CDbl(Numerico(lblSubTotal.Text)) + (CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOREAL))) * (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL))))))
                'End If
                flexVentas.set_TextMatrix(flexVentas.Row, C_COLSUBTOTALADICIONAL, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLSUBTOTALADICIONAL))) + (CDec(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) * (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL))))))
                flexVentas.set_TextMatrix(flexVentas.Row, C_COLDESCUENTOADICIONAL, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLDESCUENTOADICIONAL))) + ((CDbl(Numerico(.get_TextMatrix(I, C_ColDESCUENTO))) + CDbl(Numerico(.get_TextMatrix(I, C_ColPROMOCION)))) * (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL))))))
                flexVentas.set_TextMatrix(flexVentas.Row, C_COLIVAADICIONAL, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLIVAADICIONAL))) + (CDec(Numerico(.get_TextMatrix(I, C_COLIVAREAL))) * (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL))))))
                flexVentas.set_TextMatrix(flexVentas.Row, C_COLTOTALADICIONAL, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLTOTALADICIONAL))) + (CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOREAL))) * (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL))))))
            Next
            PorcentajeAdicional = CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLTOTALADICIONAL))) / CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_ColIMPORTE))), "#####0.0000"))
            'If Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) <> "TC" And Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) <> "F" Then
            flexVentas.set_TextMatrix(flexVentas.Row, C_COLANTICIPOADICIONAL, VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLANTICIPO))) * PorcentajeAdicional, "#####0.0000"))
            'Else
            flexVentas.set_TextMatrix(flexVentas.Row, C_COLANTICIPOADICIONAL, flexVentas.get_TextMatrix(flexVentas.Row, C_COLANTICIPO))
            'End If
            If flexVentas.get_TextMatrix(flexVentas.Row, C_ColMONEDA) = "P" Then
                TotalPesos = CDbl(VB6.Format(CDbl(lblSubTotal.Text) * CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_ColTIPOCAMBIO))), "#####0.000000"))
                TotalPesos = CDbl(VB6.Format(TotalPesos, "#####0.00"))
                RedondeoPesos = ModCorporativo.RedondeoUnidadFinal(CDec(TotalPesos), CDbl(gcurRedondeo))

                RedondeoDolares = RedondeoPesos / CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_ColTIPOCAMBIO)))
                flexVentas.set_TextMatrix(flexVentas.Row, C_COLREDONDEOADICIONAL, VB6.Format(RedondeoDolares, "#####0.0000"))
                lblRedondeo.Text = VB6.Format(RedondeoDolares, "###,##0.00")
                lblTotal.Text = VB6.Format(CDbl(Numerico(lblSubTotal.Text)) + CDbl(Numerico(lblRedondeo.Text)), "###,##0.00")
                lblSubTotal.Text = VB6.Format(lblSubTotal.Text, "###,##0.00")
                lblTotalPesos.Text = VB6.Format(TotalPesos + RedondeoPesos, "#####0.0")
                lblTotalPesos.Text = VB6.Format(lblTotalPesos.Text, "###,##0.00")
            ElseIf flexVentas.get_TextMatrix(flexVentas.Row, C_ColMONEDA) = "D" Then
                RedondeoDolares = ModCorporativo.RedondeoUnidadFinal(CDec(lblSubTotal.Text), 1)
                flexVentas.set_TextMatrix(flexVentas.Row, C_COLREDONDEOADICIONAL, VB6.Format(RedondeoDolares, "#####0.0000"))
                lblRedondeo.Text = VB6.Format(RedondeoDolares, "###,##0.00")
                lblTotal.Text = VB6.Format(CDbl(Numerico(lblSubTotal.Text)) + CDbl(Numerico(lblRedondeo.Text)), "###,##0.00")
                lblSubTotal.Text = VB6.Format(lblSubTotal.Text, "###,##0.00")
                lblTotalPesos.Text = VB6.Format(CDbl((lblTotal).Text) * CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_ColTIPOCAMBIO))), "#####0.0")
                lblTotalPesos.Text = VB6.Format(lblTotalPesos.Text, "###,##0.00")
            End If
            '''        'RedondeoDolares = Format(RedondeoPesos / Numerico(flexVentas.TextMatrix(flexVentas.Row, C_COLTIPOCAMBIO)), "#####0.0000")
            '''        RedondeoDolares = RedondeoPesos / Numerico(flexVentas.TextMatrix(flexVentas.Row, C_ColTIPOCAMBIO))
            '''        flexVentas.TextMatrix(flexVentas.Row, C_COLREDONDEOADICIONAL) = Format(RedondeoDolares, "#####0.0000")
            '''        'If Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLFACTURAPTOVTA)) = "" Then
            '''            lblRedondeo = Format(RedondeoDolares, "###,##0.00")
            '''            'lblRedondeo = RedondeoDolares
            '''            lblTotal = Format(CDbl(Numerico(lblSubTotal)) + CDbl(Numerico(lblRedondeo)), "###,##0.00")
            '''            lblSubTotal = Format(lblSubTotal, "###,##0.00")
            '''            'lblTotalPesos = Format(Numerico(lblTotal) * flexVentas.TextMatrix(flexVentas.Row, C_COLTIPOCAMBIO), "###,##0.0")
            '''            lblTotalPesos = Format(TotalPesos + RedondeoPesos, "#####0.0")
            '''            lblTotalPesos = Format(lblTotalPesos, "###,##0.00")
            '''        'Else
            '''        '    lblRedondeo = "0.00"
            '''        '    lblTotal = "0.00"
            '''        '    lblSubTotal = "0.00"
            '''        '    lblTotalPesos = "0.00"
            '''        '    lblTotalPesos = "0.00"
            '''        'End If
        End With
    End Sub

    Sub CuentaPartidas()
        Dim I As Integer
        Dim J As Integer
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODIGOARTICULO)) = "" Then Exit Sub
                numPartidas = numPartidas + 1
                If numPartidas > 10 Then
                    MsgBox("La factura no puede tener mas de 10 partidas, no es posible incluir este folio en la partida.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    For J = 0 To 1
                        flexVentas.Col = J
                        flexVentas.CellBackColor = System.Drawing.ColorTranslator.FromOle(&H8000000E)
                        '                    .CellForeColor = &H8000000E
                    Next
                    flexVentas.Col = 0
                    flexVentas.set_TextMatrix(flexVentas.Row, C_COLFACTURAPTOVTA, "")
                    ColorAnteFolio = "Bln"
                    DesCuentaPartidas()
                    flexVentas.Focus()
                End If
            Next
        End With
    End Sub

    Sub DesCuentaPartidas()
        Dim I As Integer
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODIGOARTICULO)) = "" Then Exit Sub
                numPartidas = numPartidas - 1
            Next
        End With
    End Sub

    Sub EncabezadoVentas()
        With flexVentas
            .Row = 0
            .Col = C_COLFOLIOVENTA
            .set_ColWidth(C_COLFOLIOVENTA, 0, 1650)
            .set_ColWidth(C_ColIMPORTE, 0, 1450)
            .set_ColWidth(C_ColMONEDA, 0, 0)
            .set_ColWidth(C_COLESTADOFOLIO, 0, 0)
            .set_ColWidth(C_COLPORCIVA, 0, 0)
            .set_ColWidth(C_ColTIPOCAMBIO, 0, 0)
            .set_ColWidth(C_COLANTICIPO, 0, 0)
            .set_ColWidth(C_COLFOLIOADICIONAL, 0, 0)
            .set_ColWidth(C_COLMETODO, 0, 0)
            .set_ColWidth(C_COLSUBTOTALADICIONAL, 0, 0)
            .set_ColWidth(C_COLDESCUENTOADICIONAL, 0, 0)
            .set_ColWidth(C_COLIVAADICIONAL, 0, 0)
            .set_ColWidth(C_COLTOTALADICIONAL, 0, 0)
            .set_ColWidth(C_COLREDONDEOADICIONAL, 0, 0)
            .set_ColWidth(C_COLANTICIPOADICIONAL, 0, 0)
            .set_ColWidth(C_COLESTATUSADICIONAL, 0, 0)
            .set_ColWidth(C_COLFOLIOFACTURA, 0, 0)
            .set_ColWidth(C_COLGRABADO, 0, 0)
            .set_ColWidth(C_COLFOLIOEXCLUIDO, 0, 0)
            .set_ColWidth(C_COLINCLUIRFACTURA, 0, 0)
            .set_ColWidth(C_ColCONDICION, 0, 0)
            .set_ColWidth(C_COLFECHADEVENTA, 0, 0)
            .set_ColWidth(C_ColCODSUCURSAL, 0, 0)
            .set_ColWidth(C_COLCODCAJA, 0, 0)
            .set_ColWidth(C_COLCODVENDEDOR, 0, 0)
            .set_ColWidth(C_COLCODCLIENTE, 0, 0)
            .set_ColWidth(C_COLFACTURAPTOVTA, 0, 0)
            .set_ColWidth(C_COLINCFACTURA, 0, 0)
            .set_ColWidth(C_COLTIENEDEVOLUCION, 0, 0)
            .set_ColWidth(C_COLCAMBIOS, 0, 0)
            .set_ColWidth(C_COLTIPOMOVTO, 0, 0)
            .Col = C_COLFOLIOVENTA
            .Row = 0
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Folio"
            .Col = C_ColIMPORTE
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Importe"
            .Row = 1
            .Col = C_COLFOLIOVENTA
            .Rows = 12
        End With
    End Sub

    Sub EncabezadoDetalle()
        With flexDetalleVenta
            .set_ColWidth(C_COLCODIGOARTICULO, 0, 1000)
            .set_ColWidth(C_COLCANTIDADDEVOL, 0, 500)
            .set_ColWidth(C_COLCANTIDAD, 0, 500)
            .set_ColWidth(C_COLPORCDESCUENTO, 0, 800)
            .set_ColWidth(C_COLPORCPROMOCION, 0, 800)
            .set_ColWidth(C_ColPRECIOPUBLICO, 0, 1250)
            .set_ColWidth(C_COLIMPORTESINDESCUENTO, 0, 1255)
            .set_ColWidth(C_COLIMPORTECONDESCUENTO, 0, 1250)
            .set_ColWidth(C_COLNUEVADESCRIPCION, 0, 0)
            .set_ColWidth(C_COLNUEVOPRECIOPUBLICO, 0, 1250)
            .set_ColWidth(C_COLNUEVOIMPORTESINDESCUENTO, 0, 1255)
            .set_ColWidth(C_COLNUEVOIMPORTECONDESCUENTO, 0, 1250)
            .set_ColWidth(C_COLDESCRIPCION, 0, 0)
            .set_ColWidth(C_ColPRECIOLISTASINIVA, 0, 0)
            .set_ColWidth(C_COLPRECIOREAL, 0, 0)
            .set_ColWidth(C_COLIVAREAL, 0, 0)
            .set_ColWidth(C_COLMODIFICADO, 0, 0)
            .set_ColWidth(C_COLPORCDESCTO, 0, 0)
            .set_ColWidth(C_COLPORCPROM, 0, 0)
            .set_ColWidth(C_ColDESCUENTO, 0, 0)
            .set_ColWidth(C_ColPROMOCION, 0, 0)
            .set_ColWidth(C_COLFOLIOAGREGADO, 0, 0)
            .set_ColWidth(C_COLFECHAFOLIOAGREGADO, 0, 0)
            .set_ColWidth(C_COLEXCLUIDO, 0, 0)
            .set_ColWidth(C_COLDESCRIPCIONFAMILIA, 0, 0)
            .set_ColWidth(C_COLNUMPARTIDA, 0, 0)
            .set_ColWidth(C_COLALMACEN, 0, 0)
            .set_ColWidth(C_COLPRECIOLISTAADICIONAL, 0, 0)
            .set_ColWidth(C_COLPORCENTAJEADICIONAL, 0, 0)
            .set_ColWidth(C_COLGRAB, 0, 0)
            .set_ColWidth(C_COLMODIFICADOTAG, 0, 0)
            .Row = 0

            .Col = 0
            .CellFontBold = True
            .Text = "Detalle de Ventas"
            .Col = 1
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas"
            .Col = 2
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas"
            .Col = 3
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas"
            .Col = 4
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas"
            .Col = 5
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas"
            .Col = 6
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas"
            .Col = 7
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas"
            .Col = 8
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "A N A L I S I S"
            .Col = 9
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "A N A L I S I S"
            .Col = 10
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "A N A L I S I S"
            .Col = 11
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "A N A L I S I S"
            '.MergeCells = flexMergeFree
            .set_MergeRow(0, True)
            .Row = 1
            .Col = C_COLCODIGOARTICULO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Articulo"
            .Col = C_COLCANTIDADDEVOL
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Dev"
            .Col = C_COLCANTIDAD
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Cant"
            .Col = C_COLPORCDESCUENTO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Descto"
            .Col = C_COLPORCPROMOCION
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Prom"
            .Col = C_ColPRECIOPUBLICO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Precio Pub"
            .Col = C_COLIMPORTESINDESCUENTO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Importe"
            .Col = C_COLIMPORTECONDESCUENTO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Impte-Descto"
            .Col = C_COLNUEVOPRECIOPUBLICO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Precio Pub"
            .Col = C_COLNUEVOIMPORTESINDESCUENTO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Importe"
            .Col = C_COLNUEVOIMPORTECONDESCUENTO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Impte-Descto"
            .Rows = 12
            .Row = 2
            .Col = 0
        End With
    End Sub

    Sub EncabezadoVentasPendientes()
        With flexVentasPendientes
            .set_Cols(0, 15)
            .set_ColWidth(C_ColFOLIO, 0, 1650)
            .set_ColWidth(C_COLFECHAVENTA, 0, 1200)
            .set_ColWidth(C_COLCODARTICULO, 0, 1000)
            .set_ColWidth(C_COLDESCARTICULO, 0, 0)
            .set_ColWidth(C_COLDEVOL, 0, 500)
            .set_ColWidth(C_COLCANT, 0, 500)
            .set_ColWidth(C_COLDESCTO, 0, 800)
            .set_ColWidth(C_COLPROM, 0, 800)
            .set_ColWidth(C_COLPRECIOPUB, 0, 1350)
            .set_ColWidth(C_COLIMPTE, 0, 1350)
            .set_ColWidth(C_COLIMPTEDESCTO, 0, 1350)
            .set_ColWidth(C_COLPORDESCTO, 0, 0)
            .set_ColWidth(C_COLPORPROM, 0, 0)
            .set_ColWidth(C_COLDESCFAMILIA, 0, 0)
            .set_ColWidth(C_COLNPARTIDA, 0, 0)
            .set_ColAlignment(C_COLFECHAVENTA, 4)
            .set_ColAlignment(C_COLCODARTICULO, 7)
            .set_ColAlignment(C_COLDESCARTICULO, 7)
            .set_ColAlignment(C_COLDESCTO, 7)
            .set_ColAlignment(C_COLPROM, 7)
            .set_ColAlignment(C_COLPRECIOPUB, 7)
            .set_ColAlignment(C_COLIMPTE, 7)
            .set_ColAlignment(C_COLIMPTEDESCTO, 7)
            .Row = 0
            .Col = 0
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            .Col = 1
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            .Col = 2
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            .Col = 3
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            .Col = 4
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            .Col = 5
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            .Col = 6
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            .Col = 7
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            .Col = 8
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            .Col = 9
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            .Col = 10
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Detalle de Ventas Pendientes"
            '.MergeCells = MSHierarchicalFlexGridLib.MergeCellsSettings.flexMergeFree
            .set_MergeRow(0, True)
            .Row = 1
            .Col = C_ColFOLIO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Folio Venta"
            .Col = C_COLFECHAVENTA
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Fecha Venta"
            .Col = C_COLCODARTICULO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Articulo"
            .Col = C_COLDEVOL
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Dev"
            .Col = C_COLCANT
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Cant"
            .Col = C_COLDESCTO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Descto"
            .Col = C_COLPROM
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Prom"
            .Col = C_COLPRECIOPUB
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Precio Pub"
            .Col = C_COLIMPTE
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Importe"
            .Col = C_COLIMPTEDESCTO
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Impte-Descto"
            .Row = 2
            .Col = 0
        End With
    End Sub

    Function HayCambios() As Boolean
        Dim I As Integer
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If (Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) = "TC" Or Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) = "F") And Trim(.get_TextMatrix(.Row, C_COLMODIFICADO)) = "M" Then
                    HayCambios = True
                    Exit Function
                End If
                If Trim(.get_TextMatrix(I, C_COLMODIFICADO)) = "M" Then
                    HayCambios = True
                    Exit Function
                End If
                If Trim(.get_TextMatrix(I, C_COLEXCLUIDO)) = "EXCLUIDO" And Trim(.get_TextMatrix(I, C_COLGRAB)) = "" Then
                    HayCambios = True
                    Exit Function
                End If
            Next
        End With
    End Function

    Sub InicializaVariables()
        mblnSalir = False
        mblnNuevo = False
        FueraChange = False
        gblnCambiosAnalisis = False
        mblnCambiarMetodo = False
        mblnFactura = False
        blnFueraCell = False
        intCodSucursal = 0
        tecla = 0
        Fecha = ""
        CodSucursal = 0
        DescSucursal = ""
        ColorAnte = ""
        ColorAnteFolio = ""
        FolioAdicional = ""
        gintCodRFC = 0
        gstrNombreCliente = ""
        gstrRFCCliente = ""
        numPartidas = 0
        GenerarFactPtoVenta = False
        DesgloseIva = 0
    End Sub

    Sub Limpiar()
        If FoliosPendientes() Then
            MsgBox("No es posible limpiar la pantalla, ya ha generado algun(s) folios adicionales" & vbNewLine & "  Para poder limpiar la pantalla debera generar la factura correspondiente", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        Nuevo()
        dtpFechaVenta.Focus()
    End Sub

    Sub InicializaImporte()
        Dim I As Integer
        lblSubTotal.Text = "0.00"
        lblRedondeo.Text = "0.00"
        lblTotal.Text = "0.00"
        lblTotalPesos.Text = "0.00"
        flexVentas.set_TextMatrix(flexVentas.Row, C_COLSUBTOTALADICIONAL, "")
        flexVentas.set_TextMatrix(flexVentas.Row, C_COLDESCUENTOADICIONAL, "")
        flexVentas.set_TextMatrix(flexVentas.Row, C_COLIVAADICIONAL, "")
        flexVentas.set_TextMatrix(flexVentas.Row, C_COLTOTALADICIONAL, "")
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODIGOARTICULO)) <> "" And Trim(.get_TextMatrix(I, C_COLGRAB)) = "" Then
                    .set_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO, "0.00")
                    .set_TextMatrix(I, C_COLNUEVOIMPORTESINDESCUENTO, "0.00")
                    .set_TextMatrix(I, C_COLNUEVOIMPORTECONDESCUENTO, "0.00")
                    .set_TextMatrix(I, C_ColPRECIOLISTASINIVA, "")
                    .set_TextMatrix(I, C_COLPRECIOREAL, "")
                    .set_TextMatrix(I, C_COLIVAREAL, "")
                    .set_TextMatrix(I, C_COLMODIFICADO, "")
                End If
            Next
        End With
    End Sub

    Sub NuevaFactura()
        dtpFechaVenta.Value = Today
        FueraChange = True
        txtCodSucursal.Text = ""
        FueraChange = False
        FueraChange = True
        dbcSucursal.Text = ""
        FueraChange = False
        optManual.Checked = True
        optPorcentual.Checked = False
        txtPorcentaje.Text = "0"
        txtPorcentaje.Enabled = False
        LimpiarGridVentas()
        LimpiarGridDetalle()
        LimpiarGridPendientes()
        FueraChange = True
        txtDescripcion.Text = ""
        FueraChange = False
        txtDescripcion.ReadOnly = True
        cmdGenerarFactura.Enabled = True
        cmdImpresionTickets.Enabled = False
        cmdImprimirFactura.Enabled = False
        chkDoctoCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkDoctoCliente.Enabled = True
        lblSubTotal.Text = "0.00"
        lblRedondeo.Text = "0.00"
        lblTotal.Text = "0.00"
        lblTotalPesos.Text = "0.00"
        lblImporteSubTotal.Text = "0.00"
        lblImporteRedondeo.Text = "0.00"
        lblImporteTotal.Text = "0.00"
        lblDescuento.Text = "0.00"
        lblIva.Text = "0.00"
        lblSubTot.Text = "0.00"
        lblFactura.Text = "F00" & (Year(dtpFechaVenta.Value) + "0000") & (Month(dtpFechaVenta.Value) + "00") & ((dtpFechaVenta.Value) + "00") & "000000"
        chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked
        InicializaVariables()
        dtpFechaVenta.Enabled = True
        dbcSucursal.Enabled = True
        DeterminaRango()
    End Sub

    Sub DeterminaRango()
        'If dtpFechaVenta.Value = 1 Then
        '    DtpDesde.Value = DateAdd(Microsoft.VisualBasic.DateInterval.Month, -1, CDate(dtpFechaVenta.Value))
        'Else
        '    DtpDesde.Value = (DateSerial(Year(CDate(dtpFechaVenta.Value)), Month(CDate(dtpFechaVenta.Value)), 1) + "dd/MMM/yyyy")
        'End If
        DtpDesde.Tag = DtpDesde.Value
        DtpHasta.Value = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(dtpFechaVenta.Value))
        DtpHasta.Tag = DtpHasta.Value
    End Sub

    Sub Nuevo()
        dtpFechaVenta.Value = Today
        FueraChange = True
        txtCodSucursal.Text = ""
        dbcSucursal.Text = ""
        optManual.Checked = True
        optPorcentual.Checked = False
        txtPorcentaje.Text = "0"
        txtPorcentaje.Enabled = False
        txtFolioFactura.Text = ""
        FueraChange = False
        dtpFechaRegistro.Value = Today
        lblDescripcion.Text = ""
        lblEstadoFolio.Text = ""
        lblMoneda.Text = ""
        lblSubTotal.Text = "0.00"
        lblRedondeo.Text = "0.00"
        lblTotal.Text = "0.00"
        lblTotalPesos.Text = "0.00"
        lblImporteSubTotal.Text = "0.00"
        lblImporteRedondeo.Text = "0.00"
        lblImporteTotal.Text = "0.00"
        lblDescuento.Text = "0.00"
        lblIva.Text = "0.00"
        lblSubTot.Text = "0.00"
        lblFactura.Text = "F00" & (Year(dtpFechaVenta.Value) + "0000") & (Month(dtpFechaVenta.Value) + "00") & ((dtpFechaVenta.Value) + "00") & "000000"
        lblCantidad.Text = ""
        LimpiarGridVentas()
        LimpiarGridDetalle()
        LimpiarGridPendientes()
        FueraChange = True
        txtDescripcion.Text = ""
        FueraChange = False
        txtDescripcion.ReadOnly = True
        cmdGenerarFactura.Enabled = True
        cmdImpresionTickets.Enabled = False
        cmdImprimirFactura.Enabled = False
        txtFacturaAdicional.Text = ""
        InicializaVariables()
        chkDoctoCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkDoctoCliente.Enabled = True
        chkDesglosarIva.Enabled = True
        chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked
        cmdDatosFiscales.Enabled = False
        dtpFechaVenta.Enabled = True
        dbcSucursal.Enabled = True
        DeterminaRango()
    End Sub

    Sub LimpiarGridVentas()
        Dim I As Integer
        With flexVentas
            .Rows = 2
            .Rows = 12
            .RemoveItem(1)
        End With
    End Sub

    Sub LimpiarGridDetalle()
        With flexDetalleVenta
            .Rows = 3
            .Rows = 13
            .RemoveItem(2)
        End With
    End Sub

    Sub LimpiarGridPendientes()
        With flexVentasPendientes
            .Rows = 3
            .Rows = 13
            .RemoveItem(2)
        End With
    End Sub

    Sub CargaVentas()
        Dim RsAux As ADODB.Recordset
        Dim Descuento As Double
        Dim SubTotal As Double
        Dim Iva As Double
        Dim Total As Double
        Dim TotalPesos As Double
        Dim RedondeoDolares As Double
        Dim RedondeoPesos As Double
        Dim Sql As String
        On Error GoTo Err_Renamed

        gStrSql = "SELECT VtaCab.FolioVenta,VtaCab.Total,VtaCab.Redondeo,VtaCab.Moneda,VtaCab.FolioFactura,VtaCab.PorcIva,VtaCab.TipoCambio,VtaCab.Anticipo," & "SUM(CASE WHEN ISNULL(FP.EsTarjeta,0) = 0 THEN 0 ELSE 1 END) AS EsTarjeta,Vtacab.Cantidad,ISNULL(VtaCab.CantidadDevol,0) AS CantidadDevol,VtaCab.TotalAdicional,VtaCab.RedondeoAdicional,VtaCab.EstatusAdicional,VtaCab.Descuento,VtaCab.Iva,VtaCab.SubTotal,VtaCab.Condicion,VtaCab.CodSucursal,VtaCab.CodCaja,VtaCab.CodVendedor,VtaCab.CodCliente,VtaCab.TipoMovto,SUM(CASE WHEN ISNULL(FP.EsTarjeta,0) = 1 AND ISNULL(FP.EsDolar,0) = 0 THEN Ing.Importe/Ing.TipoCambio WHEN ISNULL(FP.EsTarjeta,0) = 1 AND ISNULL(FP.EsDolar,0) = 1 THEN Ing.Importe ELSE 0 END) AS IngTarjeta " & "FROM ((SELECT VtaCab.FolioVenta,VtacAB.Total,VtaCab.Redondeo,VTACAB.Moneda,VTACAB.FolioFactura,VTACAB.PorcIva,VTACAB.TipoCambio,VTACAB.Anticipo,SUM(VtaDet.Cantidad) as Cantidad,Dev.CantidadDevol,VtaCab.TotalAdicional,VtaCab.RedondeoAdicional,VtaCab.EstatusAdicional,VtaCab.Descuento,VtaCab.Iva,VtaCab.SubTotal,VtaCab.Condicion,VtaCab.CodSucursal,VtaCab.CodCaja,VtaCab.CodVendedor,VtaCab.CodCliente,VtaCab.TipoMovto " & "FROM MovimientosVentasCab VtaCab INNER JOIN MovimientosVentasDet VtaDet ON VtaCab.FolioVenta = VtaDet.FolioVenta LEFT OUTER JOIN (select devcab.folioventa,SUM(DEVDET.cantidaddevol) AS CANTIDADDEVOL from devolucionescab devcab inner join devolucionesdet devdet ON DEVCAB.FOLIODEVOLUCION = DEVDET.FOLIODEVOLUCION where devcab.estatus <> 'C' GROUP BY DEVCAB.FOLIOVENTA) DEV ON VTACAB.FOLIOVENTA = DEV.FOLIOVENTA " & "WHERE vtacab.Estatus <> 'C' AND vtacab.TipoMovto = 'V' AND vtacab.FechaVenta = '" & VB6.Format(dtpFechaVenta.Value, "mm/dd/yyyy") & "' AND vtacab.CodSucursal = " & Numerico(txtCodSucursal.Text) & " GROUP BY VtaCab.FolioVenta,VtaCab.Total,VtaCab.Redondeo,VtaCab.Moneda,VtaCab.FolioFactura,VtaCab.PorcIva,VtaCab.TipoCambio,VtaCab.Anticipo,DEV.CANTIDADDEVOL,VtaCab.TotalAdicional,VtaCab.RedondeoAdicional,VtaCab.EstatusAdicional,VtaCab.Descuento,VtaCab.Iva,VtaCab.SubTotal,VtaCab.Condicion,VtaCab.CodSucursal,VtaCab.CodCaja,VtaCab.CodVendedor,VtaCab.CodCliente,VtaCab.TipoMovto) " & "UNION " & "(SELECT APT.FolioVenta,VtaCab.Total,VtaCab.Redondeo,VtaCab.Moneda,VtaCab.FolioFactura,VtaCab.PorcIva,VtaCab.TipoCambio,VtaCab.Anticipo,sum(APT.Cantidad) as cantidad,ISNULL(DEV.CANTIDADDEVOL,0) AS CantidadDevol,VtaCab.TotalAdicional,VtaCab.RedondeoAdicional,VtaCab.EstatusAdicional,VtaCab.Descuento,VtaCab.Iva,VtaCab.SubTotal,VtaCab.Condicion,VtaCab.CodSucursal,VtaCab.CodCaja,VtaCab.CodVendedor,VtaCab.CodCliente,VtaCab.TipoMovto " & "FROM (SELECT * FROM DBO.VW_APARTADOS WHERE FECHAVTAAP = '" & VB6.Format(dtpFechaVenta.Value, "mm/dd/yyyy") & "' AND ESTATUSAPT = 'S' AND ESTATUS <> 'C') APT " & "Inner Join MovimientosVentasCab VtaCab ON APT.FolioVenta = VtaCab.FolioVenta LEFT OUTER JOIN (select devcab.folioventa,SUM(DEVDET.cantidaddevol) AS CANTIDADDEVOL from devolucionescab devcab " & "inner join devolucionesdet devdet ON DEVCAB.FOLIODEVOLUCION = DEVDET.FOLIODEVOLUCION where devcab.estatus <> 'C' " & "GROUP BY DEVCAB.FOLIOVENTA) DEV ON APT.FolioVenta = DEV.FOLIOVENTA AND VtaCab.FolioVenta = DEV.FolioVenta " & "WHERE VtaCab.TipoMovto = 'A' AND VtaCab.Estatus <> 'C' AND VtaCab.CodSucursal = " & Numerico(txtCodSucursal.Text) & " " & "GROUP BY APT.FolioVenta,/*VtaCab.FechaVenta,*/VtaCab.Total,VtaCab.Redondeo,VtaCab.Moneda,VtaCab.FolioFactura,VtaCab.PorcIva,VtaCab.TipoCambio,VtaCab.Anticipo,DEV.CANTIDADDEVOL,VtaCab.TotalAdicional,VtaCab.RedondeoAdicional,VtaCab.EstatusAdicional,VtaCab.Descuento,VtaCab.Iva,VtaCab.SubTotal,VtaCab.Condicion,VtaCab.CodSucursal,VtaCab.CodCaja,VtaCab.CodVendedor,VtaCab.CodCliente,VtaCab.TipoMovto)) VtaCab " & "Left Outer Join IngresosFormaDePago Ing ON VtaCab.FolioVenta = Ing.FolioMovto " & "Left Outer Join CatFormasPago FP ON Ing.CodFormaPago = FP.CodFormaPago " & "/*WHERE VtaCab.FechaVenta = '" & VB6.Format(dtpFechaVenta.Value, "mm/dd/yyyy") & "' AND VtaCab.CodSucursal = " & Numerico(txtCodSucursal.Text) & " AND VtaCab.Estatus <> 'C' AND VtaCab.TipoMovto = 'V' " & "AND VtaCab.FolioAdicional = ''*/" & "WHERE (VTACAB.CANTIDAD - ISNULL(VTACAB.CANTIDADDEVOL,0)) > 0 " & "GROUP BY VtaCab.FolioVenta,VtaCab.Total,VtaCab.Redondeo,VtaCab.Moneda,VtaCab.FolioFactura,VtaCab.PorcIva,VtaCab.TipoCambio,VtaCab.Anticipo,Vtacab.Cantidad,VtaCab.CantidadDevol,VtaCab.TotalAdicional,VtaCab.RedondeoAdicional,VtaCab.EstatusAdicional,VtaCab.Descuento,VtaCab.Iva,VtaCab.SubTotal,VtaCab.Condicion,VtaCab.CodSucursal,VtaCab.CodCaja,VtaCab.CodVendedor,VtaCab.CodCliente,VtaCab.TipoMovto"

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))

        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            lblImporteTotal.Text = "0.00"
            lblImporteSubTotal.Text = "0.00"
            lblImporteRedondeo.Text = "0.00"
            lblDescuento.Text = "0.00"
            lblSubTot.Text = "0.00"
            lblIva.Text = "0.00"
            With flexVentas
                If RsGral.RecordCount > 10 Then
                    .Rows = RsGral.RecordCount + 2
                End If
                .Row = 1
                Do While Not RsGral.EOF
                    .set_TextMatrix(.Row, C_COLFOLIOVENTA, Trim(RsGral.Fields("FolioVenta").Value))
                    .set_TextMatrix(.Row, C_ColIMPORTE, VB6.Format(RsGral.Fields("Total").Value, "###,##0.00"))
                    .set_TextMatrix(.Row, C_ColMONEDA, RsGral.Fields("Moneda").Value)
                    .set_TextMatrix(.Row, C_COLTIPOMOVTO, Trim(RsGral.Fields("TipoMovto").Value))
                    If Trim(RsGral.Fields("FolioFactura").Value) <> "" Then
                        .set_TextMatrix(.Row, C_COLESTADOFOLIO, "F")
                    ElseIf RsGral.Fields("EsTarjeta").Value > 0 Then  'And RsGral!Condicion = "CO" And RsGral!TipoMovto = "V"
                        If System.Math.Round(RsGral.Fields("IngTarjeta").Value, 1) >= System.Math.Round(RsGral.Fields("Total").Value + RsGral.Fields("Redondeo").Value, 1) Then
                            .set_TextMatrix(.Row, C_COLESTADOFOLIO, "TC")
                            .set_TextMatrix(.Row, C_COLINCFACTURA, "S")
                            If RsGral.Fields("CantidadDevol").Value > 0 Then
                                .set_TextMatrix(.Row, C_COLTIENEDEVOLUCION, "S")
                            End If
                        Else
                            .set_TextMatrix(.Row, C_COLESTADOFOLIO, "N")
                        End If
                    Else
                        .set_TextMatrix(.Row, C_COLESTADOFOLIO, "N")
                    End If
                    If .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "N" Then
                        Sql = "SELECT DISTINCT FolioAdicional FROM MovimientosVentasDet " & "WHERE FolioVenta = '" & Trim(.get_TextMatrix(.Row, C_COLFOLIOVENTA)) & "' AND FolioAdicional <> ''"
                        ModEstandar.BorraCmd()
                        Cmd.CommandText = "dbo.Up_Select_Datos"
                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                        RsAux = Cmd.Execute
                        If RsAux.RecordCount > 0 Then
                            flexVentas.set_TextMatrix(flexVentas.Row, C_COLFOLIOADICIONAL, Trim(RsAux.Fields("FolioAdicional").Value))
                        End If
                    End If
                    .set_TextMatrix(.Row, C_COLPORCIVA, RsGral.Fields("PorcIva").Value)
                    .set_TextMatrix(.Row, C_ColTIPOCAMBIO, RsGral.Fields("TipoCambio").Value)
                    .set_TextMatrix(.Row, C_COLANTICIPO, RsGral.Fields("Anticipo").Value)
                    .set_TextMatrix(.Row, C_ColCONDICION, RsGral.Fields("Condicion").Value)
                    .set_TextMatrix(.Row, C_COLFECHADEVENTA, dtpFechaVenta.Value)
                    .set_TextMatrix(.Row, C_ColCODSUCURSAL, RsGral.Fields("CodSucursal").Value)
                    .set_TextMatrix(.Row, C_COLCODCAJA, RsGral.Fields("CodCaja").Value)
                    .set_TextMatrix(.Row, C_COLCODVENDEDOR, RsGral.Fields("CodVendedor").Value)
                    .set_TextMatrix(.Row, C_COLCODCLIENTE, RsGral.Fields("CodCliente").Value)

                    If (Trim(RsGral.Fields("EstatusAdicional").Value) <> "" And Trim(RsGral.Fields("EstatusAdicional").Value) <> "O") Then
                        .set_TextMatrix(.Row, C_COLTOTALADICIONAL, RsGral.Fields("TotalAdicional").Value)
                        .set_TextMatrix(.Row, C_COLREDONDEOADICIONAL, RsGral.Fields("RedondeoAdicional").Value)
                        .set_TextMatrix(.Row, C_COLGRABADO, "S")
                        .set_TextMatrix(.Row, C_COLESTATUSADICIONAL, Trim(RsGral.Fields("EstatusAdicional").Value))
                    ElseIf (Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) = "TC" Or Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) = "F") And Trim(RsGral.Fields("EstatusAdicional").Value) = "" Then
                        Sql = "SELECT VtaCab.FolioVenta,VtaCab.Anticipo,VtaDet.CodArticulo,CatArt.CodAlmacenOrigen,CatArt.DescArticulo," & "SUM(ISNULL(DevDet.CantidadDevol,0)) AS CantidadDevuelta,VtaDet.Cantidad,VtaDet.PorcDescuentos," & "VtaDet.PorcPromociones,VtaDet.PrecioLista,VtaDet.NumPartida," & "VtaDet.PrecioLista * VtaDet.Cantidad AS Importe,VtaDet.PrecioReal * VtaDet.Cantidad AS ImporteDescto," & "CatArt.CodGrupo,ISNULL(Cf.DescFamilia,'') AS DescFamilia," & "VtaDet.ImptePromociones,VtaDet.ImpteDescuentos,VtaDet.PrecioLista,VtaDet.PrecioListaSinIva," & "VtaDet.PrecioReal,VtaDet.IvaReal,VtaDet.PorcAdicional,VtaDet.DescArticuloAdicional,VtaDet.FolioAdicional,VtaDet.CantidadAdicional,VtaDet.EstatusAdicional,VtaDet.FolioAdicional " & "FROM MovimientosVentasCab VtaCab INNER JOIN MovimientosVentasDet VtaDet " & "ON VtaCab.FolioVenta = VtaDet.FolioVenta " & "INNER JOIN CatArticulos CatArt ON VtaDet.CodArticulo = CatArt.CodArticulo " & "LEFT OUTER JOIN CatFamilias Cf ON CatArt.CodGrupo = Cf.CodGrupo AND CatArt.CodFamilia = Cf.CodFamilia " & "LEFT OUTER JOIN DevolucionesCab DevCab ON VtaDet.FolioVenta = DevCab.FolioVenta " & "LEFT OUTER JOIN DevolucionesDet DevDet ON DevCab.FolioDevolucion = DevDet.FolioDevolucion AND VtaDet.CodArticulo = DevDet.CodArticulo " & "WHERE VtaCab.FolioVenta ='" & RsGral.Fields("FolioVenta").Value & "' AND VtaCab.Estatus <> 'C' AND ISNULL(DevCab.Estatus,'') <> 'C' " & "GROUP BY VtaCab.FolioVenta,VtaCab.Anticipo,VtaDet.CodArticulo,CatArt.CodAlmacenOrigen,CatArt.DescArticulo," & "VtaDet.Cantidad,VtaDet.PorcDescuentos,VtaDet.PorcPromociones,VtaDet.PrecioLista,VtaDet.NumPartida," & "VtaDet.PrecioLista * VtaDet.Cantidad,VtaDet.PrecioReal * VtaDet.Cantidad," & "CatArt.CodGrupo,ISNULL(Cf.DescFamilia,''),VtaDet.ImptePromociones,VtaDet.ImpteDescuentos,VtaDet.PrecioLista,VtaDet.PrecioListaSinIva," & "VtaDet.PrecioReal,VtaDet.IvaReal,VtaDet.PorcAdicional,VtaDet.DescArticuloAdicional,VtaDet.FolioAdicional,VtaDet.CantidadAdicional,VtaDet.EstatusAdicional"

                        ModEstandar.BorraCmd()
                        Cmd.CommandText = "dbo.Up_Select_Datos"
                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                        RsAux = Cmd.Execute
                        If RsAux.RecordCount > 0 Then
                            Total = 0
                            SubTotal = 0
                            Descuento = 0
                            Iva = 0
                            RedondeoDolares = 0
                            RedondeoPesos = 0
                            flexVentas.set_TextMatrix(flexVentas.Row, C_COLFOLIOADICIONAL, Trim(RsAux.Fields("FolioAdicional").Value))
                            Do While Not RsAux.EOF
                                If (RsAux.Fields("Cantidad").Value - RsAux.Fields("CantidadDevuelta").Value) > 0 Then
                                    SubTotal = SubTotal + (RsAux.Fields("PrecioListaSinIva").Value * (RsAux.Fields("Cantidad").Value - RsAux.Fields("CantidadDevuelta").Value))
                                    Descuento = Descuento + ((RsAux.Fields("ImptePromociones").Value + RsAux.Fields("ImpteDescuentos").Value) * (RsAux.Fields("Cantidad").Value - RsAux.Fields("CantidadDevuelta").Value))
                                    Iva = Iva + (RsAux.Fields("IvaReal").Value * (RsAux.Fields("Cantidad").Value - RsAux.Fields("CantidadDevuelta").Value))
                                End If
                                RsAux.MoveNext()
                            Loop
                            If SubTotal > 0 Then
                                Total = ((SubTotal - Descuento) + Iva)
                                TotalPesos = CDbl(VB6.Format(Total * RsGral.Fields("TipoCambio").Value, "#####0.000000"))
                                TotalPesos = CDbl(VB6.Format(TotalPesos, "#####0.00"))
                                RedondeoPesos = ModCorporativo.RedondeoUnidadFinal(CDec(TotalPesos), CDbl(gcurRedondeo))
                                RedondeoDolares = CDbl(VB6.Format(RedondeoPesos / RsGral.Fields("TipoCambio").Value, "#####0.0000"))
                                .set_TextMatrix(.Row, C_COLTOTALADICIONAL, Total)
                                .set_TextMatrix(.Row, C_COLSUBTOTALADICIONAL, SubTotal)
                                .set_TextMatrix(.Row, C_COLDESCUENTOADICIONAL, Descuento)
                                .set_TextMatrix(.Row, C_COLIVAADICIONAL, Iva)
                                .set_TextMatrix(.Row, C_COLREDONDEOADICIONAL, RedondeoDolares)
                                .set_TextMatrix(.Row, C_COLINCLUIRFACTURA, "S")
                                .set_TextMatrix(.Row, C_COLANTICIPOADICIONAL, RsGral.Fields("Anticipo").Value)
                                If .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "TC" Then
                                    'CalculaImportesFactura
                                    mblnFactura = True
                                End If
                            End If
                        Else
                        End If
                    ElseIf Trim(RsGral.Fields("EstatusAdicional").Value) = "O" Then
                        .set_TextMatrix(.Row, C_COLTOTALADICIONAL, 0)
                        .set_TextMatrix(.Row, C_COLREDONDEOADICIONAL, 0)
                        .set_TextMatrix(.Row, C_COLGRABADO, "S")
                        .set_TextMatrix(.Row, C_COLESTATUSADICIONAL, "O")
                        .set_TextMatrix(.Row, C_COLFOLIOEXCLUIDO, "EXCLUIDO")
                        .Col = 0
                        .CellBackColor = lblExcluido.BackColor
                        .CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                        .Col = 1
                        .CellBackColor = lblExcluido.BackColor
                        .CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                    Else
                        .set_TextMatrix(.Row, C_COLGRABADO, "")
                    End If
                    RsGral.MoveNext()
                    If Not RsGral.EOF Then
                        .Row = .Row + 1
                    End If
                Loop
                .Row = 1
                .Col = 0
                CargarVentasPendientes()
                cmdGenerarFactura.Enabled = True
                cmdImpresionTickets.Enabled = False
                cmdImprimirFactura.Enabled = False
                '            flexVentas_EnterCell
                flexVentas.Focus()
                EnterCell()
            End With
        Else
            MsgBox("No existen ventas registradas para la sucursal " & dbcSucursal.Text & vbNewLine & "     El día " & VB6.Format(dtpFechaVenta.Value, "dd/mmm/yyyy") & ". Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            FueraChange = True
            optManual.Checked = True
            optPorcentual.Checked = False
            txtPorcentaje.Text = "0"
            txtPorcentaje.Enabled = False
            txtFolioFactura.Text = ""
            FueraChange = False
            dtpFechaRegistro.Value = Today
            lblDescripcion.Text = ""
            lblEstadoFolio.Text = ""
            lblSubTotal.Text = "0.00"
            lblRedondeo.Text = "0.00"
            lblTotal.Text = "0.00"
            lblTotalPesos.Text = "0.00"
            lblImporteSubTotal.Text = "0.00"
            lblImporteRedondeo.Text = "0.00"
            lblImporteTotal.Text = "0.00"
            lblDescuento.Text = "0.00"
            lblIva.Text = "0.00"
            lblSubTot.Text = "0.00"
            lblFactura.Text = "F00" & (Year(dtpFechaVenta.Value) + "0000") & (Month(dtpFechaVenta.Value) + "00") & ((dtpFechaVenta.Value) + "00") & "000000"
            lblCantidad.Text = ""
            LimpiarGridVentas()
            LimpiarGridDetalle()
            LimpiarGridPendientes()
            FueraChange = True
            txtDescripcion.Text = ""
            FueraChange = False
            txtDescripcion.ReadOnly = True
            cmdGenerarFactura.Enabled = True
            cmdImpresionTickets.Enabled = False
            cmdImprimirFactura.Enabled = False
            txtFacturaAdicional.Text = ""
            InicializaVariables()
            chkDoctoCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkDoctoCliente.Enabled = True
            chkDesglosarIva.Enabled = True
            chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked
            cmdDatosFiscales.Enabled = False
            'If Screen.ActiveForm.ActiveControl.Name <> "dbcSucursal" Then
            dbcSucursal.Focus()
            'End If
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub CargaDetalleAdicional()
        Dim I As Integer
        Dim TipoCambio As Decimal
        On Error GoTo Err_Renamed
        LimpiarGridDetalle()
        gStrSql = "SELECT VtaDet.FolioAdicional,VtaDet.CodArticulo,CatArt.CodAlmacenOrigen,CatArt.DescArticulo," & "VtaDet.CantidadAdicional,VtaDet.PorcPromociones,VtaDet.PorcDescuentos," & "VtaDet.PrecioLista,(VtaDet.PrecioLista * VtaDet.CantidadAdicional) AS Importe," & "(VtaDet.PrecioReal * VtaDet.CantidadAdicional) AS ImpteDescto,PrecioListaAdicional," & "(VtaDet.PrecioListaAdicional * VtaDet.CantidadAdicional) AS ImporteAdicional," & "(VtaDet.PrecioRealAdicional * VtaDet.CantidadAdicional) AS ImpteDesctoAdicional," & "VtaDet.PorcAdicional,VtaDet.DescArticuloAdicional,VtaDet.PrecioRealAdicional,VtaDet.RedondeoAdicional,VtaDet.TipoCambioAdicional " & "FROM MovimientosVentasCab VtaCab INNER JOIN MovimientosVentasDet VtaDet ON VtaCab.FolioVenta = VtaDet.FolioVenta " & "INNER JOIN CatArticulos CatArt ON VtaDet.CodArticulo = CatArt.CodArticulo " & "WHERE FolioAdicional = '" & flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            With flexDetalleVenta
                I = 2
                lblSubTotal.Text = "0.00"
                lblRedondeo.Text = "0.00"
                lblTotal.Text = "0.00"
                lblTotalPesos.Text = "0.00"
                lblRedondeo.Text = RsGral.Fields("RedondeoAdicional").Value
                TipoCambio = RsGral.Fields("TipoCambioAdicional").Value
                Do While Not RsGral.EOF
                    .set_TextMatrix(I, C_COLCODIGOARTICULO, RsGral.Fields("CodArticulo").Value & "-" & RsGral.Fields("CodAlmacenOrigen").Value)
                    .set_TextMatrix(I, C_COLCANTIDADDEVOL, 0)
                    .set_TextMatrix(I, C_COLCANTIDAD, RsGral.Fields("CantidadAdicional").Value)
                    .set_TextMatrix(I, C_COLPORCDESCUENTO, VB6.Format(RsGral.Fields("PorcDescuentos").Value, "###,##0.00") & "%")
                    .set_TextMatrix(I, C_COLPORCPROMOCION, VB6.Format(RsGral.Fields("PorcPromociones").Value, "###,##0.00") & "%")
                    .set_TextMatrix(I, C_ColPRECIOPUBLICO, VB6.Format(RsGral.Fields("PrecioLista").Value, "###,##0.00"))
                    .set_TextMatrix(I, C_COLIMPORTESINDESCUENTO, VB6.Format(RsGral.Fields("importe").Value, "###,##0.00"))
                    .set_TextMatrix(I, C_COLIMPORTECONDESCUENTO, VB6.Format(RsGral.Fields("ImpteDescto").Value, "###,##0.00"))
                    .set_TextMatrix(I, C_COLDESCRIPCION, Trim(RsGral.Fields("DescArticulo").Value))
                    .set_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO, VB6.Format(RsGral.Fields("PrecioListaAdicional").Value, "###,##0.00"))
                    .set_TextMatrix(I, C_COLNUEVOIMPORTESINDESCUENTO, VB6.Format(RsGral.Fields("ImporteAdicional").Value, "###,##0.00"))
                    .set_TextMatrix(I, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format(RsGral.Fields("ImpteDesctoAdicional").Value, "###,##0.00"))
                    .set_TextMatrix(I, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescArticuloAdicional").Value))
                    .set_TextMatrix(I, C_COLPORCENTAJEADICIONAL, RsGral.Fields("PorcAdicional").Value)
                    .set_TextMatrix(I, C_COLGRAB, "S")
                    lblSubTotal.Text = CStr(CDbl(Numerico(lblSubTotal.Text)) + (RsGral.Fields("PrecioRealAdicional").Value * RsGral.Fields("CantidadAdicional").Value))
                    If I = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    I = I + 1
                    RsGral.MoveNext()
                Loop
                If RsGral.RecordCount < 12 Then
                    .Rows = 12
                End If
                lblTotal.Text = CStr(CDbl(Numerico(lblSubTotal.Text)) + CDbl(Numerico(lblRedondeo.Text)))
                lblTotalPesos.Text = VB6.Format(CDbl(Numerico(lblTotal.Text)) * TipoCambio, "#####0.0")
                lblSubTotal.Text = VB6.Format(lblSubTotal.Text, "###,##0.00")
                lblRedondeo.Text = VB6.Format(lblRedondeo.Text, "###,##0.00")
                lblTotal.Text = VB6.Format(lblTotal.Text, "###,##0.00")
                lblTotalPesos.Text = VB6.Format(lblTotalPesos.Text, "###,##0.00")
            End With
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub CargaDetalle()
        Dim I As Integer
        On Error GoTo Err_Renamed
        LimpiarGridDetalle()

        gStrSql = "SELECT VtaDet.FolioVenta,VtaDet.CodArticulo,CatArt.CodAlmacenOrigen,CatArt.DescArticulo," & "SUM(ISNULL(DevDet.CantidadDevol,0)) AS CantidadDevuelta,VtaDet.Cantidad,VtaDet.PorcDescuentos," & "VtaDet.PorcPromociones,VtaDet.PrecioLista,VtaDet.NumPartida," & "VtaDet.PrecioLista * VtaDet.Cantidad AS Importe,VtaDet.PrecioReal * VtaDet.Cantidad AS ImporteDescto," & "CatArt.CodGrupo,ISNULL(Cf.DescFamilia,'') AS DescFamilia," & "VtaDet.ImptePromocionesAdicional,VtaDet.ImpteDescuentosAdicional,VtaDet.PrecioListaAdicional,VtaDet.PrecioListaSinIvaAdicional," & "VtaDet.PrecioRealAdicional,VtaDet.IvaRealAdicional,VtaDet.PorcAdicional,VtaDet.DescArticuloAdicional,VtaDet.FolioAdicional,VtaDet.CantidadAdicional,VtaDet.EstatusAdicional,VtaDet.PrecioReal " & "FROM /*MovimientosVentasCab VtaCab,FULL JOIN*/ MovimientosVentasDet VtaDet " & "/*ON VtaCab.FolioVenta = VtaDet.FolioVenta*/ " & "INNER JOIN CatArticulos CatArt ON VtaDet.CodArticulo = CatArt.CodArticulo " & "LEFT OUTER JOIN CatFamilias Cf ON CatArt.CodGrupo = Cf.CodGrupo AND CatArt.CodFamilia = Cf.CodFamilia " & "LEFT OUTER JOIN DevolucionesCab DevCab ON VtaDet.FolioVenta = DevCab.FolioVenta " & "LEFT OUTER JOIN DevolucionesDet DevDet ON DevCab.FolioDevolucion = DevDet.FolioDevolucion AND VtaDet.CodArticulo = DevDet.CodArticulo " & "WHERE (VtaDet.FolioVenta ='" & flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA) & "' OR VtaDet.FolioAdicional = '" & flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOADICIONAL) & "' AND VtaDet.FolioAdicional <> '') AND /*VtaCab.Estatus <> 'C' AND (VtaDet.EstatusAdicional  '' OR VtaDet.EstatusAdicional = 'O') AND*/ ISNULL(DevCab.Estatus,'') <> 'C' " & "GROUP BY VtaDet.FolioVenta,VtaDet.CodArticulo,CatArt.CodAlmacenOrigen,CatArt.DescArticulo," & "VtaDet.Cantidad,VtaDet.PorcDescuentos,VtaDet.PorcPromociones,VtaDet.PrecioLista,VtaDet.NumPartida," & "VtaDet.PrecioLista * VtaDet.Cantidad,VtaDet.PrecioReal * VtaDet.Cantidad," & "CatArt.CodGrupo,ISNULL(Cf.DescFamilia,''),VtaDet.ImptePromocionesAdicional,VtaDet.ImpteDescuentosAdicional,VtaDet.PrecioListaAdicional,VtaDet.PrecioListaSinIvaAdicional," & "VtaDet.PrecioRealAdicional,VtaDet.IvaRealAdicional,VtaDet.PorcAdicional,VtaDet.DescArticuloAdicional,VtaDet.FolioAdicional,VtaDet.CantidadAdicional,VtaDet.EstatusAdicional,VtaDet.PrecioReal"

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            With flexDetalleVenta
                If RsGral.RecordCount > 10 Then
                    .Rows = RsGral.RecordCount + 2
                End If
                .Row = 2
                Do While Not RsGral.EOF
                    If RsGral.Fields("FolioVenta").Value <> flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA) Then
                        .set_TextMatrix(.Row, C_COLFOLIOAGREGADO, RsGral.Fields("FolioVenta").Value) 'Quiere decir que se agregaron partidas de otros folios
                    End If
                    .set_TextMatrix(.Row, C_COLCODIGOARTICULO, RsGral.Fields("CodArticulo").Value & "-" & RsGral.Fields("CodAlmacenOrigen").Value)
                    .set_TextMatrix(.Row, C_COLCANTIDADDEVOL, RsGral.Fields("CantidadDevuelta").Value)
                    '                If RsGral!cantidaddevuelta > 0 Then
                    '                    flexVentas.TextMatrix(flexVentas.Row, C_COLTIENEDEVOLUCION) = "S"
                    '                End If
                    .set_TextMatrix(.Row, C_COLCANTIDAD, RsGral.Fields("Cantidad").Value)
                    .set_TextMatrix(.Row, C_COLPORCDESCUENTO, VB6.Format(RsGral.Fields("PorcDescuentos").Value, "###,##0.00") & "%")
                    .set_TextMatrix(.Row, C_COLPORCDESCTO, RsGral.Fields("PorcDescuentos").Value)
                    .set_TextMatrix(.Row, C_COLPORCPROMOCION, VB6.Format(RsGral.Fields("PorcPromociones").Value, "###,##0.00") & "%")
                    .set_TextMatrix(.Row, C_COLPORCPROM, RsGral.Fields("PorcPromociones").Value)
                    .set_TextMatrix(.Row, C_ColPRECIOPUBLICO, VB6.Format(RsGral.Fields("PrecioLista").Value, "###,##0.00"))
                    .set_TextMatrix(.Row, C_COLIMPORTESINDESCUENTO, VB6.Format(RsGral.Fields("importe").Value, "###,##0.00"))
                    .set_TextMatrix(.Row, C_COLIMPORTECONDESCUENTO, VB6.Format(RsGral.Fields("importedescto").Value, "###,##0.00"))
                    .set_TextMatrix(.Row, C_COLDESCRIPCION, Trim(RsGral.Fields("DescArticulo").Value))
                    .set_TextMatrix(.Row, C_COLNUMPARTIDA, RsGral.Fields("NumPartida").Value)
                    If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) = "N" And Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLGRABADO)) = "S" And RsGral.Fields("EstatusAdicional").Value <> "O" Then
                        .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescArticuloAdicional").Value))
                        .set_TextMatrix(.Row, C_COLDESCRIPCIONFAMILIA, Trim(RsGral.Fields("DescArticuloAdicional").Value))
                        .set_TextMatrix(.Row, C_COLNUEVOPRECIOPUBLICO, VB6.Format(RsGral.Fields("PrecioListaAdicional").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTESINDESCUENTO, VB6.Format(RsGral.Fields("PrecioListaAdicional").Value * RsGral.Fields("CantidadAdicional").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format(RsGral.Fields("PrecioRealAdicional").Value * RsGral.Fields("CantidadAdicional").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLALMACEN, RsGral.Fields("CodAlmacenOrigen").Value)
                        .set_TextMatrix(.Row, C_COLPORCENTAJEADICIONAL, RsGral.Fields("PorcAdicional").Value)
                        .set_TextMatrix(.Row, C_COLGRAB, "S")
                        .set_TextMatrix(.Row, C_ColPROMOCION, RsGral.Fields("ImptePromocionesAdicional").Value)
                        .set_TextMatrix(.Row, C_ColDESCUENTO, RsGral.Fields("ImpteDescuentosAdicional").Value)
                        .set_TextMatrix(.Row, C_COLPRECIOLISTAADICIONAL, RsGral.Fields("PrecioListaAdicional").Value)
                        .set_TextMatrix(.Row, C_ColPRECIOLISTASINIVA, RsGral.Fields("PrecioListaSinIvaAdicional").Value)
                        .set_TextMatrix(.Row, C_COLPRECIOREAL, RsGral.Fields("PrecioRealAdicional").Value)
                        .set_TextMatrix(.Row, C_COLIVAREAL, RsGral.Fields("IvaRealAdicional").Value)
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLESTATUSADICIONAL, "V")
                    ElseIf Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) = "N" And Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLGRABADO)) = "" Then
                        If RsGral.Fields("CodGrupo").Value = gCODJOYERIA Or RsGral.Fields("CodGrupo").Value = gCODVARIOS Then
                            If Trim(RsGral.Fields("DescArticuloAdicional").Value) = "" Then
                                .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescFamilia").Value))
                            Else
                                .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescArticuloAdicional").Value))
                            End If
                            .set_TextMatrix(.Row, C_COLDESCRIPCIONFAMILIA, Trim(RsGral.Fields("DescFamilia").Value))
                        ElseIf RsGral.Fields("CodGrupo").Value = gCODRELOJERIA Then
                            .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, "RELOJ")
                            .set_TextMatrix(.Row, C_COLDESCRIPCIONFAMILIA, "RELOJ")
                        End If
                        .set_TextMatrix(.Row, C_COLNUEVOPRECIOPUBLICO, "0.00")
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTESINDESCUENTO, "0.00")
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTECONDESCUENTO, "0.00")
                        .set_TextMatrix(.Row, C_COLALMACEN, RsGral.Fields("CodAlmacenOrigen").Value)
                        lblSubTotal.Text = "0.00"
                        lblRedondeo.Text = "0.00"
                        lblTotal.Text = "0.00"
                        lblTotalPesos.Text = "0.00"
                    ElseIf Trim(RsGral.Fields("EstatusAdicional").Value) = "O" Then
                        .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescFamilia").Value))
                        .set_TextMatrix(.Row, C_COLDESCRIPCIONFAMILIA, Trim(RsGral.Fields("DescFamilia").Value))
                        .set_TextMatrix(.Row, C_COLNUEVOPRECIOPUBLICO, "0.00")
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTESINDESCUENTO, "0.00")
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTECONDESCUENTO, "0.00")
                        .set_TextMatrix(.Row, C_COLEXCLUIDO, "EXCLUIDO")
                        .set_TextMatrix(.Row, C_COLGRAB, "S")
                        .set_TextMatrix(.Row, C_COLALMACEN, RsGral.Fields("CodAlmacenOrigen").Value)
                        For I = 0 To 11
                            .Col = I
                            .CellBackColor = lblExcluido.BackColor
                            .CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                        Next
                        .Col = 0
                    ElseIf Trim(RsGral.Fields("FolioAdicional").Value) <> "" And Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) <> "TC" Then
                        .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescArticuloAdicional").Value))
                        .set_TextMatrix(.Row, C_COLDESCRIPCIONFAMILIA, Trim(RsGral.Fields("DescArticuloAdicional").Value))
                        .set_TextMatrix(.Row, C_COLNUEVOPRECIOPUBLICO, VB6.Format(RsGral.Fields("PrecioListaAdicional").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTESINDESCUENTO, VB6.Format(RsGral.Fields("PrecioListaAdicional").Value * RsGral.Fields("CantidadAdicional").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format(RsGral.Fields("PrecioRealAdicional").Value * RsGral.Fields("CantidadAdicional").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLALMACEN, RsGral.Fields("CodAlmacenOrigen").Value)
                        .set_TextMatrix(.Row, C_COLPORCENTAJEADICIONAL, RsGral.Fields("PorcAdicional").Value)
                        .set_TextMatrix(.Row, C_COLGRAB, "S")
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLESTATUSADICIONAL, "V")
                    ElseIf Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) = "TC" And Trim(RsGral.Fields("FolioAdicional").Value) <> "" Then
                        .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescArticuloAdicional").Value))
                        .set_TextMatrix(.Row, C_COLDESCRIPCIONFAMILIA, Trim(RsGral.Fields("DescArticuloAdicional").Value))
                        .set_TextMatrix(.Row, C_COLNUEVOPRECIOPUBLICO, VB6.Format(RsGral.Fields("PrecioLista").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTESINDESCUENTO, VB6.Format(RsGral.Fields("PrecioLista").Value * RsGral.Fields("Cantidad").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format(RsGral.Fields("PrecioReal").Value * RsGral.Fields("Cantidad").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLALMACEN, RsGral.Fields("CodAlmacenOrigen").Value)
                        .set_TextMatrix(.Row, C_COLPORCENTAJEADICIONAL, RsGral.Fields("PorcAdicional").Value)

                        .set_TextMatrix(.Row, C_COLGRAB, "S")
                        .set_TextMatrix(.Row, C_ColPROMOCION, RsGral.Fields("ImptePromocionesAdicional").Value)
                        .set_TextMatrix(.Row, C_ColDESCUENTO, RsGral.Fields("ImpteDescuentosAdicional").Value)
                        .set_TextMatrix(.Row, C_COLPRECIOLISTAADICIONAL, RsGral.Fields("PrecioListaAdicional").Value)
                        .set_TextMatrix(.Row, C_ColPRECIOLISTASINIVA, RsGral.Fields("PrecioListaSinIvaAdicional").Value)
                        .set_TextMatrix(.Row, C_COLPRECIOREAL, RsGral.Fields("PrecioRealAdicional").Value)
                        .set_TextMatrix(.Row, C_COLIVAREAL, RsGral.Fields("IvaRealAdicional").Value)

                    ElseIf Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) = "TC" And Trim(RsGral.Fields("FolioAdicional").Value) = "" Then
                        If RsGral.Fields("CodGrupo").Value = gCODJOYERIA Or RsGral.Fields("CodGrupo").Value = gCODVARIOS Then
                            If Trim(RsGral.Fields("DescArticuloAdicional").Value) = "" Then
                                .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescFamilia").Value))
                            Else
                                .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescArticuloAdicional").Value))
                            End If
                            .set_TextMatrix(.Row, C_COLDESCRIPCIONFAMILIA, Trim(RsGral.Fields("DescFamilia").Value))
                        ElseIf RsGral.Fields("CodGrupo").Value = gCODRELOJERIA Then
                            If Trim(RsGral.Fields("DescArticuloAdicional").Value) = "" Then
                                .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, "RELOJ")
                            Else
                                .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescArticuloAdicional").Value))
                            End If
                            .set_TextMatrix(.Row, C_COLDESCRIPCIONFAMILIA, "RELOJ")
                        End If
                        ''' ***** ojo folio pagado totalmente con tarjeta *****
                        ''' al parecer solo entran apartados pagados con diferentes formas de pago a esta opcion
                        ''' y ventas credito sin anticipo...
                        .set_TextMatrix(.Row, C_COLNUEVOPRECIOPUBLICO, "0.00")
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTESINDESCUENTO, "0.00")
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTECONDESCUENTO, "0.00")

                        .set_TextMatrix(.Row, C_COLNUEVOPRECIOPUBLICO, VB6.Format(RsGral.Fields("PrecioLista").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTESINDESCUENTO, VB6.Format(RsGral.Fields("PrecioLista").Value * RsGral.Fields("Cantidad").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format(RsGral.Fields("PrecioReal").Value * RsGral.Fields("Cantidad").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, C_COLALMACEN, RsGral.Fields("CodAlmacenOrigen").Value)
                        .set_TextMatrix(.Row, C_COLPORCENTAJEADICIONAL, RsGral.Fields("PorcAdicional").Value)
                        .set_TextMatrix(.Row, C_COLGRAB, "S")
                        .set_TextMatrix(.Row, C_ColPROMOCION, RsGral.Fields("ImptePromocionesAdicional").Value)
                        .set_TextMatrix(.Row, C_ColDESCUENTO, RsGral.Fields("ImpteDescuentosAdicional").Value)
                        .set_TextMatrix(.Row, C_COLPRECIOLISTAADICIONAL, RsGral.Fields("PrecioListaAdicional").Value)
                        .set_TextMatrix(.Row, C_ColPRECIOLISTASINIVA, RsGral.Fields("PrecioListaSinIvaAdicional").Value)
                        .set_TextMatrix(.Row, C_COLPRECIOREAL, RsGral.Fields("PrecioRealAdicional").Value)
                        .set_TextMatrix(.Row, C_COLIVAREAL, RsGral.Fields("IvaRealAdicional").Value)
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLESTATUSADICIONAL, "V")
                    Else
                        If RsGral.Fields("CodGrupo").Value = gCODJOYERIA Or RsGral.Fields("CodGrupo").Value = gCODVARIOS Then
                            .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, Trim(RsGral.Fields("DescFamilia").Value))
                            .set_TextMatrix(.Row, C_COLDESCRIPCIONFAMILIA, Trim(RsGral.Fields("DescFamilia").Value))
                        ElseIf RsGral.Fields("CodGrupo").Value = gCODRELOJERIA Then
                            .set_TextMatrix(.Row, C_COLNUEVADESCRIPCION, "RELOJ")
                            .set_TextMatrix(.Row, C_COLDESCRIPCIONFAMILIA, "RELOJ")
                        End If
                        .set_TextMatrix(.Row, C_COLNUEVOPRECIOPUBLICO, "0.00")
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTESINDESCUENTO, "0.00")
                        .set_TextMatrix(.Row, C_COLNUEVOIMPORTECONDESCUENTO, "0.00")
                        .set_TextMatrix(.Row, C_COLALMACEN, RsGral.Fields("CodAlmacenOrigen").Value)
                        lblSubTotal.Text = "0.00"
                        lblRedondeo.Text = "0.00"
                        lblTotal.Text = "0.00"
                        lblTotalPesos.Text = "0.00"
                    End If
                    RsGral.MoveNext()
                    If Not RsGral.EOF Then
                        .Row = .Row + 1
                    End If
                Loop
                .Row = 2
                lblSubTotal.Text = Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLTOTALADICIONAL))
                lblRedondeo.Text = Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLREDONDEOADICIONAL))
                lblTotal.Text = CStr(CDbl(Numerico(lblSubTotal.Text)) + CDbl(Numerico(lblRedondeo.Text)))
                lblTotalPesos.Text = CStr(System.Math.Round(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLTOTALADICIONAL))) * CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_ColTIPOCAMBIO))), 1) + System.Math.Round(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLREDONDEOADICIONAL))) * CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_ColTIPOCAMBIO))), 1))
                lblSubTotal.Text = VB6.Format(lblSubTotal.Text, "###,##0.00")
                lblRedondeo.Text = VB6.Format(lblRedondeo.Text, "###,##0.00")
                lblTotal.Text = VB6.Format(lblTotal.Text, "###,##0.00")
                lblTotalPesos.Text = VB6.Format(lblTotalPesos.Text, "###,##0.00")
            End With
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub CargarVentasPendientes()
        On Error GoTo Err_Renamed
        gStrSql = "SELECT VtaCab.FolioVenta AS 'FolioVenta',RTRIM(Dbo.FormatFecha(VtaCab.FechaVenta,10)) AS 'FechaVenta',CAST(VtaDet.CodArticulo AS Varchar(5)) + '-' + CAST(CatArt.CodAlmacenOrigen AS Varchar(5)) AS Articulo,RTRIM(CatArt.DescArticulo) AS DescArticulo,SUM(ISNULL(DevDet.CantidadDevol,0)) AS Dev,VtaDet.Cantidad AS Cant,DBO.FormatCantidad(VtaDet.PorcDescuentos) + '%' AS Descto,DBO.FormatCantidad(VtaDet.PorcPromociones) + '%' AS Prom,DBO.FormatCantidad(VtaDet.PrecioLista) AS 'PrecioPub',DBO.formatCantidad(VtaDet.PrecioLista * VtaDet.Cantidad) AS Importe,DBO.FormatCantidad(VtaDet.PrecioReal * VtaDet.Cantidad) AS 'ImpteDescto',RTRIM(CASE WHEN CatArt.CodGrupo = " & gCODJOYERIA & " OR CatArt.CodGrupo = " & gCODVARIOS & " THEN Cf.DescFamilia ELSE 'RELOJ' END) AS DescFamilia," & "VtaDet.PorcDescuentos,VtaDet.PorcPromociones,VtaDet.NumPartida,SUM(CASE WHEN ISNULL(FP.EsTarjeta,0) = 0 THEN 0 ELSE 1 END) AS EsTarjeta,SUM(CASE WHEN (ISNULL(FP.EsDevolucion,0) = 1 AND ISNULL(FP.EsDocumentoInterno,0) = 0) THEN 1 ELSE 0 END) AS EsNotaCredito,VtaCab.Condicion,VtaCab.TipoMovto " & "FROM ((SELECT VtaCab.FolioVenta,VtaCab.FechaVenta,VtaCab.Total,VtaCab.Moneda,VtaCab.FolioFactura,VtaCab.Condicion,VtaCab.TipoMovto " & "FROM MovimientosVentasCab VtaCab INNER JOIN MovimientosVentasDet VtaDet ON VtaCab.FolioVenta = VtaDet.FolioVenta WHERE VtaCab.Estatus <> 'C' AND VtaCab.TipoMovto = 'V' AND (VtaCab.FechaVenta BETWEEN '" & VB6.Format(DtpDesde.Value, "mm/dd/yyyy") & "' AND '" & VB6.Format(DtpHasta.Value, "mm/dd/yyyy") & "') AND VtaCab.CodSucursal = " & Numerico(txtCodSucursal.Text) & " AND VtaCab.FolioFactura = '' AND VtaDet.FolioAdicional = '' AND VtaDet.EstatusAdicional = '' AND VtaDet.FolioFactura = '') " & "UNION " & "(SELECT VtaCab.FolioVenta,APT.FechaVtaAp AS FechaVenta,VtaCab.Total,VtaCab.Moneda,VtaCab.FolioFactura,VtaCab.Condicion,VtaCab.TipoMovto " & "FROM (SELECT * FROM DBO.VW_APARTADOS WHERE (FECHAVTAAP BETWEEN '" & VB6.Format(DtpDesde.Value, "mm/dd/yyyy") & "' AND '" & VB6.Format(DtpHasta.Value, "mm/dd/yyyy") & "') AND ESTATUSAPT = 'S' AND ESTATUS <> 'C') APT " & "INNER JOIN (SELECT * FROM MovimientosVentasCab WHERE TipoMovto = 'A' AND Estatus <> 'C') VtaCab " & "ON APT.FolioVenta = VtaCab.FolioVenta INNER JOIN MovimientosVentasDet VtaDet ON APT.FolioVenta = VtaDet.FolioVenta AND VtaCab.FolioVenta = VtaDet.FolioVenta " & "WHERE APT.CodSucursal = " & Numerico(txtCodSucursal.Text) & " AND VtaCab.FolioFactura = '' AND VtaDet.FolioAdicional = '' AND VtaDet.EstatusAdicional = '' AND VtaDet.FolioFactura = '' " & "GROUP BY VtaCab.FolioVenta,APT.FechaVtaAp,VtaCab.Total,VtaCab.Moneda,VtaCab.FolioFactura,vtacab.condicion,vtacab.tipomovto " & ")) VtaCab LEFT OUTER JOIN IngresosFormaDePago Ing ON VtaCab.FolioVenta = Ing.FolioMovto LEFT OUTER JOIN CatFormasPago FP ON Ing.CodFormaPago = FP.CodFormaPago " & "INNER JOIN (SELECT * FROM MovimientosVentasDet WHERE FolioAdicional = '' AND EstatusAdicional = '' AND FolioFactura = '') VtaDet ON VtaCab.FolioVenta = VtaDet.FolioVenta " & "INNER JOIN CatArticulos CatArt ON VtaDet.CodArticulo = CatArt.CodArticulo " & "LEFT OUTER JOIN CatFamilias Cf ON CatArt.CodGrupo = Cf.CodGrupo AND CatArt.CodFamilia = Cf.CodFamilia " & "LEFT OUTER JOIN DevolucionesCab DevCab ON VtaCab.FolioVenta = DevCab.FolioVenta " & "LEFT OUTER JOIN DevolucionesDet DevDet ON DevCab.FolioDevolucion = DevDet.FolioDevolucion AND VtaDet.CodArticulo = DevDet.CodArticulo " & "WHERE ISNULL(DevCab.Estatus,'') <> 'C' " & "/*AND CatArt.CodAlmacenOrigen <> 0 AND (ISNULL(VtaDet.Cantidad,0) - ISNULL(DevDet.CantidadDevol,0)) > 0*/ " & "GROUP BY VtaCab.FolioVenta,VtaCab.FechaVenta,VtaDet.CodArticulo,CatArt.CodAlmacenOrigen,CatArt.DescArticulo," & "VtaDet.Cantidad,VtaDet.PorcDescuentos,VtaDet.PorcPromociones,VtaDet.NumPartida," & "VtaDet.PrecioLista,VtaDet.PrecioLista,VtaDet.Cantidad," & "VtaDet.PrecioReal,CatArt.CodGrupo,Cf.DescFamilia,VtaDet.PorcDescuentos,VtaDet.PorcPromociones,vTAcAB.CONDICION,vtacab.tipomovto " & "/*HAVING (VtaDet.Cantidad - SUM(ISNULL(DevDet.CantidadDevol,0))) > 0 AND SUM(CASE WHEN ISNULL(FP.EsTarjeta,0) = 0 THEN 0 ELSE 1 END) = 0*/ " & "ORDER BY VtaCab.FolioVenta,VtaCab.FechaVenta DESC"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            LimpiarGridPendientes()
            With flexVentasPendientes
                .Row = 2
                Do While Not RsGral.EOF
                    If (RsGral.Fields("cant").Value - RsGral.Fields("dev").Value) > 0 Then
                        'If ((RsGral!EsTarjeta = 0 Or RsGral!EsTarjeta > 0)) Or Not (RsGral!TipoMovto = "V" And RsGral!Condicion = "CO") Then
                        .set_TextMatrix(.Row, C_ColFOLIO, Trim(RsGral.Fields("FolioVenta").Value))
                        .set_TextMatrix(.Row, C_COLFECHAVENTA, RsGral.Fields("FechaVenta").Value)
                        .set_TextMatrix(.Row, C_COLCODARTICULO, RsGral.Fields("Articulo").Value)
                        .set_TextMatrix(.Row, C_COLDESCARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
                        .set_TextMatrix(.Row, C_COLDEVOL, RsGral.Fields("dev").Value)
                        .set_TextMatrix(.Row, C_COLCANT, RsGral.Fields("cant").Value)
                        .set_TextMatrix(.Row, C_COLDESCTO, RsGral.Fields("Descto").Value)
                        .set_TextMatrix(.Row, C_COLPROM, RsGral.Fields("Prom").Value)
                        .set_TextMatrix(.Row, C_COLPRECIOPUB, RsGral.Fields("PrecioPub").Value)
                        .set_TextMatrix(.Row, C_COLIMPTE, RsGral.Fields("importe").Value)
                        .set_TextMatrix(.Row, C_COLIMPTEDESCTO, RsGral.Fields("ImpteDescto").Value)
                        .set_TextMatrix(.Row, C_COLPORDESCTO, RsGral.Fields("PorcDescuentos").Value)
                        .set_TextMatrix(.Row, C_COLPORPROM, RsGral.Fields("PorcPromociones").Value)
                        .set_TextMatrix(.Row, C_COLDESCFAMILIA, RsGral.Fields("DescFamilia").Value)
                        .set_TextMatrix(.Row, C_COLNPARTIDA, RsGral.Fields("NumPartida").Value)
                        If .Row = .Rows - 1 Then
                            .Rows = .Rows + 1
                        End If
                        .Row = .Row + 1
                        'End If
                    End If
                    RsGral.MoveNext()
                Loop
                .Row = 2
            End With
        Else
            LimpiarGridPendientes()
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub EnterCell()
        If System.Drawing.ColorTranslator.ToOle(flexVentas.CellBackColor) = System.Drawing.ColorTranslator.ToOle(lblExcluido.BackColor) Then
            ColorAnteFolio = "Exc"
        ElseIf System.Drawing.ColorTranslator.ToOle(flexVentas.CellBackColor) = System.Drawing.ColorTranslator.ToOle(lblFacturar.BackColor) Then
            ColorAnteFolio = "FAC"
        Else
            ColorAnteFolio = "Bln"
        End If
        '    flexVentas.Col = 0
        '    flexVentas.CellBackColor = flexVentas.BackColorSel
        '    flexVentas.CellForeColor = &H80000009
        '    flexVentas.Col = 1
        '    flexVentas.CellBackColor = flexVentas.BackColorSel
        '    flexVentas.CellForeColor = &H80000009
        '    DoEvents
        With flexVentas
            If .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "F" Then
                lblEstadoFolio.Text = "FACTURADO"
                txtDescripcion.ReadOnly = True
            ElseIf .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "TC" Then
                lblEstadoFolio.Text = "PAG TARJ CRED"
                txtDescripcion.ReadOnly = False
            ElseIf .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "N" And Trim(.get_TextMatrix(.Row, C_COLFOLIOEXCLUIDO)) = "" And Trim(.get_TextMatrix(.Row, C_COLGRABADO)) = "" Then
                lblEstadoFolio.Text = "AFECTABLE"
                txtDescripcion.ReadOnly = False
            ElseIf .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "N" And Trim(.get_TextMatrix(.Row, C_COLFOLIOEXCLUIDO)) = "" And Trim(.get_TextMatrix(.Row, C_COLGRABADO)) = "S" Then
                lblEstadoFolio.Text = "AFECTADO"
                txtDescripcion.ReadOnly = False
            ElseIf .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "N" And Trim(.get_TextMatrix(.Row, C_COLFOLIOEXCLUIDO)) = "EXCLUIDO" Then
                lblEstadoFolio.Text = "EXCLUIDO"
                txtDescripcion.ReadOnly = False
            ElseIf .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "" Then
                lblEstadoFolio.Text = ""
                txtDescripcion.ReadOnly = True
            End If
            lblMoneda.Text = IIf(Trim(.get_TextMatrix(.Row, C_ColMONEDA)) = "P", "PESOS", IIf(Trim(.get_TextMatrix(.Row, C_ColMONEDA)) = "D", "DOLARES", ""))
            If Trim(.get_TextMatrix(.Row, C_COLFOLIOVENTA)) <> "" And Not mblnNuevo Then
                LimpiarGridDetalle()
                CargaDetalle()
                'CargarVentasPendientes
                '            DoEvents
            ElseIf Trim(.get_TextMatrix(.Row, C_COLFOLIOVENTA)) <> "" And mblnNuevo Then
                LimpiarGridDetalle()
                CargaDetalleAdicional()
                '            DoEvents
            Else
                LimpiarGridDetalle()
                lblSubTotal.Text = "0.00"
                lblRedondeo.Text = "0.00"
                lblTotal.Text = "0.00"
                lblTotalPesos.Text = "0.00"
            End If
            FueraChange = True
            txtDescripcion.Text = ""
            FueraChange = False
        End With
        txtPorcentaje.Text = "0"
    End Sub

    Sub EnviaPartida(ByRef Ren As Integer)
        With flexDetalleVenta
            .set_TextMatrix(Ren, C_COLCODIGOARTICULO, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLCODARTICULO))
            .set_TextMatrix(Ren, C_COLDESCRIPCION, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLDESCARTICULO))
            .set_TextMatrix(Ren, C_COLCANTIDADDEVOL, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLDEVOL))
            .set_TextMatrix(Ren, C_COLCANTIDAD, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLCANT))
            .set_TextMatrix(Ren, C_COLPORCDESCUENTO, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLDESCTO))
            .set_TextMatrix(Ren, C_COLPORCPROMOCION, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLPROM))
            .set_TextMatrix(Ren, C_ColPRECIOPUBLICO, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLPRECIOPUB))
            .set_TextMatrix(Ren, C_COLIMPORTESINDESCUENTO, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLIMPTE))
            .set_TextMatrix(Ren, C_COLIMPORTECONDESCUENTO, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLIMPTEDESCTO))
            .set_TextMatrix(Ren, C_COLFOLIOAGREGADO, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_ColFOLIO))
            .set_TextMatrix(Ren, C_COLFECHAFOLIOAGREGADO, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLFECHAVENTA))
            .set_TextMatrix(Ren, C_COLNUEVOPRECIOPUBLICO, "0.00")
            .set_TextMatrix(Ren, C_COLNUEVOIMPORTESINDESCUENTO, "0.00")
            .set_TextMatrix(Ren, C_COLNUEVOIMPORTECONDESCUENTO, "0.00")
            .set_TextMatrix(Ren, C_COLPORCDESCTO, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLPORDESCTO))
            .set_TextMatrix(Ren, C_COLPORCPROM, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLPORPROM))
            .set_TextMatrix(Ren, C_COLNUEVADESCRIPCION, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLDESCFAMILIA))
            .set_TextMatrix(Ren, C_COLNUMPARTIDA, flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLNPARTIDA))
            If flexVentasPendientes.Rows - 1 > 12 Then
                flexVentasPendientes.RemoveItem((flexVentasPendientes.Row))
            Else
                flexVentasPendientes.RemoveItem((flexVentasPendientes.Row))
                flexVentasPendientes.Rows = 12
            End If
        End With
    End Sub

    Sub RegresaPartida(ByRef Ren As Integer)
        With flexVentasPendientes
            .set_TextMatrix(Ren, C_ColFOLIO, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLFOLIOAGREGADO))
            .set_TextMatrix(Ren, C_COLFECHAVENTA, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLFECHAFOLIOAGREGADO))
            .set_TextMatrix(Ren, C_COLCODARTICULO, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLCODIGOARTICULO))
            .set_TextMatrix(Ren, C_COLDESCARTICULO, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLDESCRIPCION))
            .set_TextMatrix(Ren, C_COLDEVOL, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLCANTIDADDEVOL))
            .set_TextMatrix(Ren, C_COLCANT, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLCANTIDAD))
            .set_TextMatrix(Ren, C_COLDESCTO, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLPORCDESCUENTO))
            .set_TextMatrix(Ren, C_COLPROM, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLPORCPROMOCION))
            .set_TextMatrix(Ren, C_COLPRECIOPUB, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_ColPRECIOPUBLICO))
            .set_TextMatrix(Ren, C_COLIMPTE, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLIMPORTESINDESCUENTO))
            .set_TextMatrix(Ren, C_COLIMPTEDESCTO, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLIMPORTECONDESCUENTO))
            .set_TextMatrix(Ren, C_COLDESCFAMILIA, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLNUEVADESCRIPCION))
            .set_TextMatrix(Ren, C_COLPORDESCTO, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLPORCDESCTO))
            .set_TextMatrix(Ren, C_COLPORPROM, flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLPORCPROM))
            If flexDetalleVenta.Rows - 1 > 12 Then
                flexDetalleVenta.RemoveItem((flexDetalleVenta.Row))
            Else
                flexDetalleVenta.RemoveItem((flexDetalleVenta.Row))
                flexDetalleVenta.Rows = 12
            End If
            flexVentasPendientes.Sort = 8
        End With
    End Sub

    Sub ExcluirPartida()
        Dim I As Integer
        Dim Col As Integer
        With flexDetalleVenta
            Col = .Col
            If Trim(.get_TextMatrix(.Row, C_COLEXCLUIDO)) = "" And Trim(.get_TextMatrix(.Row, C_COLGRAB)) = "" Then
                If Trim(.get_TextMatrix(.Row, 0)) <> "" Then
                    For I = 0 To 11
                        .Col = I
                        .CellBackColor = lblExcluido.BackColor
                        .CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                    Next
                    ColorAnte = "Exc"
                    .set_TextMatrix(.Row, C_COLNUEVOPRECIOPUBLICO, "0.00")
                    .set_TextMatrix(.Row, C_COLNUEVOIMPORTESINDESCUENTO, "0.00")
                    .set_TextMatrix(.Row, C_COLNUEVOIMPORTECONDESCUENTO, "0.00")
                    .set_TextMatrix(.Row, C_ColPRECIOLISTASINIVA, "")
                    .set_TextMatrix(.Row, C_COLPRECIOREAL, "")
                    .set_TextMatrix(.Row, C_COLIVAREAL, "")
                    .set_TextMatrix(.Row, C_COLMODIFICADOTAG, .get_TextMatrix(.Row, C_COLMODIFICADO))
                    .set_TextMatrix(.Row, C_COLMODIFICADO, "")
                    .set_TextMatrix(.Row, C_ColDESCUENTO, "")
                    .set_TextMatrix(.Row, C_ColPROMOCION, "")
                    .set_TextMatrix(.Row, C_COLEXCLUIDO, "EXCLUIDO")
                    CalculaImportes()
                End If
            ElseIf Trim(.get_TextMatrix(.Row, C_COLEXCLUIDO)) = "EXCLUIDO" And Trim(.get_TextMatrix(.Row, C_COLGRAB)) = "" Then
                If Trim(.get_TextMatrix(.Row, 0)) <> "" Then
                    For I = 0 To 11
                        .Col = I
                        .CellBackColor = flexDetalleVenta.BackColor
                        .CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                    Next
                    ColorAnte = "Exc"
                    .set_TextMatrix(.Row, C_COLEXCLUIDO, "")
                End If
            Else
                Exit Sub
            End If
            .Col = Col
            flexDetalleVenta_EnterCell(flexDetalleVenta, New System.EventArgs())
        End With
    End Sub

    Sub ExcluirFolio()
        Dim I As Integer
        Dim J As Integer
        Dim Col As Integer
        With flexVentas
            Col = .Col
            If Trim(.get_TextMatrix(.Row, C_COLFOLIOEXCLUIDO)) = "" Then
                If Trim(.get_TextMatrix(.Row, 0)) <> "" Then
                    For I = 0 To 1
                        .Col = I
                        .CellBackColor = lblExcluido.BackColor
                        .CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                    Next
                    .Col = 0
                    ColorAnteFolio = "Exc"
                    .set_TextMatrix(.Row, C_COLFOLIOEXCLUIDO, "EXCLUIDO")
                End If
                With flexDetalleVenta
                    For I = 2 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" Then
                            .set_TextMatrix(I, C_COLEXCLUIDO, "EXCLUIDO")
                            For J = 0 To 11
                                .Row = I
                                .Col = J
                                .CellBackColor = lblExcluido.BackColor
                                .CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                            Next
                        End If
                    Next
                    .Row = 2
                    .Col = 0
                End With
                Select Case MsgBox("¿Realmente desea excluir este folio?" & vbNewLine & "!!!ADVERTENCIA¡¡¡ Una vez excluido no podra desexcluirlo.", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        GuardarFolio()
                        Exit Sub
                    Case MsgBoxResult.No
                        If Trim(.get_TextMatrix(.Row, 0)) <> "" Then
                            For I = 0 To 1
                                .Col = I
                                .CellBackColor = System.Drawing.ColorTranslator.FromOle(&H8000000E)
                                '.CellForeColor = &H80000009
                            Next
                            .Col = 0
                            ColorAnteFolio = "Bln"
                            .set_TextMatrix(.Row, C_COLFOLIOEXCLUIDO, "")

                        End If
                        With flexDetalleVenta
                            For I = 2 To .Rows - 1
                                If Trim(.get_TextMatrix(I, 0)) <> "" Then
                                    .set_TextMatrix(I, C_COLEXCLUIDO, "")
                                    For J = 0 To 11
                                        .Row = I
                                        .Col = J
                                        .CellBackColor = flexVentas.BackColor
                                        .CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                                    Next
                                End If
                                .set_TextMatrix(I, C_COLMODIFICADO, .get_TextMatrix(I, C_COLMODIFICADOTAG))
                            Next
                            .Row = 2
                            .Col = 0
                        End With
                        Exit Sub
                End Select
            End If
            .Col = Col
        End With
    End Sub

    Function EstaExcluido() As Boolean
        Dim I As Integer
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLEXCLUIDO)) <> "EXCLUIDO" And Trim(.get_TextMatrix(I, C_COLCODIGOARTICULO)) <> "" Then
                    EstaExcluido = False
                    Exit Function
                End If
            Next
            EstaExcluido = True
        End With
    End Function

    Function EstaVacia() As Boolean
        Dim I As Integer
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" Then
                    EstaVacia = False
                    Exit Function
                End If
            Next
            EstaVacia = True
        End With
    End Function

    Sub GeneraFolioAdicional()
        On Error GoTo Err_Renamed
        Dim TipoMovto As String
        Dim CajaSucursal As String
        Dim Fecha As String
        Dim Consecutivo As Integer
        TipoMovto = (flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA) + 1)
        CajaSucursal = Mid(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA), 2, 4)
        Fecha = Mid(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA), 6, 8)
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
            FolioAdicional = TipoMovto & CajaSucursal & Fecha & VB6.Format(Consecutivo, "0000")
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub ImprimirFactura()
        On Error GoTo Merr
        Dim strImpresora As String
        Dim strNomEmpresa As String
        Dim Archivo As Integer
        Dim I As Integer
        Dim TotalPesos As Double
        Dim CoordY As Integer
        Dim NombreCliente As String
        Dim Domicilio As String
        Dim Ciudad As String
        Dim Rfc As String
        Dim SubTotal As Decimal
        Dim Iva As Decimal
        Dim Total As Decimal

        '''X Columna
        '''Y Renglon

        strImpresora = gstrRutaImpresora
        If Not ModEstandar.BuscarImpresora(strImpresora) Then
            MsgBox("Impresora Incorrecta")
        End If
        gStrSql = "SELECT * FROM CatRFC WHERE Codrfc = " & CodCliente
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            NombreCliente = Trim(RsGral.Fields("DescClienteRFC").Value)
            Domicilio = Trim(RsGral.Fields("Domicilio").Value)
            Ciudad = Trim(RsGral.Fields("Ciudad").Value)
            Rfc = Trim(RsGral.Fields("Rfc").Value)
        End If
        SubTotal = CDec(VB6.Format((CDec(Numerico(lblSubTot.Text)) - CDec(Numerico(lblDescuento.Text))) * TipoCambio, "###,##0.0"))
        SubTotal = CDec(VB6.Format(SubTotal, "###,##0.00"))
        Iva = CDec(VB6.Format(CDec(Numerico(lblIva.Text)) * TipoCambio, "###,##0.0"))
        Iva = CDec(VB6.Format(Iva, "###,##0.00"))
        Total = CDec(Numerico(lblImporteSubTotal.Text)) + CDec(Numerico(lblImporteRedondeo.Text))
        Total = CDec(VB6.Format(CDec(Total) * TipoCambio, "###,##0.0"))
        Total = CDec(VB6.Format(Total, "###,##0.00"))
        With Printer
            '.ScaleMode = vbMillimeters
            .FontName = "Courier New"
            .Orientation = 1
            .Height = 140 'Cambia el tamaño de la hoja en la impresora
        End With
        '''             acutal  Anterior  original
        Printer.CurrentX = 170
        Printer.CurrentY = 15
        Printer.Print(Trim(txtFolioFactura.Text))
        Printer.CurrentX = 25
        Printer.CurrentY = 24
        Printer.Print(Trim(NombreCliente))
        Printer.CurrentX = 25
        Printer.CurrentY = 29
        Printer.Print(Trim(Domicilio))
        Printer.CurrentX = 25
        Printer.CurrentY = 36
        Printer.Print(Trim(Ciudad))
        Printer.CurrentX = 125
        Printer.CurrentY = 36
        Printer.Print(Trim(Rfc))
        Printer.CurrentX = 163
        Printer.CurrentY = 35
        Printer.Print(((dtpFechaRegistro.Value)))
        Printer.CurrentX = 176
        Printer.CurrentY = 35
        Printer.Print(ModEstandar.MesLetra(dtpFechaRegistro.Value, False))
        Printer.CurrentX = 200
        Printer.CurrentY = 35
        Printer.Print((dtpFechaRegistro.Value.Year))
        CoordY = 51
        Printer.CurrentX = 52
        Printer.CurrentY = CoordY
        Printer.Print(DescEspecial)
        CoordY = 110

        If DesgloseIva = 1 Then
            Printer.CurrentX = 181
            Printer.CurrentY = CoordY
            SubTotal = CDec(VB6.Format((CDec(Numerico(lblSubTot.Text)) - CDec(Numerico(lblDescuento.Text)) + CDec(Numerico(lblImporteRedondeo.Text))) * TipoCambio, "###,##0.0"))
            'SubTotal = Format(SubTotal, "###,##0.00")
            Printer.Print((Space(15) & VB6.Format(SubTotal, "###,##0.00")))
            CoordY = CoordY + 6
            Printer.CurrentX = 181
            Printer.CurrentY = CoordY
            Printer.Print((Space(15) & Format(Iva, "###,##0.00")))
        Else
            Printer.CurrentX = 181
            Printer.CurrentY = CoordY
            ' Printer.Print Right(Space(15) & Format((SubTotal + Iva), "###,##0.00"), 13)
            SubTotal = CDec(VB6.Format((CDec(Numerico(lblSubTot.Text)) - CDec(Numerico(lblDescuento.Text)) + CDec(Numerico(lblImporteRedondeo.Text))) * TipoCambio, "###,##0.0"))
            Printer.Print((Space(15) & VB6.Format(SubTotal + Iva, "###,##0.00")))
            CoordY = CoordY + 6
            Printer.CurrentX = 181
            Printer.CurrentY = CoordY
            'Printer.Print Right(Space(15) & Format(Iva, "###,##0.00"), 13)
        End If
        CoordY = CoordY + 1
        Printer.CurrentX = 7
        Printer.CurrentY = CoordY
        Printer.Print("(PAGO EN UNA SOLA EXHIBICION)")
        CoordY = CoordY + 6
        Printer.CurrentX = 7
        Printer.CurrentY = CoordY
        Printer.Print(ModEstandar.ConLetra(CDbl(Numerico(VB6.Format(Total, "#####0.00"))), True, CStr(1)))
        Printer.CurrentX = 181
        Printer.CurrentY = CoordY
        Printer.Print((Space(15) & (Total + "###,##0.00") + 13))
        Printer.EndDoc()

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function UsoMetodo() As Boolean
        Dim I As Integer
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) <> 0 And Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLGRABADO)) = "" And Trim(.get_TextMatrix(I, C_COLEXCLUIDO)) = "" Then
                    If flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) <> "TC" And flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) <> "F" Then
                        UsoMetodo = True
                        Exit Function
                    End If
                End If
            Next
        End With
        UsoMetodo = False
    End Function

    Private Sub chkDoctoCliente_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDoctoCliente.CheckStateChanged
        If chkDoctoCliente.CheckState = System.Windows.Forms.CheckState.Checked Then
            cmdDatosFiscales.Enabled = True
        ElseIf chkDoctoCliente.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cmdDatosFiscales.Enabled = False
        End If
    End Sub

    Private Sub cmdDatosFiscales_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDatosFiscales.Click
        Me.Enabled = False
        frmFactDatosFiscales.Show()
    End Sub

    Private Sub cmdGenerarFactura_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGenerarFactura.Click
        If chkDoctoCliente.CheckState = System.Windows.Forms.CheckState.Checked Then
            If Not ExistenFoliosMarcados() Then
                MsgBox("No ha seleccionado ningun folio para la factura del punto de venta, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Sub
            End If
            If gintCodRFC = 0 And Trim(gstrNombreCliente) = "" And Trim(gstrRFCCliente) = "" Then
                MsgBox("No se han proporcionado los datos fiscales del cliente, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Sub
            Else
                If Not FoliosPendientesPtoVenta() Then Exit Sub
                GenerarFacturaPuntoVenta()
                Exit Sub
            End If
        Else
            If Not FoliosPendientesCorporativo() Then Exit Sub
            GuardarFactura()
        End If
    End Sub

    Private Sub cmdImpresionTickets_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdImpresionTickets.Click
        Dim I As Integer
        'On Error GoTo Err_Renamed
        frmFactAnalisisVentasImpresionTickets.frmFactAnalisisVentasImpresionTickets_Load(New Object, New EventArgs)

        With frmFactAnalisisVentasImpresionTickets.flexTickets
            .Rows = flexVentas.Rows
            For I = 1 To flexVentas.Rows - 1
                .set_TextMatrix(I, 0, flexVentas.get_TextMatrix(I, C_COLFOLIOVENTA))
                .set_TextMatrix(I, 1, flexVentas.get_TextMatrix(I, C_ColIMPORTE))
                .set_TextMatrix(I, 2, flexVentas.get_TextMatrix(I, C_COLCODCAJA))
                .set_TextMatrix(I, 3, flexVentas.get_TextMatrix(I, C_ColCODSUCURSAL))
                .set_TextMatrix(I, 4, flexVentas.get_TextMatrix(I, C_ColCONDICION))
                .set_TextMatrix(I, 5, flexVentas.get_TextMatrix(I, C_ColMONEDA))
                .set_TextMatrix(I, 6, Trim(flexVentas.get_TextMatrix(I, C_COLTIPOMOVTO)))
            Next I
        End With

        frmFactAnalisisVentasImpresionTickets.ShowDialog()
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub cmdImprimirFactura_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdImprimirFactura.Click
        ImprimirFactura()
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        If Me.FoliosPendientes() Then
            MsgBox("No es posible consultar otra sucursal, ya ha generado algun(s) folios adicionales" & vbNewLine & "  Para poder consultar otra sucursal debera generar la factura correspondiente", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            FueraChange = True
            dbcSucursal.Text = DescSucursal
            FueraChange = False
            Exit Sub
        End If
        If Trim(dbcSucursal.Text) = "" Or Trim(dbcSucursal.Text) <> Trim(DescSucursal) Then
            FueraChange = True
            txtCodSucursal.Text = ""
            optManual.Checked = True
            optPorcentual.Checked = False
            txtPorcentaje.Text = "0"
            txtPorcentaje.Enabled = False
            txtFolioFactura.Text = ""
            FueraChange = False
            dtpFechaRegistro.Value = Today
            lblDescripcion.Text = ""
            lblEstadoFolio.Text = ""
            lblSubTotal.Text = "0.00"
            lblRedondeo.Text = "0.00"
            lblTotal.Text = "0.00"
            lblTotalPesos.Text = "0.00"
            lblImporteSubTotal.Text = "0.00"
            lblImporteRedondeo.Text = "0.00"
            lblImporteTotal.Text = "0.00"
            lblDescuento.Text = "0.00"
            lblIva.Text = "0.00"
            lblSubTot.Text = "0.00"
            lblFactura.Text = "F00" & VB6.Format(Year(dtpFechaVenta.Value), "0000") & VB6.Format(Month(dtpFechaVenta.Value), "00") & VB6.Format((dtpFechaVenta.Value), "00") & "000000"
            lblCantidad.Text = ""
            LimpiarGridVentas()
            LimpiarGridDetalle()
            LimpiarGridPendientes()
            FueraChange = True
            txtDescripcion.Text = ""
            FueraChange = False
            txtDescripcion.ReadOnly = True
            cmdGenerarFactura.Enabled = True
            cmdImpresionTickets.Enabled = False
            cmdImprimirFactura.Enabled = False
            txtFacturaAdicional.Text = ""
            InicializaVariables()
            chkDoctoCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkDoctoCliente.Enabled = True
            chkDesglosarIva.Enabled = True
            chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked
            cmdDatosFiscales.Enabled = False
            dtpFechaVenta.Enabled = True
            dbcSucursal.Enabled = True
        End If
        gStrSql = "SELECT CodAlmacen,RTRIM(LTRIM(DescAlmacen)) AS DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,RTRIM(LTRIM(DescAlmacen)) AS DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursal)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            dtpFechaVenta.Focus()
        End If
    End Sub

    Private Sub dbcSucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcSucursal.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.TabIndex < dbcSucursal.TabIndex Then Exit Sub
        '    If frmFactAnalisisVentas.FoliosPendientes() Then
        '        MsgBox "No es posible consultar otra sucursal, ya ha generado algun(s) folios adicionales" & vbNewLine & _
        ''               "  Para poder consultar otra sucursal debera generar la factura correspondiente", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '        FueraChange = True
        '        dbcSucursal.text = DescSucursal
        '        FueraChange = False
        '        Exit Sub
        '    End If
        FueraChange = True
        gStrSql = "SELECT CodAlmacen,RTRIM(LTRIM(DescAlmacen)) AS DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
        If intCodSucursal <> 0 Then
            txtCodSucursal.Text = VB6.Format(intCodSucursal, "000")
        End If

        If Trim(dbcSucursal.Text) = "" Then
            FueraChange = False
            Exit Sub
        End If

        If CDbl(Numerico(txtCodSucursal.Text)) <> 0 And Trim(dbcSucursal.Text) <> "" Then
            If CDbl(Numerico(txtCodSucursal.Text)) <> CodSucursal Or Trim(dbcSucursal.Text) <> DescSucursal Or CDate(VB6.Format(dtpFechaVenta.Value, "dd/MM/yyyy")) <> CDate(VB6.Format(IIf(Fecha = "", "01/01/1900", Fecha), "dd/MM/yyyy")) Then
                lblFactura.Text = "F" & VB6.Format(txtCodSucursal.Text, "00") & VB6.Format(Year(dtpFechaVenta.Value), "0000") & VB6.Format(Month(dtpFechaVenta.Value), "00") & VB6.Format((dtpFechaVenta.Value), "00") & "000000"
                txtFolioFactura.Text = ""
                mblnNuevo = False
                'flexVentas.SetFocus
                FueraChange = True
                optManual.Checked = True
                optPorcentual.Checked = False
                txtPorcentaje.Text = "0"
                txtPorcentaje.Enabled = False
                txtFolioFactura.Text = ""
                FueraChange = False
                dtpFechaRegistro.Value = Today
                lblDescripcion.Text = ""
                lblEstadoFolio.Text = ""
                lblSubTotal.Text = "0.00"
                lblRedondeo.Text = "0.00"
                lblTotal.Text = "0.00"
                lblTotalPesos.Text = "0.00"
                lblImporteSubTotal.Text = "0.00"
                lblImporteRedondeo.Text = "0.00"
                lblImporteTotal.Text = "0.00"
                lblDescuento.Text = "0.00"
                lblIva.Text = "0.00"
                lblSubTot.Text = "0.00"
                lblFactura.Text = "F00" & (Year(dtpFechaVenta.Value) + "0000") & (Month(dtpFechaVenta.Value) + "00") & ((dtpFechaVenta.Value) + "00") & "000000"
                lblCantidad.Text = ""
                LimpiarGridVentas()
                LimpiarGridDetalle()
                LimpiarGridPendientes()
                FueraChange = True
                txtDescripcion.Text = ""
                FueraChange = False
                txtDescripcion.ReadOnly = True
                cmdGenerarFactura.Enabled = True
                cmdImpresionTickets.Enabled = False
                cmdImprimirFactura.Enabled = False
                txtFacturaAdicional.Text = ""
                InicializaVariables()
                chkDoctoCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkDoctoCliente.Enabled = True
                chkDesglosarIva.Enabled = True
                chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked
                cmdDatosFiscales.Enabled = False
                CargaVentas()
                Fecha = VB6.Format(dtpFechaVenta.Value, "dd/mmm/yyyy")
                CodSucursal = CShort(Numerico(txtCodSucursal.Text))
                DescSucursal = Trim(dbcSucursal.Text)
                dtpFechaVenta.Enabled = True
                dbcSucursal.Enabled = True
                FueraChange = False
                Exit Sub
            End If
        End If
        FueraChange = False
    End Sub

    Private Sub DtpDesde_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DtpDesde.Leave
        If dbcSucursal.Enabled = False Then Exit Sub
        If DtpDesde.Value > DtpHasta.Value Then
            MsgBox("La fecha inicial no puede ser mayor que la fecha final. " & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            DtpDesde.Value = DtpDesde.Tag
            If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "DtpHasta" Then
                DtpDesde.Focus()
            End If
            Exit Sub
        End If
        DtpDesde.Tag = DtpDesde.Value
    End Sub

    Private Sub dtpFechaVenta_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaVenta.CursorChanged
        If Me.FoliosPendientes() Then
            MsgBox("No es posible consultar otra fecha, ya ha generado algun(s) folios adicionales" & vbNewLine & "  Para poder consultar otra fecha debera generar la factura correspondiente", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dtpFechaVenta.Value = Fecha
            Exit Sub
        End If
        '    If Numerico(txtCodSucursal) <> 0 And Trim(dbcSucursal.text) <> "" Then
        '        If Format(dtpFechaVenta, "dd/mmm/yyyy") <> Fecha Or Numerico(txtCodSucursal) <> CodSucursal Or Trim(dbcSucursal) <> DescSucursal Then
        '            FueraChange = True
        '            txtFolioFactura = ""
        '            FueraChange = False
        '            lblSubTotal.Caption = "0.00"
        '            lblRedondeo.Caption = "0.00"
        '            lblTotal.Caption = "0.00"
        '            lblTotalPesos.Caption = "0.00"
        '            lblImporteSubTotal = "0.00"
        '            lblImporteRedondeo = "0.00"
        '            lblImporteTotal = "0.00"
        '            lblDescuento = "0.00"
        '            lblIva = "0.00"
        '            lblSubTot = "0.00"
        '            lblFactura.Caption = "F00" & Format(Year(dtpFechaVenta), "0000") & Format(Month(dtpFechaVenta), "00") & Format(Day(dtpFechaVenta), "00") & "000000"
        '            LimpiarGridVentas
        '            CargaVentas
        '            Fecha = Format(dtpFechaVenta, "dd/mmm/yyyy")
        '            CodSucursal = Numerico(txtCodSucursal)
        '            DescSucursal = Trim(dbcSucursal.text)
        '        End If
        '    End If
    End Sub

    Private Sub dtpFechaVenta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaVenta.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaVenta_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaVenta.Leave
        On Error GoTo Err_Renamed
        '    If frmFactAnalisisVentas.FoliosPendientes() Then
        '        MsgBox "No es posible consultar otra fecha, ya ha generado algun(s) folios adicionales" & vbNewLine & _
        ''               "  Para poder consultar otra fecha debera generar la factura correspondiente", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '        dtpFechaVenta = Fecha
        '        Exit Sub
        '    End If
        If CDbl(Numerico(txtCodSucursal.Text)) <> 0 And Trim(dbcSucursal.Text) <> "" Then
            If CDbl(Numerico(txtCodSucursal.Text)) <> CodSucursal Or Trim(dbcSucursal.Text) <> DescSucursal Or CDate(VB6.Format(dtpFechaVenta.Value, "dd/MM/yyyy")) <> CDate(VB6.Format(IIf(Fecha = "", "01/01/1900", Fecha), "dd/MM/yyyy")) Then
                lblFactura.Text = "F" & VB6.Format(txtCodSucursal.Text, "00") & VB6.Format(Year(dtpFechaVenta.Value), "0000") & VB6.Format(Month(dtpFechaVenta.Value), "00") & VB6.Format((dtpFechaVenta.Value), "00") & "000000"
                txtFolioFactura.Text = ""
                mblnNuevo = False
                'flexVentas.SetFocus
                FueraChange = True
                optManual.Checked = True
                optPorcentual.Checked = False
                txtPorcentaje.Text = "0"
                txtPorcentaje.Enabled = False
                txtFolioFactura.Text = ""
                FueraChange = False
                dtpFechaRegistro.Value = Today
                lblDescripcion.Text = ""
                lblEstadoFolio.Text = ""
                lblSubTotal.Text = "0.00"
                lblRedondeo.Text = "0.00"
                lblTotal.Text = "0.00"
                lblTotalPesos.Text = "0.00"
                lblImporteSubTotal.Text = "0.00"
                lblImporteRedondeo.Text = "0.00"
                lblImporteTotal.Text = "0.00"
                lblDescuento.Text = "0.00"
                lblIva.Text = "0.00"
                lblSubTot.Text = "0.00"
                lblFactura.Text = "F00" & (Year(dtpFechaVenta.Value) + "0000") & (Month(dtpFechaVenta.Value) + "00") & (dtpFechaVenta.Value + "00") & "000000"
                lblCantidad.Text = ""
                LimpiarGridVentas()
                LimpiarGridDetalle()
                LimpiarGridPendientes()
                FueraChange = True
                txtDescripcion.Text = ""
                FueraChange = False
                txtDescripcion.ReadOnly = True
                cmdGenerarFactura.Enabled = True
                cmdImpresionTickets.Enabled = False
                cmdImprimirFactura.Enabled = False
                txtFacturaAdicional.Text = ""
                InicializaVariables()
                chkDoctoCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkDoctoCliente.Enabled = True
                chkDesglosarIva.Enabled = True
                chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked
                cmdDatosFiscales.Enabled = False
                DeterminaRango()
                CargaVentas()
                Fecha = VB6.Format(dtpFechaVenta.Value, "dd/mmm/yyyy")
                CodSucursal = CShort(Numerico(txtCodSucursal.Text))
                DescSucursal = Trim(dbcSucursal.Text)
                dtpFechaVenta.Enabled = True
                dbcSucursal.Enabled = True
                FueraChange = False
                Exit Sub
            End If
        End If
        FueraChange = False
        DeterminaRango()
        'CargarVentasPendientes
Err_Renamed:
        If Err.Number <> 0 Then ModErrores.Errores()
    End Sub

    Private Sub dtpHasta_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DtpHasta.Leave
        If dbcSucursal.Enabled = False Then Exit Sub
        If DtpHasta.Value >= dtpFechaVenta.Value Then
            MsgBox("La fecha final no puede ser mayor o igual que la fecha de venta. " & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            DtpHasta.Value = DtpHasta.Tag
            DtpHasta.Focus()
            Exit Sub
        End If
        DtpHasta.Tag = DtpHasta.Value
        If CDbl(Numerico(txtCodSucursal.Text)) <> 0 Then
            CargarVentasPendientes()
        End If
    End Sub

    Private Sub flexDetalleVenta_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalleVenta.ClickEvent
        flexDetalleVenta.Row = flexDetalleVenta.Row
        flexDetalleVenta.Col = flexDetalleVenta.Col
    End Sub

    Private Sub flexDetalleVenta_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalleVenta.DblClick
        With flexVentas
            If Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) <> "TC" And Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) <> "F" And Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) <> "" And Trim(.get_TextMatrix(.Row, C_COLGRABADO)) = "" And Not mblnNuevo Then
                ExcluirPartida()
                If EstaExcluido() Then
                    ExcluirFolio()
                    flexDetalleVenta.Focus()
                    flexDetalleVenta_EnterCell(flexDetalleVenta, New System.EventArgs())
                End If
            End If
        End With
    End Sub

    Private Sub flexDetalleVenta_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalleVenta.EnterCell
        txtPorcentaje.Text = VB6.Format(Numerico(flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLPORCENTAJEADICIONAL)), "###,##0")
        If System.Drawing.ColorTranslator.ToOle(flexDetalleVenta.CellBackColor) = System.Drawing.ColorTranslator.ToOle(lblExcluido.BackColor) Then
            ColorAnte = "Exc"
        Else
            ColorAnte = "Bln"
        End If
        flexDetalleVenta.CellBackColor = flexDetalleVenta.BackColorSel
        flexDetalleVenta.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000009)
        If flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLCODIGOARTICULO) = "" Then
            txtDescripcion.ReadOnly = True
        End If
        lblDescripcion.Text = flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLDESCRIPCION)
        lblCantidad.Text = CStr(CDbl(Numerico(flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLCANTIDAD))) - CDbl(Numerico(flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLCANTIDADDEVOL))))
        FueraChange = True
        txtDescripcion.Text = flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLNUEVADESCRIPCION)
        FueraChange = False
        'Or Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLGRABADO)) <> "" Or Trim(flexDetalleVenta.TextMatrix(flexDetalleVenta.Row, C_COLGRAB)) <> ""
        If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) = "" Or (CDbl(Numerico(flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLCANTIDAD))) - CDbl(Numerico(flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLCANTIDADDEVOL)))) = 0 Then
            txtDescripcion.ReadOnly = True
        Else
            txtDescripcion.ReadOnly = False
        End If
        If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) = "F" Then
            txtDescripcion.ReadOnly = True
        End If
        If flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLEXCLUIDO) = "EXCLUIDO" Then
            txtDescripcion.ReadOnly = True
        End If
    End Sub

    Private Sub flexDetalleVenta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalleVenta.Enter
        flexDetalleVenta_EnterCell(flexDetalleVenta, New System.EventArgs())
        '    If flexVentas.TextMatrix(flexVentas.Row, C_COLEXCLUIDO) <> "EXCLUIDO" Then
        '        flexDetalleVenta_Click
        '        flexDetalleVenta.CellBackColor = flexDetalleVenta.BackColorSel
        '        flexDetalleVenta.CellForeColor = &H80000009
        '    End If
    End Sub

    Private Sub flexDetalleVenta_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexDetalleVenta.KeyDownEvent
        Dim I As Integer
        If eventArgs.keyCode = System.Windows.Forms.Keys.Space Then
            flexDetalleVenta_LeaveCell(flexDetalleVenta, New System.EventArgs())
            flexDetalleVenta_KeyPressEvent(flexDetalleVenta, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
        End If
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete Then
            flexDetalleVenta_DblClick(flexDetalleVenta, New System.EventArgs())
        End If
        If eventArgs.keyCode = System.Windows.Forms.Keys.F4 Then
            With flexVentasPendientes
                'If Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLGRABADO)) = "S" Then Exit Sub
                If Trim(flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLFOLIOAGREGADO)) = "" Then Exit Sub
                CargaDetalle()
                CargarVentasPendientes()
                flexDetalleVenta.Focus()
            End With
        End If
    End Sub

    Private Sub flexDetalleVenta_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexDetalleVenta.KeyPressEvent
        With flexDetalleVenta
            'If flexVentas.TextMatrix(flexVentas.Row, C_COLGRABADO) = "S" Then Exit Sub          'And .TextMatrix(.Row, C_COLGRAB) = "" And Not mblnNuevo
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape And optManual.Checked = True And flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) <> "F" And flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) <> "TC" And Trim(.get_TextMatrix(.Row, C_COLEXCLUIDO)) = "" And (CDbl(Numerico(.get_TextMatrix(.Row, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(.Row, C_COLCANTIDADDEVOL)))) > 0 Then
                flexDetalleVenta.CellBackColor = flexDetalleVenta.BackColor
                flexDetalleVenta.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                .Col = C_COLNUEVOPRECIOPUBLICO
                If .get_TextMatrix(.Row, 0) <> "" Then
                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                    MSHFlexGridEdit(flexDetalleVenta, txtFlex, eventArgs.keyAscii)
                    CambiarFormatoTxtenCaptura()
                    If Len(Trim(txtFlex.Text)) = 1 Then
                        'System.Windows.Forms.SendKeys.Send("{RIGHT}")
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub flexDetalleVenta_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalleVenta.LeaveCell
        If ColorAnte = "Bln" Then
            flexDetalleVenta.CellBackColor = flexDetalleVenta.BackColor
            flexDetalleVenta.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
        ElseIf ColorAnte = "Exc" Then
            flexDetalleVenta.CellBackColor = lblExcluido.BackColor
            flexDetalleVenta.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
        End If
    End Sub

    Private Sub flexDetalleVenta_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalleVenta.Leave
        txtPorcentaje.Text = "0"
        If ColorAnte = "Bln" Then
            flexDetalleVenta.CellBackColor = flexDetalleVenta.BackColor
            flexDetalleVenta.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
        ElseIf ColorAnte = "Exc" Then
            flexDetalleVenta.CellBackColor = lblExcluido.BackColor
            flexDetalleVenta.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
        End If
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtDescripcion" Then
            lblCantidad.Text = ""
            lblDescripcion.Text = ""
            FueraChange = True
            txtDescripcion.Text = ""
            FueraChange = False
        End If
    End Sub

    Private Sub flexVentas_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentas.DblClick
        If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) <> "F" And Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) <> "TC" And Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTATUSADICIONAL)) <> "O" And Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA)) <> "" And Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLGRABADO)) = "" And Not mblnNuevo Then
            ExcluirFolio()
        End If
    End Sub

    Private Sub flexVentas_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentas.EnterCell
        If blnFueraCell Then
            flexVentas.Row = RenAct
            blnFueraCell = False
            Exit Sub
        End If
        If System.Drawing.ColorTranslator.ToOle(flexVentas.CellBackColor) = System.Drawing.ColorTranslator.ToOle(lblExcluido.BackColor) Then
            ColorAnteFolio = "Exc"
        ElseIf System.Drawing.ColorTranslator.ToOle(flexVentas.CellBackColor) = System.Drawing.ColorTranslator.ToOle(lblFacturar.BackColor) Then
            ColorAnteFolio = "FAC"
        Else
            ColorAnteFolio = "Bln"
        End If
        '    flexVentas.Col = 0
        '    flexVentas.CellBackColor = flexVentas.BackColorSel
        '    flexVentas.CellForeColor = &H80000009
        '    flexVentas.Col = 1
        '    flexVentas.CellBackColor = flexVentas.BackColorSel
        '    flexVentas.CellForeColor = &H80000009
        '    DoEvents
        With flexVentas
            If .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "F" Then
                lblEstadoFolio.Text = "FACTURADO"
                txtDescripcion.ReadOnly = True
            ElseIf .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "TC" Then
                lblEstadoFolio.Text = "PAG TARJ CRED"
                txtDescripcion.ReadOnly = False
            ElseIf .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "N" And Trim(.get_TextMatrix(.Row, C_COLFOLIOEXCLUIDO)) = "" And Trim(.get_TextMatrix(.Row, C_COLGRABADO)) = "" Then
                lblEstadoFolio.Text = "AFECTABLE"
                txtDescripcion.ReadOnly = False
            ElseIf .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "N" And Trim(.get_TextMatrix(.Row, C_COLFOLIOEXCLUIDO)) = "" And Trim(.get_TextMatrix(.Row, C_COLGRABADO)) = "S" Then
                lblEstadoFolio.Text = "AFECTADO"
                'Antes true
                txtDescripcion.ReadOnly = False
            ElseIf .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "N" And Trim(.get_TextMatrix(.Row, C_COLFOLIOEXCLUIDO)) = "EXCLUIDO" Then
                lblEstadoFolio.Text = "EXCLUIDO"
                txtDescripcion.ReadOnly = False
            ElseIf .get_TextMatrix(.Row, C_COLESTADOFOLIO) = "" Then
                lblEstadoFolio.Text = ""
                txtDescripcion.ReadOnly = True
            End If
            lblMoneda.Text = IIf(Trim(.get_TextMatrix(.Row, C_ColMONEDA)) = "P", "PESOS", IIf(Trim(.get_TextMatrix(.Row, C_ColMONEDA)) = "D", "DOLARES", ""))
            If Trim(.get_TextMatrix(.Row, C_COLFOLIOVENTA)) <> "" And Not mblnNuevo Then
                LimpiarGridDetalle()
                CargaDetalle()
                CargarVentasPendientes()
                '            DoEvents
            ElseIf Trim(.get_TextMatrix(.Row, C_COLFOLIOVENTA)) <> "" And mblnNuevo Then
                LimpiarGridDetalle()
                CargaDetalleAdicional()
                '            DoEvents
            Else
                LimpiarGridDetalle()
                lblSubTotal.Text = "0.00"
                lblRedondeo.Text = "0.00"
                lblTotal.Text = "0.00"
                lblTotalPesos.Text = "0.00"
            End If
            FueraChange = True
            txtDescripcion.Text = ""
            FueraChange = False
        End With
        txtPorcentaje.Text = "0"
    End Sub

    Private Sub flexVentas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentas.Enter
        If flexVentas.Row = RenAnterior Then Exit Sub
        Pon_Tool()
        lblDescripcion.Text = ""
        txtPorcentaje.Text = "0"
    End Sub

    Private Sub flexVentas_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexVentas.KeyDownEvent
        Dim I As Integer
        Dim Col As Integer
        If eventArgs.keyCode = System.Windows.Forms.Keys.F6 Then
            With flexVentas
                If Trim(.get_TextMatrix(.Row, C_COLFOLIOVENTA)) = "" Then Exit Sub
                'If Trim(.TextMatrix(.Row, C_COLGRABADO)) = "S" Then Exit Sub
                If Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) = "F" Then Exit Sub
                Col = .Col
                If Trim(.get_TextMatrix(.Row, C_COLFACTURAPTOVTA)) = "" Then
                    For I = 0 To 1
                        .Col = I
                        .CellBackColor = lblFacturar.BackColor
                        .CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                    Next
                    .Col = 0
                    .set_TextMatrix(.Row, C_COLFACTURAPTOVTA, "S")
                    If Trim(.get_TextMatrix(.Row, C_COLINCFACTURA)) = "S" Then
                        .set_TextMatrix(.Row, C_COLINCFACTURA, "N")
                        'RecalculaImportesFactura
                        'GenerarFactPtoVenta = True
                        If Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) = "TC" Then
                            .set_TextMatrix(.Row, C_COLCAMBIOS, "S")
                        End If
                        mblnFactura = False
                    End If
                    ColorAnteFolio = "FAC"
                    CuentaPartidas()
                Else
                    If Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) = "F" Then Exit Sub
                    'Trim(.TextMatrix(.Row, C_COLESTADOFOLIO)) = "TC" And
                    If Trim(.get_TextMatrix(.Row, C_COLFACTURAPTOVTA)) = "S" Then
                        For I = 0 To 1
                            .Col = I
                            .CellBackColor = System.Drawing.ColorTranslator.FromOle(&H8000000E)
                            '.CellForeColor = &H8000000E
                        Next
                        .Col = 0
                        .set_TextMatrix(.Row, C_COLFACTURAPTOVTA, "")
                        If Trim(.get_TextMatrix(.Row, C_COLINCFACTURA)) = "N" Then
                            .set_TextMatrix(.Row, C_COLINCFACTURA, "S")
                            'RecalculaImportesFactura
                            'GenerarFactPtoVenta = True
                            If Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) = "TC" Then
                                .set_TextMatrix(.Row, C_COLCAMBIOS, "")
                            End If
                            mblnFactura = True
                        End If
                        ColorAnteFolio = "Bln"
                        DesCuentaPartidas()
                        Exit Sub
                    End If
                    '                If Trim(.TextMatrix(.Row, C_COLFOLIOADICIONAL)) <> "" Then Exit Sub
                    For I = 0 To 1
                        .Col = I
                        .CellBackColor = System.Drawing.ColorTranslator.FromOle(&H8000000E)
                        '.CellForeColor = &H8000000E
                    Next
                    .Col = 0
                    .set_TextMatrix(.Row, C_COLFACTURAPTOVTA, "")
                    ColorAnteFolio = "Bln"
                    DesCuentaPartidas()
                End If
            End With
        End If
    End Sub

    Private Sub flexVentas_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentas.LeaveCell
        Dim I As Integer
        If HayCambios() And Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOEXCLUIDO)) = "" Then
            Select Case MsgBox("¿Deseas guardar los cambios hechos a este folio de venta?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    If Not GuardarFolio() Then
                        blnFueraCell = True
                        RenAct = flexVentas.Row
                        Exit Sub
                    End If
                Case MsgBoxResult.No
                    With flexVentas
                        If Trim(.get_TextMatrix(.Row, C_COLESTADOFOLIO)) <> "TC" Then
                            .set_TextMatrix(.Row, C_COLSUBTOTALADICIONAL, "")
                            .set_TextMatrix(.Row, C_COLDESCUENTOADICIONAL, "")
                            .set_TextMatrix(.Row, C_COLIVAADICIONAL, "")
                            .set_TextMatrix(.Row, C_COLTOTALADICIONAL, "")
                            .set_TextMatrix(.Row, C_COLREDONDEOADICIONAL, "")
                        End If
                        CargarVentasPendientes()
                    End With
            End Select
        Else
            If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFACTURAPTOVTA)) = "S" Then
                With flexDetalleVenta
                    For I = 2 To .Rows - 1
                        If Trim(.get_TextMatrix(I, C_COLCODARTICULO)) <> "" Then
                            If Trim(.get_TextMatrix(I, C_COLFOLIOAGREGADO)) <> "" Then
                                numPartidas = numPartidas - 1
                            End If
                        End If
                    Next
                End With
            End If
        End If
        If ColorAnteFolio = "Bln" Then
            '        flexVentas.Col = 0
            '        flexVentas.CellBackColor = flexVentas.BackColor
            '        flexVentas.CellForeColor = &H80000008
            '        flexVentas.Col = 1
            '        flexVentas.CellBackColor = flexVentas.BackColor
            '        flexVentas.CellForeColor = &H80000008
            '        DoEvents
        ElseIf Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOEXCLUIDO)) = "EXCLUIDO" Then
            flexVentas.Col = 0
            flexVentas.CellBackColor = lblExcluido.BackColor
            flexVentas.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
            flexVentas.Col = 1
            flexVentas.CellBackColor = lblExcluido.BackColor
            flexVentas.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
            System.Windows.Forms.Application.DoEvents()
        ElseIf Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFACTURAPTOVTA)) = "S" Then
            flexVentas.Col = 0
            flexVentas.CellBackColor = lblFacturar.BackColor
            flexVentas.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
            flexVentas.Col = 1
            flexVentas.CellBackColor = lblFacturar.BackColor
            flexVentas.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
            System.Windows.Forms.Application.DoEvents()
        End If
    End Sub

    Private Sub flexVentas_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentas.Leave
        RenAnterior = flexVentas.Row
    End Sub

    Private Sub flexVentas_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentas.Scroll
        flexVentas.Row = flexVentas.Row
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub flexVentasPendientes_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentasPendientes.DblClick
        Dim I As Integer
        If Trim(flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, 0)) = "" Then Exit Sub
        If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOVENTA)) = "" Then Exit Sub
        If flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) = "F" Then Exit Sub
        If flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) = "TC" Then Exit Sub
        If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLFOLIOEXCLUIDO)) = "EXCLUIDO" Then Exit Sub
        'If flexVentas.TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) = "N" And Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLFOLIOEXCLUIDO)) = "" And (Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLGRABADO)) = "" Or Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLGRABADO)) = "S") Then Exit Sub
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODIGOARTICULO)) = "" Then
                    EnviaPartida((I))
                    Exit Sub
                End If
            Next
            .Rows = .Rows + 1
            I = .Rows - 1
            EnviaPartida((I))
        End With
    End Sub

    Private Sub flexVentasPendientes_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentasPendientes.EnterCell
        flexVentasPendientes.CellBackColor = flexVentasPendientes.BackColorSel
        flexVentasPendientes.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000009)
        lblDesc.Text = flexVentasPendientes.get_TextMatrix(flexVentasPendientes.Row, C_COLDESCARTICULO)
    End Sub

    Private Sub flexVentasPendientes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentasPendientes.Enter
        flexVentasPendientes_EnterCell(flexVentasPendientes, New System.EventArgs())
    End Sub

    Private Sub flexVentasPendientes_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexVentasPendientes.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.F5 Then
            flexVentasPendientes_DblClick(flexVentasPendientes, New System.EventArgs())
        End If
    End Sub

    Private Sub flexVentasPendientes_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentasPendientes.LeaveCell
        flexVentasPendientes.CellBackColor = flexVentasPendientes.BackColor
        flexVentasPendientes.CellForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
    End Sub

    Private Sub flexVentasPendientes_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVentasPendientes.Leave
        flexVentasPendientes_LeaveCell(flexVentasPendientes, New System.EventArgs())
        lblDesc.Text = ""
    End Sub

    Private Sub frmFactAnalisisVentas_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmFactAnalisisVentas_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmFactAnalisisVentas_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'If System.Windows.Forms.Form.ActiveForm.Name <> "frmFactAnalisisVentas" Then
        '    Exit Sub
        'End If
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                '            If Screen.ActiveForm.ActiveControl.Name = "dbcSucursal" Then
                '                flexVentas.SetFocus
                '                Exit Sub
                '            End If
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "optManual" Or System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "optPorcentual" Then
                    If dbcSucursal.Enabled = False Then
                        mblnSalir = True
                        Me.Close()
                        Exit Sub
                    End If
                End If
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dtpFechaVenta" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
            Case System.Windows.Forms.Keys.Delete
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "flexVentas" Then
                    flexVentas_DblClick(flexVentas, New System.EventArgs())
                End If
        End Select
    End Sub

    Private Sub frmFactAnalisisVentas_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Public Sub frmFactAnalisisVentas_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) - 300)
        dtpFechaVenta.MinDate = C_FECHAINICIAL
        dtpFechaVenta.MaxDate = C_FECHAFINAL
        flexVentas.Clear()
        EncabezadoVentas()
        flexDetalleVenta.Clear()
        EncabezadoDetalle()
        flexVentasPendientes.Clear()
        flexVentasPendientes.Rows = 12
        EncabezadoVentasPendientes()
        Nuevo()
    End Sub

    Private Sub frmFactAnalisisVentas_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        '            System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmFactAnalisisVentas_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Dim bytR As Byte
        If Me.FoliosPendientes() Then
            MsgBox("No es posible salir, ya ha generado algun(s) folios adicionales" & vbNewLine & "  Para poder salir debera generar la factura correspondiente", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            'Cancel = 1
            Exit Sub
        End If
        'Pregunta por los cambios en la factura
        '    If mblnNuevo And Trim(txtFacturaAdicional.text) <> txtFacturaAdicional.Tag Then
        '        bytR = MsgBox("Se han hecho cambios a la factura" & vbNewLine _
        ''            & "¿Desea guardar los cambios?", vbYesNoCancel + vbQuestion)
        '        Select Case bytR
        '        Case vbYes
        '            Cancel = Not GuardarCambiosFactura
        '        Case vbCancel
        '            Cancel = True
        '        End Select
        '    End If

        'If CBool(Cancel) Then Exit Sub

        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        'MDIMenuPrincipalCorpo.mnuFacturacionOpc(0).Enabled = True
    End Sub

    Private Sub optManual_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optManual.CheckedChanged
        If eventSender.Checked Then
            If FueraChange Then Exit Sub
            If UsoMetodo() And Not mblnCambiarMetodo Then
                Select Case MsgBox("Usted Esta Aplicando el Método Porcentual, ¿Desea Cambiar al Método Manual?" & Chr(13) & "     !!!ADVERTENCIA¡¡¡ los Movimientos Que ha Realizado se Cancelaran. ", MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        txtPorcentaje.Text = "0"
                        txtPorcentaje.Enabled = False
                        InicializaImporte()
                    Case MsgBoxResult.No
                        txtPorcentaje.Enabled = True
                        mblnCambiarMetodo = True
                        System.Windows.Forms.SendKeys.Send("{DOWN}")
                        Exit Sub
                    Case MsgBoxResult.Cancel
                        txtPorcentaje.Enabled = True
                        mblnCambiarMetodo = True
                        System.Windows.Forms.SendKeys.Send("{DOWN}")
                        Exit Sub
                End Select
            Else
                txtPorcentaje.Text = "0"
                txtPorcentaje.Enabled = False
            End If
            mblnCambiarMetodo = False
        End If
    End Sub

    Private Sub optManual_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optManual.Enter
        Pon_Tool()
    End Sub

    Private Sub optPorcentual_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPorcentual.CheckedChanged
        If eventSender.Checked Then
            If FueraChange Then Exit Sub
            If UsoMetodo() And Not mblnCambiarMetodo Then
                Select Case MsgBox("Usted Esta Aplicando el Método Manual, ¿Desea Cambiar al Método Porcentual?" & Chr(13) & "     !!!ADVERTENCIA¡¡¡ los Movimientos Que ha Realizado se Cancelaran. ", MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        txtPorcentaje.Enabled = True
                        InicializaImporte()
                    Case MsgBoxResult.No
                        txtPorcentaje.Enabled = False
                        mblnCambiarMetodo = True
                        'System.Windows.Forms.SendKeys.Send("{UP}")
                        Exit Sub
                    Case MsgBoxResult.Cancel
                        txtPorcentaje.Enabled = False
                        mblnCambiarMetodo = True
                        'System.Windows.Forms.SendKeys.Send("{UP}")
                        Exit Sub
                End Select
            Else
                txtPorcentaje.Enabled = True
            End If
            mblnCambiarMetodo = False
        End If
    End Sub

    Private Sub optPorcentual_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPorcentual.Enter
        Pon_Tool()
    End Sub

    'Private Sub txtCodSucursal_Change()
    '    If FueraChange Then Exit Sub
    '    If FoliosPendientes() Then
    '        MsgBox "No es posible consultar otra sucursal, ya ha generado algun(s) folios adicionales" & vbNewLine & _
    ''               "  Para poder consultar otra sucursal debera generar la factura correspondiente", vbOKOnly + vbInformation, gstrNombCortoEmpresa
    '        FueraChange = True
    '        txtCodSucursal = Format(CodSucursal, "000")
    '        FueraChange = False
    '        Exit Sub
    '    End If
    '    FueraChange = True
    '    dbcSucursal.text = ""
    '    FueraChange = False
    '    If Numerico(txtCodSucursal) = 0 Then
    '        Nuevo
    '    End If
    'End Sub

    'Private Sub txtCodSucursal_GotFocus()
    '    Pon_Tool
    'End Sub

    'Private Sub txtCodsucursal_KeyPress(KeyAscii As Integer)
    '    ModEstandar.gp_CampoNumerico KeyAscii
    'End Sub

    'Private Sub txtCodSucursal_LostFocus()
    '    If Numerico(txtCodSucursal) <> 0 Then
    '        If Numerico(txtCodSucursal) <> CodSucursal Then
    '            FueraChange = True
    '            txtFolioFactura = ""
    '            lblFactura.Caption = "F" & Format(txtCodSucursal, "00") & Format(Year(dtpFechaVenta), "0000") & Format(Month(dtpFechaVenta), "00") & Format(Day(dtpFechaVenta), "00") & "000000"
    '            Fecha = Format(dtpFechaVenta, "dd/mmm/yyyy")
    '            CodSucursal = Numerico(txtCodSucursal)
    '            DescSucursal = Trim(dbcSucursal.text)
    '            mblnNuevo = False
    '            BuscaSucursal
    '            FueraChange = False
    '            flexVentas.SetFocus
    '        End If
    '    End If
    'End Sub

    Private Sub txtDescripcion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.TextChanged
        If FueraChange Then Exit Sub
        flexDetalleVenta.set_TextMatrix(flexDetalleVenta.Row, C_COLNUEVADESCRIPCION, txtDescripcion.Text)
        If Trim(flexDetalleVenta.get_TextMatrix(flexDetalleVenta.Row, C_COLNUEVADESCRIPCION)) <> "" Then
            flexDetalleVenta.set_TextMatrix(flexDetalleVenta.Row, C_COLMODIFICADO, "M")
        Else
            flexDetalleVenta.set_TextMatrix(flexDetalleVenta.Row, C_COLMODIFICADO, "")
        End If
    End Sub

    Private Sub txtDescripcion_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Leave
        lblCantidad.Text = ""
        lblDescripcion.Text = ""
        FueraChange = True
        txtDescripcion.Text = ""
        FueraChange = False
    End Sub

    Private Sub txtFacturaAdicional_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFacturaAdicional.TextChanged
        'gblnCambiosAnalisis = True
    End Sub

    Private Sub txtFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Enter
        SelTextoTxt(txtFlex)
        Pon_Tool()
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            With flexDetalleVenta
                '            If CCur(Numerico(txtFlex)) > CCur(Numerico(.TextMatrix(.Row, C_ColPRECIOPUBLICO))) Then
                '                MsgBox "El nuevo precio público no debe ser mayor que el real, Favor de verificar...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                '                txtFlex = "0.00"
                '                Exit Sub
                '            End If
                .Text = VB6.Format(txtFlex.Text, "###,##0.00")
                .set_TextMatrix(.Row, C_COLPRECIOLISTAADICIONAL, VB6.Format(txtFlex.Text, "###,##0.0000"))
                If CDbl(Numerico(txtFlex.Text)) <> 0 And Trim(.get_TextMatrix(.Row, C_COLNUEVADESCRIPCION)) <> "" Then
                    .set_TextMatrix(.Row, C_COLMODIFICADO, "M")
                Else
                    .set_TextMatrix(.Row, C_COLMODIFICADO, "")
                End If
                txtFlex.Text = ""
                txtFlex.Visible = False
                CalculaImportes()
                .Col = .Col + 1
                .Focus()
            End With
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
            txtFlex.Visible = False
            txtFlex.Text = ""
            flexDetalleVenta.Focus()
        End If
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtFlex.Text, KeyAscii, 13, 2, (txtFlex.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Private Sub txtFolioFactura_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioFactura.TextChanged
        If FueraChange Then Exit Sub
        If mblnNuevo Then
            NuevaFactura()
            mblnNuevo = True
        End If
    End Sub

    Private Sub txtFolioFactura_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioFactura.Enter
        strControlActual = UCase("txtFolioFactura")
        Pon_Tool()
        SelTextoTxt(txtFolioFactura)
    End Sub

    Private Sub txtFolioFactura_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioFactura.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        If (FoliosPendientes()) And Trim(txtFolioFactura.Text) <> "" Then
            MsgBox("No es posible consultar facturas, ya ha generado algun(os) folios adicionales" & vbNewLine & "  Para poder consultar facturas debera generar la factura correspondiente", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            If Trim(txtFolioFactura.Text) <> "" Then
                LimpiarGridDetalle()
                LimpiarGridPendientes()
                LimpiarGridVentas()
                LlenaDatos()
            End If
        End If
    End Sub

    Private Sub txtPorcentaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcentaje.Enter
        Pon_Tool()
        SelTextoTxt(txtPorcentaje)
    End Sub

    Private Sub txtPorcentaje_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPorcentaje.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPorcentaje_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcentaje.Leave
        Dim I As Integer
        Dim Total As Double
        Dim TotalPesos As Double
        Dim RedondeoDolares As Double
        Dim RedondeoPesos As Double
        Dim PorcentajeAdicional As Double
        Dim Inicializa As Boolean
        Inicializa = False
        If EstaVacia() Then Exit Sub
        If CDbl(Numerico(txtPorcentaje.Text)) = 0 Then Exit Sub
        txtPorcentaje.Text = VB6.Format(Numerico(txtPorcentaje.Text), "###,##0")
        PorcentajeAdicional = (1 - CDbl(VB6.Format(CDbl(Numerico(txtPorcentaje.Text)) / 100, "#####0.0000")))
        With flexDetalleVenta
            For I = 2 To .Rows - 1
                'If Trim(.TextMatrix(I, C_COLGRAB)) <> "" Then Exit Sub   'And Trim(flexVentas.TextMatrix(flexVentas.Row, C_COLGRABADO)) = ""
                If Trim(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO)) <> "" And Trim(.get_TextMatrix(I, C_COLEXCLUIDO)) <> "EXCLUIDO" And (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) > 0 And flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) <> "F" And flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO) <> "TC" Then
                    If Not Inicializa Then
                        lblSubTotal.Text = CStr(0)
                        lblRedondeo.Text = CStr(0)
                        lblTotal.Text = CStr(0)
                        lblTotalPesos.Text = CStr(0)
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLSUBTOTALADICIONAL, "")
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLDESCUENTOADICIONAL, "")
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLIVAADICIONAL, "")
                        flexVentas.set_TextMatrix(flexVentas.Row, C_COLTOTALADICIONAL, "")
                        Inicializa = True
                    End If
                    .set_TextMatrix(I, C_COLPRECIOLISTAADICIONAL, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOPUBLICO))) * PorcentajeAdicional, "###,##0.0000"))
                    .set_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOPUBLICO))) * PorcentajeAdicional, "###,##0.00"))
                    .set_TextMatrix(I, C_COLNUEVOIMPORTESINDESCUENTO, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) * CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))), "###,##0.00"))
                    If CDbl(Numerico(.get_TextMatrix(I, C_COLPORCDESCTO))) <> 0 Then
                        .set_TextMatrix(I, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) * ((CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) * CDbl(1 - CDbl(VB6.Format(Numerico(CStr(CDbl(.get_TextMatrix(I, C_COLPORCDESCTO)) / 100)), "#####0.0000"))))), "###,##0.00"))
                    ElseIf CDbl(Numerico(.get_TextMatrix(I, C_COLPORCPROM))) <> 0 Then
                        .set_TextMatrix(I, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) * ((CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) * CDbl(1 - CDbl(VB6.Format(Numerico(CStr(CDbl(.get_TextMatrix(I, C_COLPORCPROM)) / 100)), "#####0.0000"))))), "###,##0.00"))
                    Else
                        .set_TextMatrix(I, C_COLNUEVOIMPORTECONDESCUENTO, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL)))) * ((CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) * CDbl(1 - CDbl(VB6.Format(Numerico(CStr(CDbl(.get_TextMatrix(I, C_COLPORCDESCTO)) / 100)), "#####0.0000"))))), "###,##0.00"))
                    End If
                    .set_TextMatrix(I, C_ColPRECIOLISTASINIVA, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) / (1 + CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000"))), "#####0.0000"))
                    If CDbl(Numerico(.get_TextMatrix(I, C_COLPORCDESCTO))) <> 0 Then
                        .set_TextMatrix(I, C_ColDESCUENTO, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) * CDbl(VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPORCDESCTO))) / 100, "#####0.0000")) / (1 + CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000"))), "#####0.0000"))
                        .set_TextMatrix(I, C_COLIVAREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColDESCUENTO)))) * CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000")), "#####0.0000"))
                        .set_TextMatrix(I, C_COLPRECIOREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColDESCUENTO)))) + CDbl(Numerico(.get_TextMatrix(I, C_COLIVAREAL))), "#####0.0000"))
                    ElseIf CDbl(Numerico(.get_TextMatrix(I, C_COLPORCPROM))) <> 0 Then
                        .set_TextMatrix(I, C_ColPROMOCION, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLNUEVOPRECIOPUBLICO))) * CDbl(VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPORCPROM))) / 100, "#####0.0000")) / (1 + CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000"))), "#####0.0000"))
                        .set_TextMatrix(I, C_COLIVAREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColPROMOCION)))) * CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000")), "#####0.0000"))
                        .set_TextMatrix(I, C_COLPRECIOREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColPROMOCION)))) + CDbl(Numerico(.get_TextMatrix(I, C_COLIVAREAL))), "#####0.0000"))
                    Else
                        .set_TextMatrix(I, C_ColDESCUENTO, 0)
                        .set_TextMatrix(I, C_ColPROMOCION, 0)
                        .set_TextMatrix(I, C_COLIVAREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColDESCUENTO)))) * CDbl(VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLPORCIVA))) / 100, "#####0.0000")), "#####0.0000"))
                        .set_TextMatrix(I, C_COLPRECIOREAL, VB6.Format((CDbl(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) - CDbl(Numerico(.get_TextMatrix(I, C_ColDESCUENTO)))) + CDbl(Numerico(.get_TextMatrix(I, C_COLIVAREAL))), "#####0.0000"))
                    End If
                    .set_TextMatrix(I, C_COLPORCENTAJEADICIONAL, (100 - (PorcentajeAdicional * 100)))
                    lblSubTotal.Text = CStr(CDbl(Numerico(lblSubTotal.Text)) + (CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOREAL))) * (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL))))))
                    flexVentas.set_TextMatrix(flexVentas.Row, C_COLSUBTOTALADICIONAL, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLSUBTOTALADICIONAL))) + (CDec(Numerico(.get_TextMatrix(I, C_ColPRECIOLISTASINIVA))) * (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL))))))
                    flexVentas.set_TextMatrix(flexVentas.Row, C_COLDESCUENTOADICIONAL, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLDESCUENTOADICIONAL))) + ((CDbl(Numerico(.get_TextMatrix(I, C_ColDESCUENTO))) + CDbl(Numerico(.get_TextMatrix(I, C_ColPROMOCION)))) * (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL))))))
                    flexVentas.set_TextMatrix(flexVentas.Row, C_COLIVAADICIONAL, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLIVAADICIONAL))) + (CDec(Numerico(.get_TextMatrix(I, C_COLIVAREAL))) * (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL))))))
                    flexVentas.set_TextMatrix(flexVentas.Row, C_COLTOTALADICIONAL, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLTOTALADICIONAL))) + (CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOREAL))) * (CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))) - CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDADDEVOL))))))
                    If .get_TextMatrix(I, C_COLNUEVADESCRIPCION) <> "" Then
                        .set_TextMatrix(I, C_COLMODIFICADO, "M")
                    Else
                        .set_TextMatrix(I, C_COLMODIFICADO, "")
                    End If
                End If
            Next
            If Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) <> "TC" And Trim(flexVentas.get_TextMatrix(flexVentas.Row, C_COLESTADOFOLIO)) <> "F" Then
                flexVentas.set_TextMatrix(flexVentas.Row, C_COLANTICIPOADICIONAL, VB6.Format(CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_COLANTICIPO))) * PorcentajeAdicional, "#####0.0000"))
            Else
                flexVentas.set_TextMatrix(flexVentas.Row, C_COLANTICIPOADICIONAL, flexVentas.get_TextMatrix(flexVentas.Row, C_COLANTICIPO))
            End If
            TotalPesos = CDbl(VB6.Format(CDbl(lblSubTotal.Text) * CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_ColTIPOCAMBIO))), "#####0.000000"))
            TotalPesos = CDbl(VB6.Format(TotalPesos, "#####0.00"))
            RedondeoPesos = ModCorporativo.RedondeoUnidadFinal(CDec(TotalPesos), CDbl(gcurRedondeo))
            RedondeoDolares = CDbl(VB6.Format(RedondeoPesos / CDbl(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, C_ColTIPOCAMBIO))), "#####0.0000"))
            flexVentas.set_TextMatrix(flexVentas.Row, C_COLREDONDEOADICIONAL, RedondeoDolares)
            lblRedondeo.Text = VB6.Format(RedondeoDolares, "###,##0.00")
            lblTotal.Text = VB6.Format(CDbl(Numerico(lblSubTotal.Text)) + CDbl(Numerico(lblRedondeo.Text)), "###,##0.00")
            lblSubTotal.Text = VB6.Format(lblSubTotal.Text, "###,##0.00")
            'lblTotalPesos = Format(Numerico(lblTotal) * flexVentas.TextMatrix(flexVentas.Row, C_COLTIPOCAMBIO), "###,##0.0")
            lblTotalPesos.Text = VB6.Format(TotalPesos + RedondeoPesos, "#####0.0")
            lblTotalPesos.Text = VB6.Format(lblTotalPesos.Text, "###,##0.00")
        End With
    End Sub

    'Guarda los cambios hechos a la factura (hecho para guardar los cambios que se pudieran
    'hacer en la factura adicional despues de consultar el folio de factura
    Private Function GuardarCambiosFactura() As Boolean
        On Error GoTo Merr '~Å~
        Dim blnTransaccion As Boolean
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Cnn.BeginTrans()
        blnTransaccion = True
        'Guardar la Factura
        ModStoredProcedures.PR_IME_Facturas(Trim(txtFolioFactura.Text), "1", Trim(txtCodSucursal.Text), Str(Caja), "01/01/1900", "N", "", "", "", "", C_DOLAR, "", "", "", "", "", "", "", "V", "01/01/1900", "0", "", "0", "0", "0", Trim(txtFacturaAdicional.Text), "", CStr(0), C_MODIFICACION, CStr(0))
        Cmd.Execute()
        Cnn.CommitTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        blnTransaccion = False
        GuardarCambiosFactura = True
        Exit Function
Merr:
        If blnTransaccion Then Cnn.RollbackTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MostrarError("Ha ocurrido un error al intentar guardar los cambios en la factura")
    End Function

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs)
        ImprimirFactura()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs)
        GuardarFactura()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs)

    End Sub
End Class