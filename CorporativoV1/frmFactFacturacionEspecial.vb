Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6

Public Class frmFactFacturacionEspecial
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             FACTURACION ESPECIAL                                                                         *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      LUNES 16 DE JUNIO DE 2003                                                                    *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkDesglosarIva As System.Windows.Forms.CheckBox
    Public WithEvents cmdAbcRfc As System.Windows.Forms.Button
    Public WithEvents txtTotalFinal As System.Windows.Forms.TextBox
    Public WithEvents txtRedondeo As System.Windows.Forms.TextBox
    Public WithEvents txtSubTotal As System.Windows.Forms.TextBox
    Public WithEvents txtIva As System.Windows.Forms.TextBox
    Public WithEvents txtTotal As System.Windows.Forms.TextBox
    Public WithEvents _Label1_12 As System.Windows.Forms.Label
    Public WithEvents _Label1_11 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    Public WithEvents _Label1_8 As System.Windows.Forms.Label
    Public WithEvents _Label1_9 As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents txtTipoCambio As System.Windows.Forms.TextBox
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents txtCiudad As System.Windows.Forms.TextBox
    Public WithEvents txtCodigo As System.Windows.Forms.TextBox
    Public WithEvents txtRFC As System.Windows.Forms.TextBox
    Public WithEvents txtCP As System.Windows.Forms.TextBox
    Public WithEvents txtColonia As System.Windows.Forms.TextBox
    Public WithEvents txtDomicilio As System.Windows.Forms.TextBox
    Public WithEvents txtNombreCliente As System.Windows.Forms.TextBox
    Public WithEvents _Label1_10 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_4 As System.Windows.Forms.Label
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents _optCredito_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optContado_0 As System.Windows.Forms.RadioButton
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents txtFolioFactura As System.Windows.Forms.TextBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optContado As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optCredito As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray

    Dim mblnSalir As Boolean
    Dim mblnContado As Boolean
    Dim mblnCredito As Boolean
    Dim mstrCorporativo As String
    Dim mblnNuevo As Boolean
    Dim mblnCambiosEnCodigo As Boolean
    Dim mblnPierdeFoco As Boolean
    Dim mblnCancelar As Boolean
    Dim Redondeo As Double
    Dim DesgloseIva As Byte
    Dim CodSucursal As Integer
    Dim CodCaja As Integer
    Public WithEvents brnBuscar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnSalir As Button
    '''Const TipoMovto = 1
    Const TipoMovto As Integer = 5
    Public strControlActual As String 'Nombre del control actual


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFactFacturacionEspecial))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtTotalFinal = New System.Windows.Forms.TextBox()
        Me.txtRedondeo = New System.Windows.Forms.TextBox()
        Me.txtSubTotal = New System.Windows.Forms.TextBox()
        Me.txtIva = New System.Windows.Forms.TextBox()
        Me.txtTotal = New System.Windows.Forms.TextBox()
        Me.txtTipoCambio = New System.Windows.Forms.TextBox()
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me.txtCiudad = New System.Windows.Forms.TextBox()
        Me.txtCodigo = New System.Windows.Forms.TextBox()
        Me.txtRFC = New System.Windows.Forms.TextBox()
        Me.txtCP = New System.Windows.Forms.TextBox()
        Me.txtColonia = New System.Windows.Forms.TextBox()
        Me.txtDomicilio = New System.Windows.Forms.TextBox()
        Me.txtNombreCliente = New System.Windows.Forms.TextBox()
        Me._optCredito_1 = New System.Windows.Forms.RadioButton()
        Me._optContado_0 = New System.Windows.Forms.RadioButton()
        Me.txtFolioFactura = New System.Windows.Forms.TextBox()
        Me.chkDesglosarIva = New System.Windows.Forms.CheckBox()
        Me.cmdAbcRfc = New System.Windows.Forms.Button()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me._Label1_12 = New System.Windows.Forms.Label()
        Me._Label1_11 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me._Label1_8 = New System.Windows.Forms.Label()
        Me._Label1_9 = New System.Windows.Forms.Label()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me._Label1_10 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_4 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optContado = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optCredito = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.brnBuscar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optContado, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optCredito, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtTotalFinal
        '
        Me.txtTotalFinal.AcceptsReturn = True
        Me.txtTotalFinal.BackColor = System.Drawing.SystemColors.Window
        Me.txtTotalFinal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTotalFinal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTotalFinal.Location = New System.Drawing.Point(69, 109)
        Me.txtTotalFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTotalFinal.MaxLength = 0
        Me.txtTotalFinal.Name = "txtTotalFinal"
        Me.txtTotalFinal.ReadOnly = True
        Me.txtTotalFinal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotalFinal.Size = New System.Drawing.Size(91, 20)
        Me.txtTotalFinal.TabIndex = 22
        Me.txtTotalFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTotalFinal, "Total Final.")
        '
        'txtRedondeo
        '
        Me.txtRedondeo.AcceptsReturn = True
        Me.txtRedondeo.BackColor = System.Drawing.SystemColors.Window
        Me.txtRedondeo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRedondeo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRedondeo.Location = New System.Drawing.Point(69, 84)
        Me.txtRedondeo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRedondeo.MaxLength = 0
        Me.txtRedondeo.Name = "txtRedondeo"
        Me.txtRedondeo.ReadOnly = True
        Me.txtRedondeo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRedondeo.Size = New System.Drawing.Size(91, 20)
        Me.txtRedondeo.TabIndex = 21
        Me.txtRedondeo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtRedondeo, "Redondeo.")
        '
        'txtSubTotal
        '
        Me.txtSubTotal.AcceptsReturn = True
        Me.txtSubTotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtSubTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSubTotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubTotal.Location = New System.Drawing.Point(69, 12)
        Me.txtSubTotal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSubTotal.MaxLength = 0
        Me.txtSubTotal.Name = "txtSubTotal"
        Me.txtSubTotal.ReadOnly = True
        Me.txtSubTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSubTotal.Size = New System.Drawing.Size(91, 20)
        Me.txtSubTotal.TabIndex = 18
        Me.txtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtSubTotal, "SubTotal de la Factura.")
        '
        'txtIva
        '
        Me.txtIva.AcceptsReturn = True
        Me.txtIva.BackColor = System.Drawing.SystemColors.Window
        Me.txtIva.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtIva.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtIva.Location = New System.Drawing.Point(69, 36)
        Me.txtIva.Margin = New System.Windows.Forms.Padding(2)
        Me.txtIva.MaxLength = 0
        Me.txtIva.Name = "txtIva"
        Me.txtIva.ReadOnly = True
        Me.txtIva.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtIva.Size = New System.Drawing.Size(91, 20)
        Me.txtIva.TabIndex = 19
        Me.txtIva.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtIva, "Iva.")
        '
        'txtTotal
        '
        Me.txtTotal.AcceptsReturn = True
        Me.txtTotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTotal.Location = New System.Drawing.Point(69, 60)
        Me.txtTotal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTotal.MaxLength = 0
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.ReadOnly = True
        Me.txtTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotal.Size = New System.Drawing.Size(91, 20)
        Me.txtTotal.TabIndex = 20
        Me.txtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTotal, "Total de la Factura.")
        '
        'txtTipoCambio
        '
        Me.txtTipoCambio.AcceptsReturn = True
        Me.txtTipoCambio.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipoCambio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambio.Location = New System.Drawing.Point(485, 76)
        Me.txtTipoCambio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTipoCambio.MaxLength = 0
        Me.txtTipoCambio.Name = "txtTipoCambio"
        Me.txtTipoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambio.Size = New System.Drawing.Size(115, 20)
        Me.txtTipoCambio.TabIndex = 6
        Me.txtTipoCambio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambio, "Tipo de Cambio del Peso con Respecto al Dolar.")
        '
        '_optMoneda_1
        '
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_1.Location = New System.Drawing.Point(85, 15)
        Me._optMoneda_1.Margin = New System.Windows.Forms.Padding(2)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(64, 19)
        Me._optMoneda_1.TabIndex = 5
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Text = "Dolares"
        Me.ToolTip1.SetToolTip(Me._optMoneda_1, "Moneda con la que Paga : Dolares")
        Me._optMoneda_1.UseVisualStyleBackColor = False
        '
        '_optMoneda_0
        '
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_0.Location = New System.Drawing.Point(24, 17)
        Me._optMoneda_0.Margin = New System.Windows.Forms.Padding(2)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(57, 15)
        Me._optMoneda_0.TabIndex = 4
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me._optMoneda_0, "Moneda con la que Paga : Pesos")
        Me._optMoneda_0.UseVisualStyleBackColor = False
        '
        'txtCiudad
        '
        Me.txtCiudad.AcceptsReturn = True
        Me.txtCiudad.BackColor = System.Drawing.SystemColors.Window
        Me.txtCiudad.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCiudad.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCiudad.Location = New System.Drawing.Point(54, 61)
        Me.txtCiudad.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCiudad.MaxLength = 30
        Me.txtCiudad.Name = "txtCiudad"
        Me.txtCiudad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCiudad.Size = New System.Drawing.Size(154, 20)
        Me.txtCiudad.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.txtCiudad, "Ciudad del Cliente.")
        '
        'txtCodigo
        '
        Me.txtCodigo.AcceptsReturn = True
        Me.txtCodigo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodigo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodigo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodigo.Location = New System.Drawing.Point(54, 15)
        Me.txtCodigo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodigo.MaxLength = 5
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodigo.Size = New System.Drawing.Size(50, 20)
        Me.txtCodigo.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtCodigo, "Codigo del Cliente.")
        '
        'txtRFC
        '
        Me.txtRFC.AcceptsReturn = True
        Me.txtRFC.BackColor = System.Drawing.SystemColors.Window
        Me.txtRFC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRFC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRFC.Location = New System.Drawing.Point(389, 63)
        Me.txtRFC.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRFC.MaxLength = 15
        Me.txtRFC.Name = "txtRFC"
        Me.txtRFC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRFC.Size = New System.Drawing.Size(172, 20)
        Me.txtRFC.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.txtRFC, "RFC del Cliente.")
        '
        'txtCP
        '
        Me.txtCP.AcceptsReturn = True
        Me.txtCP.BackColor = System.Drawing.SystemColors.Window
        Me.txtCP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCP.Location = New System.Drawing.Point(255, 62)
        Me.txtCP.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCP.MaxLength = 10
        Me.txtCP.Name = "txtCP"
        Me.txtCP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCP.Size = New System.Drawing.Size(78, 20)
        Me.txtCP.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.txtCP, "Codigo Postal del Cliente.")
        '
        'txtColonia
        '
        Me.txtColonia.AcceptsReturn = True
        Me.txtColonia.BackColor = System.Drawing.SystemColors.Window
        Me.txtColonia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtColonia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtColonia.Location = New System.Drawing.Point(389, 39)
        Me.txtColonia.Margin = New System.Windows.Forms.Padding(2)
        Me.txtColonia.MaxLength = 30
        Me.txtColonia.Name = "txtColonia"
        Me.txtColonia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtColonia.Size = New System.Drawing.Size(172, 20)
        Me.txtColonia.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txtColonia, "Colonia del Cliente.")
        '
        'txtDomicilio
        '
        Me.txtDomicilio.AcceptsReturn = True
        Me.txtDomicilio.BackColor = System.Drawing.SystemColors.Window
        Me.txtDomicilio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDomicilio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDomicilio.Location = New System.Drawing.Point(63, 38)
        Me.txtDomicilio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDomicilio.MaxLength = 65
        Me.txtDomicilio.Name = "txtDomicilio"
        Me.txtDomicilio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDomicilio.Size = New System.Drawing.Size(270, 20)
        Me.txtDomicilio.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.txtDomicilio, "Domicilio del Cliente.")
        '
        'txtNombreCliente
        '
        Me.txtNombreCliente.AcceptsReturn = True
        Me.txtNombreCliente.BackColor = System.Drawing.SystemColors.Window
        Me.txtNombreCliente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombreCliente.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombreCliente.Location = New System.Drawing.Point(159, 15)
        Me.txtNombreCliente.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNombreCliente.MaxLength = 40
        Me.txtNombreCliente.Name = "txtNombreCliente"
        Me.txtNombreCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombreCliente.Size = New System.Drawing.Size(402, 20)
        Me.txtNombreCliente.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtNombreCliente, "Nombre del Cliente.")
        '
        '_optCredito_1
        '
        Me._optCredito_1.BackColor = System.Drawing.SystemColors.Control
        Me._optCredito_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optCredito_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optCredito_1.Location = New System.Drawing.Point(93, 14)
        Me._optCredito_1.Margin = New System.Windows.Forms.Padding(2)
        Me._optCredito_1.Name = "_optCredito_1"
        Me._optCredito_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optCredito_1.Size = New System.Drawing.Size(68, 18)
        Me._optCredito_1.TabIndex = 3
        Me._optCredito_1.TabStop = True
        Me._optCredito_1.Text = "Crédito"
        Me.ToolTip1.SetToolTip(Me._optCredito_1, "Facturación a Credito.")
        Me._optCredito_1.UseVisualStyleBackColor = False
        '
        '_optContado_0
        '
        Me._optContado_0.BackColor = System.Drawing.SystemColors.Control
        Me._optContado_0.Checked = True
        Me._optContado_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optContado_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optContado_0.Location = New System.Drawing.Point(26, 14)
        Me._optContado_0.Margin = New System.Windows.Forms.Padding(2)
        Me._optContado_0.Name = "_optContado_0"
        Me._optContado_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optContado_0.Size = New System.Drawing.Size(68, 18)
        Me._optContado_0.TabIndex = 2
        Me._optContado_0.TabStop = True
        Me._optContado_0.Text = "Contado"
        Me.ToolTip1.SetToolTip(Me._optContado_0, "Facturación de Contado.")
        Me._optContado_0.UseVisualStyleBackColor = False
        '
        'txtFolioFactura
        '
        Me.txtFolioFactura.AcceptsReturn = True
        Me.txtFolioFactura.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioFactura.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioFactura.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioFactura.Location = New System.Drawing.Point(124, 13)
        Me.txtFolioFactura.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioFactura.MaxLength = 17
        Me.txtFolioFactura.Name = "txtFolioFactura"
        Me.txtFolioFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioFactura.Size = New System.Drawing.Size(124, 20)
        Me.txtFolioFactura.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtFolioFactura, "Folio de la Factura.")
        '
        'chkDesglosarIva
        '
        Me.chkDesglosarIva.BackColor = System.Drawing.SystemColors.Control
        Me.chkDesglosarIva.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDesglosarIva.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDesglosarIva.Location = New System.Drawing.Point(13, 361)
        Me.chkDesglosarIva.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDesglosarIva.Name = "chkDesglosarIva"
        Me.chkDesglosarIva.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDesglosarIva.Size = New System.Drawing.Size(98, 20)
        Me.chkDesglosarIva.TabIndex = 16
        Me.chkDesglosarIva.Text = "Desglosar Iva"
        Me.chkDesglosarIva.UseVisualStyleBackColor = False
        '
        'cmdAbcRfc
        '
        Me.cmdAbcRfc.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAbcRfc.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAbcRfc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAbcRfc.Location = New System.Drawing.Point(10, 389)
        Me.cmdAbcRfc.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdAbcRfc.Name = "cmdAbcRfc"
        Me.cmdAbcRfc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAbcRfc.Size = New System.Drawing.Size(104, 42)
        Me.cmdAbcRfc.TabIndex = 17
        Me.cmdAbcRfc.Text = "ABC de RFC's"
        Me.cmdAbcRfc.UseVisualStyleBackColor = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.txtTotalFinal)
        Me.Frame4.Controls.Add(Me.txtRedondeo)
        Me.Frame4.Controls.Add(Me.txtSubTotal)
        Me.Frame4.Controls.Add(Me.txtIva)
        Me.Frame4.Controls.Add(Me.txtTotal)
        Me.Frame4.Controls.Add(Me._Label1_12)
        Me.Frame4.Controls.Add(Me._Label1_11)
        Me.Frame4.Controls.Add(Me._Label1_7)
        Me.Frame4.Controls.Add(Me._Label1_8)
        Me.Frame4.Controls.Add(Me._Label1_9)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(416, 361)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(184, 140)
        Me.Frame4.TabIndex = 36
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Totales"
        '
        '_Label1_12
        '
        Me._Label1_12.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_12.Location = New System.Drawing.Point(4, 112)
        Me._Label1_12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_12.Name = "_Label1_12"
        Me._Label1_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_12.Size = New System.Drawing.Size(61, 17)
        Me._Label1_12.TabIndex = 42
        Me._Label1_12.Text = "Total Final"
        Me._Label1_12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_11
        '
        Me._Label1_11.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_11.Location = New System.Drawing.Point(4, 87)
        Me._Label1_11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_11.Name = "_Label1_11"
        Me._Label1_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_11.Size = New System.Drawing.Size(61, 17)
        Me._Label1_11.TabIndex = 41
        Me._Label1_11.Text = "Redondeo"
        Me._Label1_11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Location = New System.Drawing.Point(10, 15)
        Me._Label1_7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(55, 17)
        Me._Label1_7.TabIndex = 39
        Me._Label1_7.Text = "SubTotal "
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_8
        '
        Me._Label1_8.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_8.Location = New System.Drawing.Point(18, 39)
        Me._Label1_8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_8.Name = "_Label1_8"
        Me._Label1_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_8.Size = New System.Drawing.Size(37, 17)
        Me._Label1_8.TabIndex = 38
        Me._Label1_8.Text = "IVA "
        Me._Label1_8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_9
        '
        Me._Label1_9.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_9.Location = New System.Drawing.Point(18, 61)
        Me._Label1_9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_9.Name = "_Label1_9"
        Me._Label1_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_9.Size = New System.Drawing.Size(37, 17)
        Me._Label1_9.TabIndex = 37
        Me._Label1_9.Text = "Total "
        Me._Label1_9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(17, 244)
        Me.txtFlex.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(60, 20)
        Me.txtFlex.TabIndex = 14
        Me.txtFlex.Visible = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me._optMoneda_1)
        Me.Frame3.Controls.Add(Me._optMoneda_0)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(213, 52)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(164, 39)
        Me.Frame3.TabIndex = 33
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Moneda :"
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(13, 224)
        Me.flexDetalle.Margin = New System.Windows.Forms.Padding(2)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(587, 118)
        Me.flexDetalle.TabIndex = 15
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtCiudad)
        Me.Frame2.Controls.Add(Me.txtCodigo)
        Me.Frame2.Controls.Add(Me.txtRFC)
        Me.Frame2.Controls.Add(Me.txtCP)
        Me.Frame2.Controls.Add(Me.txtColonia)
        Me.Frame2.Controls.Add(Me.txtDomicilio)
        Me.Frame2.Controls.Add(Me.txtNombreCliente)
        Me.Frame2.Controls.Add(Me._Label1_10)
        Me.Frame2.Controls.Add(Me.Label2)
        Me.Frame2.Controls.Add(Me._Label1_6)
        Me.Frame2.Controls.Add(Me._Label1_5)
        Me.Frame2.Controls.Add(Me._Label1_4)
        Me.Frame2.Controls.Add(Me._Label1_3)
        Me.Frame2.Controls.Add(Me._Label1_2)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(15, 108)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(585, 92)
        Me.Frame2.TabIndex = 26
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Datos de Facturación"
        '
        '_Label1_10
        '
        Me._Label1_10.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_10.Location = New System.Drawing.Point(8, 61)
        Me._Label1_10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_10.Name = "_Label1_10"
        Me._Label1_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_10.Size = New System.Drawing.Size(50, 17)
        Me._Label1_10.TabIndex = 40
        Me._Label1_10.Text = "Ciudad :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 20)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(50, 17)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "Codigo :"
        '
        '_Label1_6
        '
        Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_6.Location = New System.Drawing.Point(345, 64)
        Me._Label1_6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(44, 17)
        Me._Label1_6.TabIndex = 31
        Me._Label1_6.Text = "R.F.C. :"
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_5.Location = New System.Drawing.Point(223, 64)
        Me._Label1_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(34, 17)
        Me._Label1_5.TabIndex = 30
        Me._Label1_5.Text = "C.P. :"
        '
        '_Label1_4
        '
        Me._Label1_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_4.Location = New System.Drawing.Point(344, 41)
        Me._Label1_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_4.Size = New System.Drawing.Size(50, 17)
        Me._Label1_4.TabIndex = 29
        Me._Label1_4.Text = "Colonia :"
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_3.Location = New System.Drawing.Point(8, 41)
        Me._Label1_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(59, 17)
        Me._Label1_3.TabIndex = 28
        Me._Label1_3.Text = "Domicilio :"
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_2.Location = New System.Drawing.Point(108, 20)
        Me._Label1_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(59, 17)
        Me._Label1_2.TabIndex = 27
        Me._Label1_2.Text = "Nombre :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me._optCredito_1)
        Me.Frame1.Controls.Add(Me._optContado_0)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(17, 52)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(165, 39)
        Me.Frame1.TabIndex = 25
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Condición"
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(494, 15)
        Me.dtpFecha.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(104, 20)
        Me.dtpFecha.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(12, 370)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(254, 27)
        Me.Label4.TabIndex = 35
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(393, 79)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(96, 16)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Tipo de Cambio  $"
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(449, 16)
        Me._Label1_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(48, 17)
        Me._Label1_1.TabIndex = 24
        Me._Label1_1.Text = "Fecha :"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(14, 15)
        Me._Label1_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(106, 17)
        Me._Label1_0.TabIndex = 23
        Me._Label1_0.Text = "Folio de la Factura :"
        '
        'optContado
        '
        '
        'optCredito
        '
        '
        'optMoneda
        '
        '
        'brnBuscar
        '
        Me.brnBuscar.BackColor = System.Drawing.SystemColors.Control
        Me.brnBuscar.Cursor = System.Windows.Forms.Cursors.Default
        Me.brnBuscar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.brnBuscar.Location = New System.Drawing.Point(11, 459)
        Me.brnBuscar.Margin = New System.Windows.Forms.Padding(2)
        Me.brnBuscar.Name = "brnBuscar"
        Me.brnBuscar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.brnBuscar.Size = New System.Drawing.Size(104, 42)
        Me.brnBuscar.TabIndex = 37
        Me.brnBuscar.Text = "Buscar"
        Me.brnBuscar.UseVisualStyleBackColor = False
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackColor = System.Drawing.SystemColors.Control
        Me.btnLimpiar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLimpiar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLimpiar.Location = New System.Drawing.Point(119, 459)
        Me.btnLimpiar.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLimpiar.Size = New System.Drawing.Size(104, 42)
        Me.btnLimpiar.TabIndex = 38
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnSalir
        '
        Me.btnSalir.BackColor = System.Drawing.SystemColors.Control
        Me.btnSalir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSalir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSalir.Location = New System.Drawing.Point(227, 459)
        Me.btnSalir.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSalir.Size = New System.Drawing.Size(104, 42)
        Me.btnSalir.TabIndex = 39
        Me.btnSalir.Text = "Salir"
        Me.btnSalir.UseVisualStyleBackColor = False
        '
        'frmFactFacturacionEspecial
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(619, 517)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.brnBuscar)
        Me.Controls.Add(Me.chkDesglosarIva)
        Me.Controls.Add(Me.cmdAbcRfc)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.txtFlex)
        Me.Controls.Add(Me.txtTipoCambio)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.flexDetalle)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.dtpFecha)
        Me.Controls.Add(Me.txtFolioFactura)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me._Label1_1)
        Me.Controls.Add(Me._Label1_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(148, 151)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmFactFacturacionEspecial"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Facturación Especial"
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optContado, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optCredito, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Function Cambios() As Boolean
        Dim I As Integer
        Cambios = True
        If txtCodigo.Text <> "" Then Exit Function
        If txtNombreCliente.Text <> "" Then Exit Function
        If txtDomicilio.Text <> "" Then Exit Function
        If txtColonia.Text <> "" Then Exit Function
        If txtCiudad.Text <> "" Then Exit Function
        If txtCP.Text <> "" Then Exit Function
        If txtRFC.Text <> "" Then Exit Function
        With flexDetalle
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, 0) <> "" Or .get_TextMatrix(I, 1) <> "" Or .get_TextMatrix(I, 2) <> "" Or .get_TextMatrix(I, 3) <> "" Or .get_TextMatrix(I, 4) <> "" Then Exit Function
            Next
        End With
        If CDbl(Numerico(Format(txtSubTotal.Text, "#####0.00"))) > 0 Then Exit Function
        If CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) > 0 Then Exit Function
        Cambios = False
    End Function

    Function Guardar() As Boolean
        On Error GoTo Merr
        Dim blnTransaccion As Boolean
        Dim Fecha As String
        Dim FolioFactura As String
        Dim ConsFactura As String
        Dim I As Integer
        If Not mblnNuevo Then
            Exit Function
        End If
        If Not Cambios() Then
            Exit Function
        End If
        If ValidaDatos() = False Then
            Exit Function
        End If
        If chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Checked Then
            DesgloseIva = 1
        Else
            DesgloseIva = 0
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        'Obtener la caja del almacen principal
        gStrSql = "SELECT TOP 1 CodCaja FROM CatCajas WHERE CodAlmacen = " & gintCodAlmacenGral & " ORDER BY CodCaja"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            CodCaja = RsGral.Fields("CodCaja").Value
        End If
        Fecha = ""
        Fecha = Fecha & mstrCorporativo & Format(Year(dtpFecha.Value), "0000") & Format(Month(dtpFecha.Value), "00") & Format((dtpFecha.Value), "00")
        '''gStrSql = "SELECT Prefijo,Consecutivo FROM FoliosCorporativo WHERE CodFolio =" & TipoMovto
        gStrSql = "Select CodFolio, Prefijo,  Consecutivo + 1  as Consecutivo From CatFolios Where CodFolio = " & TipoMovto & " And CodAlmacen = " & gintCodAlmacenGral & " "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            ConsFactura = Str(RsGral.Fields("Consecutivo").Value)
        End If
        FolioFactura = Trim(RsGral.Fields("Prefijo").Value) & Trim(Fecha) & Format(ConsFactura, "000000")
        '''ModStoredProcedures.PR_IMEFoliosCorporativo TipoMovto, "", "", ConsFactura, C_MODIFICACION, 0
        '''Cmd.Execute
        '''Actualiza el consecutivo de facturas para la sucursal seleccionada
        ModStoredProcedures.PR_IMECatFolios(CStr(TipoMovto), CStr(gintCodAlmacenGral), "", "", ConsFactura, C_MODIFICACION, CStr(1))
        Cmd.Execute()

        With flexDetalle
            I = 1
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, 0) <> "" And .get_TextMatrix(I, 1) <> "" And .get_TextMatrix(I, 2) <> "" And .get_TextMatrix(I, 3) <> "" And .get_TextMatrix(I, 4) <> "" Then
                    ModStoredProcedures.PR_IME_Facturas(FolioFactura, CStr(I), CStr(gintCodAlmacenGral), CStr(CodCaja), Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), "E", IIf(_optContado_0.Checked, "CO", "CR"), txtCodigo.Text, txtNombreCliente.Text, txtRFC.Text, IIf(_optMoneda_0.Checked, "P", "D"), Format(txtTipoCambio.Text, "#####0.00"), Format(txtSubTotal.Text, "#####0.00"), "0", Format(txtIva.Text, "#####0.00"), Format(txtTotal.Text, "#####0.00"), CStr(Redondeo), "0", "V", "01/01/1900", Format(.get_TextMatrix(I, 0), "0"), .get_TextMatrix(I, 1), Format(.get_TextMatrix(I, 3), "#####0.00"), Format(.get_TextMatrix(I, 4), "#####0.00"), .get_TextMatrix(I, 5), "", "C", CStr(DesgloseIva), C_INSERCION, CStr(0))
                    Cmd.Execute()
                End If
            Next
        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        txtFolioFactura.Text = FolioFactura
        MsgBox("Los Datos se Han Grabado con Exito, Se ha Generado la Factura " & txtFolioFactura.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)

        If MsgBox("Desea imprimir la factura?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, gstrCorpoNOMBREEMPRESA) Then
            ImprimirFactura()
        End If

        Limpiar()
Merr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Sub ImprimirTicket()
        On Error GoTo Merr
        Dim strImpresora As String
        Dim strNomEmpresa As String
        Dim Archivo As Integer
        Dim I As Integer
        Dim TotalPesos As Double
        strImpresora = gstrRutaImpresora
        If Not ModEstandar.BuscarImpresora(strImpresora) Then
            MsgBox("Impresora Incorrecta")
        End If
        Archivo = FreeFile()
        FileOpen(Archivo, strImpresora, OpenMode.Output)
        PrintLine(1, Chr(27) & Chr(33) & Chr(1))
        strNomEmpresa = UCase(gstrNombCortoEmpresa)
        For I = 1 To Len(strNomEmpresa)
            Select Case Mid(strNomEmpresa, I, 1)
                Case "Á"
                    strNomEmpresa = Replace(strNomEmpresa, "Á", "A")
                Case "É"
                    strNomEmpresa = Replace(strNomEmpresa, "É", "E")
                Case "Í"
                    strNomEmpresa = Replace(strNomEmpresa, "Í", "I")
                Case "Ó"
                    strNomEmpresa = Replace(strNomEmpresa, "Ó", "O")
                Case "Ú"
                    strNomEmpresa = Replace(strNomEmpresa, "Ú", "U")
            End Select
        Next
        PrintLine(Archivo, "")
        PrintLine(Archivo, Space(System.Math.Round((56 - Len(Trim(gstrNombCortoEmpresa))) / 2)) & strNomEmpresa)
        PrintLine(Archivo, Space(System.Math.Round((56 - Len(Trim(gstrCorpoDOMICILIOEMPRESA))) / 2)) & UCase(gstrCorpoDOMICILIOEMPRESA))
        PrintLine(Archivo, Space(System.Math.Round((56 - Len(Trim(gstrCorpoRFCEMPRESA))) / 2)) & UCase(gstrCorpoRFCEMPRESA))
        PrintLine(Archivo, "")
        PrintLine(Archivo, Space(System.Math.Round((56 - Len("F A C T U R A")) / 2)) & "F A C T U R A")
        PrintLine(Archivo, "")
        PrintLine(Archivo, "FECHA:   " & Format(dtpFecha.Value, "dd/MMM/yyyy") & Space(11) & "FOLIO : " & txtFolioFactura.Text)
        PrintLine(Archivo, Space(49) & IIf(_optContado_0.Checked, "CONTADO", "CREDITO"))
        PrintLine(Archivo, "CLIENTE: " & txtNombreCliente.Text)
        PrintLine(Archivo, Space(9) & (txtDomicilio.Text & Space(40)) & Space(2) & (Space(7) & IIf(_optMoneda_0.Checked, C_DESCPESOS, C_DESCDOLARES)))
        PrintLine(Archivo, Space(9) & txtRFC.Text & Space(1) & (txtCiudad.Text & Space(20)) & Space(2) & "T.C. " & (Space(5) & Format(txtTipoCambio.Text, "###,##0.00")))
        PrintLine(Archivo, "")
        PrintLine(Archivo, "========================================================")
        PrintLine(Archivo, "CANT. DESCRIPCION               PREC. UNIT.      IMPORTE")
        PrintLine(Archivo, "========================================================")
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" And Trim(.get_TextMatrix(I, 4)) <> "" Then
                    PrintLine(Archivo, (Space(16) & Format(.get_TextMatrix(I, 0), "###,##0")) & Space(2) & (.get_TextMatrix(I, 1) & Space(30)) & Space(3) & (Space(15) & Format(.get_TextMatrix(I, 3), "###,##0.00")) & Space(1) & (Space(15) & Format(.get_TextMatrix(I, 4), "###,##0.00")))
                Else
                    Exit For
                End If
            Next
        End With
        PrintLine(Archivo, Space(41) & "---------------")
        PrintLine(Archivo, Space(29) & "SubTotal    " & (Space(15) & Format(txtSubTotal.Text, "###,##0.00")))
        PrintLine(Archivo, Space(29) & "Iva         " & (Space(15) & Format(txtIva.Text, "###,##0.00")))
        PrintLine(Archivo, Space(29) & "Total       " & (Space(15) & Format(txtTotal.Text, "###,##0.00")))
        If optMoneda(1).Checked Then
            TotalPesos = CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) * CDbl(Numerico(Format(txtTipoCambio.Text, "#####0.00")))
            PrintLine(Archivo, Space(29) & "Total Pesos " & (Space(15) & Format(TotalPesos, "###,##0.00")))
        End If
        PrintLine(Archivo, "")
        If _optMoneda_0.Checked Then
            PrintLine(Archivo, ModEstandar.ConLetra(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))), True, CStr(1)))
        ElseIf optMoneda(1).Checked Then
            PrintLine(Archivo, ModEstandar.ConLetra(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))), True, CStr(2)))
        End If
        PrintLine(Archivo, "")
        PrintLine(Archivo, "")
        PrintLine(Archivo, "")
        PrintLine(Archivo, "")
        PrintLine(Archivo, "")
        PrintLine(1, Chr(27) & Chr(105))
        FileClose(Archivo)
Merr:
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
        strImpresora = gstrRutaImpresora
        If Not ModEstandar.BuscarImpresora(strImpresora) Then
            MsgBox("Impresora Incorrecta")
        End If
        With Printer
            .ScaleMode = vbMillimeters

            .FontName = "Courier New"

            .Orientation = 1

            .Height = 140 'Cambia el tamaño de la hoja en la impresora
        End With
        Printer.CurrentX = 170
        Printer.CurrentY = 15
        Printer.Print(Trim(txtFolioFactura.Text))
        Printer.CurrentX = 25
        Printer.CurrentY = 22
        Printer.Print(Trim(txtNombreCliente.Text))
        Printer.CurrentX = 25
        Printer.CurrentY = 29
        Printer.Print(Trim(txtDomicilio.Text))
        Printer.CurrentX = 25
        Printer.CurrentY = 36
        Printer.Print(Trim(txtCiudad.Text))
        Printer.CurrentX = 125
        Printer.CurrentY = 36
        Printer.Print(Trim(txtRFC.Text))
        Printer.CurrentX = 161
        Printer.CurrentY = 35
        Printer.Print(Format((dtpFecha.Value), "00"))
        Printer.CurrentX = 175
        Printer.CurrentY = 35
        Printer.Print(ModEstandar.MesLetra(dtpFecha.Value, False))
        Printer.CurrentX = 199
        Printer.CurrentY = 35
        Printer.Print((dtpFecha.Value.Year))
        CoordY = 50
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" And Trim(.get_TextMatrix(I, 4)) <> "" Then
                    Printer.CurrentX = 40
                    Printer.CurrentY = CoordY
                    Printer.Print((Space(16) & Format(.get_TextMatrix(I, 0), "###,##0")))
                    Printer.CurrentX = 50
                    Printer.CurrentY = CoordY
                    Printer.Print((Trim(.get_TextMatrix(I, 1))))
                    Printer.CurrentX = 158
                    Printer.CurrentY = CoordY
                    Printer.Print((Space(15) & Format(.get_TextMatrix(I, 3), "###,##0.00")))
                    Printer.CurrentX = 181
                    Printer.CurrentY = CoordY
                    Printer.Print((Space(15) & Format(.get_TextMatrix(I, 4), "###,##0.00")))
                End If
                CoordY = CoordY + 6
            Next
            CoordY = 110
            If DesgloseIva = 1 Then
                Printer.CurrentX = 181
                Printer.CurrentY = CoordY
                Printer.Print((Space(15) & txtSubTotal.Text))
                CoordY = CoordY + 6
                Printer.CurrentX = 181
                Printer.CurrentY = CoordY
                Printer.Print((Space(15) & txtIva.Text))
            Else
                Printer.CurrentX = 181
                Printer.CurrentY = CoordY
                Printer.Print((Space(15) & Format(CDbl(Numerico(txtSubTotal.Text)) + CDbl(Numerico(txtIva.Text)), "###,##0.00")))
                CoordY = CoordY + 6
                Printer.CurrentX = 181
                Printer.CurrentY = CoordY
                '            Printer.Print Right(Space(15) & txtIva, 13)
            End If
            CoordY = CoordY
            Printer.CurrentX = 7
            Printer.CurrentY = CoordY
            Printer.Print("PAGO EN UNA SOLA EXHIBICION")

            CoordY = CoordY + 6 '''6
            Printer.CurrentX = 7
            Printer.CurrentY = CoordY
            If _optMoneda_0.Checked Then
                Printer.Print(ModEstandar.ConLetra(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))), True, CStr(1)))
            ElseIf optMoneda(1).Checked Then
                Printer.Print(ModEstandar.ConLetra(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))), True, CStr(2)))
            End If
            Printer.CurrentX = 181
            Printer.CurrentY = CoordY
            Printer.Print((Space(15) & txtTotal.Text))
        End With
        Printer.EndDoc()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Cancelar()
        On Error GoTo Merr
        Dim blnTransaccion As Boolean
        If mblnNuevo Then Exit Sub
        If mblnCancelar Then
            MsgBox("Esta Factura ya Fue Cancelada. ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If
        Select Case MsgBox("¿Desea Cancelar Esta Factura?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
            Case MsgBoxResult.No
                Exit Sub
            Case MsgBoxResult.Cancel
                Exit Sub
        End Select
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        ModStoredProcedures.PR_IME_Facturas(txtFolioFactura.Text, "0", CStr(CodSucursal), CStr(CodCaja), dtpFecha.Value, "", "", "0", "", "", "", "0", "0", "0", "0", "0", "0", "0", "C", CStr(Today), "0", "", "0", "0", "0", "", "", CStr(0), C_MODIFICACION, CStr(0))
        Cmd.Execute()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("La Factura " & Trim(txtFolioFactura.Text) & " ha Sido Cancelada.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        Limpiar()
Merr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Sub

    Function ValidaDatos() As Boolean
        Dim I As Integer
        Dim blnAlMenosUno As Boolean
        ValidaDatos = False
        blnAlMenosUno = False
        If Val(txtCodigo.Text) = 0 Then
            MsgBox(C_msgFALTADATO & "Codigo del Cliente.", MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodigo.Focus()
            Exit Function
        End If
        If Len(Trim(txtNombreCliente.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Nombre del Cliente.", MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtNombreCliente.Focus()
            Exit Function
        End If
        '''    If Len(Trim(txtDomicilio)) = 0 Then
        '''        MsgBox C_msgFALTADATO & "Domicilio del Cliente.", vbInformation, gstrNombCortoEmpresa
        '''        txtDomicilio.SetFocus
        '''        Exit Function
        '''    End If
        '''    If Len(Trim(txtColonia)) = 0 Then
        '''        MsgBox C_msgFALTADATO & "Colonia del Cliente.", vbInformation, gstrNombCortoEmpresa
        '''        txtColonia.SetFocus
        '''        Exit Function
        '''    End If
        '''    If Len(Trim(txtCiudad)) = 0 Then
        '''        MsgBox C_msgFALTADATO & "Ciudad del Cliente.", vbInformation, gstrNombCortoEmpresa
        '''        txtCiudad.SetFocus
        '''        Exit Function
        '''    End If
        '''    If Len(Trim(txtCP)) = 0 Then
        '''        MsgBox C_msgFALTADATO & "Codigo Postal del Cliente.", vbInformation, gstrNombCortoEmpresa
        '''        txtCP.SetFocus
        '''        Exit Function
        '''    End If
        '''    If Len(Trim(txtRFC)) = 0 Then
        '''        MsgBox C_msgFALTADATO & "Rfc del Cliente.", vbInformation, gstrNombCortoEmpresa
        '''        txtRFC.SetFocus
        '''        Exit Function
        '''    End If
        With flexDetalle
            For I = 1 To .Rows - 1
                If (Trim(.get_TextMatrix(I, 0)) <> "" Or Trim(.get_TextMatrix(I, 1)) <> "" Or Trim(.get_TextMatrix(I, 2)) <> "" Or Trim(.get_TextMatrix(I, 3)) <> "" Or Trim(.get_TextMatrix(I, 4)) <> "") Then
                    If .get_TextMatrix(I, 0) = "" Or CDbl(.get_TextMatrix(I, 0)) = 0 Then
                        MsgBox("Información Incompleta En Las Partidas de la Factura.", MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .Row = I
                        .Col = 0
                        flexDetalle.Focus()
                        Exit Function
                    End If
                    If Trim(.get_TextMatrix(I, 1)) = "" Then
                        MsgBox("Información Incompleta En Las Partidas de la Factura.", MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .Row = I
                        .Col = 1
                        flexDetalle.Focus()
                        Exit Function
                    End If
                    If Trim(.get_TextMatrix(I, 2)) = "" Or (CDbl(Numerico(.get_TextMatrix(I, 5))) > 0 And CDbl(Numerico(.get_TextMatrix(I, 2))) = 0) Then
                        MsgBox("Información Incompleta En Las Partidas de la Factura.", MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .Row = I
                        .Col = 2
                        flexDetalle.Focus()
                        Exit Function
                    End If
                    If Trim(.get_TextMatrix(I, 3)) = "" Or CDbl(Numerico(.get_TextMatrix(I, 3))) = 0 Then
                        MsgBox("Información Incompleta En Las Partidas de la Factura.", MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .Row = I
                        .Col = 3
                        flexDetalle.Focus()
                        Exit Function
                    End If
                    If Trim(.get_TextMatrix(I, 4)) = "" Or CDbl(Numerico(.get_TextMatrix(I, 4))) = 0 Then
                        MsgBox("Información Incompleta En Las Partidas de la Factura.", MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .Row = I
                        .Col = 4
                        flexDetalle.Focus()
                        Exit Function
                    End If
                Else
                    If I = 1 Then
                        MsgBox("Información Incompleta En Las Partidas de la Factura.", MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .Row = I
                        .Col = 0
                        flexDetalle.Focus()
                        Exit Function
                    ElseIf I > 1 Then
                        Exit For
                    End If
                End If
            Next
        End With
        If Val(txtTotal.Text) = 0 Then
            MsgBox(C_msgFALTADATO & "Total de la Factura, No se ha Registrado Ningun Detalle en la Factura.", MsgBoxStyle.Information, gstrNombCortoEmpresa)
            flexDetalle.Row = 1
            flexDetalle.Col = 0
            flexDetalle.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Sub Buscar()
        'On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas 
        Dim I As Integer

        'strControlActual = UCase(txtCodigo.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "TXTCODIGO"
                strCaptionForm = "Consulta de Clientes"
                'gStrSql = "SELECT RIGHT('00000'+LTRIM(CodCliente),5) AS CODIGO, DescCliente AS NOMBRE FROM CatClientes ORDER BY CodCliente"
                gStrSql = "SELECT RIGHT('00000'+LTRIM(CodRfc),5) AS CODIGO, DescClienteRfc AS NOMBRE FROM CatRfc ORDER BY CodRfc"
            Case "TXTNOMBRECLIENTE"
                strCaptionForm = "Consulta de Clientes"
                If txtNombreCliente.ReadOnly Then
                    'gStrSql = "SELECT DescCliente AS NOMBRE, RIGHT('00000'+LTRIM(CodCliente),5) AS CODIGO FROM CatClientes ORDER BY DescCliente"
                    gStrSql = "SELECT DescClienteRfc AS NOMBRE, RIGHT('00000'+LTRIM(CodRfc),5) AS CODIGO FROM CatRfc ORDER BY DescClienteRfc"
                Else
                    'gStrSql = "SELECT DescCliente AS NOMBRE, RIGHT('00000'+LTRIM(CodCliente),5) AS CODIGO FROM CatClientes WHERE DescCliente LIKE '" & Trim(txtNombreCliente) & "%' ORDER BY DescCliente"
                    gStrSql = "SELECT DescClienteRfc AS NOMBRE, RIGHT('00000'+LTRIM(CodRfc),5) AS CODIGO FROM CatRfc WHERE DescClienteRfc LIKE '" & Trim(txtNombreCliente.Text) & "%' ORDER BY DescClienteRfc"
                End If
            Case "TXTFOLIOFACTURA"
                strCaptionForm = "Buscar Facturas Especiales"
                gStrSql = "SELECT DISTINCT FolioFactura AS FACTURA,Nombre AS CLIENTE,FechaFactura AS FECHA,Total AS IMPORTE," & "CASE Estatus WHEN 'V' THEN 'VIGENTE' WHEN 'C' THEN 'CANCELADA' END AS ESTATUS FROM " & "Facturas WHERE FolioFactura LIKE '" & txtFolioFactura.Text & "%' AND TipoFactura = 'E' ORDER BY FechaFactura Desc, FolioFactura Desc "
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
        ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

        'Carga el formulario de consulta 
        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODIGO"
                    'ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                Case "TXTNOMBRECLIENTE"
                    'ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
                    .set_ColWidth(1, 0, 900) 'Columna del Código
                Case "TXTFOLIOFACTURA"
                    'ConfiguraConsultas(FrmConsultas, 11200, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 1800) 'Columna de la Factura
                    .set_ColWidth(1, 0, 4500) 'Columna del Cliente
                    .set_ColWidth(2, 0, 1500) 'Columna de la Fecha
                    .set_ColWidth(3, 0, 1900) 'Columna del Importe
                    .set_ColWidth(4, 0, 1500) 'Columna del Estatus
                    .set_ColAlignment(4, 4)
                    For I = 1 To FrmConsultas.Flexdet.Rows - 1
                        FrmConsultas.Flexdet.set_TextMatrix(I, 2, Format(FrmConsultas.Flexdet.get_TextMatrix(I, 2), "dd/MMM/yyyy"))
                        FrmConsultas.Flexdet.set_TextMatrix(I, 3, Format(FrmConsultas.Flexdet.get_TextMatrix(I, 3), "###,##0.00"))
                    Next
                    FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
                    FrmConsultas.Left = VB6.TwipsToPixelsX(1800)
            End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatosCliente()
        On Error GoTo Merr
        Dim RsAux As ADODB.Recordset
        If Val(txtCodigo.Text) = 0 Then
            Exit Sub
        End If
        'txtCodigo.Text = Format(txtCodigo.Text, "00000")
        For I = 0 To 4 - txtCodigo.TextLength
            txtCodigo.Text = String.Concat("0" + txtCodigo.Text)
        Next I

        gStrSql = "SELECT * FROM CatRfc WHERE CodRfc = " & CInt(Numerico((txtCodigo.Text)))
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsAux = Cmd.Execute
        If RsAux.RecordCount > 0 Then
            txtCodigo.ReadOnly = True
            txtNombreCliente.Text = Trim(RsAux.Fields("DescClienteRFC").Value)
            txtNombreCliente.ReadOnly = True
            txtDomicilio.Text = Trim(RsAux.Fields("Domicilio").Value)
            txtDomicilio.ReadOnly = True
            txtColonia.Text = Trim(RsAux.Fields("Colonia").Value)
            txtColonia.ReadOnly = True
            txtCiudad.Text = Trim(RsAux.Fields("Ciudad").Value)
            txtCiudad.ReadOnly = True
            txtCP.Text = Trim(RsAux.Fields("CP").Value)
            txtCP.ReadOnly = True
            txtRFC.Text = Trim(RsAux.Fields("Rfc").Value)
            txtRFC.ReadOnly = True
        Else
            MsjNoExiste("El Cliente", gstrNombCortoEmpresa)
            Limpiar()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatos()
        On Error GoTo Merr
        Dim I As Integer
        If txtFolioFactura.Text = "" Then
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT * FROM Facturas WHERE FolioFactura = '" & txtFolioFactura.Text & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFolioFactura.Text = Trim(RsGral.Fields("FolioFactura").Value)
            dtpFecha.Value = Format(RsGral.Fields("FechaFactura").Value, C_FORMATFECHAMOSTRAR)
            'dtpFecha.Enabled = False
            If RsGral.Fields("Condicion").Value = "CO" Then
                _optContado_0.Checked = True
            ElseIf RsGral.Fields("Condicion").Value = "CR" Then
                optCredito(1).Checked = True
            End If
            CodCaja = IIf(IsDBNull(RsGral.Fields("CodCaja").Value), 0, RsGral.Fields("CodCaja").Value)
            CodSucursal = RsGral.Fields("CodSucursal").Value
            txtCodigo.Text = RsGral.Fields("CodCliente").Value
            LlenaDatosCliente()
            txtNombreCliente.Text = Trim(RsGral.Fields("Nombre").Value)
            txtRFC.Text = RsGral.Fields("Rfc").Value
            If RsGral.Fields("Moneda").Value = C_PESO Then
                _optMoneda_0.Checked = True
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                optMoneda(1).Checked = True
            End If
            txtTipoCambio.Text = Format(RsGral.Fields("TipoCambio").Value, "###,##0.00")
            txtTipoCambio.ReadOnly = True
            txtSubTotal.Text = Format(RsGral.Fields("SubTotal").Value, "###,##0.00")
            txtIva.Text = Format(RsGral.Fields("Iva").Value, "###,##0.00")
            txtTotal.Text = Format(RsGral.Fields("Total").Value, "###,##0.00")
            txtRedondeo.Text = Format(RsGral.Fields("Redondeo").Value, "###,##0.00")
            txtTotalFinal.Text = Format(RsGral.Fields("Total").Value + RsGral.Fields("Redondeo").Value, "###,##0.00")
            If RsGral.Fields("DesgloseIva").Value = True Then
                DesgloseIva = 1
                chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                DesgloseIva = 0
                chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
            chkDesglosarIva.Enabled = False
            If RsGral.Fields("Estatus").Value = "C" Then
                Label4.Text = "Factura Cancelada"
                mblnCancelar = True
            End If
            I = 1
            Do While Not RsGral.EOF
                With flexDetalle
                    .set_TextMatrix(I, 0, RsGral.Fields("Cantidad").Value)
                    .set_TextMatrix(I, 1, Trim(RsGral.Fields("DescEspecial").Value))
                    .set_TextMatrix(I, 2, Format(RsGral.Fields("Precio").Value * (RsGral.Fields("PorcIvaP").Value / 100), "###,##0.00"))
                    '.set_TextMatrix(I, 2, Format(.get_TextMatrix(I, 2), "###,##0.00"))
                    .set_TextMatrix(I, 3, Format(RsGral.Fields("Precio").Value, "###,##0.00"))
                    .set_TextMatrix(I, 4, Format(RsGral.Fields("importe").Value, "###,##0.00"))
                End With
                I = I + 1
                If I > flexDetalle.Rows Then
                    flexDetalle.Rows = flexDetalle.Rows + 1
                End If
                RsGral.MoveNext()
            Loop
            Frame1.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
            ToolTip1.SetToolTip(flexDetalle, "")
            mblnNuevo = False
        Else
            MsgBox("Folio de Factura no Existe ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFolioFactura.Focus()
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub EliminarLinea()
        txtSubTotal.Text = CStr(CDbl(Numerico(Format(txtSubTotal.Text, "#####0.00"))) - CDbl(Numerico(Format(flexDetalle.get_TextMatrix(flexDetalle.Row, 4), "#####0.00"))))
        txtSubTotal.Text = Format(txtSubTotal.Text, "###,##0.00")
        txtIva.Text = CStr(CDbl(Numerico(Format(txtIva.Text, "#####0.00"))) - (CDbl(Numerico(Format(flexDetalle.get_TextMatrix(flexDetalle.Row, 6), "#####0.00"))) * CInt(Numerico(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)))))
        txtIva.Text = Format(txtIva.Text, "###,##0.00")
        txtTotal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) - (CDbl(Numerico(Format(flexDetalle.get_TextMatrix(flexDetalle.Row, 4), "#####0.00"))) + (CDbl(Numerico(Format(flexDetalle.get_TextMatrix(flexDetalle.Row, 6), "#####0.00"))) * CInt(Numerico(flexDetalle.get_TextMatrix(flexDetalle.Row, 0))))))
        txtTotal.Text = Format(txtTotal.Text, "###,##0.00")
        flexDetalle.RemoveItem((flexDetalle.Row))
        flexDetalle.Rows = flexDetalle.Rows + 1
    End Sub

    Sub Limpiar()
        On Error Resume Next
        txtFolioFactura.Text = ""
        Nuevo()
        txtFolioFactura.Focus()
    End Sub

    Sub Nuevo()
        dtpFecha.Value = Today
        dtpFecha.Enabled = True
        _optContado_0.Checked = True
        _optMoneda_0.Checked = True
        chkDesglosarIva.Enabled = True
        txtCodigo.Text = ""
        txtNombreCliente.Text = ""
        txtDomicilio.Text = ""
        txtColonia.Text = ""
        txtCiudad.Text = ""
        txtCP.Text = ""
        txtRFC.Text = ""
        flexDetalle.Clear()
        Encabezado()
        txtSubTotal.Text = "0.00"
        txtIva.Text = "0.00"
        txtTotal.Text = "0.00"
        txtTipoCambio.Text = CStr(gcurCorpoTIPOCAMBIODOLAR)
        txtTipoCambio.Text = VB6.Format(txtTipoCambio.Text, "###,##0.00")
        'txtTipoCambio.ReadOnly = False
        txtRedondeo.Text = "0.00"
        txtTotalFinal.Text = "0.00"
        Label4.Text = ""
        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
        txtCodigo.ReadOnly = False
        txtNombreCliente.ReadOnly = False
        txtDomicilio.ReadOnly = False
        txtColonia.ReadOnly = False
        txtCiudad.ReadOnly = False
        txtCP.ReadOnly = False
        txtRFC.ReadOnly = False
        ToolTip1.SetToolTip(flexDetalle, "Supr Para Eliminar una Linea.")
        chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked
        InicializaVariables()
    End Sub

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnSalir = False
        mblnContado = True
        mblnCredito = False
        mblnPierdeFoco = False
        mblnCancelar = False
        mstrCorporativo = "00"
        Redondeo = 0
        DesgloseIva = 0
    End Sub

    Sub Encabezado()
        With flexDetalle
            .Row = 0
            .Col = 0
            .set_ColWidth(0, 0, 1000)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Cantidad"
            .Col = 1
            .set_ColWidth(1, 0, 3800)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Descripción"
            .Col = 2
            .set_ColWidth(2, 0, 1300)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Iva"
            .Col = 3
            .set_ColWidth(3, 0, 1800)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Precio"
            .Col = 4
            .set_ColWidth(4, 0, 1800)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Importe"
            .Col = 5
            .set_ColWidth(5, 0, 0) 'Columna para Guardar el Porcentaje de Iva
            .Col = 6
            .set_ColWidth(6, 0, 0) 'Columna para guardar los Importes de Iva
            .Col = 7
            .set_ColWidth(7, 0, 0) 'Columna para Guardar los Importes
            .Row = 1
            .Col = 0
        End With
    End Sub

    Private Sub CambiarFormatoTxtenCaptura()
        With txtFlex
            Select Case flexDetalle.Col
                Case 0 'Cantidad
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 5
                Case 1 'Descripción
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                    .MaxLength = 100
                Case 2 'Iva
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 5
                Case 3 'Precio
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 16
            End Select
        End With
    End Sub

    Private Sub cmdAbcRfc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAbcRfc.Click
        frmABCRFC.Show()
        frmABCRFC.BringToFront()
    End Sub

    Private Sub flexDetalle_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.ClickEvent
        txtFlex.Visible = False
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        Pon_Tool()
    End Sub

    Private Sub FlexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexDetalle.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete And mblnNuevo Then
            EliminarLinea()
        End If
    End Sub

    Private Sub FlexDetalle_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexDetalle.KeyPressEvent
        Dim lonR, lonI As Integer
        If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape And mblnNuevo Then
            'Verifica si se puede capturar la fila
            If flexDetalle.Row > 1 Then
                If flexDetalle.get_TextMatrix(flexDetalle.Row - 1, 0) <> "" Then
                    For lonR = 1 To flexDetalle.Row - 1 Step 1
                        For lonI = 0 To 4 Step 1
                            If flexDetalle.get_TextMatrix(lonR, lonI) = "" Then
                                'MsgBox "Hace falta información en la captura", vbExclamation, cNomEmp
                                flexDetalle.Row = lonR
                                flexDetalle.Col = lonI
                                If flexDetalle.Col = 0 Then
                                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                                ElseIf flexDetalle.Col = 1 Then
                                    ModEstandar.gp_CampoLetras(eventArgs.keyAscii, "0123456789-_*+-/!#$%&/()=?¿¡',;.:")
                                ElseIf flexDetalle.Col = 2 Then
                                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                                ElseIf flexDetalle.Col = 3 Then
                                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                                End If
                                CambiarFormatoTxtenCaptura()
                                MSHFlexGridEdit(flexDetalle, txtFlex, eventArgs.keyAscii)
                                Exit Sub
                            End If
                        Next lonI
                    Next lonR
                Else
                    'flexDetalle.SetFocus
                    Exit Sub
                End If
            End If
            'Edita el campo sólo si es Editable
            If flexDetalle.Row >= 1 And flexDetalle.Col < 4 Then
                If flexDetalle.Col = 0 Then
                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                ElseIf flexDetalle.Col = 1 Then
                    ModEstandar.gp_CampoLetras(eventArgs.keyAscii, "0123456789-_*+-/!#$%&/()=?¿¡',;.:")
                ElseIf flexDetalle.Col = 2 Then
                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                ElseIf flexDetalle.Col = 3 Then
                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                End If
                CambiarFormatoTxtenCaptura()
                MSHFlexGridEdit(flexDetalle, txtFlex, eventArgs.keyAscii)
                If Len(Trim(txtFlex.Text)) = 1 Then
                    System.Windows.Forms.SendKeys.Send("{right}")
                End If
                If flexDetalle.Col = 2 Then
                    If flexDetalle.get_TextMatrix(flexDetalle.Row, 5) = "" Then
                        txtFlex.Text = Format(gcurCorpoTASAIVA, "##0.00")
                    Else
                        txtFlex.Text = Format(flexDetalle.get_TextMatrix(flexDetalle.Row, 5), "##0.00")
                    End If
                End If
            ElseIf flexDetalle.Col = 4 Then
                If flexDetalle.Row = 10 Then
                    MsgBox("Ha Llegado al Limite de Partidas para la Factura...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Sub
                End If
                flexDetalle.Focus()
                If flexDetalle.Row < flexDetalle.Rows - 1 Then
                    flexDetalle.Row = flexDetalle.Row + 1
                    flexDetalle.Col = 0
                Else
                    flexDetalle.Rows = flexDetalle.Rows + 1
                    flexDetalle.Row = flexDetalle.Row + 1
                    flexDetalle.TopRow = flexDetalle.Row
                    flexDetalle.Col = 0
                End If
            End If
        End If
    End Sub

    Private Sub FlexDetalle_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Scroll
        txtFlex.Visible = False
    End Sub

    Private Sub frmFactFacturacionEspecial_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmFactFacturacionEspecial_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmFactFacturacionEspecial_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtFolioFactura" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmFactFacturacionEspecial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmFactFacturacionEspecial_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        ModCorporativo.ValidarRegistrodeImpresoras()
        Encabezado()
        Nuevo()
    End Sub

    Private Sub frmFactFacturacionEspecial_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
        'Else
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0
        '        Case MsgBoxResult.No
        '            mblnSalir = False
        '            txtFolioFactura.Focus()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmFactFacturacionEspecial_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
        'MDIMenuPrincipalCorpo.mnuFacturacionOpc(1).Enabled = True
    End Sub

    Private Sub optContado_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optContado.Enter
        'Dim Index As Integer = optContado.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub optCredito_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCredito.Enter
        'Dim Index As Integer = optCredito.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub optMoneda_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoneda.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer
            '= optMoneda.GetIndex(eventSender)
            Dim RedondeoPesos As Double
            Dim RedondeoDolares As Double
            Dim TotalPesos As Double
            Dim TipoCambio As Double
            'TipoCambio = CDbl(Format(String.Concat(txtTipoCambio.Text, "#####0.00")))
            TipoCambio = Format(txtTipoCambio.Text, ",0.00")
            Select Case Index
                Case 0
                    RedondeoPesos = RedondeoUnidadFinalFEsp(CDec(Numerico(Format(txtTotal.Text, "#####0.00"))), CDbl(gcurRedondeo))
                    Redondeo = RedondeoPesos
                    txtRedondeo.Text = Format(RedondeoPesos, "###,##0.00")
                    txtTotalFinal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtRedondeo.Text, "#####0.00"))))
                    txtTotalFinal.Text = Format(txtTotalFinal.Text, "###,##0.00")
                Case 1
                    TotalPesos = CDbl(Format(CDbl(Numerico(txtTotal.Text)) * TipoCambio, "#####0.000000"))
                    TotalPesos = CDbl(Format(TotalPesos, "#####0.00"))
                    RedondeoPesos = RedondeoUnidadFinalFEsp(CDec(TotalPesos), CDbl(gcurRedondeo))
                    RedondeoDolares = CDbl(Format(RedondeoPesos / TipoCambio, "#####0.0000"))
                    Redondeo = RedondeoDolares
                    txtRedondeo.Text = Format(RedondeoDolares, "###,##0.00")
                    txtTotalFinal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtRedondeo.Text, "#####0.00"))))
                    txtTotalFinal.Text = Format(txtTotalFinal.Text, "###,##0.00")
            End Select
        End If
    End Sub

    Private Sub txtCiudad_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCiudad.Enter
        Pon_Tool()
    End Sub

    Private Sub txtCodigo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.Enter
        strControlActual = UCase("txtCodigo")
        Pon_Tool()
    End Sub

    Private Sub txtCodigo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.Leave
        LlenaDatosCliente()
    End Sub

    Private Sub txtColonia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtColonia.Enter
        Pon_Tool()
    End Sub

    Private Sub txtCP_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCP.Enter
        Pon_Tool()
    End Sub

    Private Sub txtDomicilio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.Enter
        Pon_Tool()
    End Sub

    Private Sub txtFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Enter
        SelTextoTxt(txtFlex)
        Pon_Tool()
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Dim RedondeoPesos As Double
        Dim RedondeoDolares As Double
        Dim TotalPesos As Double
        Dim TipoCambio As Double
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        With flexDetalle
            If KeyCode = System.Windows.Forms.Keys.Return Then
                Select Case .Col
                    Case 0, 1
                        If .Col = 0 And CInt(Numerico(.get_TextMatrix(.Row, 0))) <> CInt(Numerico(txtFlex.Text)) Then
                            txtIva.Text = CStr(CDbl(Numerico(Format(txtIva.Text, "#####0.00"))) - (CDbl(Numerico(Format(.get_TextMatrix(.Row, 6), "#####0.00"))) * CInt(Numerico(.get_TextMatrix(.Row, 0)))))
                            txtIva.Text = CStr(CDbl(Numerico(Format(txtIva.Text, "#####0.00"))) + (CDbl(Numerico(Format(.get_TextMatrix(.Row, 2), "#####0.00"))) * CInt(Numerico(txtFlex.Text))))
                            txtIva.Text = Format(txtIva.Text, "###,##0.00")
                            txtTotal.Text = CStr(CDbl(Numerico(Format(txtSubTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtIva.Text, "#####0.00"))))
                            txtTotal.Text = Format(txtTotal.Text, "###,##0.00")
                            TipoCambio = CDbl(Format(txtTipoCambio.Text, "#####0.00"))
                            If _optMoneda_0.Checked Then
                                RedondeoPesos = RedondeoUnidadFinalFEsp(CDec(Numerico(Format(txtTotal.Text, "#####0.00"))), CDbl(gcurRedondeo))
                                Redondeo = RedondeoPesos
                                txtRedondeo.Text = Format(RedondeoPesos, "###,##0.00")
                                txtTotalFinal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtRedondeo.Text, "#####0.00"))))
                                txtTotalFinal.Text = Format(txtTotalFinal.Text, "###,##0.00")
                            ElseIf optMoneda(1).Checked Then
                                TotalPesos = CDbl(Format(CDbl(Numerico(txtTotal.Text)) * TipoCambio, "#####0.000000"))
                                TotalPesos = CDbl(Format(TotalPesos, "#####0.00"))
                                RedondeoPesos = RedondeoUnidadFinalFEsp(CDec(TotalPesos), CDbl(gcurRedondeo))
                                RedondeoDolares = CDbl(Format(RedondeoPesos / TipoCambio, "#####0.0000"))
                                Redondeo = RedondeoDolares
                                txtRedondeo.Text = Format(RedondeoDolares, "###,##0.00")
                                txtTotalFinal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtRedondeo.Text, "#####0.00"))))
                                txtTotalFinal.Text = Format(txtTotalFinal.Text, "###,##0.00")
                            End If
                        End If
                        .Text = Trim(txtFlex.Text)
                        .Col = .Col + 1
                        txtFlex.Visible = False
                        mblnPierdeFoco = True
                        CambiarFormatoTxtenCaptura()
                        MSHFlexGridEdit(flexDetalle, txtFlex, KeyCode)
                        If flexDetalle.Col = 2 Then
                            If flexDetalle.get_TextMatrix(flexDetalle.Row, 5) = "" Then
                                txtFlex.Text = Format(gcurCorpoTASAIVA, "##0.00")
                            Else
                                txtFlex.Text = Format(flexDetalle.get_TextMatrix(flexDetalle.Row, 5), "##0.00")
                            End If
                        End If
                        If flexDetalle.Col = 1 And CDbl(Numerico(.get_TextMatrix(flexDetalle.Row, 3))) <> 0 Then
                            .set_TextMatrix(.Row, 4, Format(CDbl(Numerico(.get_TextMatrix(.Row, 0))) * CDbl(Numerico(.get_TextMatrix(.Row, 3))), "###,##0.00"))
                            If CDbl(Numerico(Format(.get_TextMatrix(.Row, 4), "#####0.00"))) <> CDbl(Numerico(Format(.get_TextMatrix(.Row, 7), "#####0.00"))) Then
                                If CDbl(Numerico(Format(.get_TextMatrix(.Row, 7), "#####0.00"))) <> 0 Then
                                    txtSubTotal.Text = CStr(CDbl(Numerico(Format(txtSubTotal.Text, "#####0.00"))) - CDbl(Numerico(Format(.get_TextMatrix(.Row, 7), "#####0.00"))))
                                End If
                                txtSubTotal.Text = CStr(CDbl(Numerico(Format(txtSubTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(.get_TextMatrix(.Row, 4), "#####0.00"))))
                                txtSubTotal.Text = Format(txtSubTotal.Text, "###,##0.00")
                                txtTotal.Text = CStr(CDbl(Numerico(Format(txtSubTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtIva.Text, "#####0.00"))))
                                txtTotal.Text = Format(txtTotal.Text, "###,##0.00")
                                TipoCambio = CDbl(Format(txtTipoCambio.Text, "#####0.00"))
                                If _optMoneda_0.Checked Then
                                    RedondeoPesos = RedondeoUnidadFinalFEsp(CDec(Numerico(Format(txtTotal.Text, "#####0.00"))), CDbl(gcurRedondeo))
                                    Redondeo = RedondeoPesos
                                    txtRedondeo.Text = Format(RedondeoPesos, "###,##0.00")
                                    txtTotalFinal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtRedondeo.Text, "#####0.00"))))
                                    txtTotalFinal.Text = Format(txtTotalFinal.Text, "###,##0.00")
                                ElseIf optMoneda(1).Checked Then
                                    TotalPesos = CDbl(Format(CDbl(Numerico(txtTotal.Text)) * TipoCambio, "#####0.000000"))
                                    TotalPesos = CDbl(Format(TotalPesos, "#####0.00"))
                                    RedondeoPesos = RedondeoUnidadFinalFEsp(CDec(TotalPesos), CDbl(gcurRedondeo))
                                    RedondeoDolares = CDbl(Format(RedondeoPesos / TipoCambio, "#####0.0000"))
                                    Redondeo = RedondeoDolares
                                    txtRedondeo.Text = Format(RedondeoDolares, "###,##0.00")
                                    txtTotalFinal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtRedondeo.Text, "#####0.00"))))
                                    txtTotalFinal.Text = Format(txtTotalFinal.Text, "###,##0.00")
                                End If
                            End If
                            .set_TextMatrix(.Row, 7, .get_TextMatrix(.Row, 4))
                        End If
                    Case 2
                        .set_TextMatrix(.Row, 5, txtFlex.Text)
                        .Text = Format(CDbl(Numerico(.get_TextMatrix(.Row, 3))) * CDbl(Format(CDbl(Numerico(.get_TextMatrix(.Row, 5))) / 100, "##0.0000")), "###,##0.00")
                        .Col = .Col + 1
                        txtFlex.Visible = False
                        mblnPierdeFoco = True
                        CambiarFormatoTxtenCaptura()
                        MSHFlexGridEdit(flexDetalle, txtFlex, KeyCode)
                        If CDbl(Numerico(Format(.get_TextMatrix(.Row, 2), "#####0.00"))) <> CDbl(Numerico(Format(.get_TextMatrix(.Row, 6), "#####0.00"))) Then
                            If CDbl(Numerico(Format(.get_TextMatrix(.Row, 6), "#####0.00"))) <> 0 Then
                                txtIva.Text = CStr(CDbl(Numerico(Format(txtIva.Text, "#####0.00"))) - (CDbl(Numerico(Format(.get_TextMatrix(.Row, 6), "#####0.00"))) * CInt(Numerico(.get_TextMatrix(.Row, 0)))))
                            End If
                            txtIva.Text = CStr(CDbl(Numerico(Format(txtIva.Text, "#####0.00"))) + (CDbl(Numerico(Format(.get_TextMatrix(.Row, 2), "#####0.00"))) * CInt(Numerico(.get_TextMatrix(.Row, 0)))))
                            txtIva.Text = Format(txtIva.Text, "###,##0.00")
                            txtTotal.Text = CStr(CDbl(Numerico(Format(txtSubTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtIva.Text, "#####0.00"))))
                            txtTotal.Text = Format(txtTotal.Text, "###,##0.00")
                            TipoCambio = CDbl(Format(txtTipoCambio.Text, "#####0.00"))
                            If _optMoneda_0.Checked Then
                                RedondeoPesos = RedondeoUnidadFinalFEsp(CDec(Numerico(Format(txtTotal.Text, "#####0.00"))), CDbl(gcurRedondeo))
                                Redondeo = RedondeoPesos
                                txtRedondeo.Text = Format(RedondeoPesos, "###,##0.00")
                                txtTotalFinal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtRedondeo.Text, "#####0.00"))))
                                txtTotalFinal.Text = Format(txtTotalFinal.Text, "###,##0.00")
                            ElseIf optMoneda(1).Checked Then
                                TotalPesos = CDbl(Format(CDbl(Numerico(txtTotal.Text)) * TipoCambio, "#####0.000000"))
                                TotalPesos = CDbl(Format(TotalPesos, "#####0.00"))
                                RedondeoPesos = RedondeoUnidadFinalFEsp(CDec(TotalPesos), CDbl(gcurRedondeo))
                                RedondeoDolares = CDbl(Format(RedondeoPesos / TipoCambio, "#####0.0000"))
                                Redondeo = RedondeoDolares
                                txtRedondeo.Text = Format(RedondeoDolares, "###,##0.00")
                                txtTotalFinal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtRedondeo.Text, "#####0.00"))))
                                txtTotalFinal.Text = Format(txtTotalFinal.Text, "###,##0.00")
                            End If
                        End If
                        .set_TextMatrix(.Row, 6, .get_TextMatrix(.Row, 2))
                    Case 3
                        If .get_TextMatrix(.Row, 5) = "" Then
                            .set_TextMatrix(.Row, 5, gcurCorpoTASAIVA)
                        End If
                        .Text = Format(txtFlex.Text, "###,##0.00")
                        .set_TextMatrix(.Row, 2, Format(CDbl(Numerico(.get_TextMatrix(.Row, 3))) * CDbl(Format(CDbl(Numerico(.get_TextMatrix(.Row, 5))) / 100, "##0.0000")), "###,##0.00"))
                        If CDbl(Numerico(Format(.get_TextMatrix(.Row, 2), "#####0.00"))) <> CDbl(Numerico(Format(.get_TextMatrix(.Row, 6), "#####0.00"))) Then
                            If CDbl(Numerico(Format(.get_TextMatrix(.Row, 6), "#####0.00"))) <> 0 Then
                                txtIva.Text = CStr(CDbl(Numerico(Format(txtIva.Text, "#####0.00"))) - (CDbl(Numerico(Format(.get_TextMatrix(.Row, 6), "#####0.00"))) * CInt(Numerico(.get_TextMatrix(.Row, 0)))))
                            End If
                            txtIva.Text = CStr(CDbl(Numerico(Format(txtIva.Text, "#####0.00"))) + (CDbl(Numerico(Format(.get_TextMatrix(.Row, 2), "#####0.00"))) * CInt(Numerico(.get_TextMatrix(.Row, 0)))))
                            txtIva.Text = Format(txtIva.Text, "###,##0.00")
                        End If
                        .set_TextMatrix(.Row, 6, .get_TextMatrix(.Row, 2))
                        .set_TextMatrix(.Row, 4, Format(CDbl(Numerico(.get_TextMatrix(.Row, 0))) * CDbl(Numerico(.get_TextMatrix(.Row, 3))), "###,##0.00"))
                        If CDbl(Numerico(.get_TextMatrix(.Row, 4))) <> CDbl(Numerico(.get_TextMatrix(.Row, 7))) Then
                            If CDbl(Numerico(.get_TextMatrix(.Row, 7))) <> 0 Then
                                txtSubTotal.Text = CStr(CDbl(Numerico(Format(txtSubTotal.Text, "#####0.00"))) - CDbl(Numerico(Format(.get_TextMatrix(.Row, 7), "#####0.00"))))
                            End If
                            txtSubTotal.Text = CStr(CDbl(Numerico(Format(txtSubTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(.get_TextMatrix(.Row, 4), "#####0.00"))))
                            txtSubTotal.Text = Format(txtSubTotal.Text, "###,##0.00")
                        End If
                        txtTotal.Text = CStr(CDbl(Numerico(Format(txtSubTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtIva.Text, "#####0.00"))))
                        txtTotal.Text = Format(txtTotal.Text, "###,##0.00")
                        TipoCambio = CDbl(Format(txtTipoCambio.Text, "#####0.00"))
                        If _optMoneda_0.Checked Then
                            RedondeoPesos = RedondeoUnidadFinalFEsp(CDec(Numerico(Format(txtTotal.Text, "#####0.00"))), CDbl(gcurRedondeo))
                            Redondeo = RedondeoPesos
                            txtRedondeo.Text = Format(RedondeoPesos, "###,##0.00")
                            txtTotalFinal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtRedondeo.Text, "#####0.00"))))
                            txtTotalFinal.Text = Format(txtTotalFinal.Text, "###,##0.00")
                        ElseIf optMoneda(1).Checked Then
                            TotalPesos = CDbl(Format(CDbl(Numerico(txtTotal.Text)) * TipoCambio, "#####0.000000"))
                            TotalPesos = CDbl(Format(TotalPesos, "#####0.00"))
                            RedondeoPesos = RedondeoUnidadFinalFEsp(CDec(TotalPesos), CDbl(gcurRedondeo))
                            RedondeoDolares = CDbl(Format(RedondeoPesos / TipoCambio, "#####0.0000"))
                            Redondeo = RedondeoDolares
                            txtRedondeo.Text = Format(RedondeoDolares, "###,##0.00")
                            txtTotalFinal.Text = CStr(CDbl(Numerico(Format(txtTotal.Text, "#####0.00"))) + CDbl(Numerico(Format(txtRedondeo.Text, "#####0.00"))))
                            txtTotalFinal.Text = Format(txtTotalFinal.Text, "###,##0.00")
                        End If
                        .set_TextMatrix(.Row, 7, .get_TextMatrix(.Row, 4))
                        .Col = .Col + 1
                        txtFlex.Visible = False
                End Select
            ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
                .Focus()
                txtFlex.Visible = False
            End If
        End With
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
            Case Else
                Select Case flexDetalle.Col
                    Case 0
                        ModEstandar.gp_CampoNumerico(KeyAscii)
                    Case 1
                        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "-_*+-/!#$%&\()=?¿¡,;.:@")
                    Case 2
                        KeyAscii = ModEstandar.MskCantidad(txtFlex.Text, KeyAscii, 2, 2, (txtFlex.SelectionStart))
                    Case 3
                        KeyAscii = ModEstandar.MskCantidad(txtFlex.Text, KeyAscii, 10, 2, (txtFlex.SelectionStart))
                End Select
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        If Not mblnPierdeFoco Then
            txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
        Else
            mblnPierdeFoco = False
        End If
    End Sub

    Private Sub txtFolioFactura_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioFactura.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtFolioFactura_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioFactura.Enter
        strControlActual = UCase("txtFolioFactura")
        SelTextoTxt(txtFolioFactura)
        Pon_Tool()
    End Sub

    Private Sub txtFolioFactura_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioFactura.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, "F")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioFactura_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioFactura.Leave
        Dim Fecha As String
        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If Trim(txtFolioFactura.Text) = "" Then
            Fecha = Format(Year(dtpFecha.Value), "0000") & Format(Month(dtpFecha.Value), "00") & Format((dtpFecha.Value), "00")
            txtFolioFactura.Text = "F00" & Fecha & "000000"
            Exit Sub
        End If
        If mblnCambiosEnCodigo = True And txtFolioFactura.Text <> "" Then 'si hubo cambios en el codigo hace la consulta
            LlenaDatos()
        End If
    End Sub

    Private Sub txtIVA_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtIva.Enter
        Pon_Tool()
    End Sub

    Private Sub txtNombreCliente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombreCliente.Enter
        strControlActual = UCase("txtNombreCliente")
        Pon_Tool()
    End Sub

    Private Sub txtRedondeo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRedondeo.Enter
        Pon_Tool()
    End Sub

    Private Sub txtRFC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRFC.Enter
        Pon_Tool()
    End Sub

    Private Sub txtSubTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSubTotal.Enter
        Pon_Tool()
    End Sub

    Private Sub txtTipoCambio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Enter
        Pon_Tool()
    End Sub

    Private Sub txtTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTotal.Enter
        Pon_Tool()
    End Sub

    Private Sub txtTotalFinal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTotalFinal.Enter
        Pon_Tool()
    End Sub

    Sub Imprime()
        ImprimirFactura()
    End Sub

    Private Function RedondeoUnidadFinalFEsp(ByRef importe As Double, ByRef Decimales As Double) As Double
        Dim ParteInf As Double
        Dim ParteSup As Double
        Dim ParteEnt As Integer
        Dim ParteDec As Double

        Dim Limite As Double
        Dim LimInfIni As Double
        Dim LimInfFin As Double
        Dim LimSupIni As Double
        Dim LimSupFin As Double

        '''Determina limite inferior y superior del rango permitido
        ParteEnt = Fix(importe)
        ParteDec = System.Math.Round(System.Math.Abs(importe - ParteEnt), 4)
        Limite = ((importe - ParteDec) \ Decimales) * Decimales

        If Limite < importe Then
            ParteInf = ((importe - ParteDec) \ Decimales) * Decimales
            ParteSup = ParteInf + Decimales
        Else
            ParteSup = ((importe - ParteDec) \ Decimales) * Decimales
            ParteInf = ParteInf - Decimales
        End If

        '''Determina los limites Ini y Fin de los limites Inf y Sup
        LimInfIni = ParteInf
        LimInfFin = ParteInf + (Decimales / 2)
        LimSupIni = ParteSup - (Decimales / 2) + 0.01
        LimSupFin = ParteSup
        importe = System.Math.Round(importe, 2)

        Select Case importe
            Case LimInfIni To LimInfFin
                RedondeoUnidadFinalFEsp = -1 * (importe - ParteInf)
            Case LimSupIni To LimSupFin
                RedondeoUnidadFinalFEsp = (ParteSup - importe)
        End Select
        RedondeoUnidadFinalFEsp = System.Math.Round(RedondeoUnidadFinalFEsp, 4)
    End Function

    Private Sub brnBuscar_Click(sender As Object, e As EventArgs) Handles brnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub
End Class