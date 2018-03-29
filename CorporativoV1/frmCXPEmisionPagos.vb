Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms
Imports System.Data
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCXPEmisionPagos
    Inherits System.Windows.Forms.Form

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents _optOrigen_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optOrigen_1 As System.Windows.Forms.RadioButton
    Public WithEvents _fraProgPago_0 As System.Windows.Forms.GroupBox
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents _fraPagos_3 As System.Windows.Forms.GroupBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents btnGenerar As System.Windows.Forms.Button
    Public WithEvents _fraPagos_0 As System.Windows.Forms.GroupBox
    Public WithEvents txtTipoCambioEuro As System.Windows.Forms.TextBox
    Public WithEvents txtTipoCambio As System.Windows.Forms.TextBox
    Public WithEvents txtAPagar As System.Windows.Forms.TextBox
    Public WithEvents txtAnticipos As System.Windows.Forms.TextBox
    Public WithEvents txtNotasCredito As System.Windows.Forms.TextBox
    Public WithEvents txtDescuentoFinanciero As System.Windows.Forms.TextBox
    Public WithEvents txtFacturas As System.Windows.Forms.TextBox
    Public WithEvents _lblPago_8 As System.Windows.Forms.Label
    Public WithEvents lblEuro As System.Windows.Forms.Label
    Public WithEvents lblDolar As System.Windows.Forms.Label
    Public WithEvents _lblPago_7 As System.Windows.Forms.Label
    Public WithEvents _lblPago_6 As System.Windows.Forms.Label
    Public WithEvents _lblPago_5 As System.Windows.Forms.Label
    Public WithEvents _lblPago_4 As System.Windows.Forms.Label
    Public WithEvents _lblPago_3 As System.Windows.Forms.Label
    Public WithEvents _fraPagos_1 As System.Windows.Forms.GroupBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents mshPagos As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents mshAnticipos As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents mshNotasCredito As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents dtpCorte As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblPago_0 As System.Windows.Forms.Label
    Public WithEvents _fraPagos_2 As System.Windows.Forms.GroupBox
    Public WithEvents lblCR As System.Windows.Forms.Label
    Public WithEvents _lblPago_9 As System.Windows.Forms.Label
    Public WithEvents _lblPago_10 As System.Windows.Forms.Label
    Public WithEvents _lblPago_2 As System.Windows.Forms.Label
    Public WithEvents _lblPago_1 As System.Windows.Forms.Label
    Public WithEvents lblProveedor As System.Windows.Forms.Label
    Public WithEvents fraPagos As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents fraProgPago As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblPago As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optOrigen As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Public Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCXPEmisionPagos))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me._fraProgPago_0 = New System.Windows.Forms.GroupBox
        Me._optOrigen_0 = New System.Windows.Forms.RadioButton
        Me._optOrigen_1 = New System.Windows.Forms.RadioButton
        Me._fraPagos_3 = New System.Windows.Forms.GroupBox
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton
        Me.txtFlex = New System.Windows.Forms.TextBox
        Me.btnGenerar = New System.Windows.Forms.Button
        Me._fraPagos_0 = New System.Windows.Forms.GroupBox
        Me._fraPagos_1 = New System.Windows.Forms.GroupBox
        Me.txtTipoCambioEuro = New System.Windows.Forms.TextBox
        Me.txtTipoCambio = New System.Windows.Forms.TextBox
        Me.txtAPagar = New System.Windows.Forms.TextBox
        Me.txtAnticipos = New System.Windows.Forms.TextBox
        Me.txtNotasCredito = New System.Windows.Forms.TextBox
        Me.txtDescuentoFinanciero = New System.Windows.Forms.TextBox
        Me.txtFacturas = New System.Windows.Forms.TextBox
        Me._lblPago_8 = New System.Windows.Forms.Label
        Me.lblEuro = New System.Windows.Forms.Label
        Me.lblDolar = New System.Windows.Forms.Label
        Me._lblPago_7 = New System.Windows.Forms.Label
        Me._lblPago_6 = New System.Windows.Forms.Label
        Me._lblPago_5 = New System.Windows.Forms.Label
        Me._lblPago_4 = New System.Windows.Forms.Label
        Me._lblPago_3 = New System.Windows.Forms.Label
        Me.dbcProveedor = New System.Windows.Forms.ComboBox
        Me.mshPagos = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
        Me.mshAnticipos = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
        Me.mshNotasCredito = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
        Me._fraPagos_2 = New System.Windows.Forms.GroupBox
        Me.dtpCorte = New System.Windows.Forms.DateTimePicker
        Me._lblPago_0 = New System.Windows.Forms.Label
        Me.lblCR = New System.Windows.Forms.Label
        Me._lblPago_9 = New System.Windows.Forms.Label
        Me._lblPago_10 = New System.Windows.Forms.Label
        Me._lblPago_2 = New System.Windows.Forms.Label
        Me._lblPago_1 = New System.Windows.Forms.Label
        Me.lblProveedor = New System.Windows.Forms.Label
        Me.fraPagos = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
        Me.fraProgPago = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
        Me.lblPago = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
        Me.optOrigen = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
        Me._fraProgPago_0.SuspendLayout()
        Me._fraPagos_3.SuspendLayout()
        Me._fraPagos_1.SuspendLayout()
        Me._fraPagos_2.SuspendLayout()
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        CType(Me.dbcProveedor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mshPagos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mshAnticipos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mshNotasCredito, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtpCorte, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraPagos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraProgPago, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPago, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optOrigen, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Text = "Emisión de pagos"
        Me.ClientSize = New System.Drawing.Size(850, 584)
        Me.Location = New System.Drawing.Point(88, 97)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ControlBox = True
        Me.Enabled = True
        Me.MinimizeBox = True
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = True
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmCXPEmisionPagos"
        Me._fraProgPago_0.Text = "Origen del Pago"
        Me._fraProgPago_0.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me._fraProgPago_0.Size = New System.Drawing.Size(201, 57)
        Me._fraProgPago_0.Location = New System.Drawing.Point(424, 8)
        Me._fraProgPago_0.TabIndex = 2
        Me._fraProgPago_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraProgPago_0.Enabled = True
        Me._fraProgPago_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraProgPago_0.Visible = True
        Me._fraProgPago_0.Padding = New System.Windows.Forms.Padding(0)
        Me._fraProgPago_0.Name = "_fraProgPago_0"
        Me._optOrigen_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optOrigen_0.Text = "Joyería"
        Me._optOrigen_0.Size = New System.Drawing.Size(57, 17)
        Me._optOrigen_0.Location = New System.Drawing.Point(24, 24)
        Me._optOrigen_0.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me._optOrigen_0, "Origen del pago")
        Me._optOrigen_0.Checked = True
        Me._optOrigen_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optOrigen_0.BackColor = System.Drawing.SystemColors.Control
        Me._optOrigen_0.CausesValidation = True
        Me._optOrigen_0.Enabled = True
        Me._optOrigen_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optOrigen_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optOrigen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optOrigen_0.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optOrigen_0.TabStop = True
        Me._optOrigen_0.Visible = True
        Me._optOrigen_0.Name = "_optOrigen_0"
        Me._optOrigen_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optOrigen_1.Text = "Personal"
        Me._optOrigen_1.Size = New System.Drawing.Size(65, 17)
        Me._optOrigen_1.Location = New System.Drawing.Point(112, 24)
        Me._optOrigen_1.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me._optOrigen_1, "Origen del Pago")
        Me._optOrigen_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optOrigen_1.BackColor = System.Drawing.SystemColors.Control
        Me._optOrigen_1.CausesValidation = True
        Me._optOrigen_1.Enabled = True
        Me._optOrigen_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optOrigen_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optOrigen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optOrigen_1.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optOrigen_1.TabStop = True
        Me._optOrigen_1.Checked = False
        Me._optOrigen_1.Visible = True
        Me._optOrigen_1.Name = "_optOrigen_1"
        Me._fraPagos_3.Text = "Moneda del Pago"
        Me._fraPagos_3.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me._fraPagos_3.Size = New System.Drawing.Size(121, 107)
        Me._fraPagos_3.Location = New System.Drawing.Point(408, 382)
        Me._fraPagos_3.TabIndex = 14
        Me._fraPagos_3.BackColor = System.Drawing.SystemColors.Control
        Me._fraPagos_3.Enabled = True
        Me._fraPagos_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraPagos_3.Visible = True
        Me._fraPagos_3.Padding = New System.Windows.Forms.Padding(0)
        Me._fraPagos_3.Name = "_fraPagos_3"
        Me._optMoneda_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMoneda_0.Text = "Dólares"
        Me._optMoneda_0.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_0.Location = New System.Drawing.Point(24, 32)
        Me._optMoneda_0.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me._optMoneda_0, "Moneda de Compra (Dólares)")
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.CausesValidation = True
        Me._optMoneda_0.Enabled = True
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Visible = True
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMoneda_1.Text = "Pesos"
        Me._optMoneda_1.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_1.Location = New System.Drawing.Point(24, 64)
        Me._optMoneda_1.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me._optMoneda_1, "Modeda de Compra (Pesos)")
        Me._optMoneda_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.CausesValidation = True
        Me._optMoneda_1.Enabled = True
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Checked = False
        Me._optMoneda_1.Visible = True
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me.txtFlex.AutoSize = False
        Me.txtFlex.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFlex.Size = New System.Drawing.Size(81, 20)
        Me.txtFlex.Location = New System.Drawing.Point(40, 144)
        Me.txtFlex.MaxLength = 50
        Me.txtFlex.TabIndex = 9
        Me.txtFlex.Visible = False
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.CausesValidation = True
        Me.txtFlex.Enabled = True
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.HideSelection = True
        Me.txtFlex.ReadOnly = False
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.Multiline = False
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtFlex.TabStop = True
        Me.txtFlex.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtFlex.Name = "txtFlex"
        Me.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnGenerar.Text = "&Generar Pago"
        Me.btnGenerar.Size = New System.Drawing.Size(145, 27)
        Me.btnGenerar.Location = New System.Drawing.Point(696, 552)
        Me.btnGenerar.TabIndex = 34
        Me.btnGenerar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGenerar.CausesValidation = True
        Me.btnGenerar.Enabled = True
        Me.btnGenerar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGenerar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGenerar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGenerar.TabStop = True
        Me.btnGenerar.Name = "btnGenerar"
        Me._fraPagos_0.Size = New System.Drawing.Size(873, 2)
        Me._fraPagos_0.Location = New System.Drawing.Point(-11, 544)
        Me._fraPagos_0.TabIndex = 33
        Me._fraPagos_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraPagos_0.Enabled = True
        Me._fraPagos_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraPagos_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraPagos_0.Visible = True
        Me._fraPagos_0.Padding = New System.Windows.Forms.Padding(0)
        Me._fraPagos_0.Name = "_fraPagos_0"
        Me._fraPagos_1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me._fraPagos_1.Size = New System.Drawing.Size(241, 233)
        Me._fraPagos_1.Location = New System.Drawing.Point(600, 296)
        Me._fraPagos_1.TabIndex = 17
        Me._fraPagos_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraPagos_1.Enabled = True
        Me._fraPagos_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraPagos_1.Visible = True
        Me._fraPagos_1.Padding = New System.Windows.Forms.Padding(0)
        Me._fraPagos_1.Name = "_fraPagos_1"
        Me.txtTipoCambioEuro.AutoSize = False
        Me.txtTipoCambioEuro.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtTipoCambioEuro.BackColor = System.Drawing.Color.FromArgb(213, 245, 213)
        Me.txtTipoCambioEuro.Size = New System.Drawing.Size(49, 21)
        Me.txtTipoCambioEuro.Location = New System.Drawing.Point(176, 30)
        Me.txtTipoCambioEuro.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.txtTipoCambioEuro, "Tipo de Cambio (de Euros a Pesos)")
        Me.txtTipoCambioEuro.AcceptsReturn = True
        Me.txtTipoCambioEuro.CausesValidation = True
        Me.txtTipoCambioEuro.Enabled = True
        Me.txtTipoCambioEuro.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambioEuro.HideSelection = True
        Me.txtTipoCambioEuro.ReadOnly = False
        Me.txtTipoCambioEuro.MaxLength = 0
        Me.txtTipoCambioEuro.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambioEuro.Multiline = False
        Me.txtTipoCambioEuro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambioEuro.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtTipoCambioEuro.TabStop = True
        Me.txtTipoCambioEuro.Visible = True
        Me.txtTipoCambioEuro.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtTipoCambioEuro.Name = "txtTipoCambioEuro"
        Me.txtTipoCambio.AutoSize = False
        Me.txtTipoCambio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtTipoCambio.BackColor = System.Drawing.Color.FromArgb(213, 245, 213)
        Me.txtTipoCambio.Size = New System.Drawing.Size(49, 21)
        Me.txtTipoCambio.Location = New System.Drawing.Point(120, 30)
        Me.txtTipoCambio.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.txtTipoCambio, "Tipo de Cambio (de Dólares a Pesos)")
        Me.txtTipoCambio.AcceptsReturn = True
        Me.txtTipoCambio.CausesValidation = True
        Me.txtTipoCambio.Enabled = True
        Me.txtTipoCambio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambio.HideSelection = True
        Me.txtTipoCambio.ReadOnly = False
        Me.txtTipoCambio.MaxLength = 0
        Me.txtTipoCambio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambio.Multiline = False
        Me.txtTipoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambio.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtTipoCambio.TabStop = True
        Me.txtTipoCambio.Visible = True
        Me.txtTipoCambio.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtTipoCambio.Name = "txtTipoCambio"
        Me.txtAPagar.AutoSize = False
        Me.txtAPagar.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtAPagar.BackColor = System.Drawing.Color.FromArgb(201, 209, 218)
        Me.txtAPagar.Size = New System.Drawing.Size(105, 21)
        Me.txtAPagar.Location = New System.Drawing.Point(120, 200)
        Me.txtAPagar.ReadOnly = True
        Me.txtAPagar.TabIndex = 32
        Me.ToolTip1.SetToolTip(Me.txtAPagar, "Total a pagar")
        Me.txtAPagar.AcceptsReturn = True
        Me.txtAPagar.CausesValidation = True
        Me.txtAPagar.Enabled = True
        Me.txtAPagar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAPagar.HideSelection = True
        Me.txtAPagar.MaxLength = 0
        Me.txtAPagar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAPagar.Multiline = False
        Me.txtAPagar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAPagar.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtAPagar.TabStop = True
        Me.txtAPagar.Visible = True
        Me.txtAPagar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtAPagar.Name = "txtAPagar"
        Me.txtAnticipos.AutoSize = False
        Me.txtAnticipos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtAnticipos.BackColor = System.Drawing.SystemColors.Info
        Me.txtAnticipos.Size = New System.Drawing.Size(105, 21)
        Me.txtAnticipos.Location = New System.Drawing.Point(120, 166)
        Me.txtAnticipos.ReadOnly = True
        Me.txtAnticipos.TabIndex = 30
        Me.ToolTip1.SetToolTip(Me.txtAnticipos, "Importe total de los Anticipos")
        Me.txtAnticipos.AcceptsReturn = True
        Me.txtAnticipos.CausesValidation = True
        Me.txtAnticipos.Enabled = True
        Me.txtAnticipos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnticipos.HideSelection = True
        Me.txtAnticipos.MaxLength = 0
        Me.txtAnticipos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAnticipos.Multiline = False
        Me.txtAnticipos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAnticipos.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtAnticipos.TabStop = True
        Me.txtAnticipos.Visible = True
        Me.txtAnticipos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtAnticipos.Name = "txtAnticipos"
        Me.txtNotasCredito.AutoSize = False
        Me.txtNotasCredito.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtNotasCredito.BackColor = System.Drawing.SystemColors.Info
        Me.txtNotasCredito.Size = New System.Drawing.Size(105, 21)
        Me.txtNotasCredito.Location = New System.Drawing.Point(120, 132)
        Me.txtNotasCredito.ReadOnly = True
        Me.txtNotasCredito.TabIndex = 28
        Me.ToolTip1.SetToolTip(Me.txtNotasCredito, "Importe total de las notas de crédito")
        Me.txtNotasCredito.AcceptsReturn = True
        Me.txtNotasCredito.CausesValidation = True
        Me.txtNotasCredito.Enabled = True
        Me.txtNotasCredito.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNotasCredito.HideSelection = True
        Me.txtNotasCredito.MaxLength = 0
        Me.txtNotasCredito.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNotasCredito.Multiline = False
        Me.txtNotasCredito.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNotasCredito.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtNotasCredito.TabStop = True
        Me.txtNotasCredito.Visible = True
        Me.txtNotasCredito.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtNotasCredito.Name = "txtNotasCredito"
        Me.txtDescuentoFinanciero.AutoSize = False
        Me.txtDescuentoFinanciero.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDescuentoFinanciero.BackColor = System.Drawing.SystemColors.Info
        Me.txtDescuentoFinanciero.Size = New System.Drawing.Size(105, 21)
        Me.txtDescuentoFinanciero.Location = New System.Drawing.Point(120, 98)
        Me.txtDescuentoFinanciero.ReadOnly = True
        Me.txtDescuentoFinanciero.TabIndex = 26
        Me.ToolTip1.SetToolTip(Me.txtDescuentoFinanciero, "Descuento Financiero")
        Me.txtDescuentoFinanciero.AcceptsReturn = True
        Me.txtDescuentoFinanciero.CausesValidation = True
        Me.txtDescuentoFinanciero.Enabled = True
        Me.txtDescuentoFinanciero.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescuentoFinanciero.HideSelection = True
        Me.txtDescuentoFinanciero.MaxLength = 0
        Me.txtDescuentoFinanciero.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescuentoFinanciero.Multiline = False
        Me.txtDescuentoFinanciero.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescuentoFinanciero.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtDescuentoFinanciero.TabStop = True
        Me.txtDescuentoFinanciero.Visible = True
        Me.txtDescuentoFinanciero.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDescuentoFinanciero.Name = "txtDescuentoFinanciero"
        Me.txtFacturas.AutoSize = False
        Me.txtFacturas.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtFacturas.BackColor = System.Drawing.SystemColors.Info
        Me.txtFacturas.Size = New System.Drawing.Size(105, 21)
        Me.txtFacturas.Location = New System.Drawing.Point(120, 64)
        Me.txtFacturas.ReadOnly = True
        Me.txtFacturas.TabIndex = 24
        Me.ToolTip1.SetToolTip(Me.txtFacturas, "Importe total de las facturas")
        Me.txtFacturas.AcceptsReturn = True
        Me.txtFacturas.CausesValidation = True
        Me.txtFacturas.Enabled = True
        Me.txtFacturas.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFacturas.HideSelection = True
        Me.txtFacturas.MaxLength = 0
        Me.txtFacturas.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFacturas.Multiline = False
        Me.txtFacturas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFacturas.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtFacturas.TabStop = True
        Me.txtFacturas.Visible = True
        Me.txtFacturas.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtFacturas.Name = "txtFacturas"
        Me._lblPago_8.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._lblPago_8.Text = "Tipo de Cambio"
        Me._lblPago_8.Size = New System.Drawing.Size(96, 13)
        Me._lblPago_8.Location = New System.Drawing.Point(16, 34)
        Me._lblPago_8.TabIndex = 20
        Me._lblPago_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_8.Enabled = True
        Me._lblPago_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblPago_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_8.UseMnemonic = True
        Me._lblPago_8.Visible = True
        Me._lblPago_8.AutoSize = False
        Me._lblPago_8.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_8.Name = "_lblPago_8"
        Me.lblEuro.Text = "Euro"
        Me.lblEuro.Size = New System.Drawing.Size(30, 13)
        Me.lblEuro.Location = New System.Drawing.Point(176, 12)
        Me.lblEuro.TabIndex = 19
        Me.lblEuro.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblEuro.BackColor = System.Drawing.SystemColors.Control
        Me.lblEuro.Enabled = True
        Me.lblEuro.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEuro.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEuro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEuro.UseMnemonic = True
        Me.lblEuro.Visible = True
        Me.lblEuro.AutoSize = True
        Me.lblEuro.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lblEuro.Name = "lblEuro"
        Me.lblDolar.Text = "Dólar"
        Me.lblDolar.Size = New System.Drawing.Size(25, 13)
        Me.lblDolar.Location = New System.Drawing.Point(120, 12)
        Me.lblDolar.TabIndex = 18
        Me.lblDolar.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblDolar.BackColor = System.Drawing.SystemColors.Control
        Me.lblDolar.Enabled = True
        Me.lblDolar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDolar.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDolar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDolar.UseMnemonic = True
        Me.lblDolar.Visible = True
        Me.lblDolar.AutoSize = True
        Me.lblDolar.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lblDolar.Name = "lblDolar"
        Me._lblPago_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._lblPago_7.Text = "A pagar"
        Me._lblPago_7.Size = New System.Drawing.Size(96, 13)
        Me._lblPago_7.Location = New System.Drawing.Point(16, 204)
        Me._lblPago_7.TabIndex = 31
        Me._lblPago_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_7.Enabled = True
        Me._lblPago_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblPago_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_7.UseMnemonic = True
        Me._lblPago_7.Visible = True
        Me._lblPago_7.AutoSize = False
        Me._lblPago_7.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_7.Name = "_lblPago_7"
        Me._lblPago_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._lblPago_6.Text = "Anticipos"
        Me._lblPago_6.Size = New System.Drawing.Size(96, 13)
        Me._lblPago_6.Location = New System.Drawing.Point(16, 170)
        Me._lblPago_6.TabIndex = 29
        Me._lblPago_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_6.Enabled = True
        Me._lblPago_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblPago_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_6.UseMnemonic = True
        Me._lblPago_6.Visible = True
        Me._lblPago_6.AutoSize = False
        Me._lblPago_6.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_6.Name = "_lblPago_6"
        Me._lblPago_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._lblPago_5.Text = "Notas de Crédito"
        Me._lblPago_5.Size = New System.Drawing.Size(96, 13)
        Me._lblPago_5.Location = New System.Drawing.Point(16, 136)
        Me._lblPago_5.TabIndex = 27
        Me._lblPago_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_5.Enabled = True
        Me._lblPago_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblPago_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_5.UseMnemonic = True
        Me._lblPago_5.Visible = True
        Me._lblPago_5.AutoSize = False
        Me._lblPago_5.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_5.Name = "_lblPago_5"
        Me._lblPago_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._lblPago_4.Text = "Descto. Financiero"
        Me._lblPago_4.Size = New System.Drawing.Size(96, 13)
        Me._lblPago_4.Location = New System.Drawing.Point(16, 102)
        Me._lblPago_4.TabIndex = 25
        Me._lblPago_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_4.Enabled = True
        Me._lblPago_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblPago_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_4.UseMnemonic = True
        Me._lblPago_4.Visible = True
        Me._lblPago_4.AutoSize = False
        Me._lblPago_4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_4.Name = "_lblPago_4"
        Me._lblPago_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._lblPago_3.Text = "Pagos de Facturas"
        Me._lblPago_3.Size = New System.Drawing.Size(96, 13)
        Me._lblPago_3.Location = New System.Drawing.Point(16, 68)
        Me._lblPago_3.TabIndex = 23
        Me._lblPago_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_3.Enabled = True
        Me._lblPago_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblPago_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_3.UseMnemonic = True
        Me._lblPago_3.Visible = True
        Me._lblPago_3.AutoSize = False
        Me._lblPago_3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_3.Name = "_lblPago_3"
        'dbcProveedor.OcxState = CType(resources.GetObject("dbcProveedor.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dbcProveedor.Size = New System.Drawing.Size(265, 21)
        Me.dbcProveedor.Location = New System.Drawing.Point(152, 16)
        Me.dbcProveedor.TabIndex = 1
        Me.dbcProveedor.Name = "dbcProveedor"
        mshPagos.OcxState = CType(resources.GetObject("mshPagos.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mshPagos.Size = New System.Drawing.Size(834, 137)
        Me.mshPagos.Location = New System.Drawing.Point(8, 72)
        Me.mshPagos.TabIndex = 8
        Me.mshPagos.Name = "mshPagos"
        mshAnticipos.OcxState = CType(resources.GetObject("mshAnticipos.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mshAnticipos.Size = New System.Drawing.Size(353, 137)
        Me.mshAnticipos.Location = New System.Drawing.Point(8, 392)
        Me.mshAnticipos.TabIndex = 13
        Me.mshAnticipos.Name = "mshAnticipos"
        mshNotasCredito.OcxState = CType(resources.GetObject("mshNotasCredito.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mshNotasCredito.Size = New System.Drawing.Size(521, 137)
        Me.mshNotasCredito.Location = New System.Drawing.Point(8, 232)
        Me.mshNotasCredito.TabIndex = 11
        Me.mshNotasCredito.Name = "mshNotasCredito"
        Me._fraPagos_2.Enabled = False
        Me._fraPagos_2.Size = New System.Drawing.Size(193, 50)
        Me._fraPagos_2.Location = New System.Drawing.Point(632, 8)
        Me._fraPagos_2.TabIndex = 35
        Me._fraPagos_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraPagos_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraPagos_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraPagos_2.Visible = True
        Me._fraPagos_2.Padding = New System.Windows.Forms.Padding(0)
        Me._fraPagos_2.Name = "_fraPagos_2"
        'dtpCorte.OcxState = CType(resources.GetObject("dtpCorte.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dtpCorte.Size = New System.Drawing.Size(105, 21)
        Me.dtpCorte.Location = New System.Drawing.Point(80, 16)
        Me.dtpCorte.TabIndex = 37
        Me.dtpCorte.Name = "dtpCorte"
        Me._lblPago_0.Text = "Fecha"
        Me._lblPago_0.Size = New System.Drawing.Size(30, 13)
        Me._lblPago_0.Location = New System.Drawing.Point(32, 20)
        Me._lblPago_0.TabIndex = 36
        Me._lblPago_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblPago_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_0.Enabled = True
        Me._lblPago_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblPago_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_0.UseMnemonic = True
        Me._lblPago_0.Visible = True
        Me._lblPago_0.AutoSize = True
        Me._lblPago_0.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_0.Name = "_lblPago_0"
        Me.lblCR.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblCR.BackColor = System.Drawing.Color.FromArgb(255, 200, 145)
        Me.lblCR.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.lblCR.Size = New System.Drawing.Size(19, 19)
        Me.lblCR.Location = New System.Drawing.Point(152, 48)
        Me.lblCR.TabIndex = 6
        Me.lblCR.Enabled = True
        Me.lblCR.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCR.UseMnemonic = True
        Me.lblCR.Visible = True
        Me.lblCR.AutoSize = False
        Me.lblCR.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCR.Name = "lblCR"
        Me._lblPago_9.Text = "Registros a tomar en cuenta al efectuar el Pago"
        Me._lblPago_9.Size = New System.Drawing.Size(225, 13)
        Me._lblPago_9.Location = New System.Drawing.Point(176, 51)
        Me._lblPago_9.TabIndex = 7
        Me._lblPago_9.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblPago_9.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_9.Enabled = True
        Me._lblPago_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblPago_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_9.UseMnemonic = True
        Me._lblPago_9.Visible = True
        Me._lblPago_9.AutoSize = True
        Me._lblPago_9.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_9.Name = "_lblPago_9"
        Me._lblPago_10.Text = "Pagos ..."
        Me._lblPago_10.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me._lblPago_10.Size = New System.Drawing.Size(52, 13)
        Me._lblPago_10.Location = New System.Drawing.Point(16, 50)
        Me._lblPago_10.TabIndex = 5
        Me._lblPago_10.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblPago_10.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_10.Enabled = True
        Me._lblPago_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_10.UseMnemonic = True
        Me._lblPago_10.Visible = True
        Me._lblPago_10.AutoSize = True
        Me._lblPago_10.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_10.Name = "_lblPago_10"
        Me._lblPago_2.Text = "Notas de crédito ..."
        Me._lblPago_2.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me._lblPago_2.Size = New System.Drawing.Size(111, 13)
        Me._lblPago_2.Location = New System.Drawing.Point(16, 216)
        Me._lblPago_2.TabIndex = 10
        Me._lblPago_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblPago_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_2.Enabled = True
        Me._lblPago_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_2.UseMnemonic = True
        Me._lblPago_2.Visible = True
        Me._lblPago_2.AutoSize = True
        Me._lblPago_2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_2.Name = "_lblPago_2"
        Me._lblPago_1.Text = "Anticipos ..."
        Me._lblPago_1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me._lblPago_1.Size = New System.Drawing.Size(69, 13)
        Me._lblPago_1.Location = New System.Drawing.Point(16, 376)
        Me._lblPago_1.TabIndex = 12
        Me._lblPago_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblPago_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblPago_1.Enabled = True
        Me._lblPago_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPago_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPago_1.UseMnemonic = True
        Me._lblPago_1.Visible = True
        Me._lblPago_1.AutoSize = True
        Me._lblPago_1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblPago_1.Name = "_lblPago_1"
        Me.lblProveedor.Text = "Proveedor / Acreedor"
        Me.lblProveedor.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.lblProveedor.Size = New System.Drawing.Size(124, 13)
        Me.lblProveedor.Location = New System.Drawing.Point(16, 20)
        Me.lblProveedor.TabIndex = 0
        Me.lblProveedor.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblProveedor.BackColor = System.Drawing.SystemColors.Control
        Me.lblProveedor.Enabled = True
        Me.lblProveedor.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblProveedor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblProveedor.UseMnemonic = True
        Me.lblProveedor.Visible = True
        Me.lblProveedor.AutoSize = True
        Me.lblProveedor.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lblProveedor.Name = "lblProveedor"
        Me.fraPagos.SetIndex(_fraPagos_3, CType(3, Short))
        Me.fraPagos.SetIndex(_fraPagos_0, CType(0, Short))
        Me.fraPagos.SetIndex(_fraPagos_1, CType(1, Short))
        Me.fraPagos.SetIndex(_fraPagos_2, CType(2, Short))
        Me.fraProgPago.SetIndex(_fraProgPago_0, CType(0, Short))
        Me.lblPago.SetIndex(_lblPago_8, CType(8, Short))
        Me.lblPago.SetIndex(_lblPago_7, CType(7, Short))
        Me.lblPago.SetIndex(_lblPago_6, CType(6, Short))
        Me.lblPago.SetIndex(_lblPago_5, CType(5, Short))
        Me.lblPago.SetIndex(_lblPago_4, CType(4, Short))
        Me.lblPago.SetIndex(_lblPago_3, CType(3, Short))
        Me.lblPago.SetIndex(_lblPago_0, CType(0, Short))
        Me.lblPago.SetIndex(_lblPago_9, CType(9, Short))
        Me.lblPago.SetIndex(_lblPago_10, CType(10, Short))
        Me.lblPago.SetIndex(_lblPago_2, CType(2, Short))
        Me.lblPago.SetIndex(_lblPago_1, CType(1, Short))
        Me.optMoneda.SetIndex(_optMoneda_0, CType(0, Short))
        Me.optMoneda.SetIndex(_optMoneda_1, CType(1, Short))
        Me.optOrigen.SetIndex(_optOrigen_0, CType(0, Short))
        Me.optOrigen.SetIndex(_optOrigen_1, CType(1, Short))
        CType(Me.optOrigen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPago, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraProgPago, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraPagos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtpCorte, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mshNotasCredito, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mshAnticipos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mshPagos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dbcProveedor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Controls.Add(_fraProgPago_0)
        Me.Controls.Add(_fraPagos_3)
        Me.Controls.Add(txtFlex)
        Me.Controls.Add(btnGenerar)
        Me.Controls.Add(_fraPagos_0)
        Me.Controls.Add(_fraPagos_1)
        Me.Controls.Add(dbcProveedor)
        Me.Controls.Add(mshPagos)
        Me.Controls.Add(mshAnticipos)
        Me.Controls.Add(mshNotasCredito)
        Me.Controls.Add(_fraPagos_2)
        Me.Controls.Add(lblCR)
        Me.Controls.Add(_lblPago_9)
        Me.Controls.Add(_lblPago_10)
        Me.Controls.Add(_lblPago_2)
        Me.Controls.Add(_lblPago_1)
        Me.Controls.Add(lblProveedor)
        Me._fraProgPago_0.Controls.Add(_optOrigen_0)
        Me._fraProgPago_0.Controls.Add(_optOrigen_1)
        Me._fraPagos_3.Controls.Add(_optMoneda_0)
        Me._fraPagos_3.Controls.Add(_optMoneda_1)
        Me._fraPagos_1.Controls.Add(txtTipoCambioEuro)
        Me._fraPagos_1.Controls.Add(txtTipoCambio)
        Me._fraPagos_1.Controls.Add(txtAPagar)
        Me._fraPagos_1.Controls.Add(txtAnticipos)
        Me._fraPagos_1.Controls.Add(txtNotasCredito)
        Me._fraPagos_1.Controls.Add(txtDescuentoFinanciero)
        Me._fraPagos_1.Controls.Add(txtFacturas)
        Me._fraPagos_1.Controls.Add(_lblPago_8)
        Me._fraPagos_1.Controls.Add(lblEuro)
        Me._fraPagos_1.Controls.Add(lblDolar)
        Me._fraPagos_1.Controls.Add(_lblPago_7)
        Me._fraPagos_1.Controls.Add(_lblPago_6)
        Me._fraPagos_1.Controls.Add(_lblPago_5)
        Me._fraPagos_1.Controls.Add(_lblPago_4)
        Me._fraPagos_1.Controls.Add(_lblPago_3)
        Me._fraPagos_2.Controls.Add(dtpCorte)
        Me._fraPagos_2.Controls.Add(_lblPago_0)
        Me._fraProgPago_0.ResumeLayout(False)
        Me._fraPagos_3.ResumeLayout(False)
        Me._fraPagos_1.ResumeLayout(False)
        Me._fraPagos_2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub


    Const C_PAGAR As String = "P"

    Const P_RENENCABEZADO As Short = 0
    Const S_RENENCABEZADO As Short = 1

    Const P_COLFACTURA As Short = 0
    Const P_COLPAGO As Short = 1
    Const P_COLFECHAFACTURA As Short = 2
    Const P_COLFECHAVENCTO As Short = 3
    Const P_COLFECHAPAGO As Short = 4
    Const P_COLIMPORTE As Short = 5
    Const P_COLPAGOS As Short = 6
    Const P_COLSALDO As Short = 7
    Const P_COLIMPORTEPAGO As Short = 8
    Const P_COLMONEDA As Short = 9
    Const P_COLDESCTOPORC As Short = 10
    Const P_COLDESCTOFIN As Short = 11
    Const P_COLSUBTOTALDF As Short = 12
    Const P_COLIVADF As Short = 13
    Const P_COLAPAGAR As Short = 14
    Const P_COLNUMPARTIDA As Short = 15
    Const P_COLESTATUS As Short = 16
    Const P_COLDESCTOFINPORC As Short = 17

    Const A_COLFOLIO As Short = 0
    Const A_COLFECHA As Short = 1
    Const A_COLIMPORTE As Short = 2
    Const A_COLESTATUS As Short = 3
    Const A_COLMONEDA As Short = 4

    Const N_COLFOLIO As Short = 0
    Const N_COLFECHA As Short = 1
    Const N_COLFACTURA As Short = 2
    Const N_COLTOTAL As Short = 3
    Const N_COLTIPO As Short = 4
    Const N_COLESTATUS As Short = 5
    Const N_COLMONEDA As Short = 6

    Dim cMonedaPago As String

    Dim mblnFueraChange As Boolean
    Dim mintCodProveedor As Integer
    Dim tecla As Short

    Dim mblnLoad As Boolean 'En el Load del formulario sí debe inicializar todos los grid

    Dim mblnSalir As Boolean

    Public Sub ActualizaCantidades()
        On Error GoTo Merr

        Dim nFacturas As Decimal
        Dim nDescuentoFinanciero As Decimal
        Dim nNotasCredito As Decimal
        Dim nAnticipos As Decimal
        Dim nAPagar As Decimal

        Dim nImporteConv As Decimal

        Dim nColImporte As Decimal
        Dim nColPagos As Decimal
        Dim nColSaldo As Decimal
        Dim nColImportePago As Decimal
        Dim nColDesctoFin As Decimal
        Dim nColDesctoPorc As Decimal
        Dim nColSubTotal As Decimal
        Dim nColIva As Decimal
        Dim nColAPagar As Decimal

        Dim cMonedaOrigen As String
        Dim cMonedaDestino As String
        Dim nTipoCambio As String
        Dim nTipoCambioE As String

        Dim I As Integer

        nFacturas = 0
        nDescuentoFinanciero = 0
        nNotasCredito = 0
        nAnticipos = 0
        nAPagar = 0
        nColImporte = 0
        nColSubTotal = 0
        nColIva = 0
        nColAPagar = 0

        cMonedaDestino = IIf(Me.optMoneda(0).Checked, C_DOLAR, C_PESO)
        nTipoCambio = CStr(CDec(Numerico((Me.txtTipoCambio.Text))))
        nTipoCambioE = CStr(CDec(Numerico((Me.txtTipoCambioEuro.Text))))

        If CDbl(nTipoCambio) = 0 Or CDbl(nTipoCambioE) = 0 Then
            Exit Sub
        End If

        With Me.mshPagos
            For I = 2 To .Rows - 2
                If Trim(.get_TextMatrix(I, P_COLPAGO)) = "" Then
                    Exit For
                End If
                'Convertir cantidades a la moneda indicada
                cMonedaOrigen = VB.Left(Trim(.get_TextMatrix(I, P_COLMONEDA)), 1)

                nColImporte = CDec(Numerico(.get_TextMatrix(I, P_COLIMPORTE)))
                nColPagos = CDec(Numerico(.get_TextMatrix(I, P_COLPAGOS)))
                nColSaldo = CDec(Numerico(.get_TextMatrix(I, P_COLSALDO)))
                nColImportePago = CDec(Numerico(.get_TextMatrix(I, P_COLIMPORTEPAGO)))
                nColDesctoFin = CDec(Numerico(.get_TextMatrix(I, P_COLDESCTOFIN)))
                If nColImportePago <> 0 Then
                    nColDesctoPorc = (nColDesctoFin * 100) / nColImportePago
                    nColIva = System.Math.Round(nColDesctoFin * (gcurCorpoTASAIVA / 100), 4)
                    nColSubTotal = System.Math.Round(nColDesctoFin - nColIva, 4)
                Else
                    nColDesctoPorc = 0
                    nColIva = 0
                    nColSubTotal = 0
                End If
                nColAPagar = nColImportePago - nColDesctoFin

                If Trim(.get_TextMatrix(I, P_COLESTATUS)) = C_PAGAR Then
                    'Convertir todo a pesos
                    If cMonedaOrigen = C_DOLAR Then
                        nImporteConv = nColImportePago * CDbl(nTipoCambio)
                        nDescuentoFinanciero = nDescuentoFinanciero + (nColDesctoFin * CDbl(nTipoCambio))
                    ElseIf cMonedaOrigen = C_PESO Then
                        nImporteConv = nColImportePago
                        nDescuentoFinanciero = nDescuentoFinanciero + nColDesctoFin
                    Else 'C_EURO
                        nImporteConv = nColImportePago * CDbl(nTipoCambioE)
                        nDescuentoFinanciero = nDescuentoFinanciero + (nColDesctoFin * CDbl(nTipoCambioE))
                    End If
                    nFacturas = nFacturas + nImporteConv
                End If
                .set_TextMatrix(I, P_COLDESCTOPORC, nColDesctoPorc)
                .set_TextMatrix(I, P_COLSUBTOTALDF, nColSubTotal)
                .set_TextMatrix(I, P_COLIVADF, nColIva)
                .set_TextMatrix(I, P_COLAPAGAR, VB6.Format(nColAPagar, "###,###,##0.00"))
            Next I
        End With

        '        If cMonedaOrigen <> cMonedaDestino Then
        '            Select Case cMonedaOrigen
        '                Case C_PESO 'Cambia el importe de Pesos - Dólares
        '                    If nTipoCambio = 0 Then
        '                        nDescuentoFinanciero = 0
        '                    Else
        '                        nDescuentoFinanciero = nDescuentoFinanciero / nTipoCambio
        '                        nFacturas = nFacturas / nTipoCambio
        '                    End If
        '                Case C_DOLAR 'Cambia el importe de Dólares - Pesos
        '                    nDescuentoFinanciero = nDescuentoFinanciero * nTipoCambio
        '                    nFacturas = nFacturas * nTipoCambio
        '                Case Else 'Cambia el importe a euros
        '                    If cMonedaDestino = C_PESO Then 'De Pesos - Euros
        '                    ElseIf cMonedaDestino = C_DOLAR Then 'De Dólares - Euros
        '                    End If
        '            End Select
        '        End If


        With Me.mshNotasCredito
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, N_COLFOLIO)) = "" Then
                    Exit For
                End If
                'Convertir cantidades a la moneda indicada
                cMonedaOrigen = VB.Left(Trim(.get_TextMatrix(I, N_COLMONEDA)), 1)

                If Trim(.get_TextMatrix(I, N_COLESTATUS)) = C_PAGAR Then
                    'Convertir todo a pesos
                    If cMonedaOrigen = C_DOLAR Then
                        nImporteConv = CDec(Numerico(.get_TextMatrix(I, N_COLTOTAL))) * CDbl(nTipoCambio)
                    ElseIf cMonedaOrigen = C_PESO Then
                        nImporteConv = CDec(Numerico(.get_TextMatrix(I, N_COLTOTAL)))
                    Else 'C_EURO
                        nImporteConv = CDec(Numerico(.get_TextMatrix(I, N_COLTOTAL))) * CDbl(nTipoCambioE)
                    End If
                    nNotasCredito = nNotasCredito + nImporteConv
                End If
            Next I
        End With
        With Me.mshAnticipos
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, A_COLFOLIO)) = "" Then
                    Exit For
                End If
                'Convertir cantidades a la moneda indicada
                cMonedaOrigen = VB.Left(Trim(.get_TextMatrix(I, A_COLMONEDA)), 1)

                If Trim(.get_TextMatrix(I, A_COLESTATUS)) = C_PAGAR Then
                    'Convertir todo a pesos
                    If cMonedaOrigen = C_DOLAR Then
                        nImporteConv = CDec(Numerico(.get_TextMatrix(I, A_COLIMPORTE))) * CDbl(nTipoCambio)
                    ElseIf cMonedaOrigen = C_PESO Then
                        nImporteConv = CDec(Numerico(.get_TextMatrix(I, A_COLIMPORTE)))
                    Else 'C_EURO
                        nImporteConv = CDec(Numerico(.get_TextMatrix(I, A_COLIMPORTE))) * CDbl(nTipoCambioE)
                    End If
                    nAnticipos = nAnticipos + nImporteConv
                End If
            Next I
        End With

        'Todo está en pesos, ahora necesito convertir todo a la moneda indicada

        If cMonedaDestino = C_DOLAR Then
            If CDbl(nTipoCambio) > 0 Then
                nFacturas = nFacturas / CDbl(nTipoCambio)
                nDescuentoFinanciero = nDescuentoFinanciero / CDbl(nTipoCambio)
                nNotasCredito = nNotasCredito / CDbl(nTipoCambio)
                nAnticipos = nAnticipos / CDbl(nTipoCambio)
            Else
                nFacturas = 0
                nDescuentoFinanciero = 0
                nNotasCredito = 0
                nAnticipos = 0
            End If
        End If

        nAPagar = nFacturas - (nDescuentoFinanciero + nNotasCredito + nAnticipos)

        Me.txtFacturas.Text = VB6.Format(nFacturas, "###,###,##0.00")
        Me.txtDescuentoFinanciero.Text = VB6.Format(nDescuentoFinanciero, "###,###,##0.00")
        Me.txtNotasCredito.Text = VB6.Format(nNotasCredito, "###,###,##0.00")
        Me.txtAnticipos.Text = VB6.Format(nAnticipos, "###,###,##0.00")
        Me.txtAPagar.Text = VB6.Format(nAPagar, "###,###,##0.00")

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub PonerColorNotas()
        Dim I As Integer
        Dim J As Integer
        Dim Ctl As System.Windows.Forms.Control
        Dim nCol As Integer
        With Me.mshNotasCredito
            nCol = .Col
            Select Case Trim(.get_TextMatrix(.Row, N_COLESTATUS))
                Case C_PAGAR
                    Ctl = lblCR
                Case ""
                    Ctl = mshNotasCredito
            End Select
            For J = 0 To 6
                .Col = J
                .CellBackColor = BackColor
            Next
            .Col = nCol
        End With
    End Sub

    Sub PonerColorAnticipos()
        Dim I As Integer
        Dim J As Integer
        Dim Ctl As System.Windows.Forms.Control
        Dim nCol As Integer
        With Me.mshAnticipos
            nCol = .Col
            Select Case Trim(.get_TextMatrix(.Row, A_COLESTATUS))
                Case C_PAGAR
                    Ctl = lblCR
                Case ""
                    Ctl = mshAnticipos
            End Select
            For J = 0 To 4
                .Col = J
                .CellBackColor = BackColor
            Next
            .Col = nCol
        End With
    End Sub

    Sub PonerColor()
        Dim I As Integer
        Dim J As Integer
        Dim Ctl As System.Windows.Forms.Control
        Dim nCol As Integer
        With Me.mshPagos
            nCol = .Col
            Select Case Trim(.get_TextMatrix(.Row, P_COLESTATUS))
                Case C_PAGAR
                    Ctl = lblCR
                Case ""
                    Ctl = mshPagos
            End Select
            For J = 0 To 14
                .Col = J
                'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.BackColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .CellBackColor = BackColor
            Next
            .Col = nCol
        End With
    End Sub

    Public Sub EncabezadoPagos()
        On Error GoTo Merr
        Dim LnContador As Short
        Dim I As Integer

        With Me.mshPagos
            If Not mblnLoad Then
                .Rows = 3
                .Rows = 12
                .set_Cols(0, 18)
                .RemoveItem((2))
                Exit Sub
            End If

            .set_Cols(0, 18)
            .Clear()

            .set_ColWidth(P_COLFACTURA, 0, 1400)
            .set_ColWidth(P_COLPAGO, 0, 1515)
            .set_ColWidth(P_COLFECHAFACTURA, 0, 1035)
            .set_ColWidth(P_COLFECHAVENCTO, 0, 1035)
            .set_ColWidth(P_COLFECHAPAGO, 0, 1035)
            .set_ColWidth(P_COLIMPORTE, 0, 900)
            .set_ColWidth(P_COLPAGOS, 0, 900)
            .set_ColWidth(P_COLSALDO, 0, 900)
            .set_ColWidth(P_COLIMPORTEPAGO, 0, 950)
            .set_ColWidth(P_COLMONEDA, 0, 700)
            .set_ColWidth(P_COLDESCTOPORC, 0, 0)
            .set_ColWidth(P_COLDESCTOFIN, 0, 900)
            .set_ColWidth(P_COLSUBTOTALDF, 0, 0)
            .set_ColWidth(P_COLIVADF, 0, 0)
            .set_ColWidth(P_COLAPAGAR, 0, 900)
            .set_ColWidth(P_COLNUMPARTIDA, 0, 0)
            .set_ColWidth(P_COLESTATUS, 0, 0)
            .set_ColWidth(P_COLDESCTOFINPORC, 0, 0)

            .set_TextMatrix(P_RENENCABEZADO, P_COLFACTURA, "")
            .set_TextMatrix(P_RENENCABEZADO, P_COLPAGO, "")
            .set_TextMatrix(P_RENENCABEZADO, P_COLFECHAFACTURA, "Fechas")
            .set_TextMatrix(P_RENENCABEZADO, P_COLFECHAVENCTO, "Fechas")
            .set_TextMatrix(P_RENENCABEZADO, P_COLFECHAPAGO, "Fechas")
            .set_TextMatrix(P_RENENCABEZADO, P_COLIMPORTE, "Importe Real")
            .set_TextMatrix(P_RENENCABEZADO, P_COLPAGOS, "Importe Real")
            .set_TextMatrix(P_RENENCABEZADO, P_COLSALDO, "Importe Real")

            .Row = P_RENENCABEZADO
            For LnContador = 0 To P_COLSALDO
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterBottom
                .CellFontBold = False
            Next LnContador

            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)

            .MergeCells = MSHierarchicalFlexGridLib.MergeCellsSettings.flexMergeFree
            .set_MergeRow(0, True)

            .set_TextMatrix(S_RENENCABEZADO, P_COLFACTURA, "Factura")
            .set_TextMatrix(S_RENENCABEZADO, P_COLPAGO, "Pago")
            .set_TextMatrix(S_RENENCABEZADO, P_COLFECHAFACTURA, "Factura")
            .set_TextMatrix(S_RENENCABEZADO, P_COLFECHAVENCTO, "Vencto.")
            .set_TextMatrix(S_RENENCABEZADO, P_COLFECHAPAGO, "Pago")
            .set_TextMatrix(S_RENENCABEZADO, P_COLIMPORTE, "Importe")
            .set_TextMatrix(S_RENENCABEZADO, P_COLPAGOS, "Pagos")
            .set_TextMatrix(S_RENENCABEZADO, P_COLSALDO, "Saldo")
            .set_TextMatrix(S_RENENCABEZADO, P_COLIMPORTEPAGO, "Impte. Pago")
            .set_TextMatrix(S_RENENCABEZADO, P_COLMONEDA, "Moneda")
            .set_TextMatrix(S_RENENCABEZADO, P_COLDESCTOFIN, "Descto.")
            .set_TextMatrix(S_RENENCABEZADO, P_COLAPAGAR, "A pagar...")

            'Colocar los textos de los encabezados centrados
            .Row = S_RENENCABEZADO
            For LnContador = 0 To P_COLESTATUS
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterBottom
                .CellFontBold = False
            Next LnContador

            .Rows = 11
            .Col = 0
            .Row = 2
            .TopRow = 2
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub EncabezadoAnticipos()
        On Error GoTo Merr
        Dim LnContador As Short
        Dim I As Integer

        With Me.mshAnticipos
            If Not mblnLoad Then
                .Rows = 2
                .Rows = 12
                .set_Cols(0, 5)
                .RemoveItem((1))
                Exit Sub
            End If
            .set_Cols(0, 5)
            .Clear()

            .set_ColWidth(A_COLFOLIO, 0, 1500)
            .set_ColWidth(A_COLFECHA, 0, 1250)
            .set_ColWidth(A_COLIMPORTE, 0, 1250)
            .set_ColWidth(A_COLESTATUS, 0, 0)
            .set_ColWidth(A_COLMONEDA, 0, 960)

            .set_TextMatrix(P_RENENCABEZADO, A_COLFOLIO, "Folio")
            .set_TextMatrix(P_RENENCABEZADO, A_COLFECHA, "Fecha")
            .set_TextMatrix(P_RENENCABEZADO, A_COLIMPORTE, "Importe")
            .set_TextMatrix(P_RENENCABEZADO, A_COLMONEDA, "Moneda")

            .Row = P_RENENCABEZADO
            For LnContador = 0 To A_COLMONEDA
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterBottom
                .CellFontBold = False
            Next LnContador

            .Rows = 11
            .Col = 0
            .Row = 1
            .TopRow = 1
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub EncabezadoNotas()
        On Error GoTo Merr
        Dim LnContador As Short
        Dim I As Integer

        With Me.mshNotasCredito
            If Not mblnLoad Then
                .Rows = 2
                .Rows = 12
                .set_Cols(0, 7)
                .RemoveItem((1))
                Exit Sub
            End If
            .set_Cols(0, 7)
            .Clear()

            .set_ColWidth(N_COLFOLIO, 0, 1900)
            .set_ColWidth(N_COLFECHA, 0, 1170)
            .set_ColWidth(N_COLFACTURA, 0, 1500)
            .set_ColWidth(N_COLTOTAL, 0, 1250)
            .set_ColWidth(N_COLTIPO, 0, 720)
            .set_ColWidth(N_COLESTATUS, 0, 0)
            .set_ColWidth(N_COLMONEDA, 0, 960)

            .set_TextMatrix(P_RENENCABEZADO, N_COLFOLIO, "Folio")
            .set_TextMatrix(P_RENENCABEZADO, N_COLFECHA, "Fecha")
            .set_TextMatrix(P_RENENCABEZADO, N_COLFACTURA, "Factura")
            .set_TextMatrix(P_RENENCABEZADO, N_COLTOTAL, "Total")
            .set_TextMatrix(P_RENENCABEZADO, N_COLTIPO, "Tipo")
            .set_TextMatrix(P_RENENCABEZADO, N_COLMONEDA, "Moneda")

            .Row = P_RENENCABEZADO
            For LnContador = 0 To N_COLMONEDA
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterBottom
                .CellFontBold = False
            Next LnContador

            .Rows = 11
            .Col = 0
            .Row = 1
            .TopRow = 1
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function ValidaDatos() As Boolean
        Select Case True
            Case CDec(Numerico((Me.txtAPagar.Text))) = 0
                MsgBox("El importe del pago no debe quedar en ceros", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Case CDec(Numerico((Me.txtAPagar.Text))) < 0
                MsgBox("El importe del pago no debe ser menor o igual a cero", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Public Sub Limpiar()
        Nuevo()
        mblnFueraChange = True
        'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Me.dbcProveedor.Text = ""
        Me.dbcProveedor.Tag = ""
        mintCodProveedor = 0
        mblnFueraChange = False
        If Not frmBancosProcesoDiarioRegistrodePagos.blnEmisionPagos Then
            If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Me.dbcProveedor.Focus()
        End If
    End Sub

    Public Sub Nuevo()
        On Error GoTo Merr

        If mblnLoad Then
            Call Me.EncabezadoPagos()
            Call Me.EncabezadoNotas()
            Call Me.EncabezadoAnticipos()
            mblnLoad = False
        Else
            If Trim(Me.mshPagos.get_TextMatrix(2, P_COLPAGO)) <> "" Then
                Call Me.EncabezadoPagos()
            End If
            If Trim(Me.mshNotasCredito.get_TextMatrix(1, N_COLFOLIO)) <> "" Then
                Call Me.EncabezadoNotas()
            End If
            If Trim(Me.mshAnticipos.get_TextMatrix(1, A_COLFOLIO)) <> "" Then
                Call Me.EncabezadoAnticipos()
            End If
        End If

        Me.txtTipoCambio.Text = VB6.Format(gcurCorpoTIPOCAMBIODOLAR, "###,##0.00")
        Me.txtTipoCambio.Tag = Me.txtTipoCambio.Text

        Me.txtTipoCambioEuro.Text = VB6.Format(gcurCorpoTIPOCAMBIOEURO, "###,##0.00")
        Me.txtTipoCambioEuro.Tag = Me.txtTipoCambioEuro.Text

        Me.txtFacturas.Text = "0.00"
        Me.txtFacturas.Tag = Me.txtFacturas.Text
        Me.txtDescuentoFinanciero.Text = "0.00"
        Me.txtDescuentoFinanciero.Tag = Me.txtDescuentoFinanciero.Text
        Me.txtNotasCredito.Text = "0.00"
        Me.txtNotasCredito.Tag = Me.txtNotasCredito.Text
        Me.txtAnticipos.Text = "0.00"
        Me.txtAnticipos.Tag = Me.txtAnticipos.Text
        Me.txtAPagar.Text = "0.00"
        Me.txtAPagar.Tag = Me.txtAPagar.Text

        Me.optMoneda(0).Checked = True
        Me.optMoneda(1).Checked = False
        cMonedaPago = C_DOLAR

        Call Me.ActualizaCantidades()

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub LlenaDatos()
        On Error GoTo Merr
        Dim rsLocal As ADODB.Recordset
        Dim I As Integer
        If mintCodProveedor = 0 Then
            Limpiar()
            Exit Sub
        End If
        gStrSql = " SELECT a.FolioProgramacionP, a.NumPartida, a.CodProvAcreed, a.TipoFacturaCxP, a.TipoGasto, " & " (select sum(c.TotalPago) from Pagos c where c.FolioProgramacionP = a.FolioProgramacionP and c.NumPartida = a.NumPartida and c.Estatus <> 'C') as ImportePagos, " & " a.FolioFactura, a.FechaFactura, b.FechaVencto, a.FechaPago, a.TotalPago, a.Moneda, a.TipoCambio, a.TipoCambioE, " & " a.DescuentoFinanciero, a.SubtotalDF, a.IvaDF, a.Estatus, a.FechaCancel, a.TipoPagoProg, " & " a.Efectivo " & " FROM ProgramacionPagos a LEFT OUTER JOIN CxPFacturas b ON a.CodProvAcreed = b.CodProvAcreed " & " and b.Estatus <> 'C'  and  Ltrim(rTrim(a.FolioFactura)) = Ltrim(Rtrim(b.FolioFactura)) " & " WHERE a.CodProvAcreed = " & mintCodProveedor & " and a.TipoGasto = '" & IIf(Me.optOrigen(0).Checked, C_TIPOGASTOJOYERIA, C_TIPOGASTOPERSONAL) & "' " & " And a.Estatus = 'V'"
        '''02 SEP 2004
        '''eliminar la fecha de pago para que muestre todas las que estan pendientes ( con saldo )
        '''" And a.FechaPago <= '" & Format(Me.dtpCorte.Value, "mm/dd/yyyy") & "' and a.Estatus = 'V' "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            If Trim(Me.mshPagos.get_TextMatrix(2, P_COLPAGO)) <> "" Then
                Call Me.EncabezadoPagos()
            End If
            If Trim(Me.mshNotasCredito.get_TextMatrix(1, N_COLFOLIO)) <> "" Then
                Call Me.EncabezadoNotas()
            End If
            If Trim(Me.mshAnticipos.get_TextMatrix(1, A_COLFOLIO)) <> "" Then
                Call Me.EncabezadoAnticipos()
            End If
            With Me.mshPagos
                If rsLocal.RecordCount >= 11 Then
                    .Rows = rsLocal.RecordCount + 3
                Else
                    .Rows = 14
                End If
                rsLocal.MoveFirst()
                For I = 1 To rsLocal.RecordCount
                    .set_TextMatrix(I + 1, P_COLFACTURA, rsLocal.Fields("FolioFactura").Value)
                    .set_TextMatrix(I + 1, P_COLPAGO, rsLocal.Fields("FolioProgramacionP").Value)
                    .set_TextMatrix(I + 1, P_COLFECHAFACTURA, IIf(rsLocal.Fields("FechaFactura").Value = #1/1/1900#, "", VB6.Format(rsLocal.Fields("FechaFactura").Value, "dd/MMM/yyyy")))
                    .set_TextMatrix(I + 1, P_COLFECHAVENCTO, IIf(IsDBNull(rsLocal.Fields("FechaVencto").Value), "", VB6.Format(rsLocal.Fields("FechaVencto").Value, "dd/MMM/yyyy")))
                    .set_TextMatrix(I + 1, P_COLFECHAPAGO, VB6.Format(rsLocal.Fields("FechaPago").Value, "dd/MMM/yyyy"))
                    .set_TextMatrix(I + 1, P_COLIMPORTE, VB6.Format(rsLocal.Fields("TotalPago").Value, "###,###,##0.00"))
                    .set_TextMatrix(I + 1, P_COLPAGOS, VB6.Format(IIf(IsDBNull(rsLocal.Fields("ImportePagos").Value), 0, rsLocal.Fields("ImportePagos").Value), "###,###,##0.00"))
                    .set_TextMatrix(I + 1, P_COLSALDO, VB6.Format(IIf(IsDBNull(rsLocal.Fields("ImportePagos").Value), rsLocal.Fields("TotalPago").Value, rsLocal.Fields("TotalPago").Value - rsLocal.Fields("ImportePagos").Value), "###,###,##0.00"))
                    .set_TextMatrix(I + 1, P_COLIMPORTEPAGO, "0.00")
                    .set_TextMatrix(I + 1, P_COLDESCTOFIN, "0.00")
                    .set_TextMatrix(I + 1, P_COLDESCTOPORC, "0.00")
                    .set_TextMatrix(I + 1, P_COLSUBTOTALDF, "0.00")
                    .set_TextMatrix(I + 1, P_COLIVADF, "0.00")
                    .set_TextMatrix(I + 1, P_COLAPAGAR, "0.00")
                    .set_TextMatrix(I + 1, P_COLNUMPARTIDA, rsLocal.Fields("NumPartida").Value)
                    .set_TextMatrix(I + 1, P_COLMONEDA, Trim(IIf(rsLocal.Fields("Moneda").Value = C_DOLAR, "DOL", IIf(rsLocal.Fields("Moneda").Value = C_PESO, "PES", "EUR"))))
                    .set_TextMatrix(I + 1, P_COLESTATUS, "")
                    .set_TextMatrix(I + 1, P_COLDESCTOFINPORC, (rsLocal.Fields("DescuentoFinanciero").Value / 100))
                    rsLocal.MoveNext()
                Next I
            End With
            'Llenar el Grid de Notas de crédito
            gStrSql = "SELECT * FROM NotasCreditoCab WHERE CodProvAcreed = " & mintCodProveedor & " and Estatus = 'V'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            rsLocal = Cmd.Execute
            If rsLocal.RecordCount > 0 Then
                With Me.mshNotasCredito
                    rsLocal.MoveFirst()
                    For I = 1 To rsLocal.RecordCount
                        .set_TextMatrix(I, N_COLFOLIO, Trim(rsLocal.Fields("FolioNotaCredito").Value))
                        .set_TextMatrix(I, N_COLFECHA, VB6.Format(rsLocal.Fields("FechaNotaCredito").Value, "dd/MMM/yyyy"))
                        .set_TextMatrix(I, N_COLFACTURA, Trim(rsLocal.Fields("FolioFactura").Value))
                        .set_TextMatrix(I, N_COLTOTAL, VB6.Format(rsLocal.Fields("Total").Value, "###,###,##0.00"))
                        .set_TextMatrix(I, N_COLTIPO, IIf(rsLocal.Fields("TipoNotaCredito").Value = C_TIPONOTADEVOLUCION, "DEVOL.", "BONIF."))
                        .set_TextMatrix(I, N_COLMONEDA, Trim(IIf(rsLocal.Fields("Moneda").Value = C_DOLAR, C_DESCDOLARES, IIf(rsLocal.Fields("Moneda").Value = C_PESO, C_DESCPESOS, C_DESCEUROS))))
                        If I = .Rows - 1 Then .Rows = .Rows + 1
                        rsLocal.MoveNext()
                    Next I
                End With
            Else
                If Trim(Me.mshNotasCredito.get_TextMatrix(1, N_COLFOLIO)) <> "" Then
                    Call Me.EncabezadoNotas()
                End If
            End If
            'Llenar el grid de Anticipos
            gStrSql = "SELECT * FROM Anticipos WHERE CodProvAcreed = " & mintCodProveedor & " and Estatus = '" & C_STVIGENTE & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            rsLocal = Cmd.Execute
            If rsLocal.RecordCount > 0 Then
                rsLocal.MoveFirst()
                With Me.mshAnticipos
                    For I = 1 To rsLocal.RecordCount
                        .set_TextMatrix(I, A_COLFOLIO, Trim(rsLocal.Fields("FolioAnticipo").Value))
                        .set_TextMatrix(I, A_COLFECHA, VB6.Format(rsLocal.Fields("FechaAnticipo").Value, "dd/MMM/yyyy"))
                        .set_TextMatrix(I, A_COLIMPORTE, VB6.Format(rsLocal.Fields("Total").Value, "###,###,##0.00"))
                        .set_TextMatrix(I, A_COLMONEDA, Trim(IIf(rsLocal.Fields("Moneda").Value = C_DOLAR, C_DESCDOLARES, IIf(rsLocal.Fields("Moneda").Value = C_PESO, C_DESCPESOS, C_DESCEUROS))))
                        If I = .Rows - 1 Then .Rows = .Rows + 1
                        rsLocal.MoveNext()
                    Next I
                End With
            Else
                If Trim(Me.mshAnticipos.get_TextMatrix(1, A_COLFOLIO)) <> "" Then
                    Call Me.EncabezadoAnticipos()
                End If
            End If
            Call ActualizaCantidades()
        Else
            'MsgBox "El proveedor/acreedor que seleccionó no tiene ningún pago pendiente a la fecha, de " & IIf(Me.optOrigen(0).Value, "Joyería", "índole Personal"), vbOKOnly + vbInformation, gstrNombCortoEmpresa
            Nuevo()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub btnGenerar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnGenerar.Click
        If Not ValidaDatos() Then
            Exit Sub
        Else
            Me.Enabled = False
            frmBancosProcesoDiarioRegistrodePagos.Tag = UCase(Me.Name)
            frmBancosProcesoDiarioRegistrodePagos.blnEmisionPagos = True
            Call frmBancosProcesoDiarioRegistrodePagos.LlenaDatosPagos(mintCodProveedor, Trim(Me.dbcProveedor.Text), CDec(Numerico((Me.txtAPagar.Text))), cMonedaPago)
        End If
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        If Trim(dbcProveedor.Text) = "" Then
            Limpiar()
        End If

        lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcProveedor)

        If Me.dbcProveedor.SelectedItem <> "" Then
            dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        Pon_Tool()
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed ORDER BY descProvAcreed"
        ModDCombo.DCGotFocus(gStrSql, dbcProveedor)
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcProveedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        Aux = mintCodProveedor
        mintCodProveedor = 0
        ModDCombo.DCLostFocus(dbcProveedor, gStrSql, mintCodProveedor)
        If Aux <> mintCodProveedor Then
            Nuevo()
            Call Me.LlenaDatos()
        End If
    End Sub

    'Private Sub dbcProveedor_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '    Dim Aux As String
    '    Aux = Trim(Me.dbcProveedor.text)
    '    If Me.dbcProveedor.SelectedItem <> 0 Then
    '        dbcProveedor_LostFocus
    '        Aux = Trim(Me.dbcProveedor.text)
    '    End If
    '    Me.dbcProveedor.text = Aux
    'End Sub

    Private Sub frmCXPEmisionPagos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCXPEmisionPagos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPEmisionPagos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                If Me.ActiveControl.Name = "MSHPAGOS" Then
                    If Me.mshPagos.Col <> P_COLIMPORTEPAGO And Me.mshPagos.Col <> P_COLDESCTOFIN Then
                        ModEstandar.AvanzarTab(Me)
                    End If
                Else
                    ModEstandar.AvanzarTab(Me)
                End If
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmCXPEmisionPagos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.gp_CampoMayusculas(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPEmisionPagos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me, MDIMenuPrincipalCorpo)
        Me.dtpCorte.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
        mblnLoad = True
        Call Me.Nuevo()
    End Sub

    Private Sub frmCXPEmisionPagos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If mblnSalir Then
            mblnSalir = False
            Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Sale del Formulario
                    Cancel = 0
                Case MsgBoxResult.No 'No sale del formulario
                    Me.dbcProveedor.Focus()
                    ModEstandar.SelTxt()
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCXPEmisionPagos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'UPGRADE_NOTE: Object frmCXPEmisionPagos may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'Me = Nothing
    End Sub

    Private Sub mshAnticipos_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshAnticipos.DblClick
        mshAnticipos_KeyPressEvent(mshAnticipos, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(32))
    End Sub

    Private Sub mshAnticipos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshAnticipos.Enter
        Pon_Tool()
    End Sub

    Private Sub mshAnticipos_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles mshAnticipos.KeyPressEvent
        With Me.mshAnticipos
            Select Case eventArgs.keyAscii
                Case 32
                    If Trim(.get_TextMatrix(.Row, A_COLFOLIO)) = "" Then
                        Exit Sub
                    End If
                    If Trim(.get_TextMatrix(.Row, A_COLESTATUS)) = "" Then
                        .set_TextMatrix(.Row, A_COLESTATUS, C_PAGAR)
                    ElseIf Trim(.get_TextMatrix(.Row, A_COLESTATUS)) = C_PAGAR Then
                        .set_TextMatrix(.Row, A_COLESTATUS, "")
                    End If
                    Call Me.PonerColorAnticipos()
                    Call Me.ActualizaCantidades()
            End Select
        End With
    End Sub

    Private Sub mshNotasCredito_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshNotasCredito.DblClick
        mshNotasCredito_KeyPressEvent(mshNotasCredito, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(32))
    End Sub

    Private Sub mshNotasCredito_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshNotasCredito.Enter
        Pon_Tool()
    End Sub

    Private Sub mshNotasCredito_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles mshNotasCredito.KeyPressEvent
        With Me.mshNotasCredito
            Select Case eventArgs.keyAscii
                Case 32
                    If Trim(.get_TextMatrix(.Row, N_COLFOLIO)) = "" Then
                        Exit Sub
                    End If
                    If Trim(.get_TextMatrix(.Row, N_COLESTATUS)) = "" Then
                        .set_TextMatrix(.Row, N_COLESTATUS, C_PAGAR)
                    ElseIf Trim(.get_TextMatrix(.Row, N_COLESTATUS)) = C_PAGAR Then
                        .set_TextMatrix(.Row, N_COLESTATUS, "")
                    End If
                    Call Me.PonerColorNotas()
                    Call Me.ActualizaCantidades()
            End Select
        End With
    End Sub

    Private Sub mshPagos_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshPagos.DblClick
        mshPagos_KeyPressEvent(mshPagos, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(32))
    End Sub

    Private Sub mshPagos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshPagos.Enter
        Pon_Tool()
    End Sub

    Private Sub mshPagos_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles mshPagos.KeyPressEvent
        With Me.mshPagos
            Select Case eventArgs.keyAscii
                Case 13
                    If .get_TextMatrix(.Row, P_COLPAGO) = "" Then
                        Exit Sub
                    End If
                    Select Case .Col
                        Case P_COLIMPORTEPAGO
                            Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            Me.txtFlex.BackColor = .CellBackColor
                            ModEstandar.MSHFlexGridEdit(mshPagos, txtFlex, eventArgs.keyAscii)
                            ModEstandar.SelTextoTxt((Me.txtFlex))
                        Case P_COLDESCTOFIN
                            Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            Me.txtFlex.BackColor = .CellBackColor
                            ModEstandar.MSHFlexGridEdit(mshPagos, txtFlex, eventArgs.keyAscii)
                            ModEstandar.SelTextoTxt((Me.txtFlex))
                        Case P_COLAPAGAR
                            .Row = .Row + 1
                            If Trim(.get_TextMatrix(.Row, P_COLPAGO)) = "" Then
                                .Col = 0
                            Else
                                .Col = P_COLIMPORTEPAGO
                            End If
                    End Select
                Case 32
                    If .get_TextMatrix(.Row, P_COLPAGO) = "" Then
                        Exit Sub
                    End If
                    If Trim(.get_TextMatrix(.Row, P_COLESTATUS)) = "" Then
                        'Cambia el estatus a pagar
                        .set_TextMatrix(.Row, P_COLESTATUS, C_PAGAR)
                        'Pone el importe del saldo como sugerencia para el pago, pone la columna del grid
                        'en la columna del pago
                        .Col = P_COLIMPORTEPAGO
                        .set_TextMatrix(.Row, P_COLIMPORTEPAGO, VB6.Format(CDec(Numerico(.get_TextMatrix(.Row, P_COLSALDO))), "###,###,##0.00"))
                        Call Me.PonerColor()
                        Call Me.ActualizaCantidades()
                        'Activa el txtFlex con la cantidad del pago sugerido
                        Me.txtFlex.Text = .get_TextMatrix(.Row, P_COLIMPORTEPAGO)
                        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                        Me.txtFlex.BackColor = .CellBackColor
                        ModEstandar.MSHFlexGridEdit(mshPagos, txtFlex, eventArgs.keyAscii)
                        ModEstandar.SelTextoTxt((Me.txtFlex))
                    ElseIf Trim(.get_TextMatrix(.Row, P_COLESTATUS)) = C_PAGAR Then
                        .set_TextMatrix(.Row, P_COLESTATUS, "")
                        Call Me.PonerColor()
                        Call Me.ActualizaCantidades()
                    End If
                Case 27
                Case Else
                    Select Case .Col
                        Case P_COLIMPORTEPAGO
                            Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            Me.txtFlex.BackColor = .CellBackColor
                            ModEstandar.MSHFlexGridEdit(mshPagos, txtFlex, eventArgs.keyAscii)
                            If Len(Me.txtFlex.Text) <> 1 Then
                                ModEstandar.SelTextoTxt((Me.txtFlex))
                            End If
                        Case P_COLDESCTOFIN
                            Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            Me.txtFlex.BackColor = .CellBackColor
                            ModEstandar.MSHFlexGridEdit(mshPagos, txtFlex, eventArgs.keyAscii)
                            If Len(Me.txtFlex.Text) <> 1 Then
                                ModEstandar.SelTextoTxt((Me.txtFlex))
                            End If
                    End Select
            End Select
        End With
    End Sub

    Private Sub mshPagos_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshPagos.Scroll
        txtFlex.Visible = False
    End Sub

    'UPGRADE_WARNING: Event optMoneda.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optMoneda_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoneda.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optMoneda.GetIndex(eventSender)
            Select Case Index
                Case 0
                    cMonedaPago = C_DOLAR
                Case Else
                    cMonedaPago = C_PESO
            End Select
            Call Me.ActualizaCantidades()
        End If
    End Sub

    'UPGRADE_WARNING: Event optOrigen.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optOrigen_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optOrigen.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optOrigen.GetIndex(eventSender)
            Dim aintCodProveedor As Integer
            Dim cDescProveedor As String
            aintCodProveedor = mintCodProveedor
            'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
            cDescProveedor = Trim(Me.dbcProveedor.Text)
            LlenaDatos()
            mblnFueraChange = True
            mintCodProveedor = aintCodProveedor
            'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
            Me.dbcProveedor.Text = Trim(cDescProveedor)
            mblnFueraChange = False
        End If
    End Sub

    Private Sub txtAnticipos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnticipos.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub txtAnticipos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAnticipos.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtAnticipos.Text = VB6.Format(Numerico((Me.txtAnticipos.Text)), "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtAnticipos.Text), KeyAscii, 9, 2, (Me.txtAnticipos.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtAnticipos_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnticipos.Leave
        Me.txtAnticipos.Text = VB6.Format(Numerico((Me.txtAnticipos.Text)), "###,###,##0.00")
    End Sub

    Private Sub txtAPagar_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAPagar.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub txtAPagar_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAPagar.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtAPagar.Text = VB6.Format(Numerico((Me.txtAPagar.Text)), "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtAPagar.Text), KeyAscii, 9, 2, (Me.txtAPagar.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtAPagar_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAPagar.Leave
        Me.txtAPagar.Text = VB6.Format(Numerico((Me.txtAPagar.Text)), "###,###,##0.00")
    End Sub

    Private Sub txtDescuentoFinanciero_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescuentoFinanciero.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub txtDescuentoFinanciero_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDescuentoFinanciero.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtDescuentoFinanciero.Text = VB6.Format(Numerico((Me.txtDescuentoFinanciero.Text)), "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtDescuentoFinanciero.Text), KeyAscii, 9, 2, (Me.txtDescuentoFinanciero.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDescuentoFinanciero_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescuentoFinanciero.Leave
        Me.txtDescuentoFinanciero.Text = VB6.Format(Numerico((Me.txtDescuentoFinanciero.Text)), "###,###,##0.00")
    End Sub

    Private Sub txtFacturas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFacturas.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub txtFacturas_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFacturas.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtFacturas.Text = VB6.Format(Numerico((Me.txtFacturas.Text)), "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtFacturas.Text), KeyAscii, 9, 2, (Me.txtFacturas.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFacturas_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFacturas.Leave
        Me.txtFacturas.Text = VB6.Format(Numerico((Me.txtFacturas.Text)), "###,###,##0.00")
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Dim nCol As Object
        Dim nRen As Integer
        Dim nIva As Object
        Dim nIVAImporte As Object
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        With Me.mshPagos
            'UPGRADE_WARNING: Couldn't resolve default property of object nCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nCol = .Col
            nRen = .Row
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    Call ActualizaCantidades()
                    Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left 'Alinear a la izquierda
                    txtFlex.Visible = False
                    txtFlex.Text = ""
                    .Focus()
                Case System.Windows.Forms.Keys.Return
                    Select Case .Col
                        Case P_COLIMPORTEPAGO
                            If CDec(Numerico((Me.txtFlex.Text))) > 0 Then
                                .set_TextMatrix(.Row, P_COLESTATUS, C_PAGAR)
                                .set_TextMatrix(.Row, P_COLIMPORTEPAGO, VB6.Format(Numerico((Me.txtFlex.Text)), "###,###,##0.00"))
                                .set_TextMatrix(.Row, P_COLDESCTOFIN, VB6.Format(CDec(Numerico(.get_TextMatrix(.Row, P_COLIMPORTEPAGO))) * CDec(Numerico(.get_TextMatrix(.Row, P_COLDESCTOFINPORC))), "###,###,##0.00"))
                                'Calcular el descuento financiero sugerido
                                .Col = P_COLDESCTOFIN
                            Else
                                .set_TextMatrix(.Row, P_COLIMPORTEPAGO, "0.00")
                                .set_TextMatrix(.Row, P_COLDESCTOFIN, "0.00")
                                .set_TextMatrix(.Row, P_COLESTATUS, "")
                                .Col = P_COLDESCTOFIN
                            End If
                            Call Me.PonerColor()
                            Call Me.ActualizaCantidades()
                        Case P_COLDESCTOFIN
                            .set_TextMatrix(.Row, P_COLDESCTOFIN, VB6.Format(Numerico((Me.txtFlex.Text)), "###,###,##0.00"))
                            .Col = P_COLAPAGAR
                    End Select
                    Call ActualizaCantidades()
                    If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then .Focus()
                    Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left 'Alinear a la izquierda
                    txtFlex.Text = ""
                    txtFlex.Visible = False
            End Select
        End With
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtFlex.Text = VB6.Format(Numerico((Me.txtFlex.Text)), "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtFlex.Text), KeyAscii, 9, 2, (Me.txtFlex.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Private Sub txtNotasCredito_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNotasCredito.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub txtNotasCredito_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNotasCredito.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtNotasCredito.Text = VB6.Format(Numerico((Me.txtNotasCredito.Text)), "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtNotasCredito.Text), KeyAscii, 9, 2, (Me.txtNotasCredito.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtNotasCredito_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNotasCredito.Leave
        Me.txtNotasCredito.Text = VB6.Format(Numerico((Me.txtNotasCredito.Text)), "###,###,##0.00")
    End Sub

    'UPGRADE_WARNING: Event txtTipoCambio.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtTipoCambio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.TextChanged
        Call Me.ActualizaCantidades()
    End Sub

    Private Sub txtTipoCambio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub txtTipoCambio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambio.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtTipoCambio.Text = VB6.Format(Numerico((Me.txtTipoCambio.Text)), "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtTipoCambio.Text), KeyAscii, 9, 2, (Me.txtTipoCambio.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Leave
        Me.txtTipoCambio.Text = VB6.Format(Numerico((Me.txtTipoCambio.Text)), "###,###,##0.00")
    End Sub

    'UPGRADE_WARNING: Event txtTipoCambioEuro.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtTipoCambioEuro_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.TextChanged
        Call Me.ActualizaCantidades()
    End Sub

    Private Sub txtTipoCambioEuro_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub txtTipoCambioEuro_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambioEuro.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtTipoCambioEuro.Text = VB6.Format(Numerico((Me.txtTipoCambioEuro.Text)), "###,###,##0.00")
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
End Class