<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAbcCuentasBancarias
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtCtaBancaria As System.Windows.Forms.TextBox
    Public WithEvents dbcBanco2 As AxMSDataListLib.AxDataCombo
    Public WithEvents _lblCuentasBancarias_1 As System.Windows.Forms.Label
    Public WithEvents _lblCuentasBancarias_0 As System.Windows.Forms.Label
    Public WithEvents fraBancos As System.Windows.Forms.GroupBox
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents cboLetraFolios As System.Windows.Forms.ComboBox
    Public WithEvents _optTipoCuenta_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optTipoCuenta_0 As System.Windows.Forms.RadioButton
    Public WithEvents txtSaldoInicial As System.Windows.Forms.TextBox
    Public WithEvents txtConsecutivodeChq As System.Windows.Forms.TextBox
    Public WithEvents txtCuentaHabiente As System.Windows.Forms.TextBox
    Public WithEvents txtSucursal As System.Windows.Forms.TextBox
    Public WithEvents _lblCuentasBancarias_7 As System.Windows.Forms.Label
    Public WithEvents _lblCuentasBancarias_6 As System.Windows.Forms.Label
    Public WithEvents _lblCuentasBancarias_5 As System.Windows.Forms.Label
    Public WithEvents _lblCuentasBancarias_4 As System.Windows.Forms.Label
    Public WithEvents _lblCuentasBancarias_3 As System.Windows.Forms.Label
    Public WithEvents _lblCuentasBancarias_2 As System.Windows.Forms.Label
    Public WithEvents fraInformacionGeneral As System.Windows.Forms.GroupBox
    Public WithEvents lblCuentasBancarias As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optTipoCuenta As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAbcCuentasBancarias))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCtaBancaria = New System.Windows.Forms.TextBox()
        Me._optTipoCuenta_1 = New System.Windows.Forms.RadioButton()
        Me._optTipoCuenta_0 = New System.Windows.Forms.RadioButton()
        Me.txtSaldoInicial = New System.Windows.Forms.TextBox()
        Me.txtConsecutivodeChq = New System.Windows.Forms.TextBox()
        Me.txtCuentaHabiente = New System.Windows.Forms.TextBox()
        Me.txtSucursal = New System.Windows.Forms.TextBox()
        Me.fraBancos = New System.Windows.Forms.GroupBox()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me.dbcBanco2 = New AxMSDataListLib.AxDataCombo()
        Me._lblCuentasBancarias_1 = New System.Windows.Forms.Label()
        Me._lblCuentasBancarias_0 = New System.Windows.Forms.Label()
        Me.fraInformacionGeneral = New System.Windows.Forms.GroupBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me.cboLetraFolios = New System.Windows.Forms.ComboBox()
        Me._lblCuentasBancarias_7 = New System.Windows.Forms.Label()
        Me._lblCuentasBancarias_6 = New System.Windows.Forms.Label()
        Me._lblCuentasBancarias_5 = New System.Windows.Forms.Label()
        Me._lblCuentasBancarias_4 = New System.Windows.Forms.Label()
        Me._lblCuentasBancarias_3 = New System.Windows.Forms.Label()
        Me._lblCuentasBancarias_2 = New System.Windows.Forms.Label()
        Me.lblCuentasBancarias = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optTipoCuenta = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.fraBancos.SuspendLayout()
        CType(Me.dbcBanco2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraInformacionGeneral.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.lblCuentasBancarias, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optTipoCuenta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCtaBancaria
        '
        Me.txtCtaBancaria.AcceptsReturn = True
        Me.txtCtaBancaria.BackColor = System.Drawing.SystemColors.Window
        Me.txtCtaBancaria.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCtaBancaria.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCtaBancaria.Location = New System.Drawing.Point(136, 64)
        Me.txtCtaBancaria.MaxLength = 16
        Me.txtCtaBancaria.Name = "txtCtaBancaria"
        Me.txtCtaBancaria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCtaBancaria.Size = New System.Drawing.Size(200, 20)
        Me.txtCtaBancaria.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCtaBancaria, "Número de Cuenta")
        '
        '_optTipoCuenta_1
        '
        Me._optTipoCuenta_1.BackColor = System.Drawing.SystemColors.Control
        Me._optTipoCuenta_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoCuenta_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTipoCuenta.SetIndex(Me._optTipoCuenta_1, CType(1, Short))
        Me._optTipoCuenta_1.Location = New System.Drawing.Point(200, 32)
        Me._optTipoCuenta_1.Name = "_optTipoCuenta_1"
        Me._optTipoCuenta_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoCuenta_1.Size = New System.Drawing.Size(81, 21)
        Me._optTipoCuenta_1.TabIndex = 3
        Me._optTipoCuenta_1.TabStop = True
        Me._optTipoCuenta_1.Text = "Inversiones"
        Me.ToolTip1.SetToolTip(Me._optTipoCuenta_1, "Tipo de Cuenta")
        Me._optTipoCuenta_1.UseVisualStyleBackColor = False
        '
        '_optTipoCuenta_0
        '
        Me._optTipoCuenta_0.BackColor = System.Drawing.SystemColors.Control
        Me._optTipoCuenta_0.Checked = True
        Me._optTipoCuenta_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoCuenta_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTipoCuenta.SetIndex(Me._optTipoCuenta_0, CType(0, Short))
        Me._optTipoCuenta_0.Location = New System.Drawing.Point(120, 32)
        Me._optTipoCuenta_0.Name = "_optTipoCuenta_0"
        Me._optTipoCuenta_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoCuenta_0.Size = New System.Drawing.Size(81, 21)
        Me._optTipoCuenta_0.TabIndex = 2
        Me._optTipoCuenta_0.TabStop = True
        Me._optTipoCuenta_0.Text = "Normal"
        Me.ToolTip1.SetToolTip(Me._optTipoCuenta_0, "Tipo de Cuenta")
        Me._optTipoCuenta_0.UseVisualStyleBackColor = False
        '
        'txtSaldoInicial
        '
        Me.txtSaldoInicial.AcceptsReturn = True
        Me.txtSaldoInicial.BackColor = System.Drawing.SystemColors.Window
        Me.txtSaldoInicial.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSaldoInicial.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSaldoInicial.Location = New System.Drawing.Point(128, 160)
        Me.txtSaldoInicial.MaxLength = 0
        Me.txtSaldoInicial.Name = "txtSaldoInicial"
        Me.txtSaldoInicial.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSaldoInicial.Size = New System.Drawing.Size(105, 20)
        Me.txtSaldoInicial.TabIndex = 9
        Me.txtSaldoInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtSaldoInicial, "Saldo Inicial")
        '
        'txtConsecutivodeChq
        '
        Me.txtConsecutivodeChq.AcceptsReturn = True
        Me.txtConsecutivodeChq.BackColor = System.Drawing.SystemColors.Window
        Me.txtConsecutivodeChq.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConsecutivodeChq.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConsecutivodeChq.Location = New System.Drawing.Point(376, 160)
        Me.txtConsecutivodeChq.MaxLength = 6
        Me.txtConsecutivodeChq.Name = "txtConsecutivodeChq"
        Me.txtConsecutivodeChq.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConsecutivodeChq.Size = New System.Drawing.Size(89, 20)
        Me.txtConsecutivodeChq.TabIndex = 10
        Me.txtConsecutivodeChq.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtConsecutivodeChq, "Consecutivo de Cheques")
        '
        'txtCuentaHabiente
        '
        Me.txtCuentaHabiente.AcceptsReturn = True
        Me.txtCuentaHabiente.BackColor = System.Drawing.SystemColors.Window
        Me.txtCuentaHabiente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCuentaHabiente.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCuentaHabiente.Location = New System.Drawing.Point(128, 96)
        Me.txtCuentaHabiente.MaxLength = 40
        Me.txtCuentaHabiente.Name = "txtCuentaHabiente"
        Me.txtCuentaHabiente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCuentaHabiente.Size = New System.Drawing.Size(337, 20)
        Me.txtCuentaHabiente.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtCuentaHabiente, "Cuentahabiente")
        '
        'txtSucursal
        '
        Me.txtSucursal.AcceptsReturn = True
        Me.txtSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSucursal.Location = New System.Drawing.Point(128, 64)
        Me.txtSucursal.MaxLength = 4
        Me.txtSucursal.Name = "txtSucursal"
        Me.txtSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSucursal.Size = New System.Drawing.Size(49, 20)
        Me.txtSucursal.TabIndex = 6
        Me.txtSucursal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtSucursal, "Número de Sucursal")
        '
        'fraBancos
        '
        Me.fraBancos.BackColor = System.Drawing.SystemColors.Control
        Me.fraBancos.Controls.Add(Me.dbcBanco)
        Me.fraBancos.Controls.Add(Me.txtCtaBancaria)
        Me.fraBancos.Controls.Add(Me.dbcBanco2)
        Me.fraBancos.Controls.Add(Me._lblCuentasBancarias_1)
        Me.fraBancos.Controls.Add(Me._lblCuentasBancarias_0)
        Me.fraBancos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraBancos.Location = New System.Drawing.Point(8, 8)
        Me.fraBancos.Name = "fraBancos"
        Me.fraBancos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraBancos.Size = New System.Drawing.Size(481, 105)
        Me.fraBancos.TabIndex = 11
        Me.fraBancos.TabStop = False
        '
        'dbcBanco
        '
        Me.dbcBanco.FormattingEnabled = True
        Me.dbcBanco.Location = New System.Drawing.Point(136, 20)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(200, 21)
        Me.dbcBanco.TabIndex = 14
        '
        'dbcBanco2
        '
        Me.dbcBanco2.Location = New System.Drawing.Point(360, 19)
        Me.dbcBanco2.Name = "dbcBanco2"
        Me.dbcBanco2.OcxState = CType(resources.GetObject("dbcBanco2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dbcBanco2.RowSource = Nothing
        Me.dbcBanco2.Size = New System.Drawing.Size(100, 21)
        Me.dbcBanco2.TabIndex = 0
        '
        '_lblCuentasBancarias_1
        '
        Me._lblCuentasBancarias_1.AutoSize = True
        Me._lblCuentasBancarias_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblCuentasBancarias_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCuentasBancarias.SetIndex(Me._lblCuentasBancarias_1, CType(1, Short))
        Me._lblCuentasBancarias_1.Location = New System.Drawing.Point(32, 28)
        Me._lblCuentasBancarias_1.Name = "_lblCuentasBancarias_1"
        Me._lblCuentasBancarias_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCuentasBancarias_1.Size = New System.Drawing.Size(38, 13)
        Me._lblCuentasBancarias_1.TabIndex = 12
        Me._lblCuentasBancarias_1.Text = "Banco"
        '
        '_lblCuentasBancarias_0
        '
        Me._lblCuentasBancarias_0.AutoSize = True
        Me._lblCuentasBancarias_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblCuentasBancarias_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCuentasBancarias.SetIndex(Me._lblCuentasBancarias_0, CType(0, Short))
        Me._lblCuentasBancarias_0.Location = New System.Drawing.Point(32, 68)
        Me._lblCuentasBancarias_0.Name = "_lblCuentasBancarias_0"
        Me._lblCuentasBancarias_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCuentasBancarias_0.Size = New System.Drawing.Size(86, 13)
        Me._lblCuentasBancarias_0.TabIndex = 13
        Me._lblCuentasBancarias_0.Text = "Cuenta Bancaria"
        '
        'fraInformacionGeneral
        '
        Me.fraInformacionGeneral.BackColor = System.Drawing.SystemColors.Control
        Me.fraInformacionGeneral.Controls.Add(Me.Frame1)
        Me.fraInformacionGeneral.Controls.Add(Me.cboLetraFolios)
        Me.fraInformacionGeneral.Controls.Add(Me._optTipoCuenta_1)
        Me.fraInformacionGeneral.Controls.Add(Me._optTipoCuenta_0)
        Me.fraInformacionGeneral.Controls.Add(Me.txtSaldoInicial)
        Me.fraInformacionGeneral.Controls.Add(Me.txtConsecutivodeChq)
        Me.fraInformacionGeneral.Controls.Add(Me.txtCuentaHabiente)
        Me.fraInformacionGeneral.Controls.Add(Me.txtSucursal)
        Me.fraInformacionGeneral.Controls.Add(Me._lblCuentasBancarias_7)
        Me.fraInformacionGeneral.Controls.Add(Me._lblCuentasBancarias_6)
        Me.fraInformacionGeneral.Controls.Add(Me._lblCuentasBancarias_5)
        Me.fraInformacionGeneral.Controls.Add(Me._lblCuentasBancarias_4)
        Me.fraInformacionGeneral.Controls.Add(Me._lblCuentasBancarias_3)
        Me.fraInformacionGeneral.Controls.Add(Me._lblCuentasBancarias_2)
        Me.fraInformacionGeneral.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraInformacionGeneral.Location = New System.Drawing.Point(8, 120)
        Me.fraInformacionGeneral.Name = "fraInformacionGeneral"
        Me.fraInformacionGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraInformacionGeneral.Size = New System.Drawing.Size(481, 201)
        Me.fraInformacionGeneral.TabIndex = 14
        Me.fraInformacionGeneral.TabStop = False
        Me.fraInformacionGeneral.Text = "Información General"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me._optMoneda_0)
        Me.Frame1.Controls.Add(Me._optMoneda_1)
        Me.Frame1.ForeColor = System.Drawing.Color.Black
        Me.Frame1.Location = New System.Drawing.Point(304, 16)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(161, 65)
        Me.Frame1.TabIndex = 21
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Moneda"
        '
        '_optMoneda_0
        '
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_0, CType(0, Short))
        Me._optMoneda_0.Location = New System.Drawing.Point(56, 16)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(67, 17)
        Me._optMoneda_0.TabIndex = 4
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Text = "Pesos"
        Me._optMoneda_0.UseVisualStyleBackColor = False
        '
        '_optMoneda_1
        '
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_1, CType(1, Short))
        Me._optMoneda_1.Location = New System.Drawing.Point(56, 40)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(67, 17)
        Me._optMoneda_1.TabIndex = 5
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Text = "Dolares"
        Me._optMoneda_1.UseVisualStyleBackColor = False
        '
        'cboLetraFolios
        '
        Me.cboLetraFolios.BackColor = System.Drawing.SystemColors.Window
        Me.cboLetraFolios.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLetraFolios.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLetraFolios.Location = New System.Drawing.Point(128, 128)
        Me.cboLetraFolios.Name = "cboLetraFolios"
        Me.cboLetraFolios.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLetraFolios.Size = New System.Drawing.Size(41, 21)
        Me.cboLetraFolios.Sorted = True
        Me.cboLetraFolios.TabIndex = 8
        '
        '_lblCuentasBancarias_7
        '
        Me._lblCuentasBancarias_7.AutoSize = True
        Me._lblCuentasBancarias_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblCuentasBancarias_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCuentasBancarias.SetIndex(Me._lblCuentasBancarias_7, CType(7, Short))
        Me._lblCuentasBancarias_7.Location = New System.Drawing.Point(248, 164)
        Me._lblCuentasBancarias_7.Name = "_lblCuentasBancarias_7"
        Me._lblCuentasBancarias_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCuentasBancarias_7.Size = New System.Drawing.Size(126, 13)
        Me._lblCuentasBancarias_7.TabIndex = 20
        Me._lblCuentasBancarias_7.Text = "Consecutivo de Cheques"
        '
        '_lblCuentasBancarias_6
        '
        Me._lblCuentasBancarias_6.AutoSize = True
        Me._lblCuentasBancarias_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblCuentasBancarias_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCuentasBancarias.SetIndex(Me._lblCuentasBancarias_6, CType(6, Short))
        Me._lblCuentasBancarias_6.Location = New System.Drawing.Point(24, 164)
        Me._lblCuentasBancarias_6.Name = "_lblCuentasBancarias_6"
        Me._lblCuentasBancarias_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCuentasBancarias_6.Size = New System.Drawing.Size(64, 13)
        Me._lblCuentasBancarias_6.TabIndex = 19
        Me._lblCuentasBancarias_6.Text = "Saldo Inicial"
        '
        '_lblCuentasBancarias_5
        '
        Me._lblCuentasBancarias_5.AutoSize = True
        Me._lblCuentasBancarias_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblCuentasBancarias_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCuentasBancarias.SetIndex(Me._lblCuentasBancarias_5, CType(5, Short))
        Me._lblCuentasBancarias_5.Location = New System.Drawing.Point(24, 132)
        Me._lblCuentasBancarias_5.Name = "_lblCuentasBancarias_5"
        Me._lblCuentasBancarias_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCuentasBancarias_5.Size = New System.Drawing.Size(76, 13)
        Me._lblCuentasBancarias_5.TabIndex = 18
        Me._lblCuentasBancarias_5.Text = "Letra de Folios"
        '
        '_lblCuentasBancarias_4
        '
        Me._lblCuentasBancarias_4.AutoSize = True
        Me._lblCuentasBancarias_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblCuentasBancarias_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCuentasBancarias.SetIndex(Me._lblCuentasBancarias_4, CType(4, Short))
        Me._lblCuentasBancarias_4.Location = New System.Drawing.Point(24, 100)
        Me._lblCuentasBancarias_4.Name = "_lblCuentasBancarias_4"
        Me._lblCuentasBancarias_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCuentasBancarias_4.Size = New System.Drawing.Size(82, 13)
        Me._lblCuentasBancarias_4.TabIndex = 17
        Me._lblCuentasBancarias_4.Text = "Cuentahabiente"
        '
        '_lblCuentasBancarias_3
        '
        Me._lblCuentasBancarias_3.AutoSize = True
        Me._lblCuentasBancarias_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblCuentasBancarias_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCuentasBancarias.SetIndex(Me._lblCuentasBancarias_3, CType(3, Short))
        Me._lblCuentasBancarias_3.Location = New System.Drawing.Point(24, 68)
        Me._lblCuentasBancarias_3.Name = "_lblCuentasBancarias_3"
        Me._lblCuentasBancarias_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCuentasBancarias_3.Size = New System.Drawing.Size(83, 13)
        Me._lblCuentasBancarias_3.TabIndex = 16
        Me._lblCuentasBancarias_3.Text = "No. de Sucursal"
        '
        '_lblCuentasBancarias_2
        '
        Me._lblCuentasBancarias_2.AutoSize = True
        Me._lblCuentasBancarias_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblCuentasBancarias_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCuentasBancarias.SetIndex(Me._lblCuentasBancarias_2, CType(2, Short))
        Me._lblCuentasBancarias_2.Location = New System.Drawing.Point(24, 36)
        Me._lblCuentasBancarias_2.Name = "_lblCuentasBancarias_2"
        Me._lblCuentasBancarias_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCuentasBancarias_2.Size = New System.Drawing.Size(79, 13)
        Me._lblCuentasBancarias_2.TabIndex = 15
        Me._lblCuentasBancarias_2.Text = "Tipo de cuenta"
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Location = New System.Drawing.Point(303, 346)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(93, 35)
        Me.btnLimpiar.TabIndex = 59
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(196, 346)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(93, 35)
        Me.btnEliminar.TabIndex = 58
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.Location = New System.Drawing.Point(88, 346)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(93, 35)
        Me.btnGuardar.TabIndex = 57
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'frmAbcCuentasBancarias
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(498, 395)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.fraBancos)
        Me.Controls.Add(Me.fraInformacionGeneral)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 15)
        Me.MaximizeBox = False
        Me.Name = "frmAbcCuentasBancarias"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Cuentas Bancarias"
        Me.fraBancos.ResumeLayout(False)
        Me.fraBancos.PerformLayout()
        CType(Me.dbcBanco2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraInformacionGeneral.ResumeLayout(False)
        Me.fraInformacionGeneral.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        CType(Me.lblCuentasBancarias, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optTipoCuenta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private components As System.ComponentModel.IContainer
    Friend WithEvents dbcBanco As ComboBox
    Friend WithEvents btnLimpiar As Button
    Friend WithEvents btnEliminar As Button
    Friend WithEvents btnGuardar As Button
End Class
