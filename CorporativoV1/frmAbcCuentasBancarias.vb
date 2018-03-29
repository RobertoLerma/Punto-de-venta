'**********************************************************************************************************************'
'*PROGRAMA: ABC CUENTAS BANCARIAS JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018 
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic


Public Class frmAbcCuentasBancarias

    Inherits System.Windows.Forms.Form
    ' Programa :                ABC de Cuentas Bancarias
    ' Autor :                   Paimí
    ' Fecha de Inicio:          15 de Mayo de 2003
    ' Fecha de Finalización:


    Private components As System.ComponentModel.IContainer
    Friend WithEvents dbcBanco As ComboBox
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtCtaBancaria As System.Windows.Forms.TextBox
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


    Dim mblnSALIR As Boolean 'Controla la salida con ESCAPE

    Dim mblnNuevo As Boolean
    Dim mblnCambiosEnCodigo1 As Object
    Dim mblnCambiosEnCodigo2 As Boolean

    Dim rsLocal As ADODB.Recordset 'Para buscar la descripción de Banco

    Dim mstrTipoCuenta As Object
    Dim mstrTipoCuentaTag As String

    'Variables para manejar el combo
    Public mintCodBanco As Integer
    Dim mblnFueraChange As Boolean
    Dim tecla As Integer

    'Variables para manejar el combo de Letras de Folios
    Dim rsCombo As ADODB.Recordset
    Dim cLetras As Object
    Dim cLetrasEnCombo As String
    Dim maLetra(25) As String
    Dim blnPesos As Boolean
    Public WithEvents Panel1 As Panel
    Public WithEvents Panel3 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public blnDolares As Boolean
    Public strControlActual As String 'Nombre del control actual


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.fraBancos.SuspendLayout()
        Me.fraInformacionGeneral.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.lblCuentasBancarias, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optTipoCuenta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
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
        Me.txtCtaBancaria.Size = New System.Drawing.Size(252, 20)
        Me.txtCtaBancaria.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCtaBancaria, "Número de Cuenta")
        '
        '_optTipoCuenta_1
        '
        Me._optTipoCuenta_1.BackColor = System.Drawing.Color.Silver
        Me._optTipoCuenta_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoCuenta_1.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me._optTipoCuenta_0.BackColor = System.Drawing.Color.Silver
        Me._optTipoCuenta_0.Checked = True
        Me._optTipoCuenta_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoCuenta_0.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me.fraBancos.BackColor = System.Drawing.Color.Silver
        Me.fraBancos.Controls.Add(Me.dbcBanco)
        Me.fraBancos.Controls.Add(Me.txtCtaBancaria)
        Me.fraBancos.Controls.Add(Me._lblCuentasBancarias_1)
        Me.fraBancos.Controls.Add(Me._lblCuentasBancarias_0)
        Me.fraBancos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraBancos.Location = New System.Drawing.Point(13, 13)
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
        Me.dbcBanco.Size = New System.Drawing.Size(252, 21)
        Me.dbcBanco.TabIndex = 14
        '
        '_lblCuentasBancarias_1
        '
        Me._lblCuentasBancarias_1.AutoSize = True
        Me._lblCuentasBancarias_1.BackColor = System.Drawing.Color.Silver
        Me._lblCuentasBancarias_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_1.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me._lblCuentasBancarias_0.BackColor = System.Drawing.Color.Silver
        Me._lblCuentasBancarias_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblCuentasBancarias_0.Location = New System.Drawing.Point(32, 68)
        Me._lblCuentasBancarias_0.Name = "_lblCuentasBancarias_0"
        Me._lblCuentasBancarias_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCuentasBancarias_0.Size = New System.Drawing.Size(86, 13)
        Me._lblCuentasBancarias_0.TabIndex = 13
        Me._lblCuentasBancarias_0.Text = "Cuenta Bancaria"
        '
        'fraInformacionGeneral
        '
        Me.fraInformacionGeneral.BackColor = System.Drawing.Color.Silver
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
        Me.fraInformacionGeneral.Location = New System.Drawing.Point(13, 124)
        Me.fraInformacionGeneral.Name = "fraInformacionGeneral"
        Me.fraInformacionGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraInformacionGeneral.Size = New System.Drawing.Size(481, 201)
        Me.fraInformacionGeneral.TabIndex = 14
        Me.fraInformacionGeneral.TabStop = False
        Me.fraInformacionGeneral.Text = "Información General"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.Color.Silver
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
        Me._optMoneda_0.BackColor = System.Drawing.Color.Silver
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me._optMoneda_1.BackColor = System.Drawing.Color.Silver
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me._lblCuentasBancarias_7.BackColor = System.Drawing.Color.Silver
        Me._lblCuentasBancarias_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_7.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me._lblCuentasBancarias_6.BackColor = System.Drawing.Color.Silver
        Me._lblCuentasBancarias_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_6.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me._lblCuentasBancarias_5.BackColor = System.Drawing.Color.Silver
        Me._lblCuentasBancarias_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_5.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me._lblCuentasBancarias_4.BackColor = System.Drawing.Color.Silver
        Me._lblCuentasBancarias_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_4.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me._lblCuentasBancarias_3.BackColor = System.Drawing.Color.Silver
        Me._lblCuentasBancarias_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_3.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me._lblCuentasBancarias_2.BackColor = System.Drawing.Color.Silver
        Me._lblCuentasBancarias_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCuentasBancarias_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblCuentasBancarias_2.Location = New System.Drawing.Point(24, 36)
        Me._lblCuentasBancarias_2.Name = "_lblCuentasBancarias_2"
        Me._lblCuentasBancarias_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCuentasBancarias_2.Size = New System.Drawing.Size(79, 13)
        Me._lblCuentasBancarias_2.TabIndex = 15
        Me._lblCuentasBancarias_2.Text = "Tipo de cuenta"
        '
        'optMoneda
        '
        '
        'optTipoCuenta
        '
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.fraInformacionGeneral)
        Me.Panel1.Controls.Add(Me.fraBancos)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(508, 414)
        Me.Panel1.TabIndex = 15
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(13, 330)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(481, 74)
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
        'frmAbcCuentasBancarias
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(531, 438)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 15)
        Me.MaximizeBox = False
        Me.Name = "frmAbcCuentasBancarias"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Cuentas Bancarias"
        Me.fraBancos.ResumeLayout(False)
        Me.fraBancos.PerformLayout()
        Me.fraInformacionGeneral.ResumeLayout(False)
        Me.fraInformacionGeneral.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        CType(Me.lblCuentasBancarias, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optTipoCuenta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub


    'Function GetCtaBancaria() As String
    '    Dim Clave As String
    '    Clave = ""
    '    Clave = Trim(Clave) & Trim(Me.txtCtaBancaria.Text)
    '    GetCtaBancaria = Clave
    '    Return GetCtaBancaria
    'End Function

    Function BuscaCuenta() As Boolean
        'gStrSql = "select ctaBancaria from CatCuentasBancarias where LTRIM(RTRIM(ctaBancaria)) = '" & Trim(Me.GetCtaBancaria()) & "'"
        gStrSql = "select ctaBancaria from CatCuentasBancarias where LTRIM(RTRIM(ctaBancaria)) = '" & Trim(Me.txtCtaBancaria.Text) & "'"



        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute

        If rsLocal.RecordCount > 0 Then
            BuscaCuenta = True
        Else
            BuscaCuenta = False
        End If
    End Function

    Function BuscaCuentayBanco() As Boolean
        'gStrSql = "select a.ctaBancaria , a.codBanco, b.descBanco from CatCuentasBancarias a, CatBancos b where LTRIM(RTRIM(a.ctaBancaria)) = '" & (GetCtaBancaria()) & "' and a.CodBanco = b.CodBanco and LTRIM(RTRIM(b.DescBanco)) = '" & Trim(Me.dbcBanco.Text) & "'"
        gStrSql = "select a.ctaBancaria , a.codBanco, b.descBanco from CatCuentasBancarias a, CatBancos b where LTRIM(RTRIM(a.ctaBancaria)) = '" & Trim(Me.txtCtaBancaria.Text) & "' and a.CodBanco = b.CodBanco and LTRIM(RTRIM(b.DescBanco)) = '" & Trim(Me.dbcBanco.Text) & "'"

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute

        If rsLocal.RecordCount > 0 Then
            BuscaCuentayBanco = True
        Else
            BuscaCuentayBanco = False
        End If
    End Function


    Function BuscaDato() As Boolean
        'Esta función sirve para ver si el dato actual es un dato nuevo
        'gStrSql = "select ctaBancaria from CatCuentasBancarias where LTRIM(RTRIM(ctaBancaria)) = '" & Trim(Me.GetCtaBancaria()) & "' and codBanco = " & mintCodBanco
        gStrSql = "select ctaBancaria from CatCuentasBancarias where LTRIM(RTRIM(ctaBancaria)) = '" & Trim(Me.txtCtaBancaria.Text) & "' and codBanco = " & mintCodBanco

        Trim(Me.txtCtaBancaria.Text)

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute

        If rsLocal.RecordCount > 0 Then
            BuscaDato = True
        Else
            BuscaDato = False
        End If
    End Function

    Sub LlenaCombo(ByRef cboParam As System.Windows.Forms.ComboBox)
        On Error GoTo MErr
        Dim lStrSql As String
        Dim I As Object
        Dim J As Integer

        For I = 0 To 25
            maLetra(I) = UCase(Chr(System.Windows.Forms.Keys.A + I))
        Next

        lStrSql = "SELECT LetraFolios FROM CatCuentasBancarias WHERE LetraFolios IS NOT NULL ORDER BY LetraFolios"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lStrSql))
        rsCombo = Cmd.Execute

        cLetras = ""
        If rsCombo.RecordCount > 0 Then
            rsCombo.MoveFirst()
            For I = 1 To rsCombo.RecordCount
                cLetras = cLetras & UCase(rsCombo.Fields("LetraFolios").Value)
                rsCombo.MoveNext()
            Next I
        End If

        Me.cboLetraFolios.Items.Clear()
        cLetrasEnCombo = ""
        For I = 0 To 25
            If InStr(1, cLetras, maLetra(I)) = 0 Then
                Me.cboLetraFolios.Items.Add(maLetra(I))
                cLetrasEnCombo = cLetrasEnCombo & maLetra(I)
            End If
        Next I
MErr:
        If Err.Number <> 0 Then ModErrores.Errores()
        '''Aviso:=gstrNombCortoEmpresa
    End Sub


    Public Function llenaBancos(ByRef cboParam As System.Windows.Forms.ComboBox)

        On Error GoTo MErr
        Dim lStrSql As String
        'Dim I As Object
        'Dim J As Integer

        'For I = 0 To 25
        '    maLetra(I) = UCase(Chr(System.Windows.Forms.Keys.A + I))
        'Next

        lStrSql = "SELECT codBanco, descBanco = rtrim(ltrim(descBanco)) FROM catBancos ORDER BY DescBanco"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lStrSql))
        rsCombo = Cmd.Execute

        mintCodBanco = ""
        If rsCombo.RecordCount > 0 Then
            rsCombo.MoveFirst()
            For I = 1 To rsCombo.RecordCount
                mintCodBanco = mintCodBanco & UCase(rsCombo.Fields("DescBanco").Value)
                rsCombo.MoveNext()
            Next I
        End If

        Me.dbcBanco.Items.Clear()
        'rsCombo = ""
        'For I = 0 To 25
        '    If InStr(1, rsCombo, maLetra(I)) = 0 Then
        '        Me.dbcBanco.Items.Add(maLetra(I))
        '        'cLetrasEnCombo = cLetrasEnCombo & maLetra(I)
        '    End If
        'Next I
MErr:
        If Err.Number <> 0 Then ModErrores.Errores()

        'gStrSql = "SELECT codBanco, descBanco = rtrim(ltrim(descBanco)) FROM catBancos ORDER BY DescBanco"
        'ModDCombo.DCGotFocus(gStrSql, dbcBanco)
    End Function




    Public Sub BuscaBanco()
        Dim NomBanco As String

        gStrSql = "Select codBanco, DescBanco From CatBancos Where CodBanco = " & mintCodBanco

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        rsLocal = Cmd.Execute

        If rsLocal.RecordCount > 0 Then
            mintCodBanco = rsLocal.Fields("CodBanco").Value
            mblnFueraChange = True
            Me.dbcBanco.Text = Trim(rsLocal.Fields("DescBanco").Value)
            Me.dbcBanco.Tag = Me.dbcBanco.Text
            mblnFueraChange = False
        End If
    End Sub


    Sub Buscar()
        On Error GoTo MErr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendrá el string del tag que se le mandara al fromulario de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas


        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        'strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
        strTag = UCase("FRMCORPOABCCUENTASBANCARIAS" & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        strCaptionForm = "Consulta de Cuentas Bancarias"
        Select Case strControlActual
            Case "TXTCTABANCARIA"
                If mintCodBanco = 0 Then
                    gStrSql = "SELECT a.CtaBancaria as CUENTA, b.DescBanco as BANCO, a.CuentaHabiente AS CUENTAHABIENTE, b.CodBanco as CODIGOBANCO FROM CatCuentasBancarias a, CatBancos b WHERE a.CodBanco = b.CodBanco ORDER BY a.CtaBancaria, b.DescBanco, a.CuentaHabiente"
                Else
                    gStrSql = "SELECT a.CtaBancaria as CUENTA, b.DescBanco as BANCO, a.CuentaHabiente AS CUENTAHABIENTE, b.CodBanco as CODIGOBANCO FROM CatCuentasBancarias a, CatBancos b WHERE a.CodBanco = b.CodBanco and a.CodBanco = " & mintCodBanco & " ORDER BY a.CtaBancaria, b.DescBanco, a.CuentaHabiente"
                End If
            Case "TXTCUENTAHABIENTE"
                If mintCodBanco = 0 Then
                    gStrSql = "SELECT a.CtaBancaria as CUENTA, b.DescBanco as BANCO, a.CuentaHabiente AS CUENTAHABIENTE, b.CodBanco as CODIGOBANCO FROM CatCuentasBancarias a, CatBancos b WHERE a.CodBanco = b.CodBanco and a.CuentaHabiente LIKE '" & Trim(Me.txtCuentaHabiente.Text) & "%' ORDER BY a.CuentaHabiente"
                Else
                    gStrSql = "SELECT a.CtaBancaria as CUENTA, b.DescBanco as BANCO, a.CuentaHabiente AS CUENTAHABIENTE, b.CodBanco as CODIGOBANCO FROM CatCuentasBancarias a, CatBancos b WHERE a.CodBanco = b.CodBanco and a.CodBanco = " & mintCodBanco & " and a.CuentaHabiente LIKE '" & Trim(Me.txtCuentaHabiente.Text) & "%' ORDER BY a.CuentaHabiente"
                End If
            Case "TXTSUCURSAL"
                If mintCodBanco = 0 Then
                    gStrSql = "SELECT b.DescBanco as BANCO, a.Sucursal as SUCURSAL, a.CtaBancaria as CUENTA, a.CuentaHabiente AS CUENTAHABIENTE, b.CodBanco as CODIGOBANCO FROM CatCuentasBancarias a, CatBancos b WHERE a.CodBanco = b.CodBanco and a.Sucursal LIKE '" & Trim(Me.txtSucursal.Text) & "%' ORDER BY b.DescBanco, a.Sucursal, a.CuentaHabiente"
                Else
                    gStrSql = "SELECT b.DescBanco as BANCO, a.Sucursal as SUCURSAL, a.CtaBancaria as CUENTA, a.CuentaHabiente AS CUENTAHABIENTE, b.CodBanco as CODIGOBANCO FROM CatCuentasBancarias a, CatBancos b WHERE a.CodBanco = b.CodBanco and a.CodBanco = " & mintCodBanco & " and a.Sucursal LIKE '" & Trim(Me.txtSucursal.Text) & "%' ORDER BY b.DescBanco, a.Sucursal, a.CuentaHabiente"
                End If
            Case Else
                'Sale de este sub para que no ejecute ninguna opción
                Exit Sub
        End Select

        strSQL = gStrSql 'Se hace uso de una variable temporal para el query

        'Si hubo cambios y es una modificacion entonces preguntará si desea grabar los cambios
        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se cargue la consulta
                    mblnNuevo = True
                Case MsgBoxResult.Cancel 'Cancela la consulta
                    Exit Sub
            End Select
        End If

        gStrSql = strSQL 'Se regresa el valor de la variable temporal a la variable original

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsjNoExiste(C_msgSINDATOS, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta
        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 9500, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCTABANCARIA"
                    'ConfiguraConsultas(FrmConsultas, 8500, RsGral, strTag, strCaptionForm)
                    .ScrollBars = MSHierarchicalFlexGridLib.ScrollBarsSettings.flexScrollBarVertical
                    .set_ColWidth(0, 0, 2000) 'Cuenta
                    .set_ColWidth(1, 0, 2000) 'Banco
                    .set_ColWidth(2, 0, 4500) 'Cuentahabiente
                    .set_ColWidth(3, 0, 1000) 'CodigoBanco
                Case "TXTCUENTAHABIENTE"
                    .ScrollBars = MSHierarchicalFlexGridLib.ScrollBarsSettings.flexScrollBarVertical
                    'ConfiguraConsultas(FrmConsultas, 8500, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 2000) 'Cuenta
                    .set_ColWidth(1, 0, 2000) 'Banco
                    .set_ColWidth(2, 0, 4500) 'Cuentahabiente
                    .set_ColWidth(3, 0, 1000) 'CodigoBanco
                Case "TXTSUCURSAL"
                    'ConfiguraConsultas(FrmConsultas, 9500, RsGral, strTag, strCaptionForm)
                    .ScrollBars = MSHierarchicalFlexGridLib.ScrollBarsSettings.flexScrollBarVertical
                    .set_ColWidth(0, 0, 2000) 'Banco
                    .set_ColWidth(1, 0, 1200) 'Sucursal
                    .set_ColWidth(2, 0, 2000) 'Cuenta
                    .set_ColWidth(3, 0, 4300) 'Cuentahabiente
                    .set_ColWidth(4, 0, 1) '1000 'CodigoBanco
            End Select
        End With
        FrmConsultas.ShowDialog()
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub LlenaDatos()
        On Error GoTo MErr
        'If Me.GetCtaBancaria() = "" Then
        If Trim(Me.txtCtaBancaria.Text) = "" Then
            'Limpiar
            ModEstandar.AvanzarTab(Me)
            Exit Sub
        End If
        'Me.txtCtaBancaria.text = Format(Me.txtCtaBancaria.text, "0000000000000000")

        'Me.txtCtaBancaria.Text = Format(String.Concat(Me.txtCtaBancaria.Text, "0000000000000000"))

        For i = 1 To 4 - (txtCtaBancaria.TextLength)
            txtCtaBancaria.Text = String.Concat("0" + txtCtaBancaria.Text)
        Next i

        'gStrSql = "select * from CatCuentasBancarias where CtaBancaria  ='" & Me.GetCtaBancaria() & "' and CodBanco = " & mintCodBanco
        gStrSql = "select * from CatCuentasBancarias where CtaBancaria  ='" & Trim(Me.txtCtaBancaria.Text) & "' and CodBanco = " & mintCodBanco


        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            mintCodBanco = RsGral.Fields("CodBanco").Value
            Call BuscaBanco()
            If RsGral.Fields("TipoCuenta").Value = C_NORMAL Then
                Me._optTipoCuenta_0.Checked = True
                Me._optTipoCuenta_1.Checked = False
                mstrTipoCuenta = C_NORMAL
                mstrTipoCuentaTag = C_NORMAL
            ElseIf RsGral.Fields("TipoCuenta").Value = C_INVERSION Then
                Me._optTipoCuenta_0.Checked = False
                Me._optTipoCuenta_1.Checked = True
                mstrTipoCuenta = C_INVERSION
                mstrTipoCuentaTag = C_INVERSION
            End If

            ' Me.txtSucursal.Text = RsGral.Fields("Sucursal").Value()
            'For i = 1 To 4 - (txtSucursal.TextLength)
            '    txtSucursal.Text = String.Concat("0" + txtSucursal.Text)
            'Next i

            Me.txtSucursal.Text = Format(RsGral.Fields("Sucursal").Value, "0000")
            Me.txtSucursal.Tag = Me.txtSucursal.Text

            'Me.txtSucursal.Text = Format(RsGral.Fields("Sucursal").Value, "0000")
            'Me.txtSucursal.Tag = Me.txtSucursal.Text
            'Me.txtSucursal.Tag = Format(RsGral.Fields("Sucursal").Value, "0000")


            For i = 1 To 4 - (txtSucursal.TextLength)
                txtSucursal.Text = String.Concat("0" + txtSucursal.Text)
            Next i

            Me.txtCuentaHabiente.Text = Trim(RsGral.Fields("CuentaHabiente").Value)
            Me.txtCuentaHabiente.Tag = Trim(RsGral.Fields("CuentaHabiente").Value)
            Me.cboLetraFolios.Text = Trim(RsGral.Fields("LetraFolios").Value)
            Me.cboLetraFolios.Tag = Me.cboLetraFolios.Text

            cLetrasEnCombo = cLetrasEnCombo & UCase(Me.cboLetraFolios.Text)
            Me.cboLetraFolios.Items.Add(Me.cboLetraFolios.Text)

            Me.txtSaldoInicial.Text = Format(RsGral.Fields("SaldoInicial").Value, "0.00")
            Me.txtSaldoInicial.Tag = Me.txtSaldoInicial.Text

            For i = 1 To 2 - (txtSaldoInicial.TextLength)
                txtSaldoInicial.Text = String.Concat("0" + txtSaldoInicial.Text)
            Next i


            'Me.txtConsecutivodeChq.Text = Format(RsGral.Fields("ConsecutivoChq").Value, "000000")
            'Me.txtConsecutivodeChq.Tag = Me.txtConsecutivodeChq.Text
            'Me.txtConsecutivodeChq.Tag = Format(RsGral.Fields("ConsecutivoChq").Value, "000000")


            Me.txtConsecutivodeChq.Text = Format(RsGral.Fields("ConsecutivoChq").Value, "000000")
            Me.txtConsecutivodeChq.Tag = Me.txtConsecutivodeChq.Text

            For i = 1 To 6 - (txtConsecutivodeChq.TextLength)
                txtConsecutivodeChq.Text = String.Concat("0" + txtConsecutivodeChq.Text)
            Next i



            If RsGral.Fields("Moneda").Value = C_PESO Then
                _optMoneda_0.Checked = True
                blnPesos = True
                blnDolares = False
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                _optMoneda_1.Checked = True
                blnDolares = True
                blnPesos = False
            End If
        Else
            If Not BuscaDato() Then
                MsgBox("No existe la Cuenta Bancaria en el Banco indicado", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            End If
            Limpiar()
        End If
        mblnCambiosEnCodigo1 = False
        mblnCambiosEnCodigo2 = False
        mblnNuevo = False
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()

    End Sub

    Public Sub Eliminar()
        On Error GoTo MErr
        Dim I As Integer
        Dim blntransaction As Boolean
        'gStrSql = "select ctaBancaria from CatCuentasBancarias where LTRIM(RTRIM(ctaBancaria)) = '" & Trim(Me.GetCtaBancaria()) & "' and codBanco = " & mintCodBanco
        gStrSql = "select ctaBancaria from CatCuentasBancarias where LTRIM(RTRIM(ctaBancaria)) = '" & Trim(Me.txtCtaBancaria.Text) & "' and codBanco = " & mintCodBanco
        'gStrSql = "select ctaBancaria from CatCuentasBancarias where LTRIM(RTRIM(ctaBancaria)) = '" & Trim(Me.txtCtaBancaria.Text) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            MsgBox("Proporcione una Cuenta Bancaria y un Banco existentes", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If
        'Preguntar si desea borrar el registro
        If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) = MsgBoxResult.No Then
            Exit Sub
        End If
        Cnn.BeginTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        blntransaction = True
        'ModStoredProcedures.PR_IMECatCuentasBancarias(Str(mintCodBanco), Trim(Me.GetCtaBancaria), Trim(mstrTipoCuenta), Trim(Me.txtSucursal.Text), Trim(Me.txtCuentaHabiente.Text), Trim(Me.cboLetraFolios.Text), Trim(Me.txtSaldoInicial.Text), Trim(Me.txtConsecutivodeChq.Text), IIf(_optMoneda_0.Checked, C_PESO, C_DOLAR), C_ELIMINACION, CStr(0))
        ModStoredProcedures.PR_IMECatCuentasBancarias(Str(mintCodBanco), Trim(Me.txtCtaBancaria.Text), Trim(mstrTipoCuenta), Trim(Me.txtSucursal.Text), Trim(Me.txtCuentaHabiente.Text), Trim(Me.cboLetraFolios.Text), Trim(Me.txtSaldoInicial.Text), Trim(Me.txtConsecutivodeChq.Text), IIf(_optMoneda_0.Checked, C_PESO, C_DOLAR), C_ELIMINACION, CStr(0))

        Cmd.Execute()
        Cnn.CommitTrans()
        blntransaction = False
        For I = 0 To 3
            Me.txtCtaBancaria.Text = ""
        Next I
        mblnCambiosEnCodigo1 = False
        mblnCambiosEnCodigo2 = False
        mblnNuevo = False
        Limpiar()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
MErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            If blntransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Sub Nuevo()
        On Error GoTo MErr
        Dim I As Integer
        If Not mblnNuevo Then
            mblnFueraChange = True
            Me.dbcBanco.Text = ""
            Me.dbcBanco.Tag = ""
            mintCodBanco = 0
            mblnFueraChange = False
        End If
        If mblnCambiosEnCodigo2 Then
            For I = 0 To 3
                Me.txtCtaBancaria.Text = ""
            Next I
            mblnCambiosEnCodigo2 = False
        End If

        txtCtaBancaria.Text = ""
        Me._optTipoCuenta_0.Checked = True
        Me._optTipoCuenta_1.Checked = False
        mstrTipoCuenta = C_NORMAL
        mstrTipoCuentaTag = C_NORMAL
        Me.txtSucursal.Text = ""
        Me.txtSucursal.Tag = ""
        Me.txtCuentaHabiente.Text = ""
        Me.txtCuentaHabiente.Tag = ""
        Me.cboLetraFolios.Text = ""
        Me.cboLetraFolios.Tag = ""
        Me.txtSaldoInicial.Text = "0.00"
        Me.txtSaldoInicial.Tag = "0.00"
        Me.txtConsecutivodeChq.Text = "0"
        Me.txtConsecutivodeChq.Tag = "0"
        _optMoneda_0.Checked = True
        _optMoneda_1.Checked = False
        blnPesos = True
        blnDolares = False
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function Cambios() As Boolean
        Select Case True
            Case mstrTipoCuenta <> mstrTipoCuentaTag
                Cambios = True
            Case Trim(Me.txtSucursal.Text) <> Trim(Me.txtSucursal.Tag)
                Cambios = True
            Case Trim(Me.txtCuentaHabiente.Text) <> Trim(Me.txtCuentaHabiente.Tag)
                Cambios = True
            Case Trim(Me.cboLetraFolios.Text) <> Trim(Me.cboLetraFolios.Tag)
                Cambios = True
            Case ModEstandar.Numerico((Me.txtSaldoInicial.Text)) <> ModEstandar.Numerico((Me.txtSaldoInicial.Tag))
                Cambios = True
            Case ModEstandar.Numerico((Me.txtConsecutivodeChq.Text)) <> ModEstandar.Numerico((Me.txtConsecutivodeChq.Tag))
                Cambios = True
            Case _optMoneda_0.Checked <> blnPesos
                Cambios = True
            Case _optMoneda_1.Checked <> blnDolares
                Cambios = True
            Case Else
                Cambios = False
        End Select
    End Function

    Function ChecaBancoCuenta() As Boolean
        Dim Moneda As String
        On Error GoTo Err_Renamed
        ChecaBancoCuenta = False
        gStrSql = "SELECT * FROM CatBancos WHERE CodBanco = " & mintCodBanco & " AND ControlInterno = 1 AND Sucursal = 0"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 1 Then
            Moneda = IIf(_optMoneda_0.Checked, C_PESO, C_DOLAR)
            gStrSql = "SELECT * FROM CatCuentasBancarias WHERE CodBanco = " & mintCodBanco & " AND Moneda = '" & Moneda & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount = 1 Then
                MsgBox("¡¡¡ATENCION!!! Este Banco ya Tiene una Cuenta en " & IIf(_optMoneda_0.Checked, "Pesos", "Dolares") & Chr(13) & "                             Favor de Verificar", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                ChecaBancoCuenta = True
            End If
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Function ValidaDatos() As Boolean
        Select Case True
            Case mintCodBanco = 0
                MsgBox(C_msgFALTADATO & "El Banco al que pertenece la cuenta", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Me.dbcBanco.Focus()
                ValidaDatos = False
            Case ChecaBancoCuenta()
                ValidaDatos = False
                If _optMoneda_0.Checked Then
                    _optMoneda_0.Focus()
                Else
                    _optMoneda_1.Focus()
                End If
            'Case Trim(Me.GetCtaBancaria()) = ""
            Case Trim(Trim(Me.txtCtaBancaria.Text) = "")
                MsgBox(C_msgFALTADATO & "El número de Cuenta Bancaria", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Me.txtCtaBancaria.Focus()
                ValidaDatos = False
                    '        Case Len(Me.GetCtaBancaria()) < 16
                    '            MsgBox "El número de Cuenta debe ser de 16 dígitos", vbInformation + vbOKOnly, gstrNombCortoEmpresa
                    '            Me.txtCtaBancaria.SetFocus
                    '            ValidaDatos = False
            Case mstrTipoCuenta = ""
                If Me._optTipoCuenta_0.Checked = True Then
                    mstrTipoCuenta = C_NORMAL
                Else
                    mstrTipoCuenta = C_INVERSION
                End If
            Case Len(Trim(Me.txtSucursal.Text)) = 0
                MsgBox(C_msgFALTADATO & "La Sucursal del Banco", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Me.txtSucursal.Focus()
                ValidaDatos = False
            Case Len(Trim(Me.txtCuentaHabiente.Text)) = 0
                MsgBox(C_msgFALTADATO & "El Nombre del Cuentahabiente", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Me.txtCuentaHabiente.Focus()
                ValidaDatos = False
            Case InStr(1, cLetrasEnCombo, Me.cboLetraFolios.Text) = 0 Or Trim(Me.cboLetraFolios.Text) = ""
                MsgBox("Introduzca una letra de las existentes en el combo", MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.cboLetraFolios.Text = ""
                Me.cboLetraFolios.Focus()
                ValidaDatos = False
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Public Function Guardar() As Boolean
        On Error GoTo MErr
        Dim blntransaction As Boolean
        Dim I As Integer
        'Valida si todos los datos han sido llenados correctamnte para poder ser guardados
        If Not ValidaDatos() Then
            Exit Function
        End If
        If Not Cambios() Then
            For I = 0 To 3
                Me.txtCtaBancaria.Text = ""
            Next I
            mblnNuevo = False
            Limpiar()
            mblnCambiosEnCodigo1 = False
            Me.dbcBanco.Focus()
            Exit Function
        End If
        If BuscaDato() Then
            mblnNuevo = False
        Else
            mblnNuevo = True
        End If
        'INICIA TRANSACCIÓN--------------------------------------
        Cnn.BeginTrans()
        blntransaction = True
        If mblnNuevo Then
            'ModStoredProcedures.PR_IMECatCuentasBancarias(Str(mintCodBanco), Trim(Me.GetCtaBancaria), Trim(mstrTipoCuenta), Trim(Me.txtSucursal.Text), Trim(Me.txtCuentaHabiente.Text), Trim(Me.cboLetraFolios.Text), Trim(Me.txtSaldoInicial.Text), Trim(Me.txtConsecutivodeChq.Text), IIf(_optMoneda_0.Checked, C_PESO, C_DOLAR), C_INSERCION, CStr(0))
            ModStoredProcedures.PR_IMECatCuentasBancarias(Str(mintCodBanco), Trim(Me.txtCtaBancaria.Text), Trim(mstrTipoCuenta), Trim(Me.txtSucursal.Text), Trim(Me.txtCuentaHabiente.Text), Trim(Me.cboLetraFolios.Text), Trim(Me.txtSaldoInicial.Text), Trim(Me.txtConsecutivodeChq.Text), IIf(_optMoneda_0.Checked, C_PESO, C_DOLAR), C_INSERCION, CStr(0))
            Cmd.Execute()
        Else
            'ModStoredProcedures.PR_IMECatCuentasBancarias(Str(mintCodBanco), Trim(Me.GetCtaBancaria), Trim(mstrTipoCuenta), Trim(Me.txtSucursal.Text), Trim(Me.txtCuentaHabiente.Text), Trim(Me.cboLetraFolios.Text), Trim(Me.txtSaldoInicial.Text), Trim(Me.txtConsecutivodeChq.Text), IIf(_optMoneda_0.Checked, C_PESO, C_DOLAR), C_MODIFICACION, CStr(0))
            ModStoredProcedures.PR_IMECatCuentasBancarias(Str(mintCodBanco), Trim(Me.txtCtaBancaria.Text), Trim(mstrTipoCuenta), Trim(Me.txtSucursal.Text), Trim(Me.txtCuentaHabiente.Text), Trim(Me.cboLetraFolios.Text), Trim(Me.txtSaldoInicial.Text), Trim(Me.txtConsecutivodeChq.Text), IIf(_optMoneda_0.Checked, C_PESO, C_DOLAR), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        Cnn.CommitTrans()
        blntransaction = False
        'TERMINA LA TRANSACCIÓN----------------------------------
        If mblnNuevo Then
            MsgBox("La Cuenta Bancaria ha sido grabada correctamente.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        End If
        mblnNuevo = False
        Nuevo()
        Guardar = True
        For I = 0 To 3
            Me.txtCtaBancaria.Text = ""
        Next I
        mblnCambiosEnCodigo1 = True
        mblnCambiosEnCodigo2 = True
        Limpiar()
        mblnCambiosEnCodigo1 = False
        mblnCambiosEnCodigo2 = False
        Me.dbcBanco.Focus()
MErr:
        If Err.Number <> 0 Then
            If blntransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub Limpiar()
        On Error Resume Next
        'Validar si hubo cambios que desee guardar
        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                    mblnNuevo = True
                Case MsgBoxResult.Cancel 'Cancela la acción de limpiar pantalla
                    Exit Sub
            End Select
        End If
        LlenaCombo(cboLetraFolios)
        mblnCambiosEnCodigo1 = True
        mblnCambiosEnCodigo2 = True
        Nuevo()
        mblnNuevo = True
        mblnCambiosEnCodigo1 = False
        mblnCambiosEnCodigo2 = False
        Me.dbcBanco.Focus()
    End Sub

    Private Sub cboLetraFolios_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cboLetraFolios.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        KeyCode = 0
    End Sub

    Private Sub cboLetraFolios_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboLetraFolios.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub cboLetraFolios_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLetraFolios.Leave
        Pon_Tool()
    End Sub

    Private Sub dbcBanco_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codBanco, descBanco = rtrim(ltrim(descBanco)) FROM catBancos Where descBanco LIKE '" & Trim(Me.dbcBanco.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcBanco)

        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox("¿Desea guardar los cambios?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    If Guardar() Then
                    End If
                    Call dbcBanco_Enter(dbcBanco, New System.EventArgs())
                Case MsgBoxResult.No
                    mblnNuevo = True
                    Limpiar()
                Case MsgBoxResult.Cancel
            End Select
        End If

        If Me.dbcBanco.Text = "" Then
            mblnCambiosEnCodigo2 = True
            Nuevo()
            dbcBanco_Enter(dbcBanco, New System.EventArgs())
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Enter
        Pon_Tool()
        gStrSql = " SELECT codBanco, descBanco = rtrim(ltrim(descBanco)) FROM catBancos ORDER BY CodBanco"
        ModDCombo.DCGotFocus(gStrSql, dbcBanco)
    End Sub



    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Leave
        Dim I As Integer
        Dim CodAnterior As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codBanco, descBanco = rtrim(ltrim(descBanco)) FROM catBancos Where descBanco Like '" & Trim(Me.dbcBanco.Text) & "%'"
        CodAnterior = mintCodBanco
        mintCodBanco = 0
        ModDCombo.DCLostFocus(dbcBanco, gStrSql, mintCodBanco)
        If mintCodBanco <> CodAnterior Then
            mblnCambiosEnCodigo2 = True
            mblnNuevo = True
            Nuevo()
        Else
            mblnCambiosEnCodigo2 = False
        End If
    End Sub

    Private Sub frmCorpoABCCuentasBancarias_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoABCCuentasBancarias_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCorpoABCCuentasBancarias_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                KeyCode = 0

                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If Trim(UCase(Me.ActiveControl.Name)) = "DBCBANCO" Then
                    'If Me.ActiveControl.Index = 0 Then
                    '    mblnSALIR = True
                    '    Me.Close()
                    'Else
                    '    ModEstandar.RetrocederTab(Me)
                    'End If
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCorpoABCCuentasBancarias_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayúscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoABCCuentasBancarias_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        LlenaCombo((Me.cboLetraFolios))
        mblnNuevo = True
        mblnCambiosEnCodigo1 = False
        mblnCambiosEnCodigo2 = False
        blnPesos = True
        blnDolares = False
    End Sub

    Private Sub frmCorpoABCCuentasBancarias_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si desea cerrar la forma y esta se encuentra minimizada, esta se restaura
        'If Not mblnSALIR Then
        '    ModEstandar.RestaurarForma(Me, False)
        '    If Cambios() And Not (mblnNuevo) Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes
        '                If Not (Guardar()) Then
        '                    Cancel = 1
        '                End If
        '            Case MsgBoxResult.No 'No hace nada y permite que se cierre el formulario
        '                mblnNuevo = True
        '                Cancel = 0
        '            Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin Guardar
        '                Cancel = 1
        '        End Select
        '    End If
        'Else 'Se quiere salir con escape
        '    mblnSALIR = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Me.dbcBanco2.Focus()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoABCCuentasBancarias_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
    End Sub

    Private Sub optMoneda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoneda.Enter
        Dim Index As Integer = optMoneda.GetIndex(eventSender)
        Select Case Index
            Case 0
                Pon_Tool()
            Case 1
                Pon_Tool()
        End Select
    End Sub

    Private Sub optTipoCuenta_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTipoCuenta.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = optTipoCuenta.GetIndex(eventSender)
            Select Case Index
                Case 0
                    mstrTipoCuenta = C_NORMAL
                Case 1
                    mstrTipoCuenta = C_INVERSION
            End Select
        End If
    End Sub

    Private Sub txtConsecutivodeChq_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConsecutivodeChq.Enter
        SelTextoTxt((Me.txtConsecutivodeChq))
        Pon_Tool()
    End Sub

    Private Sub txtConsecutivodeChq_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtConsecutivodeChq.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                'Me.txtConsecutivodeChq.Text = Format(Numerico((Me.txtConsecutivodeChq.Text)), "000000")
                Me.txtConsecutivodeChq.Text = Format(Numerico(String.Concat(Me.txtConsecutivodeChq.Text, "000000")))
        End Select
    End Sub

    Private Sub txtConsecutivodeChq_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConsecutivodeChq.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 13
                'Me.txtConsecutivodeChq.Text = Format(Me.txtConsecutivodeChq.Text, "000000")
                Me.txtConsecutivodeChq.Text = Format(Numerico(String.Concat(Me.txtConsecutivodeChq.Text, "000000")))
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtConsecutivodeChq_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConsecutivodeChq.Leave
        'Me.txtConsecutivodeChq.Text = Format(Numerico((Me.txtConsecutivodeChq.Text)), "000000")
        Me.txtConsecutivodeChq.Text = Format(Numerico(String.Concat(Me.txtConsecutivodeChq.Text, "000000")))
    End Sub

    Private Sub txtCtaBancaria_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCtaBancaria.TextChanged
        'If Len(Me.txtCtaBancaria(Index).text) = 4 Then
        '    ModEstandar.AvanzarTab Me
        'End If
        'Or Len(Me.GetCtaBancaria()) < 16
        'If Me.GetCtaBancaria() = "" Then
        If Trim(Me.txtCtaBancaria.Text) = "" Then
            mblnNuevo = True
            Nuevo()
        Else
            If BuscaCuentayBanco() Then
                LlenaDatos()
                '            If Me._optTipoCuenta_0.Value Then
                '                Me._optTipoCuenta_0.SetFocus
                '            Else
                '                Me._optTipoCuenta_1.SetFocus
                '            End If
            End If
        End If
        mblnCambiosEnCodigo1 = True
    End Sub

    Private Sub txtCtaBancaria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCtaBancaria.Enter
        strControlActual = UCase("txtCtaBancaria")
        SelTextoTxt(txtCtaBancaria, 1, 1)
        Pon_Tool()
        'If Me.GetCtaBancaria() = "" Then
        '    mblnNuevo = True
        '    Nuevo()
        'Else
        '    If BuscaCuentayBanco() Then
        '        LlenaDatos()
        '        '            If Me._optTipoCuenta_0.Value Then
        '        '                Me._optTipoCuenta_0.SetFocus
        '        '            Else
        '        '                Me._optTipoCuenta_1.SetFocus
        '        '            End If
        '    End If
        'End If
        'mblnCambiosEnCodigo1 = True
    End Sub

    Private Sub txtCtaBancaria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCtaBancaria.KeyDown
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
                    mblnNuevo = True
                Case MsgBoxResult.Cancel
                    KeyCode = 0
                    Me.txtCtaBancaria.Focus()
            End Select
        End If
    End Sub

    Private Sub txtCtaBancaria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCtaBancaria.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back Then
            KeyAscii = 0
        Else
            'Pregunta sólo si ha habido cambios
            If Cambios() And Not mblnNuevo Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        If Not Guardar() Then
                            KeyAscii = 0
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
                        mblnNuevo = True
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        KeyAscii = 0
                        Me.txtCtaBancaria.Focus()
                End Select
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCtaBancaria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCtaBancaria.Leave
        'Or Len(Me.GetCtaBancaria()) < 16
        'If ActiveControl.Text = Me.Text Then
        'If mblnCambiosEnCodigo1 = True And (Me.GetCtaBancaria() = "") Then
        '        'If Index = 3 Then
        '        Limpiar()
        '        'End If
        '    End If
        'End If
    End Sub

    Private Sub txtCuentaHabiente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCuentaHabiente.Enter
        strControlActual = UCase("txtCuentaHabiente")
        SelTextoTxt((Me.txtCuentaHabiente))
        Pon_Tool()
    End Sub

    Private Sub txtSaldoInicial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSaldoInicial.Enter
        SelTextoTxt((Me.txtSaldoInicial))
        Pon_Tool()
    End Sub

    Private Sub txtSaldoInicial_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSaldoInicial.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                'Me.txtSaldoInicial.Text = Format(Numerico((Me.txtSaldoInicial.Text)), "###,###,##0.00")
                Me.txtSaldoInicial.Text = Format(Numerico(String.Concat(Me.txtSaldoInicial.Text, ",0.00")))
        End Select
    End Sub

    Private Sub txtSaldoInicial_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSaldoInicial.Leave
        'Me.txtSaldoInicial.Text = Format(Numerico((Me.txtSaldoInicial.Text)), "###,###,##0.00")
        Me.txtSaldoInicial.Text = Format(Numerico(String.Concat(Me.txtSaldoInicial.Text, ",0.00")))
    End Sub

    Private Sub txtSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSucursal.Enter
        strControlActual = UCase("txtSucursal")
        SelTextoTxt(Me.txtSucursal, 1, 1)
        Pon_Tool()
    End Sub

    Private Sub txtSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSucursal.Leave
        Me.txtSucursal.Text = Format(String.Concat(Me.txtSucursal.Text, "0000"))
    End Sub
    Private Sub dbcBanco_KeyDown(sender As Object, e As KeyEventArgs) Handles dbcBanco.KeyDown
        'Select Case sender
        '    Case e.KeyValue
        '        mblnSALIR = True
        '        Me.Close()
        '        sender.KeyCode = 0
        'End Select
        'tecla = sender.KeyCode
    End Sub

    Private Sub dbcBanco_KeyPress(sender As Object, e As KeyPressEventArgs) Handles dbcBanco.KeyPress

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        Eliminar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub
End Class
