Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasRptReparaciones_Corpo
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''' ********************************************************************************************************************
    ''' MODIFICACION DE LA FUNCION DE ESTATUS DE REPARACIONES - SE AGREGARON PARAMETROS
    ''' 15SEP2006 - MAVF
    ''' ********************************************************************************************************************

    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkxEnviarPV As System.Windows.Forms.RadioButton
    Public WithEvents chkxEnviarTaller As System.Windows.Forms.RadioButton
    Public WithEvents chkPorTrasnferir As System.Windows.Forms.CheckBox
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents chkEntregados As System.Windows.Forms.CheckBox
    Public WithEvents chkVigente As System.Windows.Forms.CheckBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents optPorSucursal As System.Windows.Forms.RadioButton
    Public WithEvents optPorTaller As System.Windows.Forms.RadioButton
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents fraMoneda As System.Windows.Forms.GroupBox
    Public WithEvents chkTaller As System.Windows.Forms.CheckBox
    Public WithEvents dbcTaller As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_3 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_2 As System.Windows.Forms.GroupBox
    Public WithEvents chkTodas As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_0 As System.Windows.Forms.GroupBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_2 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_1 As System.Windows.Forms.GroupBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todas ... ]"
    Dim FueraChange As Boolean
    Dim mintCodSucursal As Integer
    Dim mintCodTaller As Integer
    Dim mblnSalir As Boolean
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean
    Dim tecla As Integer
    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim mblnxTransferir As Boolean


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.chkxEnviarPV = New System.Windows.Forms.RadioButton()
        Me.chkxEnviarTaller = New System.Windows.Forms.RadioButton()
        Me.chkPorTrasnferir = New System.Windows.Forms.CheckBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.chkEntregados = New System.Windows.Forms.CheckBox()
        Me.chkVigente = New System.Windows.Forms.CheckBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.optPorSucursal = New System.Windows.Forms.RadioButton()
        Me.optPorTaller = New System.Windows.Forms.RadioButton()
        Me.fraMoneda = New System.Windows.Forms.GroupBox()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me._fraVtas_2 = New System.Windows.Forms.GroupBox()
        Me.chkTaller = New System.Windows.Forms.CheckBox()
        Me.dbcTaller = New System.Windows.Forms.ComboBox()
        Me._lblVentas_3 = New System.Windows.Forms.Label()
        Me._fraVtas_0 = New System.Windows.Forms.GroupBox()
        Me.chkTodas = New System.Windows.Forms.CheckBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me._fraVtas_1 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.fraVtas = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.fraMoneda.SuspendLayout()
        Me._fraVtas_2.SuspendLayout()
        Me._fraVtas_0.SuspendLayout()
        Me._fraVtas_1.SuspendLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(12, 379)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(349, 67)
        Me.txtMensaje.TabIndex = 27
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.chkxEnviarPV)
        Me.Frame3.Controls.Add(Me.chkxEnviarTaller)
        Me.Frame3.Controls.Add(Me.chkPorTrasnferir)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(15, 288)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(253, 62)
        Me.Frame3.TabIndex = 22
        Me.Frame3.TabStop = False
        '
        'chkxEnviarPV
        '
        Me.chkxEnviarPV.BackColor = System.Drawing.SystemColors.Control
        Me.chkxEnviarPV.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkxEnviarPV.Enabled = False
        Me.chkxEnviarPV.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkxEnviarPV.Location = New System.Drawing.Point(41, 40)
        Me.chkxEnviarPV.Margin = New System.Windows.Forms.Padding(2)
        Me.chkxEnviarPV.Name = "chkxEnviarPV"
        Me.chkxEnviarPV.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkxEnviarPV.Size = New System.Drawing.Size(172, 17)
        Me.chkxEnviarPV.TabIndex = 25
        Me.chkxEnviarPV.TabStop = True
        Me.chkxEnviarPV.Text = "Por enviar al Punto de Venta"
        Me.chkxEnviarPV.UseVisualStyleBackColor = False
        '
        'chkxEnviarTaller
        '
        Me.chkxEnviarTaller.BackColor = System.Drawing.SystemColors.Control
        Me.chkxEnviarTaller.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkxEnviarTaller.Enabled = False
        Me.chkxEnviarTaller.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkxEnviarTaller.Location = New System.Drawing.Point(41, 20)
        Me.chkxEnviarTaller.Margin = New System.Windows.Forms.Padding(2)
        Me.chkxEnviarTaller.Name = "chkxEnviarTaller"
        Me.chkxEnviarTaller.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkxEnviarTaller.Size = New System.Drawing.Size(148, 21)
        Me.chkxEnviarTaller.TabIndex = 24
        Me.chkxEnviarTaller.TabStop = True
        Me.chkxEnviarTaller.Text = "Por enviar al Taller"
        Me.chkxEnviarTaller.UseVisualStyleBackColor = False
        '
        'chkPorTrasnferir
        '
        Me.chkPorTrasnferir.BackColor = System.Drawing.SystemColors.Control
        Me.chkPorTrasnferir.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPorTrasnferir.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkPorTrasnferir.Location = New System.Drawing.Point(8, 1)
        Me.chkPorTrasnferir.Margin = New System.Windows.Forms.Padding(2)
        Me.chkPorTrasnferir.Name = "chkPorTrasnferir"
        Me.chkPorTrasnferir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPorTrasnferir.Size = New System.Drawing.Size(92, 21)
        Me.chkPorTrasnferir.TabIndex = 23
        Me.chkPorTrasnferir.Text = "Por Transferir"
        Me.chkPorTrasnferir.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.chkEntregados)
        Me.Frame2.Controls.Add(Me.chkVigente)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(124, 204)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(116, 70)
        Me.Frame2.TabIndex = 16
        Me.Frame2.TabStop = False
        Me.Frame2.Text = " Estatus "
        '
        'chkEntregados
        '
        Me.chkEntregados.BackColor = System.Drawing.SystemColors.Control
        Me.chkEntregados.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEntregados.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEntregados.Location = New System.Drawing.Point(10, 41)
        Me.chkEntregados.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEntregados.Name = "chkEntregados"
        Me.chkEntregados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEntregados.Size = New System.Drawing.Size(93, 21)
        Me.chkEntregados.TabIndex = 18
        Me.chkEntregados.Text = "Entregados"
        Me.chkEntregados.UseVisualStyleBackColor = False
        '
        'chkVigente
        '
        Me.chkVigente.BackColor = System.Drawing.SystemColors.Control
        Me.chkVigente.Checked = True
        Me.chkVigente.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkVigente.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVigente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkVigente.Location = New System.Drawing.Point(10, 19)
        Me.chkVigente.Margin = New System.Windows.Forms.Padding(2)
        Me.chkVigente.Name = "chkVigente"
        Me.chkVigente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVigente.Size = New System.Drawing.Size(78, 18)
        Me.chkVigente.TabIndex = 17
        Me.chkVigente.Text = "Vigentes"
        Me.chkVigente.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optPorSucursal)
        Me.Frame1.Controls.Add(Me.optPorTaller)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(245, 204)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(101, 70)
        Me.Frame1.TabIndex = 19
        Me.Frame1.TabStop = False
        Me.Frame1.Text = " Agrupar por... "
        '
        'optPorSucursal
        '
        Me.optPorSucursal.BackColor = System.Drawing.SystemColors.Control
        Me.optPorSucursal.Checked = True
        Me.optPorSucursal.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPorSucursal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPorSucursal.Location = New System.Drawing.Point(9, 21)
        Me.optPorSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.optPorSucursal.Name = "optPorSucursal"
        Me.optPorSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPorSucursal.Size = New System.Drawing.Size(88, 18)
        Me.optPorSucursal.TabIndex = 20
        Me.optPorSucursal.TabStop = True
        Me.optPorSucursal.Text = "Por Sucursal"
        Me.optPorSucursal.UseVisualStyleBackColor = False
        '
        'optPorTaller
        '
        Me.optPorTaller.BackColor = System.Drawing.SystemColors.Control
        Me.optPorTaller.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPorTaller.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPorTaller.Location = New System.Drawing.Point(9, 41)
        Me.optPorTaller.Margin = New System.Windows.Forms.Padding(2)
        Me.optPorTaller.Name = "optPorTaller"
        Me.optPorTaller.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPorTaller.Size = New System.Drawing.Size(75, 17)
        Me.optPorTaller.TabIndex = 21
        Me.optPorTaller.TabStop = True
        Me.optPorTaller.Text = "Por Taller"
        Me.optPorTaller.UseVisualStyleBackColor = False
        '
        'fraMoneda
        '
        Me.fraMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.fraMoneda.Controls.Add(Me.optDolares)
        Me.fraMoneda.Controls.Add(Me.optPesos)
        Me.fraMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraMoneda.Location = New System.Drawing.Point(20, 204)
        Me.fraMoneda.Margin = New System.Windows.Forms.Padding(2)
        Me.fraMoneda.Name = "fraMoneda"
        Me.fraMoneda.Padding = New System.Windows.Forms.Padding(2)
        Me.fraMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMoneda.Size = New System.Drawing.Size(94, 70)
        Me.fraMoneda.TabIndex = 13
        Me.fraMoneda.TabStop = False
        Me.fraMoneda.Text = " Moneda "
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(9, 41)
        Me.optDolares.Margin = New System.Windows.Forms.Padding(2)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(72, 22)
        Me.optDolares.TabIndex = 15
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Dolares"
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Checked = True
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(9, 21)
        Me.optPesos.Margin = New System.Windows.Forms.Padding(2)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(60, 15)
        Me.optPesos.TabIndex = 14
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pesos"
        Me.optPesos.UseVisualStyleBackColor = False
        '
        '_fraVtas_2
        '
        Me._fraVtas_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_2.Controls.Add(Me.chkTaller)
        Me._fraVtas_2.Controls.Add(Me.dbcTaller)
        Me._fraVtas_2.Controls.Add(Me._lblVentas_3)
        Me._fraVtas_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_2.Location = New System.Drawing.Point(29, 137)
        Me._fraVtas_2.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_2.Name = "_fraVtas_2"
        Me._fraVtas_2.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_2.Size = New System.Drawing.Size(317, 46)
        Me._fraVtas_2.TabIndex = 9
        Me._fraVtas_2.TabStop = False
        '
        'chkTaller
        '
        Me.chkTaller.BackColor = System.Drawing.SystemColors.Control
        Me.chkTaller.Checked = True
        Me.chkTaller.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTaller.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTaller.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTaller.Location = New System.Drawing.Point(6, 0)
        Me.chkTaller.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTaller.Name = "chkTaller"
        Me.chkTaller.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTaller.Size = New System.Drawing.Size(111, 17)
        Me.chkTaller.TabIndex = 10
        Me.chkTaller.Text = "Todas los talleres"
        Me.chkTaller.UseVisualStyleBackColor = False
        '
        'dbcTaller
        '
        Me.dbcTaller.Location = New System.Drawing.Point(53, 17)
        Me.dbcTaller.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcTaller.Name = "dbcTaller"
        Me.dbcTaller.Size = New System.Drawing.Size(234, 21)
        Me.dbcTaller.TabIndex = 12
        '
        '_lblVentas_3
        '
        Me._lblVentas_3.AutoSize = True
        Me._lblVentas_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_3.Location = New System.Drawing.Point(12, 20)
        Me._lblVentas_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_3.Name = "_lblVentas_3"
        Me._lblVentas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_3.Size = New System.Drawing.Size(36, 13)
        Me._lblVentas_3.TabIndex = 11
        Me._lblVentas_3.Text = "Taller:"
        '
        '_fraVtas_0
        '
        Me._fraVtas_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_0.Controls.Add(Me.chkTodas)
        Me._fraVtas_0.Controls.Add(Me.dbcSucursal)
        Me._fraVtas_0.Controls.Add(Me._lblVentas_0)
        Me._fraVtas_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_0.Location = New System.Drawing.Point(6, 7)
        Me._fraVtas_0.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.Name = "_fraVtas_0"
        Me._fraVtas_0.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_0.Size = New System.Drawing.Size(298, 52)
        Me._fraVtas_0.TabIndex = 0
        Me._fraVtas_0.TabStop = False
        '
        'chkTodas
        '
        Me.chkTodas.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodas.Checked = True
        Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodas.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTodas.Location = New System.Drawing.Point(6, 0)
        Me.chkTodas.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodas.Name = "chkTodas"
        Me.chkTodas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodas.Size = New System.Drawing.Size(134, 18)
        Me.chkTodas.TabIndex = 1
        Me.chkTodas.Text = "Todas las sucursales"
        Me.chkTodas.UseVisualStyleBackColor = False
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(87, 20)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(188, 21)
        Me.dbcSucursal.TabIndex = 3
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_0.Location = New System.Drawing.Point(32, 23)
        Me._lblVentas_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(51, 13)
        Me._lblVentas_0.TabIndex = 2
        Me._lblVentas_0.Text = "Sucursal:"
        '
        '_fraVtas_1
        '
        Me._fraVtas_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_1.Controls.Add(Me.dtpDesde)
        Me._fraVtas_1.Controls.Add(Me.dtpHasta)
        Me._fraVtas_1.Controls.Add(Me._lblVentas_1)
        Me._fraVtas_1.Controls.Add(Me._lblVentas_2)
        Me._fraVtas_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_1.Location = New System.Drawing.Point(15, 76)
        Me._fraVtas_1.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.Name = "_fraVtas_1"
        Me._fraVtas_1.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_1.Size = New System.Drawing.Size(346, 46)
        Me._fraVtas_1.TabIndex = 4
        Me._fraVtas_1.TabStop = False
        Me._fraVtas_1.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(78, 19)
        Me.dtpDesde.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(95, 20)
        Me.dtpDesde.TabIndex = 6
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(239, 19)
        Me.dtpHasta.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(95, 20)
        Me.dtpHasta.TabIndex = 8
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_1.Location = New System.Drawing.Point(26, 23)
        Me._lblVentas_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(52, 13)
        Me._lblVentas_1.TabIndex = 5
        Me._lblVentas_1.Text = "Desde el "
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_2.Location = New System.Drawing.Point(189, 23)
        Me._lblVentas_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(46, 13)
        Me._lblVentas_2.TabIndex = 7
        Me._lblVentas_2.Text = "Hasta el"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblRpt_2.Location = New System.Drawing.Point(13, 362)
        Me._lblRpt_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 26
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(127, 460)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 103
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(12, 460)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 102
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasRptReparaciones_Corpo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(371, 505)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.fraMoneda)
        Me.Controls.Add(Me._fraVtas_2)
        Me.Controls.Add(Me._fraVtas_0)
        Me.Controls.Add(Me._fraVtas_1)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(329, 157)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmVtasRptReparaciones_Corpo"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Reparaciones"
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.fraMoneda.ResumeLayout(False)
        Me._fraVtas_2.ResumeLayout(False)
        Me._fraVtas_2.PerformLayout()
        Me._fraVtas_0.ResumeLayout(False)
        Me._fraVtas_0.PerformLayout()
        Me._fraVtas_1.ResumeLayout(False)
        Me._fraVtas_1.PerformLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub chkPorTrasnferir_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPorTrasnferir.CheckStateChanged
        Dim lValor As Boolean

        If chkPorTrasnferir.CheckState = System.Windows.Forms.CheckState.Checked Then
            chkxEnviarTaller.Enabled = True
            chkxEnviarPV.Enabled = True
            chkxEnviarTaller.Checked = True
            lValor = False
            mblnxTransferir = True
        Else
            chkxEnviarTaller.Enabled = False
            chkxEnviarTaller.Checked = False
            chkxEnviarPV.Enabled = False
            chkxEnviarPV.Checked = False
            lValor = True
            mblnxTransferir = False
        End If
        Frame2.Enabled = lValor
        Frame1.Enabled = lValor
        _fraVtas_2.Enabled = lValor
        _fraVtas_1.Enabled = lValor
        chkTaller.Enabled = lValor
        optPorSucursal.Checked = True
    End Sub

    Private Sub chkTaller_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTaller.CheckStateChanged
        If FueraChange Then Exit Sub
        Select Case chkTaller.CheckState
            Case System.Windows.Forms.CheckState.Checked
                FueraChange = True
                dbcTaller.Text = "[ Todos ... ]"
                dbcTaller.Tag = ""
                mintCodTaller = 0
                dbcTaller.Enabled = False
                FueraChange = False
            Case Else
                FueraChange = True
                dbcTaller.Text = ""
                dbcTaller.Tag = ""
                mintCodTaller = 0
                dbcTaller.Enabled = True
                FueraChange = False
        End Select
    End Sub

    Private Sub chkTodas_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodas.CheckStateChanged
        If FueraChange Then Exit Sub
        Select Case chkTodas.CheckState
            Case System.Windows.Forms.CheckState.Checked
                FueraChange = True
                dbcSucursal.Text = "[ Todas ... ]"
                dbcSucursal.Tag = ""
                mintCodSucursal = 0
                dbcSucursal.Enabled = False
                FueraChange = False
            Case Else
                FueraChange = True
                dbcSucursal.Text = ""
                dbcSucursal.Tag = ""
                mintCodSucursal = 0
                dbcSucursal.Enabled = True
                FueraChange = False
        End Select
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.CursorChanged
        '        On Error GoTo Merr
        '        Dim lStrSql As String

        '        If FueraChange Then Exit Sub
        '        lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%'"
        '        ModDCombo.DCChange(lStrSql, tecla, dbcSucursal)
        '        If Trim(dbcSucursal.Text) = "" Then
        '            mintCodSucursal = 0
        '        End If

        'Merr:
        '        If Err.Number <> 0 Then
        '            ModEstandar.MostrarError()
        '        End If
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen WHERE TipoAlmacen = 'P'"
        ModDCombo.DCGotFocus(gStrSql, dbcSucursal)
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodas.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        'Dim I As Integer
        'Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Name Then
        '    Exit Sub
        'Else
        '    If Trim(dbcSucursal.Text) = "" Or Trim(dbcSucursal.Text) = C_TODAS Then Exit Sub
        'End If
        'gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%'"
        'Aux = mintCodSucursal
        'mintCodSucursal = 0
        'ModDCombo.DCLostFocus(dbcSucursal, gStrSql, mintCodSucursal)
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = Trim(dbcSucursal.Text)
        'If dbcSucursal.SelectedItem <> 0 Then
        dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        'End If
        dbcSucursal.Text = Aux
    End Sub

    Private Sub dbcTaller_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcTaller.CursorChanged
        '        On Error GoTo Merr
        '        Dim lStrSql As String

        '        If FueraChange Then Exit Sub
        '        lStrSql = "SELECT codTaller, LTrim(RTrim(descTaller)) as DescTaller FROM catTalleres (Nolock) Where DescTaller LIKE '" & Trim(dbcTaller.Text) & "%' Order by DescTaller "
        '        ModDCombo.DCChange(lStrSql, tecla, dbcTaller)

        '        If Trim(dbcTaller.Text) = "" Then
        '            mintCodTaller = 0
        '        End If

        'Merr:
        '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcTaller_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcTaller.Enter
        Pon_Tool()
        gStrSql = "SELECT codTaller, LTrim(RTrim(descTaller)) as DescTaller FROM CatTalleres (Nolock) Order by DescTaller "
        ModDCombo.DCGotFocus(gStrSql, dbcTaller)
    End Sub

    Private Sub dbcTaller_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcTaller.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTaller.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcTaller_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcTaller.Leave
        'Dim I As Integer
        'Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Name Then
        '    Exit Sub
        'Else
        '    If Trim(dbcTaller.Text) = "" Or Trim(dbcTaller.Text) = C_TODOS Then Exit Sub
        'End If
        'gStrSql = "SELECT codTaller, LTrim(RTrim(descTaller)) as DescTaller FROM CatTalleres (Nolock) Where DescTaller LIKE '" & Trim(dbcTaller.Text) & "%' Order by DescTaller "
        'Aux = mintCodTaller
        'mintCodTaller = 0
        'ModDCombo.DCLostFocus(dbcTaller, gStrSql, mintCodTaller)
    End Sub

    Private Sub dbcTaller_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcTaller.MouseUp
        Dim Aux As String
        Aux = Trim(dbcTaller.Text)
        'If dbcTaller.SelectedItem <> 0 Then
        'dbcTaller_Leave(dbcTaller, New System.EventArgs())
        'End If
        dbcTaller.Text = Aux
    End Sub

    Private Sub dtpDesde_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpDesde.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpDesde_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpDesde.KeyPress
        mblnTecleoFechaI = True
        'msglTiempoCambioI = VB.Timer()
    End Sub

    Private Sub dtpHasta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpHasta.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpHasta_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpHasta.KeyPress
        mblnTecleoFechaF = True
        'msglTiempoCambioF = VB.Timer()
    End Sub

    Private Sub frmVtasRptReparaciones_Corpo_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        BringToFront()
    End Sub

    Private Sub frmVtasRptReparaciones_Corpo_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasRptReparaciones_Corpo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "CHKTODAS" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmVtasRptReparaciones_Corpo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasRptReparaciones_Corpo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
    End Sub

    Private Sub frmVtasRptReparaciones_Corpo_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If mblnSalir Then
        '    mblnSalir = False
        '    Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            chkTodas.Focus()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasRptReparaciones_Corpo_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Public Function DevuelveQuery() As String
        On Error GoTo Merr
        Dim lSql As String
        Dim lCamposMoneda As String
        Dim lSeleccion As String
        Dim lEstatus As String
        Dim lPeriodo As String
        Dim lxTransf As String

        lSql = ""
        lCamposMoneda = ""
        lSeleccion = ""
        lEstatus = ""
        lPeriodo = ""
        lxTransf = ""

        '''MONEDA
        If optPesos.Checked Then
            lCamposMoneda = "CostoP as Costo, ImporteP as Importe, AbonosP as Abonos, SaldoP as Saldo, "
        ElseIf optDolares.Checked Then
            lCamposMoneda = "CostoD as Costo, ImporteD as Importe, AbonosD as Abonos, SaldoD as Saldo, "
        End If
        If Not mblnxTransferir Then
            '''SUCURSAL-TALLER
            If (chkTodas.CheckState = System.Windows.Forms.CheckState.Checked And chkTaller.CheckState = System.Windows.Forms.CheckState.Checked) Then '''Todas Sucs/Todos Talleres
                lSeleccion = " (CodSucursal <> 0 ) And "
            ElseIf (chkTodas.CheckState = System.Windows.Forms.CheckState.Checked And chkTaller.CheckState = System.Windows.Forms.CheckState.Unchecked) Then  '''Todas Sucs/Un Talleres
                lSeleccion = " (CodSucursal <> 0 And CodTaller = " & mintCodTaller & ") And "
            ElseIf (chkTodas.CheckState = System.Windows.Forms.CheckState.Unchecked And chkTaller.CheckState = System.Windows.Forms.CheckState.Checked) Then  '''Una Suc/Todos Talleres
                lSeleccion = " (CodSucursal = " & mintCodSucursal & " ) And "
            ElseIf (chkTodas.CheckState = System.Windows.Forms.CheckState.Unchecked And chkTaller.CheckState = System.Windows.Forms.CheckState.Unchecked) Then  '''Una Suc/Un Taller
                lSeleccion = " (CodSucursal = " & mintCodSucursal & " And CodTaller = " & mintCodTaller & ") And "
            End If
            '''ESTATUS
            If (chkVigente.CheckState = System.Windows.Forms.CheckState.Checked) And (chkEntregados.CheckState = System.Windows.Forms.CheckState.Unchecked) Then '''VIGENTE
                lEstatus = "And Entregado = 0 "
            ElseIf (chkVigente.CheckState = System.Windows.Forms.CheckState.Unchecked) And (chkEntregados.CheckState = System.Windows.Forms.CheckState.Checked) Then  '''ENTREGADOS
                lEstatus = "And Entregado = 1 "
            End If
            lPeriodo = "FechaReparacion Between '" & VB6.Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "' And '" & VB6.Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "' "
        Else
            '''SUCURSAL-TALLER
            If (chkTodas.CheckState = System.Windows.Forms.CheckState.Checked) Then '''Todas Sucs
                lSeleccion = " (CodSucursal <> 0 ) And "
            ElseIf (chkTodas.CheckState = System.Windows.Forms.CheckState.Unchecked) Then  '''Una Suc
                lSeleccion = " (CodSucursal = " & mintCodSucursal & " ) And "
            End If

            lEstatus = ""
            lPeriodo = ""
            '''Determina si saldrán los que se deben mandar al taller o al pto de venta
            If chkxEnviarTaller.Checked = True And chkxEnviarPV.Checked = False Then
                lxTransf = " AlTaller = 1 "
            ElseIf chkxEnviarTaller.Checked = False And chkxEnviarPV.Checked = True Then
                lxTransf = " AlPtoVta = 1 "
            Else
                lxTransf = ""
            End If
        End If

        '''15SEP2006 - MAVF
        lSql = "Select CodSucursal, DescAlmacen, FolioReparacion, FechaReparacion, Nombre, CodTaller, DescTaller, Case when FechaEntregaCliente = '01/01/1900' then DateDiff(Day, FechaReparacion, '" & VB6.Format(Today, C_FORMATFECHAGUARDAR) & "') else DateDiff(Day, FechaReparacion, FechaEntregaCliente) end as Dias," & "TipoReparacion, MotivoReparacion, ObservacionesTaller, Moneda, TipoCambio, " & lCamposMoneda & "EstatusRep, dbo.ReparacionesEstatus( CodSucursal, FechaReparacion, FolioReparacion ) As EstatusAct From vw_RptReparaciones " & "Where " & lSeleccion & lPeriodo & lxTransf & "And EstatusRep <> 'Cancelado' " & lEstatus & " Order by CodSucursal, FechaReparacion, FolioReparacion "

        DevuelveQuery = lSql
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub Imprime()
        Dim rptVtasrepReparaciones_Taller As New rptVtasRepReparaciones_Taller
        Dim rptVtasRepReparaciones_Suc As New rptVtasRepReparaciones_Suc

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr
        Dim aParam(6) As Object
        Dim aValues(6) As Object
        Dim lxTransferir As String

        If Not ValidaDatos() Then
            Exit Sub
        End If

        gStrSql = DevuelveQuery()
        ModEstandar.BorraCmd()
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existen datos para el rango de fechas indicado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            rptVtasRepReparaciones_Suc.SetDataSource(frmReportes.rsReport)
            rptVtasrepReparaciones_Taller.SetDataSource(frmReportes.rsReport)
        End If


        '''Determina el nombre del archivo segun la agrupación seleccionada
        If optPorSucursal.Checked Then
            If (lxTransferir <> Nothing) Then
                pdvNum.Value = lxTransferir : pvNum.Add(pdvNum)
                rptVtasRepReparaciones_Suc.DataDefinition.ParameterFields("Tipo").ApplyCurrentValues(pvNum)
            End If

            If lxTransferir <> "" Then
                'rptVtasRepReparaciones_Suc.Text1.Suppress = True
                'rptVtasRepReparaciones_Suc.Text5.Suppress = True
            End If


            If (txtMensaje.Text <> Nothing) Then
                pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
                rptVtasRepReparaciones_Suc.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptVtasRepReparaciones_Suc.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
            End If

            If (gstrCorpoNOMBREEMPRESA <> Nothing) Then
                pdvNum.Value = gstrCorpoNOMBREEMPRESA : pvNum.Add(pdvNum)
                rptVtasRepReparaciones_Suc.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
            End If

            If (dtpDesde.Value <> Nothing) Then
                pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
                rptVtasRepReparaciones_Suc.DataDefinition.ParameterFields("FecIni").ApplyCurrentValues(pvNum)
            End If

            If (dtpHasta.Value <> Nothing) Then
                pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
                rptVtasRepReparaciones_Suc.DataDefinition.ParameterFields("FecFin").ApplyCurrentValues(pvNum)
            End If

            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptVtasRepReparaciones_Suc.DataDefinition.ParameterFields("Tipo").ApplyCurrentValues(pvNum)

            If (Me.optPesos.Checked = True Or Me.optPesos.Checked = False) Then
                pdvNum.Value = IIf(optPesos.Checked, "Importes Expresados en Pesos", "Importes Expresados en Dólares") : pvNum.Add(pdvNum)
                rptVtasRepReparaciones_Suc.DataDefinition.ParameterFields("ImportesEn").ApplyCurrentValues(pvNum)
            End If

            frmReportes.reporteActual = rptVtasRepReparaciones_Suc
            frmReportes.Show()

        Else

            If chkxEnviarTaller.Checked = True And chkxEnviarPV.Checked = False Then
                lxTransferir = "POR TRANSFERIR AL TALLER"
            ElseIf chkxEnviarTaller.Checked = False And chkxEnviarPV.Checked = True Then
                lxTransferir = "POR TRANSFERIR AL PUNTO DE VENTA"
            Else
                lxTransferir = ""
            End If


            If (txtMensaje.Text <> Nothing) Then
                pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
                rptVtasrepReparaciones_Taller.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptVtasrepReparaciones_Taller.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
            End If

            If (gstrCorpoNOMBREEMPRESA <> Nothing) Then
                pdvNum.Value = gstrCorpoNOMBREEMPRESA : pvNum.Add(pdvNum)
                rptVtasrepReparaciones_Taller.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
            End If

            If (dtpDesde.Value <> Nothing) Then
                pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
                rptVtasrepReparaciones_Taller.DataDefinition.ParameterFields("FecIni").ApplyCurrentValues(pvNum)
            End If

            If (dtpHasta.Value <> Nothing) Then
                pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
                rptVtasrepReparaciones_Taller.DataDefinition.ParameterFields("FecFin").ApplyCurrentValues(pvNum)
            End If

            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptVtasRepReparaciones_Suc.DataDefinition.ParameterFields("Tipo").ApplyCurrentValues(pvNum)

            If (Me.optPesos.Checked = True Or Me.optPesos.Checked = False) Then
                pdvNum.Value = IIf(optPesos.Checked, "Importes Expresados en Pesos", "Importes Expresados en Dólares") : pvNum.Add(pdvNum)
                rptVtasrepReparaciones_Taller.DataDefinition.ParameterFields("ImportesEn").ApplyCurrentValues(pvNum)
            End If


            frmReportes.reporteActual = rptVtasrepReparaciones_Taller
            frmReportes.Show()
        End If
        'frmReportes.Imprime(Trim(Text), aParam, aValues)

        Cmd.CommandTimeout = 90

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function ValidaDatos() As Boolean
        If mblnTecleoFechaI Then
            'Do While (msglTiempoCambioI) <= 2.1
            'Loop
            mblnTecleoFechaI = False
        End If
        If mblnTecleoFechaF Then
            'Do While (msglTiempoCambioF) <= 2.1
            'Loop
            mblnTecleoFechaF = False
        End If
        System.Windows.Forms.Application.DoEvents()

        Select Case True
            Case (chkTodas.CheckState = System.Windows.Forms.CheckState.Unchecked And mintCodSucursal = 0)
                MsgBox("Debe seleccionar la sucursal...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                ValidaDatos = False
                dbcSucursal.Focus()
            Case (dtpDesde.Value > dtpHasta.Value)
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Límite", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                ValidaDatos = False
                dtpDesde.Focus()
            Case (chkTaller.CheckState = System.Windows.Forms.CheckState.Unchecked And mintCodTaller = 0)
                MsgBox("Debe seleccionar el taller...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                ValidaDatos = False
                chkTaller.Focus()
            Case (chkVigente.CheckState = 0 And chkEntregados.CheckState = 0)
                MsgBox("Debe seleccionar al menos un estatus ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                ValidaDatos = False
                chkVigente.Focus()
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Public Sub Limpiar()
        On Error Resume Next
        Nuevo()
        chkTodas.Focus()
    End Sub

    Public Sub Nuevo()
        FueraChange = False
        chkTodas.CheckState = System.Windows.Forms.CheckState.Checked
        chkTodas_CheckStateChanged(chkTodas, New System.EventArgs())
        chkTaller.CheckState = System.Windows.Forms.CheckState.Checked
        chkTaller_CheckStateChanged(chkTaller, New System.EventArgs())
        FueraChange = True
        dtpDesde.Value = Format(Today, C_FORMATFECHAMOSTRAR)
        dtpHasta.Value = Format(Today, C_FORMATFECHAMOSTRAR)
        chkVigente.CheckState = System.Windows.Forms.CheckState.Checked
        chkEntregados.CheckState = System.Windows.Forms.CheckState.Unchecked
        optPesos.Checked = True
        optPorSucursal.Checked = True
        txtMensaje.Text = ""
        FueraChange = False
        mintCodSucursal = 0
        mblnSalir = False
        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
        mblnxTransferir = False
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class