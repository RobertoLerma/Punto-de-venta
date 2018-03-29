Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasRPTVentasSalidadeMercanciaClasifArtic
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents chkImpuesto As System.Windows.Forms.CheckBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_2 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_3 As System.Windows.Forms.GroupBox
    Public WithEvents _fraVtas_2 As System.Windows.Forms.GroupBox
    Public WithEvents _fraVtas_1 As System.Windows.Forms.GroupBox
    Public WithEvents _fraVtas_0 As System.Windows.Forms.GroupBox
    Public WithEvents chkVarios As System.Windows.Forms.CheckBox
    Public WithEvents chkRelojeria As System.Windows.Forms.CheckBox
    Public WithEvents chkJoyeria As System.Windows.Forms.CheckBox
    Public WithEvents chkTodasSuc As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents dbcJFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcJLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcJSubLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcRModelo As System.Windows.Forms.ComboBox
    Public WithEvents dbcVFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcVLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcRMarca As System.Windows.Forms.ComboBox
    Public WithEvents chkDetallar As System.Windows.Forms.CheckBox
    Public WithEvents _fraVtas_4 As System.Windows.Forms.GroupBox
    Public WithEvents _lblVentas_8 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_7 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_6 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_4 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_3 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"
    Const C_NINGUNA As String = "[ Vacío ... ]"

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean

    Dim mblnFueraChange As Boolean
    Dim mintCodSucursal As Integer
    Dim mintJFamilia As Integer
    Dim mintJLinea As Integer
    Dim mintJSubLinea As Integer
    Dim mintRMarca As Integer
    Dim mintRModelo As Integer
    Dim mintVFamilia As Integer
    Dim mintVLinea As Integer
    Dim tecla As Integer

    Dim cTablaTmp As String
    Dim cClasifArtic As String 'Esta variable sirve para mandar al reporte los parámetros seleccionadas
    Dim cCAJ As String
    Dim cCAR As String
    Dim cCAV As String
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim mblnSalir As Boolean


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me._fraVtas_3 = New System.Windows.Forms.GroupBox()
        Me.chkImpuesto = New System.Windows.Forms.CheckBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me._fraVtas_2 = New System.Windows.Forms.GroupBox()
        Me._fraVtas_1 = New System.Windows.Forms.GroupBox()
        Me._fraVtas_0 = New System.Windows.Forms.GroupBox()
        Me.chkVarios = New System.Windows.Forms.CheckBox()
        Me.chkRelojeria = New System.Windows.Forms.CheckBox()
        Me.chkJoyeria = New System.Windows.Forms.CheckBox()
        Me.chkTodasSuc = New System.Windows.Forms.CheckBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.dbcJFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcJLinea = New System.Windows.Forms.ComboBox()
        Me.dbcJSubLinea = New System.Windows.Forms.ComboBox()
        Me.dbcRModelo = New System.Windows.Forms.ComboBox()
        Me.dbcVFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcVLinea = New System.Windows.Forms.ComboBox()
        Me.dbcRMarca = New System.Windows.Forms.ComboBox()
        Me._fraVtas_4 = New System.Windows.Forms.GroupBox()
        Me.chkDetallar = New System.Windows.Forms.CheckBox()
        Me._lblVentas_8 = New System.Windows.Forms.Label()
        Me._lblVentas_7 = New System.Windows.Forms.Label()
        Me._lblVentas_6 = New System.Windows.Forms.Label()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me._lblVentas_4 = New System.Windows.Forms.Label()
        Me._lblVentas_3 = New System.Windows.Forms.Label()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.fraVtas = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me._fraVtas_3.SuspendLayout()
        Me._fraVtas_4.SuspendLayout()
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
        Me.txtMensaje.Location = New System.Drawing.Point(9, 325)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(320, 67)
        Me.txtMensaje.TabIndex = 29
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        '_fraVtas_3
        '
        Me._fraVtas_3.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_3.Controls.Add(Me.chkImpuesto)
        Me._fraVtas_3.Controls.Add(Me.dtpDesde)
        Me._fraVtas_3.Controls.Add(Me.dtpHasta)
        Me._fraVtas_3.Controls.Add(Me._lblVentas_1)
        Me._fraVtas_3.Controls.Add(Me._lblVentas_2)
        Me._fraVtas_3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_3.Location = New System.Drawing.Point(9, 226)
        Me._fraVtas_3.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_3.Name = "_fraVtas_3"
        Me._fraVtas_3.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_3.Size = New System.Drawing.Size(320, 72)
        Me._fraVtas_3.TabIndex = 22
        Me._fraVtas_3.TabStop = False
        Me._fraVtas_3.Text = "Período ..."
        '
        'chkImpuesto
        '
        Me.chkImpuesto.BackColor = System.Drawing.SystemColors.Control
        Me.chkImpuesto.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkImpuesto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkImpuesto.Location = New System.Drawing.Point(192, 17)
        Me.chkImpuesto.Margin = New System.Windows.Forms.Padding(2)
        Me.chkImpuesto.Name = "chkImpuesto"
        Me.chkImpuesto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkImpuesto.Size = New System.Drawing.Size(111, 20)
        Me.chkImpuesto.TabIndex = 27
        Me.chkImpuesto.Text = "Incluir Impuesto"
        Me.chkImpuesto.UseVisualStyleBackColor = False
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(83, 13)
        Me.dtpDesde.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(82, 20)
        Me.dtpDesde.TabIndex = 24
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(83, 37)
        Me.dtpHasta.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(80, 20)
        Me.dtpHasta.TabIndex = 26
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_1.Location = New System.Drawing.Point(30, 17)
        Me._lblVentas_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(49, 13)
        Me._lblVentas_1.TabIndex = 23
        Me._lblVentas_1.Text = "Desde el"
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_2.Location = New System.Drawing.Point(30, 37)
        Me._lblVentas_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(46, 13)
        Me._lblVentas_2.TabIndex = 25
        Me._lblVentas_2.Text = "Hasta el"
        '
        '_fraVtas_2
        '
        Me._fraVtas_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraVtas_2.Location = New System.Drawing.Point(9, 160)
        Me._fraVtas_2.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_2.Name = "_fraVtas_2"
        Me._fraVtas_2.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_2.Size = New System.Drawing.Size(289, 2)
        Me._fraVtas_2.TabIndex = 16
        Me._fraVtas_2.TabStop = False
        '
        '_fraVtas_1
        '
        Me._fraVtas_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraVtas_1.Location = New System.Drawing.Point(6, 110)
        Me._fraVtas_1.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.Name = "_fraVtas_1"
        Me._fraVtas_1.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_1.Size = New System.Drawing.Size(289, 2)
        Me._fraVtas_1.TabIndex = 10
        Me._fraVtas_1.TabStop = False
        '
        '_fraVtas_0
        '
        Me._fraVtas_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraVtas_0.Location = New System.Drawing.Point(6, 32)
        Me._fraVtas_0.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.Name = "_fraVtas_0"
        Me._fraVtas_0.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_0.Size = New System.Drawing.Size(289, 2)
        Me._fraVtas_0.TabIndex = 2
        Me._fraVtas_0.TabStop = False
        '
        'chkVarios
        '
        Me.chkVarios.BackColor = System.Drawing.SystemColors.Control
        Me.chkVarios.Checked = True
        Me.chkVarios.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkVarios.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVarios.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkVarios.Location = New System.Drawing.Point(12, 163)
        Me.chkVarios.Margin = New System.Windows.Forms.Padding(2)
        Me.chkVarios.Name = "chkVarios"
        Me.chkVarios.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVarios.Size = New System.Drawing.Size(69, 19)
        Me.chkVarios.TabIndex = 17
        Me.chkVarios.Text = "Varios"
        Me.chkVarios.UseVisualStyleBackColor = False
        '
        'chkRelojeria
        '
        Me.chkRelojeria.BackColor = System.Drawing.SystemColors.Control
        Me.chkRelojeria.Checked = True
        Me.chkRelojeria.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRelojeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkRelojeria.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkRelojeria.Location = New System.Drawing.Point(12, 113)
        Me.chkRelojeria.Margin = New System.Windows.Forms.Padding(2)
        Me.chkRelojeria.Name = "chkRelojeria"
        Me.chkRelojeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkRelojeria.Size = New System.Drawing.Size(69, 17)
        Me.chkRelojeria.TabIndex = 11
        Me.chkRelojeria.Text = "Relojería"
        Me.chkRelojeria.UseVisualStyleBackColor = False
        '
        'chkJoyeria
        '
        Me.chkJoyeria.BackColor = System.Drawing.SystemColors.Control
        Me.chkJoyeria.Checked = True
        Me.chkJoyeria.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkJoyeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkJoyeria.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkJoyeria.Location = New System.Drawing.Point(12, 41)
        Me.chkJoyeria.Margin = New System.Windows.Forms.Padding(2)
        Me.chkJoyeria.Name = "chkJoyeria"
        Me.chkJoyeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkJoyeria.Size = New System.Drawing.Size(69, 19)
        Me.chkJoyeria.TabIndex = 3
        Me.chkJoyeria.Text = "Joyería"
        Me.chkJoyeria.UseVisualStyleBackColor = False
        '
        'chkTodasSuc
        '
        Me.chkTodasSuc.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodasSuc.Checked = True
        Me.chkTodasSuc.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodasSuc.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodasSuc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodasSuc.Location = New System.Drawing.Point(12, 13)
        Me.chkTodasSuc.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodasSuc.Name = "chkTodasSuc"
        Me.chkTodasSuc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodasSuc.Size = New System.Drawing.Size(129, 16)
        Me.chkTodasSuc.TabIndex = 0
        Me.chkTodasSuc.Text = "Todas las sucursales"
        Me.chkTodasSuc.UseVisualStyleBackColor = False
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(146, 8)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(176, 21)
        Me.dbcSucursal.TabIndex = 1
        '
        'dbcJFamilia
        '
        Me.dbcJFamilia.Location = New System.Drawing.Point(146, 39)
        Me.dbcJFamilia.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcJFamilia.Name = "dbcJFamilia"
        Me.dbcJFamilia.Size = New System.Drawing.Size(176, 21)
        Me.dbcJFamilia.TabIndex = 5
        '
        'dbcJLinea
        '
        Me.dbcJLinea.Location = New System.Drawing.Point(146, 63)
        Me.dbcJLinea.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcJLinea.Name = "dbcJLinea"
        Me.dbcJLinea.Size = New System.Drawing.Size(176, 21)
        Me.dbcJLinea.TabIndex = 7
        '
        'dbcJSubLinea
        '
        Me.dbcJSubLinea.Location = New System.Drawing.Point(146, 88)
        Me.dbcJSubLinea.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcJSubLinea.Name = "dbcJSubLinea"
        Me.dbcJSubLinea.Size = New System.Drawing.Size(176, 21)
        Me.dbcJSubLinea.TabIndex = 9
        '
        'dbcRModelo
        '
        Me.dbcRModelo.Location = New System.Drawing.Point(146, 138)
        Me.dbcRModelo.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcRModelo.Name = "dbcRModelo"
        Me.dbcRModelo.Size = New System.Drawing.Size(176, 21)
        Me.dbcRModelo.TabIndex = 15
        '
        'dbcVFamilia
        '
        Me.dbcVFamilia.Location = New System.Drawing.Point(146, 167)
        Me.dbcVFamilia.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcVFamilia.Name = "dbcVFamilia"
        Me.dbcVFamilia.Size = New System.Drawing.Size(176, 21)
        Me.dbcVFamilia.TabIndex = 19
        '
        'dbcVLinea
        '
        Me.dbcVLinea.Location = New System.Drawing.Point(146, 191)
        Me.dbcVLinea.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcVLinea.Name = "dbcVLinea"
        Me.dbcVLinea.Size = New System.Drawing.Size(176, 21)
        Me.dbcVLinea.TabIndex = 21
        '
        'dbcRMarca
        '
        Me.dbcRMarca.Location = New System.Drawing.Point(146, 114)
        Me.dbcRMarca.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcRMarca.Name = "dbcRMarca"
        Me.dbcRMarca.Size = New System.Drawing.Size(176, 21)
        Me.dbcRMarca.TabIndex = 13
        '
        '_fraVtas_4
        '
        Me._fraVtas_4.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_4.Controls.Add(Me.chkDetallar)
        Me._fraVtas_4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_4.Location = New System.Drawing.Point(186, 234)
        Me._fraVtas_4.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_4.Name = "_fraVtas_4"
        Me._fraVtas_4.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_4.Size = New System.Drawing.Size(103, 7)
        Me._fraVtas_4.TabIndex = 30
        Me._fraVtas_4.TabStop = False
        Me._fraVtas_4.Visible = False
        '
        'chkDetallar
        '
        Me.chkDetallar.BackColor = System.Drawing.SystemColors.Control
        Me.chkDetallar.Checked = True
        Me.chkDetallar.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDetallar.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDetallar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDetallar.Location = New System.Drawing.Point(6, 18)
        Me.chkDetallar.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDetallar.Name = "chkDetallar"
        Me.chkDetallar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDetallar.Size = New System.Drawing.Size(85, 11)
        Me.chkDetallar.TabIndex = 31
        Me.chkDetallar.Text = "Detallar por familia"
        Me.chkDetallar.UseVisualStyleBackColor = False
        '
        '_lblVentas_8
        '
        Me._lblVentas_8.AutoSize = True
        Me._lblVentas_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_8.Location = New System.Drawing.Point(78, 184)
        Me._lblVentas_8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_8.Name = "_lblVentas_8"
        Me._lblVentas_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_8.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_8.TabIndex = 20
        Me._lblVentas_8.Text = "Línea"
        '
        '_lblVentas_7
        '
        Me._lblVentas_7.AutoSize = True
        Me._lblVentas_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_7.Location = New System.Drawing.Point(78, 164)
        Me._lblVentas_7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_7.Name = "_lblVentas_7"
        Me._lblVentas_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_7.Size = New System.Drawing.Size(39, 13)
        Me._lblVentas_7.TabIndex = 18
        Me._lblVentas_7.Text = "Familia"
        '
        '_lblVentas_6
        '
        Me._lblVentas_6.AutoSize = True
        Me._lblVentas_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_6.Location = New System.Drawing.Point(78, 133)
        Me._lblVentas_6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_6.Name = "_lblVentas_6"
        Me._lblVentas_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_6.Size = New System.Drawing.Size(42, 13)
        Me._lblVentas_6.TabIndex = 14
        Me._lblVentas_6.Text = "Modelo"
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.AutoSize = True
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_5.Location = New System.Drawing.Point(78, 114)
        Me._lblVentas_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(37, 13)
        Me._lblVentas_5.TabIndex = 12
        Me._lblVentas_5.Text = "Marca"
        '
        '_lblVentas_4
        '
        Me._lblVentas_4.AutoSize = True
        Me._lblVentas_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_4.Location = New System.Drawing.Point(78, 83)
        Me._lblVentas_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_4.Name = "_lblVentas_4"
        Me._lblVentas_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_4.Size = New System.Drawing.Size(54, 13)
        Me._lblVentas_4.TabIndex = 8
        Me._lblVentas_4.Text = "SubLínea"
        '
        '_lblVentas_3
        '
        Me._lblVentas_3.AutoSize = True
        Me._lblVentas_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_3.Location = New System.Drawing.Point(78, 63)
        Me._lblVentas_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_3.Name = "_lblVentas_3"
        Me._lblVentas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_3.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_3.TabIndex = 6
        Me._lblVentas_3.Text = "Línea"
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_0.Location = New System.Drawing.Point(78, 42)
        Me._lblVentas_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(39, 13)
        Me._lblVentas_0.TabIndex = 4
        Me._lblVentas_0.Text = "Familia"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblRpt_2.Location = New System.Drawing.Point(10, 300)
        Me._lblRpt_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 28
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(124, 407)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 32
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(9, 407)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 31
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasRPTVentasSalidadeMercanciaClasifArtic
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(390, 457)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me._fraVtas_3)
        Me.Controls.Add(Me._fraVtas_2)
        Me.Controls.Add(Me._fraVtas_1)
        Me.Controls.Add(Me._fraVtas_0)
        Me.Controls.Add(Me.dbcVLinea)
        Me.Controls.Add(Me.chkVarios)
        Me.Controls.Add(Me.chkRelojeria)
        Me.Controls.Add(Me.chkJoyeria)
        Me.Controls.Add(Me.chkTodasSuc)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.dbcJFamilia)
        Me.Controls.Add(Me.dbcJLinea)
        Me.Controls.Add(Me.dbcJSubLinea)
        Me.Controls.Add(Me.dbcRModelo)
        Me.Controls.Add(Me.dbcVFamilia)
        Me.Controls.Add(Me.dbcRMarca)
        Me.Controls.Add(Me._fraVtas_4)
        Me.Controls.Add(Me._lblVentas_8)
        Me.Controls.Add(Me._lblVentas_7)
        Me.Controls.Add(Me._lblVentas_6)
        Me.Controls.Add(Me._lblVentas_5)
        Me.Controls.Add(Me._lblVentas_4)
        Me.Controls.Add(Me._lblVentas_3)
        Me.Controls.Add(Me._lblVentas_0)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 29)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasRPTVentasSalidadeMercanciaClasifArtic"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Ventas por Clasificación de Artículo"
        Me._fraVtas_3.ResumeLayout(False)
        Me._fraVtas_3.PerformLayout()
        Me._fraVtas_4.ResumeLayout(False)
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub


    Public Sub Limpiar()
        On Error Resume Next
        Call Me.Nuevo()
        Me.chkTodasSuc.Focus()
    End Sub

    Public Sub Nuevo()
        Me.chkTodasSuc.CheckState = System.Windows.Forms.CheckState.Checked
        chkTodasSuc_CheckStateChanged(chkTodasSuc, New System.EventArgs())

        Me.chkJoyeria.CheckState = System.Windows.Forms.CheckState.Checked
        chkJoyeria_CheckStateChanged(chkJoyeria, New System.EventArgs())

        Me.chkRelojeria.CheckState = System.Windows.Forms.CheckState.Checked
        chkRelojeria_CheckStateChanged(chkRelojeria, New System.EventArgs())

        Me.chkVarios.CheckState = System.Windows.Forms.CheckState.Checked
        chkVarios_CheckStateChanged(chkVarios, New System.EventArgs())

        mintCodSucursal = 0
        mintJFamilia = 0
        mintJLinea = 0
        mintJSubLinea = 0
        mintRMarca = 0
        mintRModelo = 0
        mintVFamilia = 0
        mintVLinea = 0

        Me.dtpDesde.Value = Format(Today, "dd/MMM/yyyy")
        Me.dtpHasta.Value = Format(Today, "dd/MMM/yyyy")
        Me.chkDetallar.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked
        Me.txtMensaje.Text = ""
        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
    End Sub

    Function DevuelveQuery() As String
        On Error GoTo Err_Renamed
        Dim Sql As String
        Dim strWhere As String
        Sql = "SELECT CodSucursal,CA.DescAlmacen,SUM(Cantidad - CantidadDev) AS Cantidad," & IIf(chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "ROUND(SUM(PrecioReal * (Cantidad - CantidadDev)),2) AS Importe,ROUND(SUM(((Descuento * (1 + (PorcIva/100))) * (Cantidad - CantidadDev))),2) AS Descuento,", "ROUND(SUM((PrecioListaSinIva - Descuento) * (Cantidad - CantidadDev)),2) AS Importe,ROUND(SUM(Descuento * (Cantidad - CantidadDev)),2) as Descuento,") & "SUM(CASE WHEN NumPartida = 1 THEN Redondeo ELSE 0 END) AS Redondeo " & "FROM DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') VTA " & "INNER JOIN (SELECT * FROM CatAlmacen WHERE TipoAlmacen = 'P') CA ON VTA.CodSucursal = CA.CodAlmacen " & "INNER JOIN CatGrupos CG ON VTA.CodGrupo = CG.CodGrupo " & "WHERE (Cantidad - CantidadDev) > 0 " & IIf(mintCodSucursal <> 0, "AND CodSucursal = " & mintCodSucursal & " ", "") & "AND "
        strWhere = "("
        If chkJoyeria.CheckState = System.Windows.Forms.CheckState.Checked Then
            strWhere = strWhere & "(VTA.CodGrupo = " & gCODJOYERIA & " "
            If mintJFamilia <> 0 Then
                strWhere = strWhere & "AND VTA.CodFamilia = " & mintJFamilia & " "
                If mintJLinea <> 0 Then
                    strWhere = strWhere & "AND VTA.CodLinea = " & mintJLinea & " "
                    If mintJSubLinea <> 0 Then
                        strWhere = strWhere & "AND VTA.CodSubLinea = " & mintJSubLinea
                    End If
                End If
            End If
            strWhere = strWhere & ") " & IIf(chkRelojeria.CheckState = System.Windows.Forms.CheckState.Checked Or chkVarios.CheckState = System.Windows.Forms.CheckState.Checked, "OR ", "")
        End If
        If chkRelojeria.CheckState = System.Windows.Forms.CheckState.Checked Then
            strWhere = strWhere & "(VTA.CodGrupo = " & gCODRELOJERIA & " "
            If mintRMarca <> 0 Then
                strWhere = strWhere & "AND VTA.CodMarca = " & mintRMarca & " "
                If mintRModelo <> 0 Then
                    strWhere = strWhere & "AND VTA.CodModelo = " & mintRModelo & " "
                End If
            End If
            strWhere = strWhere & ") " & IIf(chkVarios.CheckState = System.Windows.Forms.CheckState.Checked, "OR ", "")
        End If
        If chkVarios.CheckState = System.Windows.Forms.CheckState.Checked Then
            strWhere = strWhere & "(VTA.CodGrupo = " & gCODVARIOS & " "
            If mintVFamilia <> 0 Then
                strWhere = strWhere & "AND VTA.CodFamilia = " & mintVFamilia & " "
                If mintVLinea <> 0 Then
                    strWhere = strWhere & "AND VTA.CodLinea = " & mintVLinea & " "
                End If
            End If
            strWhere = strWhere & ")"
        End If
        strWhere = strWhere & ") "
        Sql = Sql & strWhere & "GROUP BY CodSucursal,CA.DescAlmacen ORDER BY CodSucursal"
        DevuelveQuery = Sql
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub Imprime()

        Dim rptVentasSalidaDeMercanciaClasifArtic As New rptVentasSalidaDeMercanciaClasifArtic
        On Error GoTo Merr
        Dim lStrSql As String
        'Declarar vectores para almacenar los parámetros que se le enviarán al reporte
        Dim aParam(8) As Object
        Dim aValues(8) As Object
        Dim nJOYERIA As Integer
        Dim nRELOJERIA As Integer
        Dim nVARIOS As Integer

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue


        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        If Not ValidaDatos() Then
            Exit Sub
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If

        lStrSql = DevuelveQuery()

        nJOYERIA = Me.chkJoyeria.CheckState
        nRELOJERIA = Me.chkRelojeria.CheckState
        nVARIOS = Me.chkVarios.CheckState

        If nJOYERIA = 0 And nRELOJERIA = 0 And nVARIOS = 0 Then
            MsgBox("Debe elegir, por lo menos, un grupo con el cual generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Sub
        End If

        cCAJ = ""
        cCAR = ""
        cCAV = ""

        If nJOYERIA > 0 Then cCAJ = cCAJ & "Grupo : " & C_JOYERIA & " Familia : " & Trim(Me.dbcJFamilia.Text) & " Línea : " & Trim(Me.dbcJLinea.Text) & " SubLínea : " & Trim(Me.dbcJSubLinea.Text)
        If nRELOJERIA > 0 Then cCAR = cCAR & "Grupo : " & C_RELOJERIA & " Marca : " & Trim(Me.dbcRMarca.Text) & " Modelo : " & Trim(Me.dbcRModelo.Text)
        If nVARIOS > 0 Then cCAV = cCAV & "Grupo : " & C_VARIOS & " Familia : " & Trim(Me.dbcVFamilia.Text) & " Línea : " & Trim(Me.dbcVLinea.Text)

        cClasifArtic = cCAJ & vbNewLine & cCAR & vbNewLine & cCAV

        If Trim(lStrSql) = "" Then
            Exit Sub
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If

        gStrSql = lStrSql
        ModEstandar.BorraCmd()
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            MsgBox("No existen datos para el rango de fechas indicado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            rptVentasSalidaDeMercanciaClasifArtic.SetDataSource(frmReportes.rsReport)
        End If


        'aParam(1) = "Mensaje"
        'aValues(1) = Trim(Me.txtMensaje.Text)
        'aParam(2) = "dDesde"
        'aValues(2) = Me.dtpDesde.Value
        'aParam(3) = "dHasta"
        'aValues(3) = Me.dtpHasta.Value
        'aParam(4) = "Empresa"
        'aValues(4) = Trim(gstrNombCortoEmpresa)
        'aParam(5) = "IncluyeImpuesto"
        'aValues(5) = IIf(Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "** Las cantidades expresadas incluyen IVA.", "** Las cantidades expresadas NO incluyen IVA.")
        'aParam(6) = "ClasifArticJ"
        'aValues(6) = cCAJ
        'aParam(7) = "ClasifArticR"
        'aValues(7) = cCAR
        'aParam(8) = "ClasifArticV"
        'aValues(8) = cCAV


        If (txtMensaje.Text <> Nothing) Then
            pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaClasifArtic.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaClasifArtic.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        End If

        If (dtpDesde.Value <> Nothing) Then
            pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaClasifArtic.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
        End If

        If (dtpHasta.Value <> Nothing) Then
            pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaClasifArtic.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
        End If

        If (gstrNombCortoEmpresa <> Nothing) Then
            pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaClasifArtic.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
        End If

        If (chkImpuesto.CheckState <> Nothing) Then
            pdvNum.Value = IIf(Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "** Las cantidades expresadas incluyen IVA.", "** Las cantidades expresadas NO incluyen IVA.") : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaClasifArtic.DataDefinition.ParameterFields("IncluyeImpuesto").ApplyCurrentValues(pvNum)
        End If

        If (cCAJ <> Nothing) Then
            pdvNum.Value = cCAJ : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaClasifArtic.DataDefinition.ParameterFields("ClasifArticJ").ApplyCurrentValues(pvNum)
        End If

        If (cCAR <> Nothing) Then
            pdvNum.Value = cCAR : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaClasifArtic.DataDefinition.ParameterFields("ClasifArticR").ApplyCurrentValues(pvNum)
        End If

        If (cCAV <> Nothing) Then
            pdvNum.Value = cCAV : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaClasifArtic.DataDefinition.ParameterFields("ClasifArticV").ApplyCurrentValues(pvNum)
        End If

        frmReportes.reporteActual = rptVentasSalidaDeMercanciaClasifArtic 'Es el nombre del archivo que se incluyó en el proyecto
        frmReportes.Show()
        'frmReportes.Imprime(Trim(Me.Text), aParam, aValues)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Cmd.CommandTimeout = 90

Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function ValidaDatos() As Boolean
        If mblnTecleoFechaI Then
            Do While (msglTiempoCambioI) <= 2.1
            Loop
            mblnTecleoFechaI = False
        End If
        If mblnTecleoFechaF Then
            Do While (msglTiempoCambioF) <= 2.1
            Loop
            mblnTecleoFechaF = False
        End If
        System.Windows.Forms.Application.DoEvents()
        Select Case True
            Case Me.dtpDesde.Value > Me.dtpHasta.Value
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Límite", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dtpDesde.Focus()
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Private Sub chkJoyeria_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkJoyeria.CheckStateChanged
        Select Case Me.chkJoyeria.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                mintJFamilia = 0
                Me.dbcJFamilia.Text = C_TODAS
                Me.dbcJFamilia.Enabled = True
                mintJLinea = 0
                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = False
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                mintJFamilia = 0
                Me.dbcJFamilia.Text = C_NINGUNA
                Me.dbcJFamilia.Enabled = False
                mintJLinea = 0
                Me.dbcJLinea.Text = C_NINGUNA
                Me.dbcJLinea.Enabled = False
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_NINGUNA
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub chkRelojeria_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRelojeria.CheckStateChanged
        Select Case Me.chkRelojeria.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                mintRMarca = 0
                Me.dbcRMarca.Text = C_TODAS
                Me.dbcRMarca.Enabled = True
                mintRModelo = 0
                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                mintRMarca = 0
                Me.dbcRMarca.Text = C_NINGUNA
                Me.dbcRMarca.Enabled = False
                mintRModelo = 0
                Me.dbcRModelo.Text = C_NINGUNA
                Me.dbcRModelo.Enabled = False
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub chkTodasSuc_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodasSuc.CheckStateChanged
        Select Case Me.chkTodasSuc.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                Me.dbcSucursal.Text = C_TODAS
                Me.dbcSucursal.Tag = ""
                mintCodSucursal = 0
                Me.dbcSucursal.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                Me.dbcSucursal.Text = ""
                Me.dbcSucursal.Tag = ""
                mintCodSucursal = 0
                Me.dbcSucursal.Enabled = True
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub chkVarios_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkVarios.CheckStateChanged
        Select Case Me.chkVarios.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                mintVFamilia = 0
                Me.dbcVFamilia.Text = C_TODAS
                Me.dbcVFamilia.Enabled = True
                mintVLinea = 0
                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                mintVFamilia = 0
                Me.dbcVFamilia.Text = C_NINGUNA
                Me.dbcVFamilia.Enabled = False
                mintVLinea = 0
                Me.dbcVLinea.Text = C_NINGUNA
                Me.dbcVLinea.Enabled = False
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub dbcJFAmilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        lStrSql = "SELECT codFamilia, LTrim(RTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia LIKE '" & Trim(Me.dbcJFamilia.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcJFamilia))


        If Trim(Me.dbcJFamilia.Text) = "" Then
            mintJFamilia = 0
            mblnFueraChange = True
            '''Me.dbcJFamilia.text = C_TODAS
            Me.dbcJFamilia.Enabled = True
            mintJLinea = 0

            Me.dbcJLinea.Text = C_TODAS
            Me.dbcJLinea.Enabled = False
            mintJSubLinea = 0

            Me.dbcJSubLinea.Text = C_TODAS
            Me.dbcJSubLinea.Enabled = False
            mblnFueraChange = False
            '    dbcJFamilia_LostFocus
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcjFAmilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.Enter
        Pon_Tool()
        gStrSql = "SELECT codFamilia, LTrim(RTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcJFamilia))
    End Sub

    Private Sub dbcJFAmilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJFamilia.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkJoyeria.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcJFamiliaKeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJFamilia.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcJFamilia.text)
        '''    If Me.dbcJFamilia.SelectedItem <> 0 Then
        '''        dbcJFamilia_LostFocus
        '''    End If
        '''    Me.dbcJFamilia.text = Aux
    End Sub

    Private Sub dbcJFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.Leave
        Dim I As Integer
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If

        gStrSql = "SELECT codFamilia, LTrim(RTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia LIKE '" & Trim(Me.dbcJFamilia.Text) & "%'"
        Aux = mintJFamilia
        mintJFamilia = 0

        If Trim(Me.dbcJFamilia.Text) <> Trim(C_TODAS) Or Trim(Me.dbcJFamilia.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcJFamilia), gStrSql, mintJFamilia)
        End If
        If Aux <> mintJFamilia Then
            If mintJFamilia = 0 Then
                mblnFueraChange = True

                Me.dbcJFamilia.Text = C_TODAS
                Me.dbcJFamilia.Enabled = True
                mintJLinea = 0

                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = False
                mintJSubLinea = 0

                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintJLinea = 0

                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = True
                mintJSubLinea = 0

                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
                Me.dbcJLinea.Focus()
            End If
        End If

        If Trim(Me.dbcJFamilia.Text) = "" Then Me.dbcJFamilia.Text = C_TODAS
    End Sub

    Private Sub dbcJFamilia_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcJFamilia.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcJFamilia.text)
        '''    If Me.dbcJFamilia.SelectedItem <> 0 Then
        '''        dbcJFamilia_LostFocus
        '''    End If
        '''    Me.dbcJFamilia.text = Aux
    End Sub

    Private Sub dbcJLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        lStrSql = " SELECT codLinea, LTrim(RTrim(descLinea)) as DescLinea FROM CatLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFamilia & " and descLinea LIKE '" & Trim(Me.dbcJLinea.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcJLinea))


        If Trim(Me.dbcJLinea.Text) = "" Then
            mintJLinea = 0
            mblnFueraChange = True
            '''Me.dbcJLinea.text = C_TODAS
            Me.dbcJLinea.Enabled = True
            mintJSubLinea = 0

            Me.dbcJSubLinea.Text = C_TODAS
            Me.dbcJSubLinea.Enabled = False
            mblnFueraChange = False
            '    dbcJLinea_LostFocus
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcJLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.Enter
        Pon_Tool()
        gStrSql = " SELECT codLinea, LTrim(RTrim(descLinea)) as DescLinea FROM CatLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFamilia
        ModDCombo.DCGotFocus(gStrSql, dbcJLinea)
    End Sub

    Private Sub dbcJLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcJFamilia.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcJLineaKeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJLinea.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcJLinea.text)
        '''    If Me.dbcJLinea.SelectedItem <> 0 Then
        '''        dbcJLinea_LostFocus
        '''    End If
        '''    Me.dbcJLinea.text = Aux
    End Sub

    Private Sub dbcJLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.Leave
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If

        gStrSql = " SELECT codLinea, LTrim(RTrim(descLinea)) as DescLinea FROM CatLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFamilia & " and descLinea LIKE '" & Trim(Me.dbcJLinea.Text) & "%'"
        Aux = mintJLinea
        mintJLinea = 0

        If Trim(Me.dbcJLinea.Text) <> Trim(C_TODAS) Or Trim(Me.dbcJLinea.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcJLinea), gStrSql, mintJLinea)
        End If
        If Aux <> mintJLinea Then
            If mintJLinea = 0 Then
                mblnFueraChange = True

                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = True
                mintJSubLinea = 0

                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintJSubLinea = 0

                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = True
                mblnFueraChange = False
                Me.dbcJSubLinea.Focus()
            End If
        End If

        If Trim(Me.dbcJLinea.Text) = "" Then Me.dbcJLinea.Text = C_TODAS
    End Sub

    Private Sub dbcJLinea_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcJLinea.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcJLinea.text)
        '''    If Me.dbcJLinea.SelectedItem <> 0 Then
        '''        dbcJLinea_LostFocus
        '''    End If
        '''    Me.dbcJLinea.text = Aux
    End Sub

    Private Sub dbcJSubLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        lStrSql = " SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as DescSubLinea FROM CatSubLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and descSubLinea LIKE '" & Trim(Me.dbcJSubLinea.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcJSubLinea))


        If Trim(Me.dbcJSubLinea.Text) = "" Then
            mintJSubLinea = 0
            mblnFueraChange = True
            '''Me.dbcJSubLinea.text = C_TODAS
            Me.dbcJSubLinea.Enabled = True
            mblnFueraChange = False
            '    dbcJSubLinea_LostFocus
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcJSubLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.Enter
        Pon_Tool()
        gStrSql = " SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as DescSubLinea FROM CatSubLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea
        ModDCombo.DCGotFocus(gStrSql, dbcJSubLinea)
    End Sub

    Private Sub dbcJSubLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJSubLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcJLinea.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcJSubLineaKeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJSubLinea.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcJSubLinea.text)
        '''    If Me.dbcJSubLinea.SelectedItem <> 0 Then
        '''        dbcJSubLinea_LostFocus
        '''    End If
        '''    Me.dbcJSubLinea.text = Aux
    End Sub

    Private Sub dbcJSubLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.Leave
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If

        gStrSql = " SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as DescSubLinea FROM CatSubLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and descSubLinea LIKE '" & Trim(Me.dbcJSubLinea.Text) & "%'"
        Aux = mintJSubLinea
        mintJSubLinea = 0

        If Trim(Me.dbcJSubLinea.Text) <> Trim(C_TODAS) Or Trim(Me.dbcJSubLinea.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcJSubLinea), gStrSql, mintJSubLinea)
        End If
        If Aux <> mintJSubLinea Then
            If mintJSubLinea = 0 Then
                mblnFueraChange = True

                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = True
                mblnFueraChange = False
            End If
        End If

        If Trim(Me.dbcJSubLinea.Text) = "" Then Me.dbcJSubLinea.Text = C_TODAS
    End Sub

    Private Sub dbcJSubLineaMouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcJSubLinea.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcJSubLinea.text)
        '''    If Me.dbcJSubLinea.SelectedItem <> 0 Then
        '''        dbcJSubLinea_LostFocus
        '''    End If
        '''    Me.dbcJSubLinea.text = Aux
    End Sub

    Private Sub dbcRMarca_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        lStrSql = "SELECT codMarca, LTrim(RTrim(descMarca)) as descMarca FROM catMarcas Where codGrupo = " & gCODRELOJERIA & " and descMarca LIKE '" & Trim(Me.dbcRMarca.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcRMarca))


        If Trim(Me.dbcRMarca.Text) = "" Then
            mintRMarca = 0
            mblnFueraChange = True
            '''Me.dbcRMarca.text = C_TODAS
            Me.dbcRMarca.Enabled = True
            mintRModelo = 0

            Me.dbcRModelo.Text = C_TODOS
            Me.dbcRModelo.Enabled = False
            mblnFueraChange = False
            '    dbcRMarca_LostFocus
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcRMarca_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.Enter
        Pon_Tool()
        gStrSql = "SELECT codMarca, LTrim(RTrim(descMarca)) as descMarca FROM catMarcas Where codGrupo = " & gCODRELOJERIA
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcRMarca))
    End Sub

    Private Sub dbcRMarca_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRMarca.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkRelojeria.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcRMarcaKeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRMarca.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcRMarca.text)
        '''    If Me.dbcRMarca.SelectedItem <> 0 Then
        '''        dbcRMarca_LostFocus
        '''    End If
        '''    Me.dbcRMarca.text = Aux
    End Sub

    Private Sub dbcRMarca_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.Leave
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If

        gStrSql = "SELECT codMarca, LTrim(RTrim(descMarca)) as descMarca FROM catMarcas Where codGrupo = " & gCODRELOJERIA & " and descMarca LIKE '" & Trim(Me.dbcRMarca.Text) & "%'"
        Aux = mintRMarca
        mintRMarca = 0

        If Trim(Me.dbcRMarca.Text) <> Trim(C_TODAS) Or Trim(Me.dbcRMarca.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcRMarca), gStrSql, mintRMarca)
        End If

        If Aux <> mintRMarca Then
            If mintRMarca = 0 Then
                mblnFueraChange = True

                Me.dbcRMarca.Text = C_TODAS
                Me.dbcRMarca.Enabled = True
                mintRModelo = 0

                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintRModelo = 0

                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = True
                mblnFueraChange = False
                Me.dbcRModelo.Focus()
            End If
        End If

        If Trim(Me.dbcRMarca.Text) = "" Then Me.dbcRMarca.Text = C_TODAS
    End Sub

    Private Sub dbcRMarcaMouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcRMarca.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcRMarca.text)
        '''    If Me.dbcRMarca.SelectedItem <> 0 Then
        '''        dbcRMarca_LostFocus
        '''    End If
        '''    Me.dbcRMarca.text = Aux
    End Sub

    Private Sub dbcRmodelo_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        lStrSql = " SELECT codModelo, LTrim(RTrim(descModelo)) as DescModelo FROM CatModelos WHERE CodGrupo = " & gCODRELOJERIA & " and CodMarca = " & mintRMarca & " and descModelo LIKE '" & Trim(Me.dbcRModelo.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcRModelo))


        If Trim(Me.dbcRModelo.Text) = "" Then
            mintRModelo = 0
            mblnFueraChange = True
            '''Me.dbcRModelo.text = C_TODOS
            Me.dbcRModelo.Enabled = True
            mblnFueraChange = False
            '    dbcRModelo_LostFocus
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcRmodelo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.Enter
        Pon_Tool()
        gStrSql = " SELECT codModelo, LTrim(RTrim(descModelo)) as DescModelo FROM CatModelos WHERE CodGrupo = " & gCODRELOJERIA & " and CodMarca = " & mintRMarca
        ModDCombo.DCGotFocus(gStrSql, dbcRModelo)
    End Sub

    Private Sub dbcRmodelo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRModelo.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcRMarca.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcRModeloKeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRModelo.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcRModelo.text)
        '''    If Me.dbcRModelo.SelectedItem <> 0 Then
        '''        dbcRModelo_LostFocus
        '''    End If
        '''    Me.dbcRModelo.text = Aux
    End Sub

    Private Sub dbcRModelo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.Leave
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If

        gStrSql = " SELECT codModelo, LTrim(RTrim(descModelo)) as DescModelo FROM CatModelos WHERE CodGrupo = " & gCODRELOJERIA & " and CodMarca = " & mintRMarca & " and descModelo LIKE '" & Trim(Me.dbcRModelo.Text) & "%'"
        Aux = mintRModelo
        mintRModelo = 0

        If Trim(Me.dbcRModelo.Text) <> Trim(C_TODOS) Or Trim(Me.dbcRModelo.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcRModelo), gStrSql, mintRModelo)
        End If
        If Aux <> mintRModelo Then
            If mintRModelo = 0 Then
                mblnFueraChange = True

                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = True
                mblnFueraChange = False
            End If
        End If

        If Trim(Me.dbcRModelo.Text) = "" Then Me.dbcRModelo.Text = C_TODOS
    End Sub

    Private Sub dbcRModeloMouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcRModelo.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcRModelo.text)
        '''    If Me.dbcRModelo.SelectedItem <> 0 Then
        '''        dbcRModelo_LostFocus
        '''    End If
        '''    Me.dbcRModelo.text = Aux
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcSucursal)


        If Trim(Me.dbcSucursal.Text) = "" Then
            mintCodSucursal = 0
            'dbcSucursal_LostFocus
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen WHERE TipoAlmacen = 'P'"
        ModDCombo.DCGotFocus(gStrSql, dbcSucursal)
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkTodasSuc.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursalKeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcSucursal.text)
        '''    If Me.dbcSucursal.SelectedItem <> 0 Then
        '''        dbcSucursal_LostFocus
        '''    End If
        '''    Me.dbcSucursal.text = Aux
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        Dim I As Integer
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        Else

            If Trim(Me.dbcSucursal.Text) = "" Or Trim(Me.dbcSucursal.Text) = C_TODAS Then Exit Sub
        End If

        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        Aux = mintCodSucursal
        mintCodSucursal = 0
        ModDCombo.DCLostFocus((Me.dbcSucursal), gStrSql, mintCodSucursal)
    End Sub

    Private Sub dbcSucursalMouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcSucursal.text)
        '''    If Me.dbcSucursal.SelectedItem <> 0 Then
        '''        dbcSucursal_LostFocus
        '''    End If
        '''    Me.dbcSucursal.text = Aux
    End Sub

    Private Sub dbcVFamilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        lStrSql = "SELECT codFamilia, LTrim(RTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " and descFamilia LIKE '" & Trim(Me.dbcVFamilia.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcVFamilia))


        If Trim(Me.dbcVFamilia.Text) = "" Then
            mintVFamilia = 0
            mblnFueraChange = True
            '''Me.dbcVFamilia.text = C_TODAS
            Me.dbcVFamilia.Enabled = True
            mintVLinea = 0

            Me.dbcVLinea.Text = C_TODAS
            Me.dbcVLinea.Enabled = False
            mblnFueraChange = False
            '    dbcVFamilia_LostFocus
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcVFamilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.Enter
        Pon_Tool()
        gStrSql = "SELECT codFamilia, LTrim(RTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcVFamilia))
    End Sub

    Private Sub dbcVFamilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcVFamilia.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkVarios.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcVFamiliaKeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcVFamilia.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcVFamilia.text)
        '''    If Me.dbcVFamilia.SelectedItem <> 0 Then
        '''        dbcVFamilia_LostFocus
        '''    End If
        '''    Me.dbcVFamilia.text = Aux
    End Sub

    Private Sub dbcVFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.Leave
        Dim I As Integer
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If

        gStrSql = "SELECT codFamilia, LTrim(RTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " and descFamilia LIKE '" & Trim(Me.dbcVFamilia.Text) & "%'"
        Aux = mintVFamilia
        mintVFamilia = 0

        If Trim(Me.dbcVFamilia.Text) <> Trim(C_TODAS) Or Trim(Me.dbcVFamilia.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcVFamilia), gStrSql, mintVFamilia)
        End If

        If Aux <> mintVFamilia Then
            If mintVFamilia = 0 Then
                mblnFueraChange = True

                Me.dbcVFamilia.Text = C_TODAS
                Me.dbcVFamilia.Enabled = True
                mintVLinea = 0

                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintVLinea = 0

                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = True
                mblnFueraChange = False
                Me.dbcVLinea.Focus()
            End If
        End If

        If Trim(Me.dbcVFamilia.Text) = "" Then Me.dbcVFamilia.Text = C_TODAS
    End Sub

    Private Sub dbcVFamiliaMouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcVFamilia.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcVFamilia.text)
        '''    If Me.dbcVFamilia.SelectedItem <> 0 Then
        '''        dbcVFamilia_LostFocus
        '''    End If
        '''    Me.dbcVFamilia.text = Aux
    End Sub

    Private Sub dbcVLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        lStrSql = " SELECT codLinea, LTrim(RTrim(descLinea)) as DescLinea FROM CatLineas WHERE CodGrupo = " & gCODVARIOS & " and CodFamilia = " & mintVFamilia & " and descLinea LIKE '" & Trim(Me.dbcVLinea.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcVLinea))


        If Trim(Me.dbcVLinea.Text) = "" Then
            mintVLinea = 0
            mblnFueraChange = True
            '''Me.dbcVLinea.text = C_TODAS
            Me.dbcVLinea.Enabled = True
            mblnFueraChange = False
            '    dbcVLinea_LostFocus
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcVLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.Enter
        Pon_Tool()
        gStrSql = " SELECT codLinea, LTrim(RTrim(descLinea)) as DescLinea FROM CatLineas WHERE CodGrupo = " & gCODVARIOS & " and CodFamilia = " & mintVFamilia
        ModDCombo.DCGotFocus(gStrSql, dbcVLinea)
    End Sub

    Private Sub dbcVLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcVLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcVFamilia.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcVLineaKeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcVLinea.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcVLinea.text)
        '''    If Me.dbcVLinea.SelectedItem <> 0 Then
        '''        dbcVLinea_LostFocus
        '''    End If
        '''    Me.dbcVLinea.text = Aux
    End Sub

    Private Sub dbcVLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.Leave
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If

        gStrSql = " SELECT codLinea, LTrim(RTrim(descLinea)) as DescLinea FROM CatLineas WHERE CodGrupo = " & gCODVARIOS & " and CodFamilia = " & mintVFamilia & " and descLinea LIKE '" & Trim(Me.dbcVLinea.Text) & "%'"
        Aux = mintVLinea
        mintVLinea = 0

        If Trim(Me.dbcVLinea.Text) <> Trim(C_TODAS) Or Trim(Me.dbcVLinea.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcVLinea), gStrSql, mintVLinea)
        End If
        If Aux <> mintVLinea Then
            If mintVLinea = 0 Then
                mblnFueraChange = True

                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = True
                mblnFueraChange = False
            End If
        End If

        If Trim(Me.dbcVLinea.Text) = "" Then Me.dbcVLinea.Text = C_TODAS
    End Sub

    Private Sub dbcVLinea_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcVLinea.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcVLinea.text)
        '''    If Me.dbcVLinea.SelectedItem <> 0 Then
        '''        dbcVLinea_LostFocus
        '''    End If
        '''    Me.dbcVLinea.text = Aux
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

    Private Sub frmVtasRPTVentasSalidadeMercanciaClasifArtic_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaClasifArtic_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaClasifArtic_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "CHKTODASSUC" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaClasifArtic_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaClasifArtic_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Me.dtpDesde.MinDate = C_FECHAINICIAL
        Me.dtpDesde.MaxDate = C_FECHAFINAL
        Me.dtpHasta.MinDate = C_FECHAINICIAL
        Me.dtpHasta.MaxDate = C_FECHAFINAL
        Call Me.Nuevo()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaClasifArtic_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If mblnSalir Then
            mblnSalir = False
            Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Sale del Formulario
                    Cancel = 0
                Case MsgBoxResult.No 'No sale del formulario
                    Me.chkTodasSuc.Focus()
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaClasifArtic_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class