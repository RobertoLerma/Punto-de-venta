Public Class frmVtasRPTVentasSalidadeMercanciaCompara2
    Inherits System.Windows.Forms.Form

    Public components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents optAnual As System.Windows.Forms.RadioButton
    Public WithEvents optMensual As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents lblNot As System.Windows.Forms.Label
    Public WithEvents lblNuevas As System.Windows.Forms.Label
    Public WithEvents lblAnt As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents flexSucursales As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents _fraRpt_2 As System.Windows.Forms.GroupBox
    Public WithEvents chkImpuesto As System.Windows.Forms.CheckBox
    Public WithEvents txtAnio As System.Windows.Forms.TextBox
    Public WithEvents cboMes As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_2 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_1 As System.Windows.Forms.GroupBox
    Public WithEvents chkTodasSuc As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_0 As System.Windows.Forms.GroupBox
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents fraRpt As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVtasRPTVentasSalidadeMercanciaCompara))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.optAnual = New System.Windows.Forms.RadioButton()
        Me.optMensual = New System.Windows.Forms.RadioButton()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblNot = New System.Windows.Forms.Label()
        Me.lblNuevas = New System.Windows.Forms.Label()
        Me.lblAnt = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.flexSucursales = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._fraRpt_2 = New System.Windows.Forms.GroupBox()
        Me.chkImpuesto = New System.Windows.Forms.CheckBox()
        Me._fraVtas_1 = New System.Windows.Forms.GroupBox()
        Me.txtAnio = New System.Windows.Forms.TextBox()
        Me.cboMes = New System.Windows.Forms.ComboBox()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me._fraVtas_0 = New System.Windows.Forms.GroupBox()
        Me.chkTodasSuc = New System.Windows.Forms.CheckBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.fraRpt = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.fraVtas = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.flexSucursales, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._fraRpt_2.SuspendLayout()
        Me._fraVtas_1.SuspendLayout()
        Me._fraVtas_0.SuspendLayout()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_optMoneda_1
        '
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_1, CType(1, Short))
        Me._optMoneda_1.Location = New System.Drawing.Point(104, 24)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(73, 27)
        Me._optMoneda_1.TabIndex = 13
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me._optMoneda_1, "Los importes del reporte aparecerán en Pesos")
        Me._optMoneda_1.UseVisualStyleBackColor = False
        '
        '_optMoneda_0
        '
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_0, CType(0, Short))
        Me._optMoneda_0.Location = New System.Drawing.Point(16, 24)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(82, 27)
        Me._optMoneda_0.TabIndex = 12
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me._optMoneda_0, "Los importes del reporte aparecerán en dólares")
        Me._optMoneda_0.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.optAnual)
        Me.Frame3.Controls.Add(Me.optMensual)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(8, 8)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(337, 57)
        Me.Frame3.TabIndex = 25
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Tipo de Reporte"
        '
        'optAnual
        '
        Me.optAnual.BackColor = System.Drawing.SystemColors.Control
        Me.optAnual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAnual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAnual.Location = New System.Drawing.Point(200, 24)
        Me.optAnual.Name = "optAnual"
        Me.optAnual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAnual.Size = New System.Drawing.Size(93, 27)
        Me.optAnual.TabIndex = 1
        Me.optAnual.TabStop = True
        Me.optAnual.Text = "Anual"
        Me.optAnual.UseVisualStyleBackColor = False
        '
        'optMensual
        '
        Me.optMensual.BackColor = System.Drawing.SystemColors.Control
        Me.optMensual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMensual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMensual.Location = New System.Drawing.Point(72, 24)
        Me.optMensual.Name = "optMensual"
        Me.optMensual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMensual.Size = New System.Drawing.Size(105, 27)
        Me.optMensual.TabIndex = 0
        Me.optMensual.TabStop = True
        Me.optMensual.Text = "Mensual"
        Me.optMensual.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Controls.Add(Me.Label5)
        Me.Frame2.Controls.Add(Me.Label4)
        Me.Frame2.Controls.Add(Me.lblNot)
        Me.Frame2.Controls.Add(Me.lblNuevas)
        Me.Frame2.Controls.Add(Me.lblAnt)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(8, 384)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(337, 67)
        Me.Frame2.TabIndex = 18
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Estatus"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label6.Location = New System.Drawing.Point(128, 43)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(137, 21)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "No Seleccionadas"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label5.Location = New System.Drawing.Point(216, 19)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(119, 21)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Sucursales Nuevas"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(40, 19)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(137, 21)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "Sucursales Anteriores"
        '
        'lblNot
        '
        Me.lblNot.BackColor = System.Drawing.SystemColors.Window
        Me.lblNot.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNot.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNot.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblNot.Location = New System.Drawing.Point(96, 40)
        Me.lblNot.Name = "lblNot"
        Me.lblNot.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNot.Size = New System.Drawing.Size(21, 21)
        Me.lblNot.TabIndex = 21
        '
        'lblNuevas
        '
        Me.lblNuevas.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblNuevas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNuevas.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNuevas.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblNuevas.Location = New System.Drawing.Point(184, 16)
        Me.lblNuevas.Name = "lblNuevas"
        Me.lblNuevas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNuevas.Size = New System.Drawing.Size(21, 21)
        Me.lblNuevas.TabIndex = 20
        '
        'lblAnt
        '
        Me.lblAnt.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblAnt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAnt.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAnt.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblAnt.Location = New System.Drawing.Point(8, 16)
        Me.lblAnt.Name = "lblAnt"
        Me.lblAnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAnt.Size = New System.Drawing.Size(21, 21)
        Me.lblAnt.TabIndex = 19
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.flexSucursales)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(8, 200)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(337, 177)
        Me.Frame1.TabIndex = 15
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Sucursales"
        '
        'flexSucursales
        '
        Me.flexSucursales.DataSource = Nothing
        Me.flexSucursales.Location = New System.Drawing.Point(8, 16)
        Me.flexSucursales.Name = "flexSucursales"
        Me.flexSucursales.OcxState = CType(resources.GetObject("flexSucursales.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexSucursales.Size = New System.Drawing.Size(321, 152)
        Me.flexSucursales.TabIndex = 16
        '
        '_fraRpt_2
        '
        Me._fraRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_2.Controls.Add(Me._optMoneda_1)
        Me._fraRpt_2.Controls.Add(Me._optMoneda_0)
        Me._fraRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_2, CType(2, Short))
        Me._fraRpt_2.Location = New System.Drawing.Point(8, 136)
        Me._fraRpt_2.Name = "_fraRpt_2"
        Me._fraRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_2.Size = New System.Drawing.Size(193, 57)
        Me._fraRpt_2.TabIndex = 11
        Me._fraRpt_2.TabStop = False
        Me._fraRpt_2.Text = "Presentar cantidades en ..."
        '
        'chkImpuesto
        '
        Me.chkImpuesto.BackColor = System.Drawing.SystemColors.Control
        Me.chkImpuesto.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkImpuesto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkImpuesto.Location = New System.Drawing.Point(224, 160)
        Me.chkImpuesto.Name = "chkImpuesto"
        Me.chkImpuesto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkImpuesto.Size = New System.Drawing.Size(131, 27)
        Me.chkImpuesto.TabIndex = 14
        Me.chkImpuesto.Text = "Incluir Impuesto"
        Me.chkImpuesto.UseVisualStyleBackColor = False
        '
        '_fraVtas_1
        '
        Me._fraVtas_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_1.Controls.Add(Me.txtAnio)
        Me._fraVtas_1.Controls.Add(Me.cboMes)
        Me._fraVtas_1.Controls.Add(Me._lblVentas_2)
        Me._fraVtas_1.Controls.Add(Me._lblVentas_1)
        Me._fraVtas_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraVtas.SetIndex(Me._fraVtas_1, CType(1, Short))
        Me._fraVtas_1.Location = New System.Drawing.Point(8, 72)
        Me._fraVtas_1.Name = "_fraVtas_1"
        Me._fraVtas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_1.Size = New System.Drawing.Size(337, 58)
        Me._fraVtas_1.TabIndex = 6
        Me._fraVtas_1.TabStop = False
        Me._fraVtas_1.Text = "Período de Proceso ..."
        '
        'txtAnio
        '
        Me.txtAnio.AcceptsReturn = True
        Me.txtAnio.BackColor = System.Drawing.SystemColors.Window
        Me.txtAnio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAnio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnio.Location = New System.Drawing.Point(239, 25)
        Me.txtAnio.MaxLength = 0
        Me.txtAnio.Name = "txtAnio"
        Me.txtAnio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAnio.Size = New System.Drawing.Size(49, 21)
        Me.txtAnio.TabIndex = 10
        Me.txtAnio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboMes
        '
        Me.cboMes.BackColor = System.Drawing.SystemColors.Window
        Me.cboMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboMes.Items.AddRange(New Object() {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"})
        Me.cboMes.Location = New System.Drawing.Point(88, 24)
        Me.cboMes.Name = "cboMes"
        Me.cboMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboMes.Size = New System.Drawing.Size(89, 24)
        Me.cboMes.TabIndex = 8
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_2, CType(2, Short))
        Me._lblVentas_2.Location = New System.Drawing.Point(200, 28)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(33, 17)
        Me._lblVentas_2.TabIndex = 9
        Me._lblVentas_2.Text = "Año"
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_1, CType(1, Short))
        Me._lblVentas_1.Location = New System.Drawing.Point(40, 28)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(34, 17)
        Me._lblVentas_1.TabIndex = 7
        Me._lblVentas_1.Text = "Mes"
        '
        '_fraVtas_0
        '
        Me._fraVtas_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_0.Controls.Add(Me.chkTodasSuc)
        Me._fraVtas_0.Controls.Add(Me.dbcSucursal)
        Me._fraVtas_0.Controls.Add(Me._lblVentas_0)
        Me._fraVtas_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraVtas.SetIndex(Me._fraVtas_0, CType(0, Short))
        Me._fraVtas_0.Location = New System.Drawing.Point(368, 8)
        Me._fraVtas_0.Name = "_fraVtas_0"
        Me._fraVtas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_0.Size = New System.Drawing.Size(337, 57)
        Me._fraVtas_0.TabIndex = 2
        Me._fraVtas_0.TabStop = False
        '
        'chkTodasSuc
        '
        Me.chkTodasSuc.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodasSuc.Checked = True
        Me.chkTodasSuc.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodasSuc.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodasSuc.Enabled = False
        Me.chkTodasSuc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTodasSuc.Location = New System.Drawing.Point(8, 0)
        Me.chkTodasSuc.Name = "chkTodasSuc"
        Me.chkTodasSuc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodasSuc.Size = New System.Drawing.Size(145, 13)
        Me.chkTodasSuc.TabIndex = 3
        Me.chkTodasSuc.Text = "Todas las sucursales"
        Me.chkTodasSuc.UseVisualStyleBackColor = False
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(72, 20)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(249, 24)
        Me.dbcSucursal.TabIndex = 5
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_0, CType(0, Short))
        Me._lblVentas_0.Location = New System.Drawing.Point(16, 24)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(63, 17)
        Me._lblVentas_0.TabIndex = 4
        Me._lblVentas_0.Text = "Sucursal"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label7.Location = New System.Drawing.Point(8, 456)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(337, 29)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Presione la Barra Espaciadora o Haga Doble Click en el Grid Para Establecer un Es" &
    "tatus"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmVtasRPTVentasSalidadeMercanciaCompara
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(353, 488)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me._fraRpt_2)
        Me.Controls.Add(Me.chkImpuesto)
        Me.Controls.Add(Me._fraVtas_1)
        Me.Controls.Add(Me._fraVtas_0)
        Me.Controls.Add(Me.Label7)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 29)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmVtasRPTVentasSalidadeMercanciaCompara"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Comparativo de Ventas Diarias con Año Anterior"
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        CType(Me.flexSucursales, System.ComponentModel.ISupportInitialize).EndInit()
        Me._fraRpt_2.ResumeLayout(False)
        Me._fraVtas_1.ResumeLayout(False)
        Me._fraVtas_1.PerformLayout()
        Me._fraVtas_0.ResumeLayout(False)
        Me._fraVtas_0.PerformLayout()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

End Class