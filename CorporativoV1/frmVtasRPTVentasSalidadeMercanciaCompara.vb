Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasRPTVentasSalidadeMercanciaCompara
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
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
    'Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents flexSucursales As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    'Public WithEvents Frame1 As System.Windows.Forms.GroupBox
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
    'Public WithEvents _fraVtas_0 As System.Windows.Forms.GroupBox
    Public WithEvents Label7 As System.Windows.Forms.Label
    'Public WithEvents fraRpt As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    'Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray


    Const C_TODAS As String = "[ Todas ... ]"

    Dim mblnFueraChange As Boolean
    Dim mintCodSucursal As Integer
    Dim tecla As Integer
    Dim aMeses(12) As Integer
    'Dim aTablasTmp(1 To 12) As String

    Dim cTablaTmp As String

    Dim mblnSalir As Boolean
    Dim blnSucAnt As Boolean
    Dim blnSucNue As Boolean

    Const C_COLCODIGO As Integer = 0
    Const C_COLDESCRIPCION As Integer = 1
    Const C_COLESTATUS As Integer = 2

    'Estatus para las sucursales
    Const C_NOSELECCIONADA As String = "NOT"
    Const C_ANTERIOR As String = "ANT"
    Public WithEvents btnImpirmir As Button
    Public WithEvents btnSalir As Button
    Public WithEvents btnNuevo As Button
    Const C_NUEVA As String = "NVO"


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVtasRPTVentasSalidadeMercanciaCompara))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.optAnual = New System.Windows.Forms.RadioButton()
        Me.optMensual = New System.Windows.Forms.RadioButton()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblNot = New System.Windows.Forms.Label()
        Me.lblNuevas = New System.Windows.Forms.Label()
        Me.lblAnt = New System.Windows.Forms.Label()
        Me.flexSucursales = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._fraRpt_2 = New System.Windows.Forms.GroupBox()
        Me.chkImpuesto = New System.Windows.Forms.CheckBox()
        Me._fraVtas_1 = New System.Windows.Forms.GroupBox()
        Me.txtAnio = New System.Windows.Forms.TextBox()
        Me.cboMes = New System.Windows.Forms.ComboBox()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me.chkTodasSuc = New System.Windows.Forms.CheckBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnImpirmir = New System.Windows.Forms.Button()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.Frame3.SuspendLayout()
        CType(Me.flexSucursales, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._fraRpt_2.SuspendLayout()
        Me._fraVtas_1.SuspendLayout()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_optMoneda_1
        '
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_1.Location = New System.Drawing.Point(101, 20)
        Me._optMoneda_1.Margin = New System.Windows.Forms.Padding(2)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(55, 22)
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
        Me._optMoneda_0.Location = New System.Drawing.Point(35, 20)
        Me._optMoneda_0.Margin = New System.Windows.Forms.Padding(2)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(62, 22)
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
        Me.Frame3.Location = New System.Drawing.Point(6, 8)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(332, 46)
        Me.Frame3.TabIndex = 25
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Tipo de Reporte"
        '
        'optAnual
        '
        Me.optAnual.BackColor = System.Drawing.SystemColors.Control
        Me.optAnual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAnual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAnual.Location = New System.Drawing.Point(198, 17)
        Me.optAnual.Margin = New System.Windows.Forms.Padding(2)
        Me.optAnual.Name = "optAnual"
        Me.optAnual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAnual.Size = New System.Drawing.Size(70, 22)
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
        Me.optMensual.Location = New System.Drawing.Point(102, 17)
        Me.optMensual.Margin = New System.Windows.Forms.Padding(2)
        Me.optMensual.Name = "optMensual"
        Me.optMensual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMensual.Size = New System.Drawing.Size(79, 22)
        Me.optMensual.TabIndex = 0
        Me.optMensual.TabStop = True
        Me.optMensual.Text = "Mensual"
        Me.optMensual.UseVisualStyleBackColor = False
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label6.Location = New System.Drawing.Point(134, 379)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(99, 20)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "No Seleccionadas"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label5.Location = New System.Drawing.Point(206, 353)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(106, 19)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Sucursales Nuevas"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(57, 353)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(120, 19)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "Sucursales Anteriores"
        '
        'lblNot
        '
        Me.lblNot.BackColor = System.Drawing.SystemColors.Window
        Me.lblNot.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNot.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNot.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblNot.Location = New System.Drawing.Point(107, 379)
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
        Me.lblNuevas.Location = New System.Drawing.Point(183, 350)
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
        Me.lblAnt.Location = New System.Drawing.Point(36, 349)
        Me.lblAnt.Name = "lblAnt"
        Me.lblAnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAnt.Size = New System.Drawing.Size(21, 21)
        Me.lblAnt.TabIndex = 19
        '
        'flexSucursales
        '
        Me.flexSucursales.DataSource = Nothing
        Me.flexSucursales.Location = New System.Drawing.Point(6, 177)
        Me.flexSucursales.Name = "flexSucursales"
        Me.flexSucursales.OcxState = CType(resources.GetObject("flexSucursales.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexSucursales.Size = New System.Drawing.Size(332, 162)
        Me.flexSucursales.TabIndex = 16
        '
        '_fraRpt_2
        '
        Me._fraRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_2.Controls.Add(Me._optMoneda_1)
        Me._fraRpt_2.Controls.Add(Me._optMoneda_0)
        Me._fraRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraRpt_2.Location = New System.Drawing.Point(6, 110)
        Me._fraRpt_2.Margin = New System.Windows.Forms.Padding(2)
        Me._fraRpt_2.Name = "_fraRpt_2"
        Me._fraRpt_2.Padding = New System.Windows.Forms.Padding(2)
        Me._fraRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_2.Size = New System.Drawing.Size(198, 46)
        Me._fraRpt_2.TabIndex = 11
        Me._fraRpt_2.TabStop = False
        Me._fraRpt_2.Text = "Presentar cantidades en ..."
        '
        'chkImpuesto
        '
        Me.chkImpuesto.BackColor = System.Drawing.SystemColors.Control
        Me.chkImpuesto.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkImpuesto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkImpuesto.Location = New System.Drawing.Point(232, 130)
        Me.chkImpuesto.Margin = New System.Windows.Forms.Padding(2)
        Me.chkImpuesto.Name = "chkImpuesto"
        Me.chkImpuesto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkImpuesto.Size = New System.Drawing.Size(106, 22)
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
        Me._fraVtas_1.Location = New System.Drawing.Point(6, 58)
        Me._fraVtas_1.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.Name = "_fraVtas_1"
        Me._fraVtas_1.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_1.Size = New System.Drawing.Size(332, 47)
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
        Me.txtAnio.Location = New System.Drawing.Point(230, 17)
        Me.txtAnio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtAnio.MaxLength = 0
        Me.txtAnio.Name = "txtAnio"
        Me.txtAnio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAnio.Size = New System.Drawing.Size(38, 20)
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
        Me.cboMes.Location = New System.Drawing.Point(117, 17)
        Me.cboMes.Margin = New System.Windows.Forms.Padding(2)
        Me.cboMes.Name = "cboMes"
        Me.cboMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboMes.Size = New System.Drawing.Size(68, 21)
        Me.cboMes.TabIndex = 8
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_2.Location = New System.Drawing.Point(201, 20)
        Me._lblVentas_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(26, 13)
        Me._lblVentas_2.TabIndex = 9
        Me._lblVentas_2.Text = "Año"
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_1.Location = New System.Drawing.Point(81, 20)
        Me._lblVentas_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(27, 13)
        Me._lblVentas_1.TabIndex = 7
        Me._lblVentas_1.Text = "Mes"
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
        Me.dbcSucursal.Size = New System.Drawing.Size(249, 21)
        Me.dbcSucursal.TabIndex = 5
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
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
        Me.Label7.Location = New System.Drawing.Point(57, 418)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(253, 40)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Presione la Barra Espaciadora o Haga Doble Click en el Grid Para Establecer un Es" &
    "tatus"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnImpirmir
        '
        Me.btnImpirmir.Location = New System.Drawing.Point(12, 493)
        Me.btnImpirmir.Name = "btnImpirmir"
        Me.btnImpirmir.Size = New System.Drawing.Size(109, 36)
        Me.btnImpirmir.TabIndex = 26
        Me.btnImpirmir.Text = "Imprimir"
        Me.btnImpirmir.UseVisualStyleBackColor = False
        '
        'btnSalir
        '
        Me.btnSalir.BackColor = System.Drawing.SystemColors.Control
        Me.btnSalir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSalir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSalir.Location = New System.Drawing.Point(242, 492)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSalir.Size = New System.Drawing.Size(109, 36)
        Me.btnSalir.TabIndex = 74
        Me.btnSalir.Text = "&Salir"
        Me.btnSalir.UseVisualStyleBackColor = False
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(127, 492)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 73
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'frmVtasRPTVentasSalidadeMercanciaCompara
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(362, 540)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImpirmir)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me._fraRpt_2)
        Me.Controls.Add(Me.chkImpuesto)
        Me.Controls.Add(Me._fraVtas_1)
        Me.Controls.Add(Me.lblNot)
        Me.Controls.Add(Me.lblNuevas)
        Me.Controls.Add(Me.lblAnt)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.flexSucursales)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 29)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmVtasRPTVentasSalidadeMercanciaCompara"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Comparativo de Ventas Diarias con Año Anterior"
        Me.Frame3.ResumeLayout(False)
        CType(Me.flexSucursales, System.ComponentModel.ISupportInitialize).EndInit()
        Me._fraRpt_2.ResumeLayout(False)
        Me._fraVtas_1.ResumeLayout(False)
        Me._fraVtas_1.PerformLayout()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


    Function CreaTablaTemporal(ByRef NombreTabla As String) As String
        On Error GoTo Err_Renamed
        Dim Value As Integer
        Dim Tabla As String
        Randomize()
        Value = Int((10000 * Rnd()) + 1)
        Tabla = Trim(NombreTabla & CStr(Value))
        If Mid(Tabla, 3, 11) = "CompMensual" Then
            gStrSql = "CREATE TABLE " & Tabla & " ( " & " Dia             Int, " & " Fecha           SmallDateTime, " & " CodSucursal     Int, " & " DescSucursal    Char(40), " & " SaldoDiarioAn   Money, " & " AcumAnterior    Money, " & " SaldoDiarioAc   Money, " & " AcumActual      Money, " & " NTotalesP       Money, " & " NTotalesA       Money, " & " NTotalesPorc    Money, " & " HTotalesP       Money, " & " HTotalesA       Money, " & " VtaDiariaP      Money, " & " VtaDiariaA      Money, " & " VtaDiariaDif    Money)"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            Cmd.Execute()
        ElseIf Mid(Tabla, 3, 9) = "CompAnual" Then
            gStrSql = "CREATE TABLE " & Tabla & " ( " & " Mes             Int, " & " DescMes         Char(3), " & " CodSucursal     Int, " & " DescSucursal    Char(40), " & " NTotalesP       Money, " & " NTotalesA       Money, " & " TotalesP        Money, " & " TotalesA        Money, " & " PromAnt         Money, " & " PromAct         Money, " & " Dif             Money)"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            Cmd.Execute()
        End If
        CreaTablaTemporal = Tabla
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Obtener_Sucursales()
        On Error GoTo Err_Renamed
        Dim Ren As Integer
        gStrSql = "select codalmacen,descalmacen from catalmacen where tipoalmacen = 'P' order by codalmacen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            With flexSucursales
                Ren = 1
                Do While Not RsGral.EOF
                    .set_TextMatrix(Ren, C_COLCODIGO, RsGral.Fields("CodAlmacen").Value)
                    .set_TextMatrix(Ren, C_COLDESCRIPCION, Trim(RsGral.Fields("DescAlmacen").Value))
                    .set_TextMatrix(Ren, C_COLESTATUS, C_NOSELECCIONADA)
                    RsGral.MoveNext()
                    If Not RsGral.EOF Then
                        If Ren = .Rows - 1 Then
                            .Rows = .Rows + 1
                        End If
                        Ren = Ren + 1
                    End If
                Loop
                .Col = C_COLDESCRIPCION
                .Row = 1
            End With
        End If
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Sub Encabezado()
        With flexSucursales
            .Clear()
            .Rows = 11
            .set_ColWidth(C_COLCODIGO, 0, 0)
            .set_ColWidth(C_COLDESCRIPCION, 0, 4500)
            .set_ColWidth(C_COLESTATUS, 0, 0)
            .Col = C_COLDESCRIPCION
            .Row = 0
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Sucursal"
            .Row = 1
            Obtener_Sucursales()
        End With
    End Sub

    Public Sub Nuevo()
        optMensual.Checked = True
        optAnual.Checked = False
        Me.chkTodasSuc.CheckState = System.Windows.Forms.CheckState.Checked
        chkTodasSuc_CheckStateChanged(chkTodasSuc, New System.EventArgs())
        Me.cboMes.SelectedIndex = Month(Today) - 1
        Me.txtAnio.Text = CStr(Year(Today))
        Me._optMoneda_0.Checked = True
        Me._optMoneda_1.Checked = False
        Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked
        Encabezado()
    End Sub

    Public Sub Limpiar()
        On Error Resume Next
        Me.Nuevo()
        Me.optMensual.Focus()
    End Sub

    Public Sub CalculaDiaMes(ByRef nAnio As Integer)
        If (nAnio Mod 4) = 0 Then
            aMeses(2) = 29 'Febrero
        Else
            aMeses(2) = 28 'Febrero
        End If
        aMeses(1) = 31 'Enero
        aMeses(3) = 31 'Marzo
        aMeses(4) = 30 'Abril
        aMeses(5) = 31 'Mayo
        aMeses(6) = 30 'Junio
        aMeses(7) = 31 'Julio
        aMeses(8) = 31 'Agosto
        aMeses(9) = 30 'Septiembre
        aMeses(10) = 31 'Octubre
        aMeses(11) = 30 'Noviembre
        aMeses(12) = 31 'Diciembre
    End Sub

    Function ObtenerSucursalesAnt() As String
        Dim Sucursales As String
        Dim I As Integer
        With flexSucursales
            Sucursales = "("
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLESTATUS)) = C_ANTERIOR Then
                    If Trim(Sucursales) = "(" Then
                        Sucursales = Sucursales & .get_TextMatrix(I, C_COLCODIGO)
                    Else
                        Sucursales = Sucursales & "," & .get_TextMatrix(I, C_COLCODIGO)
                    End If
                End If
            Next
            Sucursales = Sucursales & ")"
        End With
        ObtenerSucursalesAnt = Sucursales
        Return ObtenerSucursalesAnt
    End Function

    Function ObtenerSucursalesNue() As String
        Dim Sucursales As String
        Dim I As Integer
        With flexSucursales
            Sucursales = "("
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLESTATUS)) = C_NUEVA Then
                    If Trim(Sucursales) = "(" Then
                        Sucursales = Sucursales & .get_TextMatrix(I, C_COLCODIGO)
                    Else
                        Sucursales = Sucursales & "," & .get_TextMatrix(I, C_COLCODIGO)
                    End If
                End If
            Next
            Sucursales = Sucursales & ")"
        End With
        ObtenerSucursalesNue = Sucursales
    End Function

    Public Function DevuelveQuery(ByRef Sucursales As String) As String
        On Error GoTo Merr
        Dim I As Integer
        Dim nMes As Integer
        Dim nAnio As Integer
        Dim nDiaMes1 As Integer ' Del mes del año pasado
        Dim nDiaMes2 As Integer ' Del mes del año indicado
        Dim DiaMes As Integer
        Dim ConImpuesto As Integer
        Dim Moneda As String
        'Dim aDiaMes(12, 2) As Integer
        'Crear una tabla temporal
        nMes = Me.cboMes.SelectedIndex + 1
        Me.CalculaDiaMes(CShort(Numerico((Me.txtAnio.Text))))
        nDiaMes1 = aMeses(nMes)
        Me.CalculaDiaMes(CShort(Numerico((Me.txtAnio.Text))) - 1)
        nDiaMes2 = aMeses(nMes)
        ConImpuesto = IIf(chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, 1, 0)
        Moneda = IIf(_optMoneda_0.Checked, "D", "P")
        If optMensual.Checked = True Then

            If nDiaMes1 = nDiaMes2 Then
                DiaMes = nDiaMes1
            ElseIf nDiaMes1 > nDiaMes2 Then
                DiaMes = nDiaMes1
            ElseIf nDiaMes2 > nDiaMes1 Then
                DiaMes = nDiaMes2
            End If
            cTablaTmp = CreaTablaTemporal("##CompMensual")
            ModStoredProcedures.PR_IME_VentasComparativoMensual(cTablaTmp, CStr(nMes), txtAnio.Text, CStr(DiaMes), CStr(ConImpuesto), Moneda, Sucursales)
            Cmd.Execute()

            DevuelveQuery = "Select * From " & cTablaTmp & " Order By CodSucursal,Dia"

        ElseIf optAnual.Checked = True Then

            cTablaTmp = CreaTablaTemporal("##CompAnual")
            ModStoredProcedures.PR_IME_VentasComparativoAnual(cTablaTmp, txtAnio.Text, CStr(ConImpuesto), Moneda, Sucursales)
            Cmd.Execute()

            DevuelveQuery = "Select * From " & cTablaTmp & " Order By CodSucursal,Mes"

        End If
        Return DevuelveQuery
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub Imprime()
        'On Error GoTo Merr
        Dim lStrSql As String
        Dim rsLocalAnt As ADODB.Recordset
        Dim rsLocalNue As ADODB.Recordset
        Dim rsSucAnt As ADODB.Recordset
        Dim rsSucNue As ADODB.Recordset
        Dim SucursalesAnt As Object
        Dim SucursalesNue As String
        Dim I As Integer
        Dim NumSucNue As Object
        Dim NumSucAnt As Integer
        'Declarar vectores para almacenar los parámetros que se le enviarán al reporte

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        If Not ValidaDatos() Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Sub
        End If

        Cmd.CommandTimeout = 300
        'Sucursales Anteriores
        SucursalesAnt = ObtenerSucursalesAnt()

        If ChecaSucursalesAnt() Then
            lStrSql = DevuelveQuery(SucursalesAnt)

            If Trim(lStrSql) = "" Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Exit Sub
            End If

            gStrSql = lStrSql
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            rsSucAnt = Cmd.Execute

            'Borrar la Tabla Temporal
            gStrSql = "DROP TABLE " & Trim(cTablaTmp)
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            Cmd.Execute()
            cTablaTmp = ""

            'Obtener el número de sucursales
            gStrSql = "SELECT Count(*) as nSucursales FROM CatAlmacen WHERE TipoAlmacen = 'P' and CodAlmacen IN" & ObtenerSucursalesAnt()
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            rsLocalAnt = Cmd.Execute

            NumSucAnt = rsLocalAnt.Fields("nSucursales").Value

            blnSucAnt = True
        Else
            blnSucAnt = False
        End If


        'ModComparativo.LlenaDatos RsGral, rsLocal!nSucursales
        'End If

        If ChecaSucursalesNue() Then
            'Sucursales Nuevas
            'SucursalesNue = ObtenerSucursalesNue
            lStrSql = DevuelveQuery(ObtenerSucursalesNue())

            If Trim(lStrSql) = "" Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Exit Sub
            End If

            gStrSql = lStrSql
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            rsSucNue = Cmd.Execute

            'Borrar la Tabla Temporal
            gStrSql = "DROP TABLE " & Trim(cTablaTmp)
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            Cmd.Execute()
            cTablaTmp = ""
            blnSucNue = True
        Else
            blnSucNue = False
        End If

        If optMensual.Checked = True Then
            If blnSucNue Then
                gStrSql = "SELECT Count(*) as nSucursales FROM CatAlmacen WHERE TipoAlmacen = 'P' and CodAlmacen IN" & ObtenerSucursalesNue()
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.UP_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                rsLocalNue = Cmd.Execute
                NumSucNue = rsLocalNue.Fields("nSucursales").Value
            End If
            ModComparativo.LlenaDatos(rsSucAnt, rsSucNue, IIf(blnSucAnt, NumSucAnt, 0), IIf(blnSucNue, NumSucNue, 0), True, blnSucAnt, blnSucNue)
        ElseIf optAnual.Checked = True Then
            If blnSucNue Then
                gStrSql = "SELECT Count(*) as nSucursales FROM CatAlmacen WHERE TipoAlmacen = 'P' and CodAlmacen IN" & ObtenerSucursalesNue()
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.UP_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                rsLocalNue = Cmd.Execute
                NumSucNue = rsLocalNue.Fields("nSucursales").Value
            End If
            ModComparativo.LlenaDatos(rsSucAnt, rsSucNue, IIf(blnSucAnt, NumSucAnt, 0), IIf(blnSucNue, NumSucNue, 0), False, blnSucAnt, blnSucNue)
        End If
        Cmd.CommandTimeout = 90

        'Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Function ChecaSucursalesAnt() As Boolean
        Dim I As Integer
        ChecaSucursalesAnt = False
        With flexSucursales
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, C_COLDESCRIPCION) = "" Then Exit Function
                If Trim(.get_TextMatrix(I, C_COLESTATUS)) = C_ANTERIOR Then
                    ChecaSucursalesAnt = True
                    Exit Function
                End If
            Next
        End With
    End Function

    Function ChecaSucursalesNue() As Boolean
        Dim I As Integer
        ChecaSucursalesNue = False
        With flexSucursales
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, C_COLDESCRIPCION) = "" Then Exit Function
                If Trim(.get_TextMatrix(I, C_COLESTATUS)) = C_NUEVA Then
                    ChecaSucursalesNue = True
                    Exit Function
                End If
            Next
        End With
    End Function

    Public Function ValidaDatos() As Boolean
        Select Case True
            Case CShort(Numerico((Me.txtAnio.Text))) < 1900 Or CShort(Numerico((Me.txtAnio.Text))) > 2075
                MsgBox("El año especificado está fuera del rango permitido ( 1900..2075 )", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Case Not ChecaSucursalesAnt() And Not ChecaSucursalesNue()
                MsgBox("No ha seleccionado ninguna sucursal, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
            Case Else
                ValidaDatos = True
        End Select
    End Function

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

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'Else
        '    If Trim(Me.dbcSucursal.Text) = "" Or Trim(Me.dbcSucursal.Text) = C_TODAS Then Exit Sub
        'End If
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        Aux = mintCodSucursal
        mintCodSucursal = 0
        ModDCombo.DCLostFocus((Me.dbcSucursal), gStrSql, mintCodSucursal)
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcSucursal.Text)
        'If Me.dbcSucursal.SelectedItem <> 0 Then
        'dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        'End If
        Me.dbcSucursal.Text = Aux
    End Sub

    Private Sub flexSucursales_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexSucursales.ClickEvent
        With flexSucursales
            .Col = C_COLDESCRIPCION
            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
        End With
    End Sub

    Private Sub flexSucursales_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexSucursales.DblClick
        Dim RenActual As Integer
        With flexSucursales
            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
            If Trim(.get_TextMatrix(.Row, C_COLDESCRIPCION)) = "" Then Exit Sub
            .Col = 1
            If Trim(.get_TextMatrix(.Row, C_COLESTATUS)) = C_NOSELECCIONADA Then
                .CellBackColor = lblAnt.BackColor
                .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
                .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
                .set_TextMatrix(.Row, C_COLESTATUS, C_ANTERIOR)
            ElseIf Trim(.get_TextMatrix(.Row, C_COLESTATUS)) = C_ANTERIOR Then
                .CellBackColor = lblNuevas.BackColor
                .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
                .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
                .set_TextMatrix(.Row, C_COLESTATUS, C_NUEVA)
            ElseIf Trim(.get_TextMatrix(.Row, C_COLESTATUS)) = C_NUEVA Then
                .CellBackColor = lblNot.BackColor
                .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
                .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
                .set_TextMatrix(.Row, C_COLESTATUS, C_NOSELECCIONADA)
            End If
        End With
    End Sub

    Private Sub flexSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexSucursales.Enter
        With flexSucursales
            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
        End With
    End Sub

    Private Sub flexSucursales_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexSucursales.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Space Then
            flexSucursales_DblClick(flexSucursales, New System.EventArgs())
        End If
    End Sub
    Private Sub frmVtasRPTVentasSalidadeMercanciaCompara_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaCompara_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaCompara_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "OPTMENSUAL" Or UCase(Me.ActiveControl.Name) = "OPTANUAL" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaCompara_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaCompara_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaCompara_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If mblnSalir Then
        '    mblnSalir = False
        '    Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaCompara_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        IsNothing(Me)
    End Sub

    Private Sub optAnual_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAnual.CheckedChanged
        If eventSender.Checked Then
            cboMes.Enabled = False
        End If
    End Sub

    Private Sub optMensual_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMensual.CheckedChanged
        If eventSender.Checked Then
            cboMes.Enabled = True
        End If
    End Sub

    Private Sub txtAnio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnio.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub txtAnio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAnio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtAnio.Text = VB6.Format(Numerico((Me.txtAnio.Text)), "###0")
        End If
        KeyAscii = ModEstandar.MskCantidad((Me.txtAnio.Text), KeyAscii, 4, 0, (Me.txtAnio.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtAnio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnio.Leave
        If CShort(Numerico((Me.txtAnio.Text))) < 1900 Then
            MsgBox("El año introducido está por debajo del rango de tiempo permitido", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Me.txtAnio.Focus()
        ElseIf CShort(Numerico((Me.txtAnio.Text))) > 2075 Then
            MsgBox("El año introducido excede el rango de tiempo permitido", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Me.txtAnio.Focus()
        End If
    End Sub

    Private Sub btnImpirmir_Click(sender As Object, e As EventArgs) Handles btnImpirmir.Click
        Imprime()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub
End Class