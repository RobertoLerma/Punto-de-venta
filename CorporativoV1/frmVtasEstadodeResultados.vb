Option Strict Off
Option Explicit On
Imports Microsoft.Office.Interop
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmVtasEstadodeResultados
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkFueraEnter As System.Windows.Forms.CheckBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents flexGastos As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents lblDesc As System.Windows.Forms.Label
    'Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents chkTodaslasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents optAnual As System.Windows.Forms.RadioButton
    Public WithEvents optMensual As System.Windows.Forms.RadioButton
    'Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents chkIncluirImpuesto As System.Windows.Forms.CheckBox
    Public WithEvents cmbAño As System.Windows.Forms.ComboBox
    Public WithEvents cmbMes As System.Windows.Forms.ComboBox
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    'Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents Line2 As System.Windows.Forms.Label
    Public WithEvents Line1 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    'Public WithEvents Frame1 As System.Windows.Forms.GroupBox


    'Variables
    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim ObjExcel As Object
    Dim objLibro As Excel.Workbook
    Dim objHoja As Excel.Worksheet
    Dim FechaInicial As String
    Dim FechaFinal As String
    Dim Fecha As String
    Dim Moneda As String
    Dim Impuesto As String
    Dim Renglon As Integer
    Dim Rango As String
    Dim CodSucursal As Integer
    Dim Año As Integer
    Dim cTablaTmpGastos As String
    Dim cTablaTmpResultados As String
    Dim Sucursales() As Integer
    Dim NumSucursales As Integer
    Dim FechaAux As Date
    Dim EjecutaExcel As Boolean
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnSalir As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnImprimir As Button
    Dim MesAcumulado As Integer
    'Dim cmd As ADODB.Command


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVtasEstadodeResultados))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.chkTodaslasSucursales = New System.Windows.Forms.CheckBox()
        Me.optAnual = New System.Windows.Forms.RadioButton()
        Me.optMensual = New System.Windows.Forms.RadioButton()
        Me.chkIncluirImpuesto = New System.Windows.Forms.CheckBox()
        Me.cmbAño = New System.Windows.Forms.ComboBox()
        Me.cmbMes = New System.Windows.Forms.ComboBox()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.chkFueraEnter = New System.Windows.Forms.CheckBox()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.flexGastos = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.Line2 = New System.Windows.Forms.Label()
        Me.Line1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        CType(Me.flexGastos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtCodSucursal
        '
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.Location = New System.Drawing.Point(94, 34)
        Me.txtCodSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucursal.MaxLength = 0
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(46, 20)
        Me.txtCodSucursal.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCodSucursal, "Codigo de la Sucursal")
        '
        'chkTodaslasSucursales
        '
        Me.chkTodaslasSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodaslasSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodaslasSucursales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodaslasSucursales.Location = New System.Drawing.Point(12, 11)
        Me.chkTodaslasSucursales.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodaslasSucursales.Name = "chkTodaslasSucursales"
        Me.chkTodaslasSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodaslasSucursales.Size = New System.Drawing.Size(146, 17)
        Me.chkTodaslasSucursales.TabIndex = 0
        Me.chkTodaslasSucursales.Text = "Todas las Sucursales"
        Me.ToolTip1.SetToolTip(Me.chkTodaslasSucursales, "Muestra Todas las Sucursales")
        Me.chkTodaslasSucursales.UseVisualStyleBackColor = False
        '
        'optAnual
        '
        Me.optAnual.BackColor = System.Drawing.SystemColors.Control
        Me.optAnual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAnual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAnual.Location = New System.Drawing.Point(15, 47)
        Me.optAnual.Margin = New System.Windows.Forms.Padding(2)
        Me.optAnual.Name = "optAnual"
        Me.optAnual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAnual.Size = New System.Drawing.Size(55, 20)
        Me.optAnual.TabIndex = 9
        Me.optAnual.TabStop = True
        Me.optAnual.Text = "Anual"
        Me.ToolTip1.SetToolTip(Me.optAnual, "Muestra un Reporte Anual")
        Me.optAnual.UseVisualStyleBackColor = False
        '
        'optMensual
        '
        Me.optMensual.BackColor = System.Drawing.SystemColors.Control
        Me.optMensual.Checked = True
        Me.optMensual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMensual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMensual.Location = New System.Drawing.Point(15, 28)
        Me.optMensual.Margin = New System.Windows.Forms.Padding(2)
        Me.optMensual.Name = "optMensual"
        Me.optMensual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMensual.Size = New System.Drawing.Size(77, 15)
        Me.optMensual.TabIndex = 8
        Me.optMensual.TabStop = True
        Me.optMensual.Text = "Mensual"
        Me.ToolTip1.SetToolTip(Me.optMensual, "Muestra un Reporte Mensual")
        Me.optMensual.UseVisualStyleBackColor = False
        '
        'chkIncluirImpuesto
        '
        Me.chkIncluirImpuesto.BackColor = System.Drawing.SystemColors.Control
        Me.chkIncluirImpuesto.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncluirImpuesto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkIncluirImpuesto.Location = New System.Drawing.Point(12, 104)
        Me.chkIncluirImpuesto.Margin = New System.Windows.Forms.Padding(2)
        Me.chkIncluirImpuesto.Name = "chkIncluirImpuesto"
        Me.chkIncluirImpuesto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncluirImpuesto.Size = New System.Drawing.Size(104, 22)
        Me.chkIncluirImpuesto.TabIndex = 5
        Me.chkIncluirImpuesto.Text = "Incluir Impuesto"
        Me.ToolTip1.SetToolTip(Me.chkIncluirImpuesto, "Muestra los Importes con Impuesto")
        Me.chkIncluirImpuesto.UseVisualStyleBackColor = False
        '
        'cmbAño
        '
        Me.cmbAño.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAño.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAño.Location = New System.Drawing.Point(256, 65)
        Me.cmbAño.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbAño.Name = "cmbAño"
        Me.cmbAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAño.Size = New System.Drawing.Size(80, 21)
        Me.cmbAño.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.cmbAño, "Año.")
        '
        'cmbMes
        '
        Me.cmbMes.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMes.Items.AddRange(New Object() {"01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"})
        Me.cmbMes.Location = New System.Drawing.Point(91, 65)
        Me.cmbMes.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbMes.Name = "cmbMes"
        Me.cmbMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMes.Size = New System.Drawing.Size(112, 21)
        Me.cmbMes.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.cmbMes, "Mes.")
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(5, 18)
        Me.optDolares.Margin = New System.Windows.Forms.Padding(2)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(65, 20)
        Me.optDolares.TabIndex = 7
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Dolares"
        Me.ToolTip1.SetToolTip(Me.optDolares, "Muestra los Importes en Dolares")
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Checked = True
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(5, 44)
        Me.optPesos.Margin = New System.Windows.Forms.Padding(2)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(65, 20)
        Me.optPesos.TabIndex = 6
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me.optPesos, "Muestra los Importes en Pesos")
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'chkFueraEnter
        '
        Me.chkFueraEnter.BackColor = System.Drawing.SystemColors.Control
        Me.chkFueraEnter.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkFueraEnter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkFueraEnter.Location = New System.Drawing.Point(458, 486)
        Me.chkFueraEnter.Name = "chkFueraEnter"
        Me.chkFueraEnter.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkFueraEnter.Size = New System.Drawing.Size(83, 26)
        Me.chkFueraEnter.TabIndex = 20
        Me.chkFueraEnter.Text = "Check1"
        Me.chkFueraEnter.UseVisualStyleBackColor = False
        Me.chkFueraEnter.Visible = False
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(12, 224)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(65, 20)
        Me.txtFlex.TabIndex = 18
        Me.txtFlex.Visible = False
        '
        'flexGastos
        '
        Me.flexGastos.DataSource = Nothing
        Me.flexGastos.Location = New System.Drawing.Point(12, 202)
        Me.flexGastos.Name = "flexGastos"
        Me.flexGastos.OcxState = CType(resources.GetObject("flexGastos.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexGastos.Size = New System.Drawing.Size(382, 151)
        Me.flexGastos.TabIndex = 10
        '
        'lblDesc
        '
        Me.lblDesc.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDesc.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDesc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblDesc.Location = New System.Drawing.Point(12, 365)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDesc.Size = New System.Drawing.Size(382, 21)
        Me.lblDesc.TabIndex = 19
        Me.lblDesc.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(144, 33)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(216, 21)
        Me.dbcSucursal.TabIndex = 2
        '
        'Line2
        '
        Me.Line2.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line2.Location = New System.Drawing.Point(8, 203)
        Me.Line2.Name = "Line2"
        Me.Line2.Size = New System.Drawing.Size(517, 1)
        Me.Line2.TabIndex = 18
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(8, 116)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(518, 1)
        Me.Line1.TabIndex = 19
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(39, 39)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(62, 12)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Sucursal :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(220, 67)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(34, 17)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Año :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(49, 67)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(38, 17)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Mes :"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optMensual)
        Me.GroupBox1.Controls.Add(Me.optAnual)
        Me.GroupBox1.Location = New System.Drawing.Point(256, 104)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(101, 80)
        Me.GroupBox1.TabIndex = 20
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Tipo de Reporte"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.optDolares)
        Me.GroupBox2.Controls.Add(Me.optPesos)
        Me.GroupBox2.Location = New System.Drawing.Point(132, 104)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(90, 80)
        Me.GroupBox2.TabIndex = 21
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Moneda"
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackColor = System.Drawing.SystemColors.Control
        Me.btnLimpiar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLimpiar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLimpiar.Location = New System.Drawing.Point(8, 452)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLimpiar.Size = New System.Drawing.Size(108, 43)
        Me.btnLimpiar.TabIndex = 85
        Me.btnLimpiar.Text = "Limpiar"
        Me.ToolTip1.SetToolTip(Me.btnLimpiar, "Registro de Clientes")
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnSalir
        '
        Me.btnSalir.BackColor = System.Drawing.SystemColors.Control
        Me.btnSalir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSalir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSalir.Location = New System.Drawing.Point(236, 452)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSalir.Size = New System.Drawing.Size(108, 43)
        Me.btnSalir.TabIndex = 84
        Me.btnSalir.Text = "Salir"
        Me.ToolTip1.SetToolTip(Me.btnSalir, "Registro de Clientes")
        Me.btnSalir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.BackColor = System.Drawing.SystemColors.Control
        Me.btnBuscar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBuscar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBuscar.Location = New System.Drawing.Point(236, 403)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBuscar.Size = New System.Drawing.Size(108, 43)
        Me.btnBuscar.TabIndex = 83
        Me.btnBuscar.Text = "Buscar"
        Me.ToolTip1.SetToolTip(Me.btnBuscar, "Registro de Clientes")
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'btnEliminar
        '
        Me.btnEliminar.BackColor = System.Drawing.SystemColors.Control
        Me.btnEliminar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEliminar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEliminar.Location = New System.Drawing.Point(122, 403)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEliminar.Size = New System.Drawing.Size(108, 43)
        Me.btnEliminar.TabIndex = 82
        Me.btnEliminar.Text = "Eliminar"
        Me.ToolTip1.SetToolTip(Me.btnEliminar, "Registro de Clientes")
        Me.btnEliminar.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(8, 403)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(108, 43)
        Me.btnGuardar.TabIndex = 81
        Me.btnGuardar.Text = "Guardar"
        Me.ToolTip1.SetToolTip(Me.btnGuardar, "Registro de Clientes")
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(122, 452)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(108, 43)
        Me.btnImprimir.TabIndex = 86
        Me.btnImprimir.Text = "Imprimir"
        Me.ToolTip1.SetToolTip(Me.btnImprimir, "Registro de Clientes")
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasEstadodeResultados
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(401, 515)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.chkTodaslasSucursales)
        Me.Controls.Add(Me.txtCodSucursal)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbMes)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmbAño)
        Me.Controls.Add(Me.chkIncluirImpuesto)
        Me.Controls.Add(Me.flexGastos)
        Me.Controls.Add(Me.txtFlex)
        Me.Controls.Add(Me.lblDesc)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(258, 127)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasEstadodeResultados"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Estado de Resultados"
        CType(Me.flexGastos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub



    Function CreaTablaTemporal(ByRef NombreTabla As String) As String
        On Error GoTo Err_Renamed
        Dim Value As Integer
        Dim Tabla As String
        Randomize()
        Value = Int((10000 * Rnd()) + 1)
        Tabla = Trim(NombreTabla & CStr(Value))
        If Mid(Tabla, 3, 10) = "Resultados" Then
            gStrSql = "CREATE TABLE " & Tabla & " ( " & " CodSucursal Int, " & " Seccion Int, " & " Mes Int, " & " Año Int, " & " TipoMovto Char(1), " & " Descripcion varchar(50), " & " Importe Money, " & " Porcentaje SmallMoney)"
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

    Function DevuelveQuery() As String
        Dim Sql As String
        Dim subsql As String
        Dim I As Object
        Dim J As Integer
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim blnExiste As Boolean
        Dim FormaQuery As Boolean

        MesAcumulado = 0
        NumSucursales = 0
        ModCorporativo.ObtenerLimitedeFechas(CShort((Trim(cmbMes.Text))), CShort(Trim(cmbAño.Text)), FechaInicial, FechaFinal)
        MesAcumulado = Month(CDate(FechaInicial))

        With flexGastos
            If Vacio() Then
                DevuelveQuery = ""
                Exit Function
            End If
            Sql = ""
            subsql = ""
            gStrSql = "Select count(*) NumSuc from catalmacen where tipoalmacen = 'P'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                NumSucursales = RsGral.Fields("numsuc").Value
                ReDim Sucursales(NumSucursales)
            End If
            gStrSql = "Select codalmacen from catalmacen where tipoalmacen = 'P'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            I = 1
            If RsGral.RecordCount > 0 Then
                Do While Not RsGral.EOF
                    Sucursales(I) = RsGral.Fields("CodAlmacen").Value
                    RsGral.MoveNext()
                    I = I + 1
                Loop
            End If
            Sql = "insert into #TablaTmp select * from ("
            For J = 1 To NumSucursales
                blnExiste = False
                FormaQuery = False
                For I = 1 To .Rows - 1
                    If Trim(.get_TextMatrix(I, 0)) <> "" Or Trim(.get_TextMatrix(I, 1)) <> "" Or Trim(.get_TextMatrix(I, 2)) <> "" Or Trim(.get_TextMatrix(I, 3)) <> "" Or Trim(.get_TextMatrix(I, 4)) <> "" Then
                        If Sucursales(J) = CDbl(Numerico(.get_TextMatrix(I, 5))) Then
                            blnExiste = True
                            If Trim(subsql) = "" And Not FormaQuery Then
                                subsql = "(select " & Numerico(.get_TextMatrix(I, 5)) & " as codsucursal,'H' as tipomovto,'Gastos' as descripcion," & "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaresconimpuesto," & "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaressinimpuesto," & "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as imppesosconimpuesto," & "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as imppesossinimpuesto," & "b.fechamovto as fecha " & "from movimientosorigenaplic a inner join movimientosbancarios b " & "on a.foliomovto = b.foliomovto " & "where a.estatus <> 'C' and ((" & IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & " and a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")", IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) = "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & ") ", "a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")"))
                                FormaQuery = True
                            ElseIf Trim(subsql) <> "" And Not FormaQuery Then
                                subsql = subsql & " union " & "(select " & Numerico(.get_TextMatrix(I, 5)) & " as codsucursal,'H' as tipomovto,'Gastos' as descripcion," & "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaresconimpuesto," & "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaressinimpuesto," & "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as imppesosconimpuesto," & "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as imppesossinimpuesto," & "b.fechamovto as fecha " & "from movimientosorigenaplic a inner join movimientosbancarios b " & "on a.foliomovto = b.foliomovto " & "where a.estatus <> 'C' and ((" & IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & " and a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")", IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) = "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & ") ", "a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")"))
                                FormaQuery = True
                            ElseIf FormaQuery Then
                                subsql = subsql & " or (" & IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & " and a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & ")", IIf(Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 3)) = "", "a.codorigenaplicr = " & Numerico(.get_TextMatrix(I, 1)) & ") ", "a.codrubro = " & Numerico(.get_TextMatrix(I, 3)) & "))"))
                            End If
                        End If
                    Else
                        If blnExiste Then
                            subsql = subsql & ") group by b.fechamovto)"
                        End If
                        Exit For
                    End If
                Next
            Next
        End With
        Sql = Sql & subsql & ") Tabla"
        DevuelveQuery = Sql
    End Function

    'Function DevuelveQuery() As String
    '    Dim sql As String
    '    Dim I As Long
    '    With flexGastos
    '        If Vacio() Then
    '            DevuelveQuery = ""
    '            Exit Function
    '        End If
    '        For I = 1 To .Rows - 1
    '            If Trim(.TextMatrix(I, 0)) <> "" Or Trim(.TextMatrix(I, 1)) <> "" Or Trim(.TextMatrix(I, 2)) <> "" Or Trim(.TextMatrix(I, 3)) <> "" Or Trim(.TextMatrix(I, 4)) <> "" Then
    '                If I = 1 Then
    '                    sql = "select * into ##TablaTmp from ((select " & Numerico(.TextMatrix(I, 5)) & " as codsucursal,'H' as tipomovto,'Gastos' as descripcion," & _
    ''                    "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaresconimpuesto," & _
    ''                    "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaressinimpuesto," & _
    ''                    "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as importepesosconimpuesto," & _
    ''                    "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as importepesossinimpuesto," & _
    ''                    "b.fechamovto as fecha " & _
    ''                    "from movimientosorigenaplic a inner join movimientosbancarios b " & _
    ''                    "on a.foliomovto = b.foliomovto " & _
    ''                    "where a.estatus <> 'C' " & IIf(Trim(.TextMatrix(I, 1)) <> "", "and a.codorigenaplicr = " & Numerico(.TextMatrix(I, 1)) & " ", "") & _
    ''                    IIf(Trim(.TextMatrix(I, 3)) <> "", "and a.codrubro = " & Numerico(.TextMatrix(I, 3)) & " ", "") & _
    ''                    "group by b.fechamovto)"
    '                Else
    '                    sql = sql & " union " & _
    ''                    "(select " & Numerico(.TextMatrix(I, 5)) & " as codsucursal,'H' as tipomovto,'Gastos' as descripcion," & _
    ''                    "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaresconimpuesto," & _
    ''                    "round(sum(case when b.moneda = 'D' then a.importe when b.moneda = 'P' then a.importe/b.tipocambio end),2) as impdolaressinimpuesto," & _
    ''                    "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as importepesosconimpuesto," & _
    ''                    "round(sum(case when b.moneda = 'P' then a.importe when b.moneda = 'D' then a.importe * b.tipocambio end),1) as importepesossinimpuesto," & _
    ''                    "b.fechamovto as fecha " & _
    ''                    "from movimientosorigenaplic a inner join movimientosbancarios b " & _
    ''                    "on a.foliomovto = b.foliomovto " & _
    ''                    "where a.estatus <> 'C' " & IIf(Trim(.TextMatrix(I, 1)) <> "", "and a.codorigenaplicr = " & Numerico(.TextMatrix(I, 1)) & " ", "") & _
    ''                    IIf(Trim(.TextMatrix(I, 3)) <> "", "and a.codrubro = " & Numerico(.TextMatrix(I, 3)) & " ", "") & _
    ''                    "group by b.fechamovto)"
    '                End If
    '            Else
    '                Exit For
    '            End If
    '        Next
    '    End With
    '    sql = sql & ") Tabla"
    '    DevuelveQuery = sql
    'End Function

    Function Vacio() As Boolean
        Dim I As Integer
        Vacio = False
        With flexGastos
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" Or Trim(.get_TextMatrix(I, 1)) <> "" Or Trim(.get_TextMatrix(I, 2)) <> "" Or Trim(.get_TextMatrix(I, 3)) <> "" Or Trim(.get_TextMatrix(I, 4)) <> "" Then
                    Exit Function
                End If
            Next
        End With
        Vacio = True
    End Function

    Function ValidaGrid() As Boolean
        Dim I As Integer
        ValidaGrid = False
        If Vacio() Then
            ValidaGrid = True
            Exit Function
        End If
        With flexGastos
            If Trim(.get_TextMatrix(1, 0)) = "" And (Trim(.get_TextMatrix(1, 0)) <> "" Or (Trim(.get_TextMatrix(1, 1)) = "" Or Trim(.get_TextMatrix(1, 2)) = "" Or Trim(.get_TextMatrix(1, 3)) = "" Or Trim(.get_TextMatrix(1, 4)) = "")) Then
                Exit Function
                'Trim(.TextMatrix(1, 0)) = "" Or Trim(.TextMatrix(1, 1)) = "" Or Trim(.TextMatrix(1, 2)) = "" Or
            ElseIf (Trim(.get_TextMatrix(1, 0)) <> "" And (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "")) And ((Trim(.get_TextMatrix(1, 3)) = "" Or Trim(.get_TextMatrix(1, 4)) = "")) Then
                Exit Function
            End If
            For I = 2 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" Or Trim(.get_TextMatrix(I, 1)) <> "" Or Trim(.get_TextMatrix(I, 2)) <> "" Or Trim(.get_TextMatrix(I, 3)) <> "" Or Trim(.get_TextMatrix(I, 4)) <> "" Then
                    If Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "" Then
                        If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Or Trim(.get_TextMatrix(I, 2)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Or Trim(.get_TextMatrix(I, 4)) = "" Then Exit Function
                    ElseIf (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "") And (Trim(.get_TextMatrix(1, 3)) = "" And Trim(.get_TextMatrix(1, 4)) = "") Then
                        If (Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Or Trim(.get_TextMatrix(I, 2)) = "") And (Trim(.get_TextMatrix(I, 3)) = "" And Trim(.get_TextMatrix(I, 4)) = "") Then Exit Function
                        If (Trim(.get_TextMatrix(I, 3)) <> "" Or Trim(.get_TextMatrix(I, 4)) <> "") Then Exit Function
                    ElseIf (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "") And (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "") Then
                        If (Trim(.get_TextMatrix(I, 1)) = "" And Trim(.get_TextMatrix(I, 2)) = "") And (Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Or Trim(.get_TextMatrix(I, 4)) = "") Then Exit Function
                        If (Trim(.get_TextMatrix(I, 1)) <> "" Or Trim(.get_TextMatrix(I, 2)) <> "") Then Exit Function
                    End If
                Else
                    ValidaGrid = True
                    Exit Function
                End If
            Next
        End With
    End Function

    Sub Imprime()
        Dim Sql As String
        Dim sql1 As Object
        Dim sql2 As String
        Dim strWhere As String
        Dim Query As String

        On Error GoTo ImprimeErr
        EjecutaExcel = False
        If chkTodaslasSucursales.CheckState = 0 Then
            If CDbl(Numerico(txtCodSucursal.Text)) = 0 Then
                MsgBox("Proporcione el código de una sucursal...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodSucursal.Focus()
                Exit Sub
            End If
            If Trim(dbcSucursal.Text) = "" Then
                MsgBox("Proporcione la descripción de la sucursal...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                dbcSucursal.Focus()
                Exit Sub
            End If
        End If
        If Not ValidaGrid() Then
            MsgBox("No se ha capturado de forma adecuada la información de las cuentas de gastos, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            flexGastos.Focus()
            Exit Sub
        End If
        If optPesos.Checked = True Then
            Moneda = "P"
        ElseIf optDolares.Checked = True Then
            Moneda = "D"
        End If
        If chkIncluirImpuesto.CheckState = 1 Then
            Impuesto = "S"
        ElseIf chkIncluirImpuesto.CheckState = 0 Then
            Impuesto = "N"
        End If
        If chkTodaslasSucursales.CheckState = 0 Then
            strWhere = "WHERE CodSucursal = " & txtCodSucursal.Text
        Else
            strWhere = ""
        End If
        Cmd.CommandTimeout = 300
        cTablaTmpGastos = CreaTablaTemporal("##TablaTmp")
        Query = DevuelveQuery()
        If Len(Query) > 8000 Then
            sql1 = (Query)
            'sql2 = (Query, Len(Convert.ToString(Query)))
        Else
            sql1 = Query
            sql2 = ""
        End If
        '    ModEstandar.BorraCmd
        '    Cmd.CommandText = "dbo.Up_Select_DatosSql"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, Sql1)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, Sql2)
        '    Set RsGral = Cmd.Execute
        '    If RsGral.RecordCount > 0 Then
        cTablaTmpResultados = CreaTablaTemporal("##Resultados")

        If optMensual.Checked = True Then
            ModCorporativo.ObtenerLimitedeFechas(CShort((Trim(cmbMes.Text))), CShort(Trim(cmbAño.Text)), FechaInicial, FechaFinal)
            ModStoredProcedures.PR_EstadodeResultados(FechaInicial, FechaFinal, Moneda, Impuesto, CStr(sql1), CStr(sql2), cTablaTmpResultados)
            Cmd.Execute()
            Sql = "SELECT * FROM " & cTablaTmpResultados & " " & strWhere & "ORDER BY CodSucursal,Mes,Año"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                EnviaExcel()
            Else
                MsgBox("No Existe Información en Este Periodo...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            End If

            '        sql = "SELECT * FROM Dbo.EstadodeResultados('" & FechaInicial & "','" & FechaFinal & "','" & Moneda & "','" & Impuesto & "') " & _
            ''        strWhere & " " & _
            ''        "ORDER BY CodSucursal,Mes,Año"
            '        ModEstandar.BorraCmd
            '        Cmd.CommandText = "dbo.Up_Select_Datos"
            '        Cmd.CommandType = adCmdStoredProc
            '        Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
            '        Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, sql)
            '        Set RsGral = Cmd.Execute
            '        If RsGral.RecordCount > 0 Then
            '            EnviaExcel
            '        Else
            '            MsgBox "No Existe Información en Este Periodo...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
            '        End If
        ElseIf optAnual.Checked = True Then
            ModStoredProcedures.PR_EstadodeResultadosAnual(CStr(cmbAño.Text), CStr(Moneda), CStr(Impuesto), CStr(sql1), CStr(sql2), CStr(cTablaTmpResultados))
            Cmd.Execute()
            Sql = "SELECT * FROM " & cTablaTmpResultados & " " & strWhere & "GROUP BY CodSucursal,Mes,Año,Seccion,TipoMovto,Descripcion,Importe,Porcentaje " & "ORDER BY CodSucursal,Mes,Año"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                EnviaExcel()
            Else
                MsgBox("No Existe Información en Este Periodo...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            End If

            'sql = "SELECT * into " & cTablaTmpResultados & " FROM Dbo.EstadoDeResultadosAnual(" & CInt(cmbAño) & ",'" & Moneda & "','" & Impuesto & "','" & cTablaTmp & "') " & _
            ''strWhere & " " & _
            '"GROUP BY CodSucursal,Mes,Año,Seccion,TipoMovto,Descripcion,Importe,Porcentaje " & _
            '"ORDER BY CodSucursal,Mes,Año"
            'ModEstandar.BorraCmd
            'Cmd.CommandText = "dbo.Up_Select_Datos"
            'Cmd.CommandType = adCmdStoredProc
            'Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
            'Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, sql)
            'Set RsGral = Cmd.Execute
            'If RsGral.RecordCount > 0 Then
            '    EnviaExcel
            'Else
            '    MsgBox "No Existe Información en Este Periodo...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
            'End If

        End If
        'Destruimos la Tabla Temporal
        gStrSql = "DROP TABLE " & Trim(cTablaTmpResultados)
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        Cmd.Execute()
        cTablaTmpResultados = ""
        cTablaTmpGastos = ""
        Cmd.CommandTimeout = 90
        Exit Sub
ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error al Imprimir : " & Err.Description, MsgBoxStyle.Exclamation, "Error de Operacion")
        FueraChange = False
        Me.Cursor = System.Windows.Forms.Cursors.Default
        MDIMenuPrincipalCorpo.Cursor = System.Windows.Forms.Cursors.Default
        If EjecutaExcel Then
            CierraInstanciasdeExcel()
        End If
    End Sub

    Private Sub CambiarFormatoTxtenCaptura()
        With txtFlex
            Select Case flexGastos.Col
                Case 0 'Descripcion de la Sucursal
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                    .MaxLength = 40
                Case 1 'Codigo del Agrupador
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 4
                Case 2 'Descripción del Agrupador
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                    .MaxLength = 40
                Case 3 'Codigo del Rubro
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 6
                Case 4 'Descripción del Rubro
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                    .MaxLength = 40
            End Select
        End With
    End Sub

    Sub CierraInstanciasdeExcel()
        objLibro.Close()
        ObjExcel.Quit()
        ObjExcel = Nothing
        objLibro = Nothing
        objHoja = Nothing
    End Sub

    Sub EnviaExcel()
        On Error GoTo Err_Renamed
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'Me.MousePointer = ccHourglass
        System.Windows.Forms.Application.DoEvents()
        If Dir(gstrCorpoDriveLocal & "\Sistema\", FileAttribute.Directory + FileAttribute.Hidden) = "" Then
            MsgBox("No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If optMensual.Checked Then
            Fecha = (FechaInicial) & "/" & Mid(FechaInicial, 6, 2) & "/" & (FechaInicial)
            FechaAux = CDate(Format(DateSerial(CInt((FechaInicial)), CInt(Mid(FechaInicial, 6, 2)), CInt((FechaInicial)))))
            If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\", FileAttribute.Directory) = "" Then
                MkDir(gstrCorpoDriveLocal & "\Sistema\Informes\")
                'SetAttr App.Path & "\Sistema\Informes\", vbHidden
            End If
            If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\ERM" & MesLetra(FechaAux) & (cmbAño.Text) & ".xls", FileAttribute.Archive) <> "" Then
                Kill(gstrCorpoDriveLocal & "\Sistema\Informes\ERM" & MesLetra(FechaAux) & (cmbAño.Text) & ".xls")
            End If
            ObjExcel = CreateObject("Excel.Application")
            objLibro = ObjExcel.Workbooks.Add
            'Set objHoja = objLibro.Worksheets.Add
            objHoja = objLibro.ActiveSheet
            ObjExcel.Visible = False
            objLibro.Sheets(1).Select()
            objHoja = objLibro.ActiveSheet
            objLibro.ActiveSheet.Name = "Edo. Resultados Mensual"
            EjecutaExcel = True
            Encabezado()
            LlenaDatos()
            objLibro.SaveAs(gstrCorpoDriveLocal & "\Sistema\Informes\ERM" & MesLetra(FechaAux) & (cmbAño.Text) & ".xls", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)
            'Me.MousePointer = ccDefault
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            System.Windows.Forms.Application.DoEvents()
            Select Case MsgBox("Se ha creado el archivo " & "ERM" & MesLetra(FechaAux) & (cmbAño.Text) & ".xls ¿Desea abrirlo?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    ObjExcel.Visible = True
                    ObjExcel = Nothing
                    objLibro = Nothing
                    objHoja = Nothing
                Case MsgBoxResult.No Or MsgBoxResult.Cancel
                    CierraInstanciasdeExcel()
            End Select
        ElseIf optAnual.Checked Then
            If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\", FileAttribute.Directory) = "" Then
                MkDir(gstrCorpoDriveLocal & "\Sistema\Informes\")
                'SetAttr App.Path & "\Sistema\Informes\", vbHidden
            End If
            If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\ERA_" & cmbAño.Text & ".xls", FileAttribute.Archive) <> "" Then
                Kill(gstrCorpoDriveLocal & "\Sistema\Informes\ERA_" & cmbAño.Text & ".xls")
            End If
            ObjExcel = CreateObject("Excel.Application")
            objLibro = ObjExcel.Workbooks.Add
            objHoja = objLibro.Worksheets.Add
            objHoja = objLibro.ActiveSheet
            ObjExcel.Visible = False
            objLibro.Sheets(1).Select()
            objHoja = objLibro.ActiveSheet
            objLibro.ActiveSheet.Name = "Edo. Resultados Anual"
            Encabezado()
            LlenaDatos()
            objLibro.SaveAs(gstrCorpoDriveLocal & "\Sistema\Informes\ERA_" & cmbAño.Text & ".xls", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)
            'Me.MousePointer = ccDefault
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            System.Windows.Forms.Application.DoEvents()
            Select Case MsgBox("Se ha creado el archivo " & "ERA_" & cmbAño.Text & ".xls ¿Desea abrirlo?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    ObjExcel.Visible = True
                    ObjExcel = Nothing
                    objLibro = Nothing
                    objHoja = Nothing
                Case MsgBoxResult.No Or MsgBoxResult.Cancel
                    CierraInstanciasdeExcel()
            End Select
        End If
Err_Renamed:
        If Err.Number = 70 Then
            MsgBox("No se puede generar en nuevo archivo " & IIf(optMensual.Checked, "Mensual", "Anual") & " hasta que el anterior este cerrado.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            'Me.MousePointer = vbDefault
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        ElseIf Err.Number <> 0 Then
            ModEstandar.MostrarError()
            'Me.MousePointer = vbDefault
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Encabezado()
        On Error GoTo Err_Renamed
        With objHoja
            .Range("B1").FormulaR1C1 = UCase(Trim(gstrCorpoNOMBREEMPRESA))
            .Range("B1:F1").Select()
            .Range("B1:F1").MergeCells = True
            .Range("B1:F1").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("B1:F1").Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With
            Fecha = (FechaInicial) & "/" & Mid(FechaInicial, 6, 2) & "/" & (FechaInicial)
            Fecha = Format(Fecha, "dd/mmm/yyyy")
            .Range("B2").FormulaR1C1 = "ESTADO DE RESULTADOS"
            .Range("B2:F2").Select()
            .Range("B2:F2").MergeCells = True
            .Range("B2:F2").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("B2:F2").Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With
            If optMensual.Checked Then
                .Range("B3").FormulaR1C1 = "Mes de " & Trim(Mid(cmbMes.Text, 6, 12)) & " del " & cmbAño.Text
            ElseIf optAnual.Checked Then
                .Range("B3").FormulaR1C1 = "Año " & cmbAño.Text
            End If
            .Range("B3:F3").Select()
            .Range("B3:F3").MergeCells = True
            .Range("B3:F3").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("B3:F3").Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range("B4").FormulaR1C1 = "COMPARATIVO"
            .Range("B4:F4").Select()
            .Range("B4:F4").MergeCells = True
            .Range("B4:F4").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("B4:F4").Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range("B6").FormulaR1C1 = "Moneda : " & IIf(Moneda = "D", "Dólares", "Pesos")
            .Range("B6:C6").Select()
            .Range("B6:C6").MergeCells = True
            .Range("B6:C6").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("B6:C6").Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range("E6").FormulaR1C1 = "Fecha : " & Format(Today, "dd/mmm/yyyy")
            .Range("E6:F6").Select()
            .Range("E6:F6").MergeCells = True
            .Range("E6:F6").HorizontalAlignment = Excel.Constants.xlRight
            With .Range("E6:F6").Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range("B7")._Default = IIf(Impuesto = "S", "* Estos importes incluyen iva.", "* Estos importes no incluyen iva.")
            .Range("B7:D7").Select()
            .Range("B7:D7").MergeCells = True
            .Range("B7:D7").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("B7:D7").Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
        End With
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            MDIMenuPrincipalCorpo.Cursor = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Conceptos()
        On Error GoTo Err_Renamed
        With objHoja
            Rango = "B" & Renglon
            .Range(Rango)._Default = "CONCEPTOS"
            Rango = "B" & Renglon & ":" & "B" & Renglon + 1
            .Range(Rango).Select()
            With .Range(Rango)
                .ColumnWidth = 25.29
                .MergeCells = True
                .Interior.ColorIndex = 15
                .HorizontalAlignment = Excel.Constants.xlCenter
                .VerticalAlignment = Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            Rango = "B" & Renglon + 2
            .Range(Rango)._Default = "SALIDA DE MERCANCIA"
            Rango = "B" & Renglon + 2 & ":" & "B" & Renglon + 2
            .Range(Rango).Select()
            With .Range(Rango)
                .MergeCells = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            Rango = "B" & Renglon + 3
            .Range(Rango)._Default = "COSTO"
            Rango = "B" & Renglon + 3 & ":" & "B" & Renglon + 3
            .Range(Rango).Select()
            With .Range(Rango)
                .MergeCells = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            Rango = "B" & Renglon + 4
            .Range(Rango)._Default = "UTILIDAD NETA"
            Rango = "B" & Renglon + 4 & ":" & "B" & Renglon + 4
            .Range(Rango).Select()
            With .Range(Rango)
                .MergeCells = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            Rango = "B" & Renglon + 5
            .Range(Rango)._Default = "GASTOS"
            Rango = "B" & Renglon + 5 & ":" & "B" & Renglon + 5
            .Range(Rango).Select()
            With .Range(Rango)
                .MergeCells = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            Rango = "B" & Renglon + 6
            .Range(Rango)._Default = "UTILIDAD BRUTA"
            Rango = "B" & Renglon + 6 & ":" & "B" & Renglon + 6
            .Range(Rango).Select()
            With .Range(Rango)
                .MergeCells = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            Rango = "B" & Renglon + 8
            .Range(Rango)._Default = "VENTAS INGRESOS"
            Rango = "B" & Renglon + 8 & ":" & "B" & Renglon + 8
            .Range(Rango).Select()
            With .Range(Rango)
                .MergeCells = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            Rango = "B" & Renglon + 9
            .Range(Rango)._Default = "COSTOS INGRESOS"
            Rango = "B" & Renglon + 9 & ":" & "B" & Renglon + 9
            .Range(Rango).Select()
            With .Range(Rango)
                .MergeCells = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            Rango = "B" & Renglon + 10
            .Range(Rango)._Default = "UTILIDAD NETA"
            Rango = "B" & Renglon + 10 & ":" & "B" & Renglon + 10
            .Range(Rango).Select()
            With .Range(Rango)
                .MergeCells = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            Rango = "B" & Renglon + 11
            .Range(Rango)._Default = "GASTOS"
            Rango = "B" & Renglon + 11 & ":" & "B" & Renglon + 11
            .Range(Rango).Select()
            With .Range(Rango)
                .MergeCells = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            Rango = "B" & Renglon + 12
            .Range(Rango)._Default = "UTILIDAD BRUTA"
            Rango = "B" & Renglon + 12 & ":" & "B" & Renglon + 12
            .Range(Rango).Select()
            With .Range(Rango)
                .MergeCells = True
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
        End With
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            MDIMenuPrincipalCorpo.Cursor = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Cuadro()
        On Error GoTo Err_Renamed
        Rango = "B" & Renglon - 2
        objHoja.Range(Rango).FormulaR1C1 = UCase(ObtenerSucursal(CodSucursal))
        If optMensual.Checked Then
            Rango = "B" & Renglon - 2 & ":" & "F" & Renglon - 2
            objHoja.Range(Rango).Select()
            objHoja.Range(Rango).MergeCells = True
            objHoja.Range(Rango).Font.Bold = True
            objHoja.Range(Rango).Font.Size = 8
            objHoja.Range(Rango).Font.Name = "Arial"
            Rango = "B" & Renglon & ":" & "F" & Renglon + 6
            With objHoja
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "B" & Renglon & ":" & "B" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon & ":" & "F" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon & ":" & "F" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon + 1 & ":" & "C" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "E" & Renglon + 1 & ":" & "E" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "B" & Renglon + 2 & ":" & "B" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon + 1 & ":" & "D" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "B" & Renglon + 8 & ":" & "F" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "B" & Renglon + 8 & ":" & "B" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon + 8 & ":" & "D" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon
                .Range(Rango).FormulaR1C1 = UCase(MesLetra(CDate(FechaAux), False))
                Rango = "C" & Renglon & ":" & "F" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "C" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                Rango = "C" & Renglon + 1 & ":" & "C" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "D" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                Rango = "D" & Renglon + 1 & ":" & "D" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "E" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                Rango = "E" & Renglon + 1 & ":" & "E" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "F" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                Rango = "F" & Renglon + 1 & ":" & "F" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
            End With
            Conceptos()
        ElseIf optAnual.Checked Then
            Rango = "B" & Renglon - 2 & ":" & "F" & Renglon - 2
            objHoja.Range(Rango).Select()
            objHoja.Range(Rango).MergeCells = True
            objHoja.Range(Rango).Font.Bold = True
            objHoja.Range(Rango).Font.Size = 8
            objHoja.Range(Rango).Font.Name = "Arial"
            Rango = "B" & Renglon & ":" & "BB" & Renglon + 6
            With objHoja
                'Dibujo el Cuadro de Arriba
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "B" & Renglon & ":" & "BB" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "B" & Renglon & ":" & "B" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon & ":" & "BB" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "B" & Renglon & ":" & "B" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon + 1 & ":" & "D" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "E" & Renglon + 1 & ":" & "F" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "G" & Renglon + 1 & ":" & "H" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "I" & Renglon + 1 & ":" & "J" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "K" & Renglon + 1 & ":" & "L" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "M" & Renglon + 1 & ":" & "N" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "O" & Renglon + 1 & ":" & "P" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "Q" & Renglon + 1 & ":" & "R" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "S" & Renglon + 1 & ":" & "T" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "U" & Renglon + 1 & ":" & "V" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "W" & Renglon + 1 & ":" & "X" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "Y" & Renglon + 1 & ":" & "Z" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AA" & Renglon + 1 & ":" & "AB" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AC" & Renglon + 1 & ":" & "AD" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AE" & Renglon + 1 & ":" & "AF" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AG" & Renglon + 1 & ":" & "AH" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AI" & Renglon + 1 & ":" & "AJ" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AK" & Renglon + 1 & ":" & "AL" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AM" & Renglon + 1 & ":" & "AN" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AO" & Renglon + 1 & ":" & "AP" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AQ" & Renglon + 1 & ":" & "AR" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AS" & Renglon + 1 & ":" & "AT" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AU" & Renglon + 1 & ":" & "AV" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AW" & Renglon + 1 & ":" & "AX" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AY" & Renglon + 1 & ":" & "AZ" & Renglon + 6
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon + 1 & ":" & "C" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "E" & Renglon + 1 & ":" & "E" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "G" & Renglon + 1 & ":" & "G" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "I" & Renglon + 1 & ":" & "I" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "K" & Renglon + 1 & ":" & "K" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "M" & Renglon + 1 & ":" & "M" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "O" & Renglon + 1 & ":" & "O" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "Q" & Renglon + 1 & ":" & "Q" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "S" & Renglon + 1 & ":" & "S" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "U" & Renglon + 1 & ":" & "U" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "W" & Renglon + 1 & ":" & "W" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "Y" & Renglon + 1 & ":" & "Y" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AA" & Renglon + 1 & ":" & "AA" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AC" & Renglon + 1 & ":" & "AC" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AE" & Renglon + 1 & ":" & "AE" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AG" & Renglon + 1 & ":" & "AG" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AI" & Renglon + 1 & ":" & "AI" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AK" & Renglon + 1 & ":" & "AK" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AM" & Renglon + 1 & ":" & "AM" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AO" & Renglon + 1 & ":" & "AO" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AQ" & Renglon + 1 & ":" & "AQ" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AS" & Renglon + 1 & ":" & "AS" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AU" & Renglon + 1 & ":" & "AU" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AW" & Renglon + 1 & ":" & "AW" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AY" & Renglon + 1 & ":" & "AY" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "BA" & Renglon + 1 & ":" & "BA" & Renglon + 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                'Dibujo el Cuadro de Abajo
                Rango = "B" & Renglon + 8 & ":" & "BB" & Renglon + 12
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "B" & Renglon + 8 & ":" & "B" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon + 8 & ":" & "D" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "E" & Renglon + 8 & ":" & "F" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "G" & Renglon + 8 & ":" & "H" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "I" & Renglon + 8 & ":" & "J" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "K" & Renglon + 8 & ":" & "L" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "M" & Renglon + 8 & ":" & "N" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "O" & Renglon + 8 & ":" & "P" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "Q" & Renglon + 8 & ":" & "R" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "S" & Renglon + 8 & ":" & "T" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "U" & Renglon + 8 & ":" & "V" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "W" & Renglon + 8 & ":" & "X" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "Y" & Renglon + 8 & ":" & "Z" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AA" & Renglon + 8 & ":" & "AB" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AC" & Renglon + 8 & ":" & "AD" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AE" & Renglon + 8 & ":" & "AF" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AG" & Renglon + 8 & ":" & "AH" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AI" & Renglon + 8 & ":" & "AJ" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AK" & Renglon + 8 & ":" & "AL" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AM" & Renglon + 8 & ":" & "AN" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AO" & Renglon + 8 & ":" & "AP" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AQ" & Renglon + 8 & ":" & "AR" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AS" & Renglon + 8 & ":" & "AT" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AU" & Renglon + 8 & ":" & "AV" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AW" & Renglon + 8 & ":" & "AX" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AY" & Renglon + 8 & ":" & "AZ" & Renglon + 12
                .Range(Rango).Select()
                With .Range(Rango)
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                ''''''''''''''''''''''''''''''''''''''''''''''''
                'Enero
                Rango = "C" & Renglon
                .Range(Rango)._Default = "ENERO"
                Rango = "C" & Renglon & ":" & "F" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "C" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "D" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "E" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "F" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Febrero
                Rango = "G" & Renglon
                .Range(Rango)._Default = "FEBRERO"
                Rango = "G" & Renglon & ":" & "J" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "G" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "H" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "I" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "J" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Marzo
                Rango = "K" & Renglon
                .Range(Rango)._Default = "MARZO"
                Rango = "K" & Renglon & ":" & "N" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "K" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "L" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "M" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "N" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Abril
                Rango = "O" & Renglon
                .Range(Rango)._Default = "ABRIL"
                Rango = "O" & Renglon & ":" & "R" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "O" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "P" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "Q" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "R" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Mayo
                Rango = "S" & Renglon
                .Range(Rango)._Default = "MAYO"
                Rango = "S" & Renglon & ":" & "V" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "S" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "T" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "U" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "V" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Junio
                Rango = "W" & Renglon
                .Range(Rango)._Default = "JUNIO"
                Rango = "W" & Renglon & ":" & "Z" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "W" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "X" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "Y" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "Z" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Julio
                Rango = "AA" & Renglon
                .Range(Rango)._Default = "JULIO"
                Rango = "AA" & Renglon & ":" & "AD" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AA" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AB" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AC" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AD" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Agosto
                Rango = "AE" & Renglon
                .Range(Rango)._Default = "AGOSTO"
                Rango = "AE" & Renglon & ":" & "AH" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AE" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AF" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AG" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AH" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Septiembre
                Rango = "AI" & Renglon
                .Range(Rango)._Default = "SEPTIEMBRE"
                Rango = "AI" & Renglon & ":" & "AL" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AI" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AJ" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AK" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AL" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Octubre
                Rango = "AM" & Renglon
                .Range(Rango)._Default = "OCTUBRE"
                Rango = "AM" & Renglon & ":" & "AP" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AM" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AN" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AO" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AP" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Noviembre
                Rango = "AQ" & Renglon
                .Range(Rango)._Default = "NOVIEMBRE"
                Rango = "AQ" & Renglon & ":" & "AT" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AQ" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AR" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AS" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AT" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Diciembre
                Rango = "AU" & Renglon
                .Range(Rango)._Default = "DICIEMBRE"
                Rango = "AU" & Renglon & ":" & "AX" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
                Rango = "AU" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AV" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AW" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AX" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                'Acumulado
                Rango = "AY" & Renglon
                .Range(Rango)._Default = "ACUMULADO"
                Rango = "AY" & Renglon & ":" & "BB" & Renglon
                .Range(Rango).Select()
                With .Range(Rango)
                    .MergeCells = True
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AY" & Renglon + 1
                .Range(Rango).FormulaR1C1 = CShort(cmbAño.Text) - 1
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "AZ" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "BA" & Renglon + 1
                .Range(Rango).FormulaR1C1 = cmbAño
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 16.86
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
                Rango = "BB" & Renglon + 1
                .Range(Rango).FormulaR1C1 = "Porcentaje"
                .Range(Rango).Select()
                With .Range(Rango)
                    .ColumnWidth = 8.3
                    .Interior.ColorIndex = 15
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    .Font.Name = "Arial"
                End With
            End With
            Conceptos()
        End If
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            MDIMenuPrincipalCorpo.Cursor = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Function LlenaSucursal() As Boolean
        On Error GoTo Err_Renamed
        LlenaSucursal = False
        gStrSql = "SELECT * FROM CatAlmacen WHERE DescAlmacen like '" & Trim(txtFlex.Text) & "%' and TipoAlmacen = 'P'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFlex.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            flexGastos.set_TextMatrix(flexGastos.Row, 5, RsGral.Fields("CodAlmacen").Value)
            LlenaSucursal = True
            txtFlex_Leave(txtFlex, New System.EventArgs())
        Else
            MsgBox("Descripción inexistente, Favor de verificar..", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            flexGastos.Col = 0
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function DescripcionAgrupador() As Boolean
        On Error GoTo Err_Renamed
        DescripcionAgrupador = False
        'If Trim(flexgastos.TextMatrix(flexgastos.Row, 3)) = "" Then
        gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR like '" & Trim(txtFlex.Text) & "%'"
        'End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFlex.Text = Trim(RsGral.Fields("DescOrigenAplicR").Value)
            flexGastos.set_TextMatrix(flexGastos.Row, 1, ("0000" & CStr(RsGral.Fields("CodOrigenAplicR").Value)))
            If Trim(flexGastos.get_TextMatrix(flexGastos.Row, 2)) <> Trim(txtFlex.Text) Then
                flexGastos.set_TextMatrix(flexGastos.Row, 3, "")
                flexGastos.set_TextMatrix(flexGastos.Row, 4, "")
            End If
            DescripcionAgrupador = True
            txtFlex_Leave(txtFlex, New System.EventArgs())
        Else
            MsgBox("Descripción Inexistente Favor de Verificar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            flexGastos.Col = 2
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function DescripcionRubro() As Boolean
        On Error GoTo Err_Renamed
        DescripcionRubro = False
        If Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) = "" Then
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos " & "WHERE DescRubro like '" & Trim(txtFlex.Text) & "%' GROUP BY CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion"
        ElseIf Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) <> "" Then
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos " & "WHERE DescRubro like '" & Trim(txtFlex.Text) & "%' and codorigaplicr = " & Numerico(Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1))) & " GROUP BY CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion"
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFlex.Text = Trim(RsGral.Fields("DescRubro").Value)
            flexGastos.set_TextMatrix(flexGastos.Row, 3, ("000000" & CStr(RsGral.Fields("CodRubro").Value)))
            txtFlex_Leave(txtFlex, New System.EventArgs())
            DescripcionRubro = True
        Else
            MsgBox("Descripción Inexistente Favor de Investigar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub LlenaDatos()
        On Error GoTo Err_Renamed
        Dim NumPartida As Integer
        Dim Acumulado As Double
        Dim Porcentaje As Double

        RsGral.MoveFirst()
        CodSucursal = RsGral.Fields("CodSucursal").Value
        Año = RsGral.Fields("Año").Value
        Renglon = 12
        Cuadro()
        NumPartida = 2
        If optMensual.Checked Then
            Do While Not RsGral.EOF
                If NumPartida = 7 Then
                    NumPartida = 8
                End If
                If NumPartida = 13 Then
                    NumPartida = 2
                End If
                With objHoja
                    If CodSucursal <> RsGral.Fields("CodSucursal").Value Then
                        CodSucursal = RsGral.Fields("CodSucursal").Value
                        Año = RsGral.Fields("Año").Value
                        Renglon = Renglon + 17
                        Cuadro()
                    End If
                    If RsGral.Fields("Año").Value = cmbAño.Text Then
                        Rango = "E" & Renglon + NumPartida
                        If RsGral.Fields("importe").Value <> 0 Then
                            .Range(Rango).NumberFormat = "###,##0.00"
                            .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                        Else
                            .Range(Rango).FormulaR1C1 = 0
                            .Range(Rango).NumberFormat = "###,##0.00"
                        End If
                        Rango = "E" & Renglon + NumPartida & ":" & "E" & Renglon + NumPartida
                        .Range(Rango).Select()
                        With .Range(Rango)
                            .HorizontalAlignment = Excel.Constants.xlRight
                            If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                .Font.Underline = True
                            End If
                            .Font.Size = 8
                            .Font.Name = "Arial"
                            If RsGral.Fields("importe").Value >= 0 Then
                                .Font.ColorIndex = 1
                            ElseIf RsGral.Fields("importe").Value < 0 Then
                                .Font.ColorIndex = 3
                            End If
                        End With
                        If NumPartida <> 2 And NumPartida <> 8 Then
                            Rango = "F" & Renglon + NumPartida
                            .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                            Rango = "F" & Renglon + NumPartida & ":" & "F" & Renglon + NumPartida
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                .Font.Size = 8
                                .Font.Name = "Arial"
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                        End If
                    ElseIf RsGral.Fields("Año").Value = (CShort(cmbAño.Text) - 1) Then
                        Rango = "C" & Renglon + NumPartida
                        If RsGral.Fields("importe").Value <> 0 Then
                            .Range(Rango).NumberFormat = "###,##0.00"
                            .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                        Else
                            .Range(Rango).FormulaR1C1 = 0
                            .Range(Rango).NumberFormat = "###,##0.00"
                        End If
                        Rango = "C" & Renglon + NumPartida & ":" & "C" & Renglon + NumPartida
                        .Range(Rango).Select()
                        With .Range(Rango)
                            .HorizontalAlignment = Excel.Constants.xlRight
                            If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                .Font.Underline = True
                            End If
                            .Font.Size = 8
                            .Font.Name = "Arial"
                            If RsGral.Fields("importe").Value >= 0 Then
                                .Font.ColorIndex = 1
                            ElseIf RsGral.Fields("importe").Value < 0 Then
                                .Font.ColorIndex = 3
                            End If
                        End With
                        If NumPartida <> 2 And NumPartida <> 8 Then
                            Rango = "D" & Renglon + NumPartida
                            .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                            Rango = "D" & Renglon + NumPartida & ":" & "D" & Renglon + NumPartida
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                        End If
                    End If
                End With
                NumPartida = NumPartida + 1
                RsGral.MoveNext()
            Loop
            objHoja.Range("A1").Select()
        ElseIf optAnual.Checked = True Then
            Do While Not RsGral.EOF
                If NumPartida = 7 Then
                    NumPartida = 8
                End If
                If NumPartida = 13 Then
                    NumPartida = 2
                End If
                With objHoja
                    If CodSucursal <> RsGral.Fields("CodSucursal").Value Then
                        CodSucursal = RsGral.Fields("CodSucursal").Value
                        Año = RsGral.Fields("Año").Value
                        Renglon = Renglon + 17
                        Cuadro()
                    End If
                    If RsGral.Fields("Año").Value = cmbAño.Text Then
                        If RsGral.Fields("Mes").Value = 1 Then
                            Rango = "E" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "F" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 2 Then
                            Rango = "I" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "J" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 3 Then
                            Rango = "M" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "N" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 4 Then
                            Rango = "Q" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "R" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 5 Then
                            Rango = "U" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                If RsGral.Fields("importe").Value >= 1000 Then
                                    .Range(Rango).NumberFormat = "###,##0.00"
                                    .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                                ElseIf RsGral.Fields("importe").Value < 1000 Then
                                    .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("importe").Value, "###,##0.00")
                                End If
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "V" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 6 Then
                            Rango = "Y" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "Z" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 7 Then
                            Rango = "AC" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 And NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AD" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 8 Then
                            Rango = "AG" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AH" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 9 Then
                            Rango = "AK" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AL" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 10 Then
                            Rango = "AO" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AP" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 11 Then
                            Rango = "AS" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AT" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 12 Then
                            Rango = "AW" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AX" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If

                            '''ACUMULADO
                            Acumulado = CalculaAcumulado(MesAcumulado, Renglon, NumPartida, True)

                            Rango = "BA" & Renglon + NumPartida
                            If Acumulado <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = Acumulado
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                .Font.Size = 8
                                If Acumulado >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf Acumulado < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            Acumulado = 0
                            If NumPartida = 3 Then
                                If .Range("BA" & (Renglon + (NumPartida - 1)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("BA" & (Renglon + NumPartida))._Default / .Range("BA" & (Renglon + (NumPartida - 1)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 4 Then
                                If .Range("BA" & (Renglon + (NumPartida - 2)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("BA" & (Renglon + NumPartida))._Default / .Range("BA" & (Renglon + (NumPartida - 2)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 5 Then
                                If .Range("BA" & (Renglon + (NumPartida - 3)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("BA" & (Renglon + NumPartida))._Default / .Range("BA" & (Renglon + (NumPartida - 3)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 6 Then
                                If .Range("BA" & (Renglon + (NumPartida - 4)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("BA" & (Renglon + NumPartida))._Default / .Range("BA" & (Renglon + (NumPartida - 4)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 9 Then
                                If .Range("BA" & (Renglon + (NumPartida - 1)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("BA" & (Renglon + NumPartida))._Default / .Range("BA" & (Renglon + (NumPartida - 1)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 10 Then
                                If .Range("BA" & (Renglon + (NumPartida - 2)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("BA" & (Renglon + NumPartida))._Default / .Range("BA" & (Renglon + (NumPartida - 2)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 11 Then
                                If .Range("BA" & (Renglon + (NumPartida - 3)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("BA" & (Renglon + NumPartida))._Default / .Range("BA" & (Renglon + (NumPartida - 3)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 12 Then
                                If .Range("BA" & (Renglon + (NumPartida - 4)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("BA" & (Renglon + NumPartida))._Default / .Range("BA" & (Renglon + (NumPartida - 4)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 3 Or NumPartida = 4 Or NumPartida = 5 Or NumPartida = 6 Or NumPartida = 9 Or NumPartida = 10 Or NumPartida = 11 Or NumPartida = 12 Then
                                Rango = "BB" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(Porcentaje, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If Porcentaje >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf Porcentaje < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        End If
                    ElseIf RsGral.Fields("Año").Value = (CShort(cmbAño.Text) - 1) Then
                        If RsGral.Fields("Mes").Value = 1 Then
                            Rango = "C" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "D" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 2 Then
                            Rango = "G" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "H" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 3 Then
                            Rango = "K" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "L" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 4 Then
                            Rango = "O" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "P" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 5 Then
                            Rango = "S" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "T" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 6 Then
                            Rango = "W" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "X" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 7 Then
                            Rango = "AA" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AB" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 8 Then
                            Rango = "AE" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AF" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 9 Then
                            Rango = "AI" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AJ" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 10 Then
                            Rango = "AM" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AN" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 11 Then
                            Rango = "AQ" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AR" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        ElseIf RsGral.Fields("Mes").Value = 12 Then
                            Rango = "AU" & Renglon + NumPartida
                            If RsGral.Fields("importe").Value <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = RsGral.Fields("importe")
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                If NumPartida = 3 Or NumPartida = 5 Or NumPartida = 9 Or NumPartida = 11 Then
                                    .Font.Underline = True
                                End If
                                .Font.Size = 8
                                If RsGral.Fields("importe").Value >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf RsGral.Fields("importe").Value < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            If NumPartida <> 2 And NumPartida <> 8 Then
                                Rango = "AV" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(RsGral.Fields("Porcentaje").Value, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If RsGral.Fields("importe").Value >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf RsGral.Fields("importe").Value < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If

                            '''ACUMULADO
                            Acumulado = CalculaAcumulado(MesAcumulado, Renglon, NumPartida, False)

                            Rango = "AY" & Renglon + NumPartida
                            If Acumulado <> 0 Then
                                .Range(Rango).NumberFormat = "###,##0.00"
                                .Range(Rango).FormulaR1C1 = Acumulado
                            Else
                                .Range(Rango).FormulaR1C1 = 0
                                .Range(Rango).NumberFormat = "###,##0.00"
                            End If
                            .Range(Rango).Select()
                            With .Range(Rango)
                                .HorizontalAlignment = Excel.Constants.xlRight
                                .Font.Size = 8
                                If Acumulado >= 0 Then
                                    .Font.ColorIndex = 1
                                ElseIf Acumulado < 0 Then
                                    .Font.ColorIndex = 3
                                End If
                            End With
                            Acumulado = 0
                            If NumPartida = 3 Then
                                If .Range("AY" & (Renglon + (NumPartida - 1)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("AY" & (Renglon + NumPartida))._Default / .Range("AY" & (Renglon + (NumPartida - 1)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 4 Then
                                If .Range("AY" & (Renglon + (NumPartida - 2)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("AY" & (Renglon + NumPartida))._Default / .Range("AY" & (Renglon + (NumPartida - 2)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 5 Then
                                If .Range("AY" & (Renglon + (NumPartida - 3)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("AY" & (Renglon + NumPartida))._Default / .Range("AY" & (Renglon + (NumPartida - 3)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 6 Then
                                If .Range("AY" & (Renglon + (NumPartida - 4)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("AY" & (Renglon + NumPartida))._Default / .Range("AY" & (Renglon + (NumPartida - 4)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 9 Then
                                If .Range("AY" & (Renglon + (NumPartida - 1)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("AY" & (Renglon + NumPartida))._Default / .Range("AY" & (Renglon + (NumPartida - 1)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 10 Then
                                If .Range("AY" & (Renglon + (NumPartida - 2)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("AY" & (Renglon + NumPartida))._Default / .Range("AY" & (Renglon + (NumPartida - 2)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 11 Then
                                If .Range("AY" & (Renglon + (NumPartida - 3)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("AY" & (Renglon + NumPartida))._Default / .Range("AY" & (Renglon + (NumPartida - 3)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 12 Then
                                If .Range("AY" & (Renglon + (NumPartida - 4)))._Default <> 0 Then
                                    Porcentaje = System.Math.Round((.Range("AY" & (Renglon + NumPartida))._Default / .Range("AY" & (Renglon + (NumPartida - 4)))._Default) * 100, 2)
                                Else
                                    Porcentaje = 0
                                End If
                            End If
                            If NumPartida = 3 Or NumPartida = 4 Or NumPartida = 5 Or NumPartida = 6 Or NumPartida = 9 Or NumPartida = 10 Or NumPartida = 11 Or NumPartida = 12 Then
                                Rango = "AZ" & Renglon + NumPartida
                                .Range(Rango).FormulaR1C1 = Format(Porcentaje, "###,##0.00") & "%"
                                .Range(Rango).Select()
                                With .Range(Rango)
                                    .HorizontalAlignment = Excel.Constants.xlRight
                                    .Font.Size = 8
                                    If Porcentaje >= 0 Then
                                        .Font.ColorIndex = 1
                                    ElseIf Porcentaje < 0 Then
                                        .Font.ColorIndex = 3
                                    End If
                                End With
                            End If
                        End If
                    End If
                End With
                NumPartida = NumPartida + 1
                RsGral.MoveNext()
            Loop
            objHoja.Range("A1").Select()
        End If

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            MDIMenuPrincipalCorpo.Cursor = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Function CalculaAcumulado(ByRef Mes As Integer, ByRef Ren As Integer, ByRef NumP As Integer, ByRef Actual As Boolean) As Decimal
        Dim lRango As String
        Dim lAcumulado As Decimal

        With objHoja
            Select Case Mes
                Case 1
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 2
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 3
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "M" & Ren + NumP Else lRango = "K" & Ren + NumP
                    'lRango = "M" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 4
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "M" & Ren + NumP Else lRango = "K" & Ren + NumP
                    'lRango = "M" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Q" & Ren + NumP Else lRango = "O" & Ren + NumP
                    'lRango = "Q" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 5
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "M" & Ren + NumP Else lRango = "K" & Ren + NumP
                    'lRango = "M" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Q" & Ren + NumP Else lRango = "O" & Ren + NumP
                    'lRango = "Q" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "U" & Ren + NumP Else lRango = "S" & Ren + NumP
                    'lRango = "U" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 6
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "M" & Ren + NumP Else lRango = "K" & Ren + NumP
                    'lRango = "M" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Q" & Ren + NumP Else lRango = "O" & Ren + NumP
                    'lRango = "Q" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "U" & Ren + NumP Else lRango = "S" & Ren + NumP
                    'lRango = "U" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Y" & Ren + NumP Else lRango = "W" & Ren + NumP
                    'lRango = "Y" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 7
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "M" & Ren + NumP Else lRango = "K" & Ren + NumP
                    'lRango = "M" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Q" & Ren + NumP Else lRango = "O" & Ren + NumP
                    'lRango = "Q" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "U" & Ren + NumP Else lRango = "S" & Ren + NumP
                    'lRango = "U" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Y" & Ren + NumP Else lRango = "W" & Ren + NumP
                    'lRango = "Y" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AC" & Ren + NumP Else lRango = "AA" & Ren + NumP
                    'lRango = "AC" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 8
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "M" & Ren + NumP Else lRango = "K" & Ren + NumP
                    'lRango = "M" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Q" & Ren + NumP Else lRango = "O" & Ren + NumP
                    'lRango = "Q" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "U" & Ren + NumP Else lRango = "S" & Ren + NumP
                    'lRango = "U" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Y" & Ren + NumP Else lRango = "W" & Ren + NumP
                    'lRango = "Y" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AC" & Ren + NumP Else lRango = "AA" & Ren + NumP
                    'lRango = "AC" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AG" & Ren + NumP Else lRango = "AE" & Ren + NumP
                    'lRango = "AG" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 9
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "M" & Ren + NumP Else lRango = "K" & Ren + NumP
                    'lRango = "M" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Q" & Ren + NumP Else lRango = "O" & Ren + NumP
                    'lRango = "Q" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "U" & Ren + NumP Else lRango = "S" & Ren + NumP
                    'lRango = "U" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Y" & Ren + NumP Else lRango = "W" & Ren + NumP
                    'lRango = "Y" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AC" & Ren + NumP Else lRango = "AA" & Ren + NumP
                    'lRango = "AC" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AG" & Ren + NumP Else lRango = "AE" & Ren + NumP
                    'lRango = "AG" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AK" & Ren + NumP Else lRango = "AI" & Ren + NumP
                    'lRango = "AK" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 10
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "M" & Ren + NumP Else lRango = "K" & Ren + NumP
                    'lRango = "M" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Q" & Ren + NumP Else lRango = "O" & Ren + NumP
                    'lRango = "Q" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "U" & Ren + NumP Else lRango = "S" & Ren + NumP
                    'lRango = "U" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Y" & Ren + NumP Else lRango = "W" & Ren + NumP
                    'lRango = "Y" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AC" & Ren + NumP Else lRango = "AA" & Ren + NumP
                    'lRango = "AC" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AG" & Ren + NumP Else lRango = "AE" & Ren + NumP
                    'lRango = "AG" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AK" & Ren + NumP Else lRango = "AI" & Ren + NumP
                    'lRango = "AK" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AO" & Ren + NumP Else lRango = "AM" & Ren + NumP
                    'lRango = "AO" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 11
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "M" & Ren + NumP Else lRango = "K" & Ren + NumP
                    'lRango = "M" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Q" & Ren + NumP Else lRango = "O" & Ren + NumP
                    'lRango = "Q" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "U" & Ren + NumP Else lRango = "S" & Ren + NumP
                    'lRango = "U" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Y" & Ren + NumP Else lRango = "W" & Ren + NumP
                    'lRango = "Y" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AC" & Ren + NumP Else lRango = "AA" & Ren + NumP
                    'lRango = "AC" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AG" & Ren + NumP Else lRango = "AE" & Ren + NumP
                    'lRango = "AG" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AK" & Ren + NumP Else lRango = "AI" & Ren + NumP
                    'lRango = "AK" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AO" & Ren + NumP Else lRango = "AM" & Ren + NumP
                    'lRango = "AO" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AS" & Ren + NumP Else lRango = "AQ" & Ren + NumP
                    'lRango = "AS" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                Case 12
                    If Actual Then lRango = "E" & Ren + NumP Else lRango = "C" & Ren + NumP
                    'lRango = "E" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "I" & Ren + NumP Else lRango = "G" & Ren + NumP
                    'lRango = "I" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "M" & Ren + NumP Else lRango = "K" & Ren + NumP
                    'lRango = "M" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Q" & Ren + NumP Else lRango = "O" & Ren + NumP
                    'lRango = "Q" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "U" & Ren + NumP Else lRango = "S" & Ren + NumP
                    'lRango = "U" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "Y" & Ren + NumP Else lRango = "W" & Ren + NumP
                    'lRango = "Y" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AC" & Ren + NumP Else lRango = "AA" & Ren + NumP
                    'lRango = "AC" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AG" & Ren + NumP Else lRango = "AE" & Ren + NumP
                    'lRango = "AG" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AK" & Ren + NumP Else lRango = "AI" & Ren + NumP
                    'lRango = "AK" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AO" & Ren + NumP Else lRango = "AM" & Ren + NumP
                    'lRango = "AO" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AS" & Ren + NumP Else lRango = "AQ" & Ren + NumP
                    'lRango = "AS" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
                    If Actual Then lRango = "AW" & Ren + NumP Else lRango = "AV" & Ren + NumP
                    'lRango = "AW" & Ren + NumP
                    lAcumulado = lAcumulado + .Range(lRango)._Default
            End Select
        End With
        CalculaAcumulado = lAcumulado
    End Function

    Function LlenaDatosAgrupador() As Boolean
        On Error GoTo Err_Renamed
        LlenaDatosAgrupador = False
        If Trim(flexGastos.get_TextMatrix(flexGastos.Row, 3)) = "" And Len(Trim(flexGastos.get_TextMatrix(flexGastos.Row, 3))) < 6 Then
            gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE CodOrigenAplicR = " & Numerico(txtFlex.Text)
        ElseIf Trim(flexGastos.get_TextMatrix(flexGastos.Row, 3)) <> "" And Len(Trim(flexGastos.get_TextMatrix(flexGastos.Row, 3))) = 6 Then
            'gStrSql = "SELECT * " & _
            ''"FROM CatOrigenAplicRecursos A, CatRubrosOrigenAplicRecursos R WHERE R.CodRubro = " & Numerico(flexgastos.TextMatrix(flexgastos.Row, 2)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR AND A.Aplicacion = '" & gstrMovimiento & "'"
            gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE CodOrigenAplicR = " & Numerico(txtFlex.Text)
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFlex.Text = ("0000" & CStr(RsGral.Fields("CodOrigenAplicR").Value))
            flexGastos.set_TextMatrix(flexGastos.Row, 2, Trim(RsGral.Fields("DescOrigenAplicR").Value))
            If Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) <> Trim(txtFlex.Text) Then
                flexGastos.set_TextMatrix(flexGastos.Row, 3, "")
                flexGastos.set_TextMatrix(flexGastos.Row, 4, "")
            End If
            LlenaDatosAgrupador = True
            txtFlex_Leave(txtFlex, New System.EventArgs())
        Else
            MsgBox("Codigo Inexistente Favor de Verificar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            flexGastos.Col = 1
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ObtenerSucursal(ByRef Sucursal As Integer) As String
        On Error GoTo Err_Renamed
        Dim RsAux As ADODB.Recordset
        gStrSql = "SELECT DescAlmacen FROM CatAlmacen WHERE CodAlmacen = " & Sucursal
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsAux = Cmd.Execute
        If RsAux.RecordCount > 0 Then
            ObtenerSucursal = Trim(RsAux.Fields("DescAlmacen").Value)
        Else
            ObtenerSucursal = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Buscar()
        'On Error GoTo Merr
        Try
            Dim strSQL As String
            Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
            Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
            Dim strControlActual As String 'Nombre del control actual
            Dim strDesc As String
            Dim I As Object
            Dim J As Integer
            If flexGastos.Row > 1 Then
                With flexGastos
                    For I = 1 To .Row - 1
                        If Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "" Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Then Exit Sub
                        ElseIf (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "") And (Trim(.get_TextMatrix(1, 3)) = "" And Trim(.get_TextMatrix(1, 4)) = "") Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Then Exit Sub
                        ElseIf (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "") And (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "") Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Then Exit Sub
                        Else
                            Exit Sub
                        End If
                    Next
                    If (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "") And (Trim(.get_TextMatrix(1, 3)) = "" And Trim(.get_TextMatrix(1, 4)) = "") Then
                        If .Col = 3 Or .Col = 4 Then Exit Sub
                    ElseIf (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "") And (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "") Then
                        If .Col = 1 Or .Col = 2 Then Exit Sub
                    End If
                End With
            End If
            If flexGastos.Col > 0 And Trim(flexGastos.get_TextMatrix(flexGastos.Row, 0)) = "" Then Exit Sub
            Me.Tag = "EDORESULTADOS"
            With flexGastos
                If .Col = 0 Then
                    strControlActual = "DESCRIPCION SUCURSAL"
                    strTag = UCase(Me.Tag) & "." & strControlActual
                ElseIf .Col = 1 Then
                    strControlActual = "CODIGO AGRUPADOR"
                    strTag = UCase(Me.Tag) & "." & strControlActual
                ElseIf .Col = 2 Then
                    strControlActual = "DESCRIPCION AGRUPADOR"
                    strTag = UCase(Me.Tag) & "." & strControlActual
                ElseIf .Col = 3 Then
                    strControlActual = "CODIGO RUBRO"
                    strTag = UCase(Me.Tag) & "." & strControlActual
                ElseIf .Col = 4 Then
                    strControlActual = "DESCRIPCION RUBRO"
                    strTag = UCase(Me.Tag) & "." & strControlActual
                End If
                If Me.ActiveControl.Name = "txtFlex" Then
                    strDesc = Trim(txtFlex.Text)
                Else
                    Exit Sub
                End If
                Select Case strControlActual
                    Case "DESCRIPCION SUCURSAL"
                        strCaptionForm = "Consulta de Sucursales"
                        If Me.ActiveControl.Name = "txtFlex" Then
                            If Trim(txtFlex.Text) = "" Then
                                gStrSql = "SELECT Descalmacen AS DESCRIPCION,RIGHT('000'+LTRIM(Codalmacen),3) AS CODIGO " & "From Catalmacen WHERE DescAlmacen LIKE '" & Trim(txtFlex.Text) & "%' and tipoalmacen = 'P' ORDER BY DescAlmacen"
                            Else
                                gStrSql = "SELECT Descalmacen AS DESCRIPCION,RIGHT('000'+LTRIM(Codalmacen),3) AS CODIGO " & "From Catalmacen WHERE tipoalmacen = 'P' ORDER BY DescAlmacen"
                            End If
                        ElseIf Me.ActiveControl.Name = "flexGastos" Then
                            gStrSql = "SELECT Descalmacen AS DESCRIPCION,RIGHT('000'+LTRIM(Codalmacen),3) AS CODIGO " & "From Catalmacen WHERE tipoalmacen = 'P' ORDER BY DescAlmacen"
                        End If
                    Case "CODIGO AGRUPADOR"
                        strCaptionForm = "Consulta de Agrupadores de Origen y Aplicación"
                        gStrSql = "SELECT RIGHT('0000' + LTRIM(CodOrigenAplicR),4) AS AGRUPADOR, DescOrigenAplicR AS DESCRIPCION " & "FROM CatOrigenAplicRecursos ORDER BY CodOrigenAplicR"
                    Case "CODIGO RUBRO"
                        strCaptionForm = "Consulta de Rubros de Origen y Aplicación"
                        If Trim(.get_TextMatrix(.Row, 1)) = "" And Len(Trim(.get_TextMatrix(.Row, 1))) < 4 Then
                            gStrSql = "SELECT RIGHT('000000' + LTRIM(CodRubro),6) AS RUBRO, DescRubro AS DESCRIPCION " & "FROM CatRubrosOrigenAplicRecursos ORDER BY CodRubro"
                        ElseIf Trim(.get_TextMatrix(.Row, 1)) <> "" And Len(Trim(.get_TextMatrix(.Row, 1))) = 4 Then
                            gStrSql = "SELECT RIGHT('000000' + LTRIM(R.CodRubro),6) AS RUBRO, R.DescRubro AS DESCRIPCION " & "FROM CatRubrosOrigenAplicRecursos R,CatOrigenAplicRecursos A WHERE A.CodOrigenAplicR = " & Numerico(.get_TextMatrix(.Row, 1)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR ORDER BY R.CodRubro"
                        End If
                    Case "DESCRIPCION AGRUPADOR"
                        strCaptionForm = "Consulta de Agrupadores de Origen y Aplicación"
                        If Trim(.get_TextMatrix(.Row, 3)) = "" And Len(Trim(.get_TextMatrix(.Row, 3))) < 6 Then
                            gStrSql = "SELECT DescOrigenAplicR AS DESCRIPCION, RIGHT('0000' + LTRIM(CodOrigenAplicR),4) AS AGRUPADOR " & "FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR LIKE '" & strDesc & "%' ORDER BY DescOrigenAplicR"
                        ElseIf Trim(.get_TextMatrix(.Row, 3)) <> "" And Len(Trim(.get_TextMatrix(.Row, 3))) = 6 Then
                            gStrSql = "SELECT A.DescOrigenAplicR AS DESCRIPCION, RIGHT('0000' + LTRIM(R.CodOrigAplicR),4) AS AGRUPADOR " & "FROM CatOrigenAplicRecursos A ,CatRubrosOrigenAplicRecursos R WHERE A.DescOrigenAplicR LIKE '" & strDesc & "%' AND R.CodRubro = " & Numerico(.get_TextMatrix(.Row, 3)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR GROUP BY R.CodOrigAplicR,A.DescOrigenAplicR ORDER BY A.DescOrigenAplicR"
                        End If
                    Case "DESCRIPCION RUBRO"
                        strCaptionForm = "Consulta de Rubros de Origen y Aplicación"
                        If Trim(.get_TextMatrix(.Row, 1)) = "" And Len(Trim(.get_TextMatrix(.Row, 1))) < 4 Then
                            gStrSql = "SELECT DescRubro AS DESCRIPCION, RIGHT('000000' + LTRIM(CodRubro),6) AS RUBRO " & "FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos WHERE DescRubro LIKE '" & strDesc & "%' AND CodOrigenAplicR = CodOrigAplicR ORDER BY DescRubro"
                        ElseIf Trim(.get_TextMatrix(.Row, 1)) <> "" And Len(Trim(.get_TextMatrix(.Row, 1))) = 4 Then
                            gStrSql = "SELECT R.DescRubro AS DESCRIPCION, RIGHT('000000' + LTRIM(R.CodRubro),6) AS RUBRO " & "FROM CatRubrosOrigenAplicRecursos R,CatOrigenAplicRecursos A WHERE R.DescRubro LIKE '" & strDesc & "%' AND A.CodOrigenAplicR = " & Numerico(.get_TextMatrix(.Row, 1)) & " AND A.CodOrigenAplicR = R.CodOrigAplicR ORDER BY DescRubro"
                        End If
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
                '    Dim f As FrmConsultas()
                '    With FrmConsultas.Flexdet
                '        Select Case strControlActual
                '            Case "DESCRIPCION SUCURSAL"
                '                Call ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                '                .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
                '                .set_ColWidth(1, 0, 900) 'Columna del Código
                '            Case "CODIGO AGRUPADOR"
                '                Call ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                '                .set_ColWidth(0, 0, 1300) 'Columna del Código Agrupador
                '                .set_ColWidth(1, 0, 4500) 'Columna de la Descripción del Agrupador
                '            Case "CODIGO RUBRO"
                '                Call ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                '                .set_ColWidth(0, 0, 1300) 'Columna del Codigo del Rubro
                '                .set_ColWidth(1, 0, 4500) 'Columna de la Descripción del Rubro
                '            Case "DESCRIPCION AGRUPADOR"
                '                Call ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                '                .set_ColWidth(0, 0, 4500) 'Columna de la Descripción del Agrupador
                '                .set_ColWidth(1, 0, 1300) 'Columna del Codigo del Agrupador
                '            Case "DESCRIPCION RUBRO"
                '                Call ConfiguraConsultas(FrmConsultas, 6000, RsGral, strTag, strCaptionForm)
                '                .set_ColWidth(0, 0, 4500) 'Columna de la Descripción del Rubro
                '                .set_ColWidth(1, 0, 1300) 'Columna del Codigo del Rubro
                '        End Select
                '    End With
            End With
            'chkFueraEnter.CheckState = System.Windows.Forms.CheckState.Checked
            'FrmConsultas.ShowDialog()
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

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
                txtCodSucursal.Text = ""
                txtCodSucursal.Focus()
                Exit Sub
            Else
                txtCodSucursal.Text = Format(txtCodSucursal.Text, "000")
                dbcSucursal.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            End If
        Else
            MsgBox("Codigo de Almacen no Existe, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodSucursal.Text = ""
            txtCodSucursal.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub EncabezadoFlex()
        With flexGastos
            .Clear()
            .Rows = 11
            .set_Cols(0, 6)
            .set_ColWidth(0, 0, 1200)
            .set_ColWidth(1, 0, 1000)
            .set_ColWidth(2, 0, 2020)
            .set_ColWidth(3, 0, 1000)
            .set_ColWidth(4, 0, 2020)
            .set_ColWidth(5, 0, 0)
            .Row = 0
            .Col = 0
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Sucursal"
            .Col = 1
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Agrup."
            .Col = 2
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Descripción"
            .Col = 3
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Rubro"
            .Col = 4
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Descripción"
            .Row = 1
            .Col = 0
        End With
    End Sub

    Sub Limpiar()
        Nuevo()
        chkTodaslasSucursales.Focus()
    End Sub

    Function LlenaDatosRubro() As Boolean
        On Error GoTo Err_Renamed
        LlenaDatosRubro = False
        If Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) = "" Then
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos " & "WHERE CodRubro = " & Numerico(txtFlex.Text) & " GROUP BY CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion"
        ElseIf Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) <> "" Then
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion FROM CatRubrosOrigenAplicRecursos,CatOrigenAplicRecursos " & "WHERE CodRubro = " & Numerico(txtFlex.Text) & " and codorigaplicr = " & Numerico(Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1))) & " GROUP BY CodOrigenAplicR,DescOrigenAplicR,CodOrigAplicR,CodRubro,DescRubro,Aplicacion"
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFlex.Text = ("000000" & CStr(RsGral.Fields("CodRubro").Value))
            flexGastos.set_TextMatrix(flexGastos.Row, 4, Trim(RsGral.Fields("DescRubro").Value))
            txtFlex_Leave(txtFlex, New System.EventArgs())
            LlenaDatosRubro = True
        Else
            MsgBox("Codigo Inexistente Favor de Investigar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFlex.Text = ""
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Nuevo()
        LlenaAños()
        cmbMes.SelectedIndex = Month(Today) - 1
        cmbAño.SelectedIndex = 0
        cmbMes.Enabled = True
        chkIncluirImpuesto.CheckState = System.Windows.Forms.CheckState.Checked
        optPesos.Checked = True
        optMensual.Checked = True
        chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtCodSucursal.Text = ""
        dbcSucursal.Text = ""
        MesAcumulado = 0
        EncabezadoFlex()
        chkFueraEnter.CheckState = System.Windows.Forms.CheckState.Unchecked
    End Sub

    Sub LlenaAños()
        Dim I As Integer
        cmbAño.Items.Clear()
        For I = Year(Today) To 2001 Step -1
            cmbAño.Items.Add(CStr(I))
        Next
    End Sub

    Private Sub chkIncluirImpuesto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIncluirImpuesto.Enter
        Pon_Tool()
    End Sub

    Private Sub chkTodaslasSucursales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.CheckStateChanged
        If chkTodaslasSucursales.CheckState = 1 Then
            txtCodSucursal.Text = ""
            txtCodSucursal.Enabled = False
            dbcSucursal.Text = ""
            dbcSucursal.Enabled = False
        ElseIf chkTodaslasSucursales.CheckState = 0 Then
            txtCodSucursal.Enabled = True
            dbcSucursal.Enabled = True
        End If
    End Sub

    Private Sub chkTodaslasSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbAño_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAño.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbAño_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmbAño.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If cmbMes.Enabled Then
                cmbMes.Focus()
            ElseIf chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
                chkTodaslasSucursales.Focus()
            ElseIf chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                dbcSucursal.Focus()
            End If
        End If
    End Sub

    Private Sub cmbMes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMes.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbMes_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmbMes.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
                chkTodaslasSucursales.Focus()
            Else
                dbcSucursal.Focus()
            End If
        End If
    End Sub

    Private Sub chkTodaslasSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkTodaslasSucursales.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodaslasSucursales.Focus()
        End If
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        'If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        'If Trim(dbcSucursal.Text) = "" Then
        '    txtCodSucursal.Text = ""
        'End If
        'gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        'DCChange(gStrSql, tecla)
        ''intCodSucursal = 0
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
            Exit Sub
        End If
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursal)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        tecla = eventArgs.KeyCode
        'If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
        txtCodSucursal.Focus()
        'End If
    End Sub

    Private Sub dbcSucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcSucursal.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcSucursal_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        'If dbcSucursal.SelectedItem <> 0 Then
        'dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        'End If
        FueraChange = True
        dbcSucursal.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        FueraChange = True
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
        If intCodSucursal <> 0 Then
            txtCodSucursal.Text = Format(String.Concat(intCodSucursal, "000"))
        End If
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        'If dbcSucursal.SelectedItem <> 0 Then
        'dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        'End If
        FueraChange = True
        dbcSucursal.Text = Aux
        FueraChange = False
    End Sub

    Private Sub flexGastos_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexGastos.ClickEvent
        txtFlex.Visible = False
    End Sub

    Private Sub flexGastos_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexGastos.DblClick
        flexGastos_KeyPressEvent(flexGastos, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub flexGastos_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexGastos.EnterCell
        With flexGastos
            If .Col = 0 Then
                lblDesc.Text = Trim(.get_TextMatrix(.Row, 0))
            ElseIf .Col = 1 Or .Col = 2 Then
                lblDesc.Text = Trim(.get_TextMatrix(.Row, 2))
            ElseIf .Col = 3 Or .Col = 4 Then
                lblDesc.Text = Trim(.get_TextMatrix(.Row, 4))
            End If
        End With
    End Sub

    Private Sub flexGastos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexGastos.Enter
        txtFlex.Visible = False
        Pon_Tool()
        flexGastos_EnterCell(flexGastos, New System.EventArgs())
    End Sub

    Function ValidaCodigos(ByRef RenActual As Integer) As Boolean
        Dim I As Integer
        ValidaCodigos = False
        With flexGastos
            For I = 1 To .Rows - 1
                If I <> RenActual Then
                    If RenActual = 1 Then
                        If Trim(flexGastos.get_TextMatrix(.Row, 1)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 2)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 3)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 4)) = "" Then Exit Function
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" And Trim(.get_TextMatrix(I, 3)) <> "" And Trim(.get_TextMatrix(I, 4)) <> "" Then
                            If Numerico(.get_TextMatrix(I, 1)) = Numerico(.get_TextMatrix(1, 1)) And Numerico(.get_TextMatrix(I, 3)) = Numerico(.get_TextMatrix(1, 3)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        ElseIf Trim(.get_TextMatrix(I, 0)) <> "" And (Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "") And (Trim(.get_TextMatrix(I, 3)) = "" And Trim(.get_TextMatrix(I, 4)) = "") Then
                            If Numerico(.get_TextMatrix(I, 1)) = Numerico(.get_TextMatrix(1, 1)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        ElseIf Trim(.get_TextMatrix(I, 0)) <> "" And (Trim(.get_TextMatrix(I, 1)) = "" And Trim(.get_TextMatrix(I, 2)) = "") And (Trim(.get_TextMatrix(I, 3)) <> "" And Trim(.get_TextMatrix(I, 4)) <> "") Then
                            If Numerico(.get_TextMatrix(I, 3)) = Numerico(.get_TextMatrix(1, 3)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        End If
                    Else
                        If Trim(flexGastos.get_TextMatrix(.Row, 1)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 2)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 3)) = "" And Trim(flexGastos.get_TextMatrix(.Row, 4)) = "" Then Exit Function
                        If Trim(flexGastos.get_TextMatrix(1, 0)) <> "" And Trim(flexGastos.get_TextMatrix(1, 1)) <> "" And Trim(flexGastos.get_TextMatrix(1, 2)) <> "" And Trim(flexGastos.get_TextMatrix(1, 3)) <> "" And Trim(flexGastos.get_TextMatrix(1, 4)) <> "" Then
                            If Numerico(.get_TextMatrix(I, 3)) = Numerico(.get_TextMatrix(.Row, 3)) And Numerico(.get_TextMatrix(I, 1)) = Numerico(.get_TextMatrix(.Row, 1)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        ElseIf (Trim(flexGastos.get_TextMatrix(1, 0)) <> "" And Trim(flexGastos.get_TextMatrix(1, 1)) <> "" And Trim(flexGastos.get_TextMatrix(1, 2)) <> "") And (Trim(flexGastos.get_TextMatrix(1, 3)) = "" And Trim(flexGastos.get_TextMatrix(1, 4)) = "") Then
                            If Numerico(.get_TextMatrix(I, 1)) = Numerico(.get_TextMatrix(.Row, 1)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        ElseIf (Trim(flexGastos.get_TextMatrix(1, 1)) = "" And Trim(flexGastos.get_TextMatrix(1, 2)) = "") And (Trim(flexGastos.get_TextMatrix(1, 0)) <> "" And Trim(flexGastos.get_TextMatrix(1, 3)) <> "" And Trim(flexGastos.get_TextMatrix(1, 4)) <> "") Then
                            If Numerico(.get_TextMatrix(I, 3)) = Numerico(.get_TextMatrix(.Row, 3)) Then
                                ValidaCodigos = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        End With
    End Function

    Sub EliminarLinea()
        Dim Ren As Integer
        With flexGastos
            If Trim(.get_TextMatrix(.Row, 0)) = "" And Trim(.get_TextMatrix(.Row, 1)) = "" And Trim(.get_TextMatrix(.Row, 2)) = "" And Trim(.get_TextMatrix(.Row, 3)) = "" And Trim(.get_TextMatrix(.Row, 4)) = "" Then Exit Sub
        End With
        Select Case MsgBox("¿Desea Eliminar Esta Informacion?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
            Case MsgBoxResult.Yes
                Ren = flexGastos.Rows
                flexGastos.RemoveItem(flexGastos.Row)
                flexGastos.Rows = Ren
                flexGastos.Focus()
            Case MsgBoxResult.No, MsgBoxResult.Cancel
                flexGastos.Focus()
                Exit Sub
        End Select
    End Sub

    Private Sub flexGastos_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexGastos.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete Then
            EliminarLinea()
        End If
    End Sub

    Private Sub flexGastos_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexGastos.KeyPressEvent
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        If chkFueraEnter.CheckState = System.Windows.Forms.CheckState.Checked Then
            chkFueraEnter.CheckState = System.Windows.Forms.CheckState.Unchecked
            Exit Sub
        End If
        Dim lonR, lonI As Integer
        Dim I As Integer
        If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then
            'Verifica si se puede capturar la fila
            If flexGastos.Row > 1 Then
                With flexGastos
                    For I = 1 To .Row - 1
                        If Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "" Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Then Exit Sub
                        ElseIf (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "") And (Trim(.get_TextMatrix(1, 3)) = "" And Trim(.get_TextMatrix(1, 4)) = "") Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 1)) = "" Then Exit Sub
                        ElseIf (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "") And (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "") Then
                            If Trim(.get_TextMatrix(I, 0)) = "" Or Trim(.get_TextMatrix(I, 3)) = "" Then Exit Sub
                        Else
                            Exit Sub
                        End If
                    Next
                    If (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 1)) <> "" And Trim(.get_TextMatrix(1, 2)) <> "") And (Trim(.get_TextMatrix(1, 3)) = "" And Trim(.get_TextMatrix(1, 4)) = "") Then
                        If .Col = 3 Or .Col = 4 Then Exit Sub
                    ElseIf (Trim(.get_TextMatrix(1, 1)) = "" And Trim(.get_TextMatrix(1, 2)) = "") And (Trim(.get_TextMatrix(1, 0)) <> "" And Trim(.get_TextMatrix(1, 3)) <> "" And Trim(.get_TextMatrix(1, 4)) <> "") Then
                        If .Col = 1 Or .Col = 2 Then Exit Sub
                    End If
                End With
            End If
            'Edita el campo sólo si es Editable
            If flexGastos.Col > 0 And Trim(flexGastos.get_TextMatrix(flexGastos.Row, 0)) = "" Then Exit Sub
            If (flexGastos.Col = 3 Or flexGastos.Col = 4) And (Trim(flexGastos.get_TextMatrix(flexGastos.Row, 0)) = "" Or Trim(flexGastos.get_TextMatrix(flexGastos.Row, 1)) = "" Or Trim(flexGastos.get_TextMatrix(flexGastos.Row, 2)) = "") Then Exit Sub
            If flexGastos.Row >= 1 And flexGastos.Col < 5 Then
                If flexGastos.Col = 0 And flexGastos.Col = 1 Or flexGastos.Col = 3 Then
                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                End If
                CambiarFormatoTxtenCaptura()
                MSHFlexGridEdit(flexGastos, txtFlex, eventArgs.keyAscii)
                If Len(Trim(txtFlex.Text)) = 1 Then
                    System.Windows.Forms.SendKeys.Send("{right}")
                End If
            End If
        ElseIf eventArgs.keyAscii = System.Windows.Forms.Keys.Escape Then
            Exit Sub
        Else
            Exit Sub
        End If
    End Sub

    Private Sub flexGastos_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexGastos.Leave
        lblDesc.Text = ""
    End Sub

    Private Sub frmVtasEstadodeResultados_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasEstadodeResultados_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasEstadodeResultados_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "chkTodaslasSucursales" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmVtasEstadodeResultados_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasEstadodeResultados_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Nuevo()
    End Sub

    Private Sub frmVtasEstadodeResultados_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        'Si se cierra el formulario y existio algun cambio en el registro se
        'informa al usuario del cabio y si desea guardar el registro, ya sea
        'que sea nuevo o un registro modificado
        If mblnSalir Then
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSalir = False
                    Cancel = 1
                    chkTodaslasSucursales.Focus()
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasEstadodeResultados_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        'Me = Nothing
    End Sub

    Private Sub optAnual_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAnual.CheckedChanged
        If eventSender.Checked Then
            cmbMes.Enabled = False
        End If
    End Sub

    Private Sub optAnual_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAnual.Enter
        Pon_Tool()
    End Sub

    Private Sub optDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub optMensual_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMensual.CheckedChanged
        If eventSender.Checked Then
            cmbMes.Enabled = True
        End If
    End Sub

    Private Sub optMensual_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMensual.Enter
        Pon_Tool()
    End Sub

    Private Sub optPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPesos.Enter
        Pon_Tool()
    End Sub

    Private Sub txtCodSucursal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.TextChanged
        dbcSucursal.Text = ""
    End Sub

    Private Sub txtCodSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Enter
        Pon_Tool()
    End Sub

    Private Sub txtCodSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodSucursal.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodaslasSucursales.Focus()
        End If
    End Sub

    Private Sub txtCodsucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodSucursal.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Leave
        If Trim(txtCodSucursal.Text) <> "" Then
            BuscaSucursal()
        End If
    End Sub

    Private Sub txtFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Enter
        SelTextoTxt(txtFlex)
        Pon_Tool()
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        With flexGastos
            If KeyCode = System.Windows.Forms.Keys.Return Then
                Select Case .Col
                    Case 0, 1, 2, 3, 4
                        If .Col = 0 And Trim(txtFlex.Text) <> "" Then
                            If LlenaSucursal() Then
                                .Text = Trim(txtFlex.Text)
                                If Not ValidaCodigos(.Row) Then
                                    .Col = 1
                                    txtFlex.Visible = False
                                    Exit Sub
                                Else
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
                                    Exit Sub
                                End If
                            Else
                                Exit Sub
                            End If
                        ElseIf .Col = 1 And Trim(txtFlex.Text) <> "" Then
                            If LlenaDatosAgrupador() Then
                                .Text = Trim(txtFlex.Text)
                                If Not ValidaCodigos(.Row) Then
                                    .Col = 3
                                Else
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 1, "")
                                    .set_TextMatrix(.Row, 2, "")
                                    System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
                                    Exit Sub
                                End If
                            Else
                                Exit Sub
                            End If
                        ElseIf .Col = 2 And Trim(txtFlex.Text) <> "" Then
                            If DescripcionAgrupador() Then
                                .Text = Trim(txtFlex.Text)
                                If Not ValidaCodigos(.Row) Then
                                    .Col = 3
                                Else
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 1, "")
                                    .set_TextMatrix(.Row, 2, "")
                                    System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
                                    Exit Sub
                                End If
                            Else
                                Exit Sub
                            End If
                        ElseIf .Col = 3 And Trim(txtFlex.Text) <> "" Then
                            If LlenaDatosRubro() Then
                                .Text = Trim(txtFlex.Text)
                                If Not ValidaCodigos(.Row) Then
                                    .Col = 0
                                Else
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 3, "")
                                    .set_TextMatrix(.Row, 4, "")
                                    System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
                                    Exit Sub
                                End If
                                txtFlex.Visible = False
                            Else
                                Exit Sub
                            End If
                        ElseIf .Col = 4 And Trim(txtFlex.Text) <> "" Then
                            If DescripcionRubro() Then
                                .Text = Trim(txtFlex.Text)
                                If Not ValidaCodigos(.Row) Then
                                    .Col = 0
                                Else
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 3, "")
                                    .set_TextMatrix(.Row, 4, "")
                                    System.Windows.Forms.Form.ActiveForm.ActiveControl.Focus()
                                    Exit Sub
                                End If
                                txtFlex.Visible = False
                            Else
                                Exit Sub
                            End If
                        ElseIf Trim(txtFlex.Text) = "" Then
                            If .Col = 0 Then
                                txtFlex.Visible = False

                                flexGastos.Focus()
                                Exit Sub
                            ElseIf .Col = 1 Or .Col = 2 Then
                                .set_TextMatrix(.Row, 1, "")
                                .set_TextMatrix(.Row, 2, "")
                                .set_TextMatrix(.Row, 3, "")
                                .set_TextMatrix(.Row, 4, "")
                                txtFlex.Visible = False
                                If ValidaCodigos(.Row) Then
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 3, "")
                                    .set_TextMatrix(.Row, 4, "")
                                    .Col = 3
                                End If
                                flexGastos.Focus()
                                Exit Sub
                            ElseIf .Col = 3 Or .Col = 4 Then
                                .set_TextMatrix(.Row, 3, "")
                                .set_TextMatrix(.Row, 4, "")
                                txtFlex.Visible = False
                                If ValidaCodigos(.Row) Then
                                    MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                    .set_TextMatrix(.Row, 1, "")
                                    .set_TextMatrix(.Row, 2, "")
                                    .Col = 1
                                End If
                                flexGastos.Focus()
                                Exit Sub
                            End If
                        End If
                End Select
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    If .Col = 1 Or .Col = 3 Then Exit Sub
                    .Row = .Row + 1
                    .TopRow = .Row
                Else
                    If .Col = 1 Or .Col = 3 Then Exit Sub
                    .Row = .Row + 1
                    If .Row > 6 Then
                        .TopRow = .Row
                    End If
                End If
                txtFlex.Visible = False
            ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
                txtFlex.Visible = False
                .Focus()
            ElseIf KeyCode = System.Windows.Forms.Keys.F3 Then
                Buscar()
            End If
        End With
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
            Case Else
                Select Case flexGastos.Col
                    Case 0
                        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
                    Case 1
                        ModEstandar.gp_CampoNumerico(KeyAscii)
                    Case 2
                        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
                    Case 3
                        ModEstandar.gp_CampoNumerico(KeyAscii)
                    Case 4
                        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
                End Select
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub


    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
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

    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class