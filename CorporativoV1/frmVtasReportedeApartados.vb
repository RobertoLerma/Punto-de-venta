Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasReportedeApartados
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkEnCancelados As System.Windows.Forms.CheckBox
    Public WithEvents chkEnVigentes As System.Windows.Forms.CheckBox
    Public WithEvents chkEnSaldados As System.Windows.Forms.CheckBox
    Public WithEvents fraAptEn As System.Windows.Forms.GroupBox
    Public WithEvents chkSoloVigentes As System.Windows.Forms.CheckBox
    Public WithEvents chkApartadosSaldados As System.Windows.Forms.CheckBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents txtCodCliente As System.Windows.Forms.TextBox
    Public WithEvents optEstadodeCuenta As System.Windows.Forms.RadioButton
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents chkTodaslasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents optReporteGeneral As System.Windows.Forms.RadioButton
    Public WithEvents dbcCliente As System.Windows.Forms.ComboBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents Line1 As System.Windows.Forms.Label
    Public WithEvents Line3 As System.Windows.Forms.Label
    Public WithEvents Line4 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Line2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label


    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim intCodCliente As Integer
    Dim rsReporte As ADODB.Recordset
    Dim rsSubReporte As ADODB.Recordset
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Public WithEvents dtpFechaFinal As DateTimePicker
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkSoloVigentes = New System.Windows.Forms.CheckBox()
        Me.chkApartadosSaldados = New System.Windows.Forms.CheckBox()
        Me.txtCodCliente = New System.Windows.Forms.TextBox()
        Me.optEstadodeCuenta = New System.Windows.Forms.RadioButton()
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.chkTodaslasSucursales = New System.Windows.Forms.CheckBox()
        Me.optReporteGeneral = New System.Windows.Forms.RadioButton()
        Me.fraAptEn = New System.Windows.Forms.GroupBox()
        Me.chkEnCancelados = New System.Windows.Forms.CheckBox()
        Me.chkEnVigentes = New System.Windows.Forms.CheckBox()
        Me.chkEnSaldados = New System.Windows.Forms.CheckBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dbcCliente = New System.Windows.Forms.ComboBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.Line1 = New System.Windows.Forms.Label()
        Me.Line3 = New System.Windows.Forms.Label()
        Me.Line4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Line2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.fraAptEn.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkSoloVigentes
        '
        Me.chkSoloVigentes.BackColor = System.Drawing.SystemColors.Control
        Me.chkSoloVigentes.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSoloVigentes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSoloVigentes.Location = New System.Drawing.Point(185, 97)
        Me.chkSoloVigentes.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSoloVigentes.Name = "chkSoloVigentes"
        Me.chkSoloVigentes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSoloVigentes.Size = New System.Drawing.Size(104, 17)
        Me.chkSoloVigentes.TabIndex = 6
        Me.chkSoloVigentes.Text = "Sólo apartados vigentes"
        Me.ToolTip1.SetToolTip(Me.chkSoloVigentes, "Muestra Apartados Saldados")
        Me.chkSoloVigentes.UseVisualStyleBackColor = False
        '
        'chkApartadosSaldados
        '
        Me.chkApartadosSaldados.BackColor = System.Drawing.SystemColors.Control
        Me.chkApartadosSaldados.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkApartadosSaldados.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkApartadosSaldados.Location = New System.Drawing.Point(42, 97)
        Me.chkApartadosSaldados.Margin = New System.Windows.Forms.Padding(2)
        Me.chkApartadosSaldados.Name = "chkApartadosSaldados"
        Me.chkApartadosSaldados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkApartadosSaldados.Size = New System.Drawing.Size(116, 17)
        Me.chkApartadosSaldados.TabIndex = 5
        Me.chkApartadosSaldados.Text = "Incluir apartados saldados"
        Me.ToolTip1.SetToolTip(Me.chkApartadosSaldados, "Muestra Apartados Saldados")
        Me.chkApartadosSaldados.UseVisualStyleBackColor = False
        '
        'txtCodCliente
        '
        Me.txtCodCliente.AcceptsReturn = True
        Me.txtCodCliente.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodCliente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodCliente.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodCliente.Location = New System.Drawing.Point(95, 163)
        Me.txtCodCliente.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodCliente.MaxLength = 5
        Me.txtCodCliente.Name = "txtCodCliente"
        Me.txtCodCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodCliente.Size = New System.Drawing.Size(56, 20)
        Me.txtCodCliente.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.txtCodCliente, "Codigo del Cliente")
        '
        'optEstadodeCuenta
        '
        Me.optEstadodeCuenta.BackColor = System.Drawing.SystemColors.Control
        Me.optEstadodeCuenta.Cursor = System.Windows.Forms.Cursors.Default
        Me.optEstadodeCuenta.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.optEstadodeCuenta.Location = New System.Drawing.Point(12, 139)
        Me.optEstadodeCuenta.Margin = New System.Windows.Forms.Padding(2)
        Me.optEstadodeCuenta.Name = "optEstadodeCuenta"
        Me.optEstadodeCuenta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optEstadodeCuenta.Size = New System.Drawing.Size(126, 17)
        Me.optEstadodeCuenta.TabIndex = 7
        Me.optEstadodeCuenta.TabStop = True
        Me.optEstadodeCuenta.Text = "Estad&o de Cuenta"
        Me.ToolTip1.SetToolTip(Me.optEstadodeCuenta, "Muestra el Reporte de Estado de Cuenta por Cliente")
        Me.optEstadodeCuenta.UseVisualStyleBackColor = False
        '
        'txtCodSucursal
        '
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.Location = New System.Drawing.Point(102, 58)
        Me.txtCodSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucursal.MaxLength = 3
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(50, 20)
        Me.txtCodSucursal.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtCodSucursal, "Codigo de la Sucursal")
        '
        'chkTodaslasSucursales
        '
        Me.chkTodaslasSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodaslasSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodaslasSucursales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodaslasSucursales.Location = New System.Drawing.Point(28, 34)
        Me.chkTodaslasSucursales.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodaslasSucursales.Name = "chkTodaslasSucursales"
        Me.chkTodaslasSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodaslasSucursales.Size = New System.Drawing.Size(175, 17)
        Me.chkTodaslasSucursales.TabIndex = 1
        Me.chkTodaslasSucursales.Text = "Todas las Sucursales"
        Me.ToolTip1.SetToolTip(Me.chkTodaslasSucursales, "Muestra Todas las Sucursales")
        Me.chkTodaslasSucursales.UseVisualStyleBackColor = False
        '
        'optReporteGeneral
        '
        Me.optReporteGeneral.BackColor = System.Drawing.SystemColors.Control
        Me.optReporteGeneral.Cursor = System.Windows.Forms.Cursors.Default
        Me.optReporteGeneral.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.optReporteGeneral.Location = New System.Drawing.Point(12, 13)
        Me.optReporteGeneral.Margin = New System.Windows.Forms.Padding(2)
        Me.optReporteGeneral.Name = "optReporteGeneral"
        Me.optReporteGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optReporteGeneral.Size = New System.Drawing.Size(181, 17)
        Me.optReporteGeneral.TabIndex = 0
        Me.optReporteGeneral.TabStop = True
        Me.optReporteGeneral.Text = "&Reporte General"
        Me.ToolTip1.SetToolTip(Me.optReporteGeneral, "Muestra el Reporte General de Apartados")
        Me.optReporteGeneral.UseVisualStyleBackColor = False
        '
        'fraAptEn
        '
        Me.fraAptEn.BackColor = System.Drawing.SystemColors.Control
        Me.fraAptEn.Controls.Add(Me.chkEnCancelados)
        Me.fraAptEn.Controls.Add(Me.chkEnVigentes)
        Me.fraAptEn.Controls.Add(Me.chkEnSaldados)
        Me.fraAptEn.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraAptEn.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.fraAptEn.Location = New System.Drawing.Point(25, 199)
        Me.fraAptEn.Margin = New System.Windows.Forms.Padding(2)
        Me.fraAptEn.Name = "fraAptEn"
        Me.fraAptEn.Padding = New System.Windows.Forms.Padding(2)
        Me.fraAptEn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraAptEn.Size = New System.Drawing.Size(304, 50)
        Me.fraAptEn.TabIndex = 11
        Me.fraAptEn.TabStop = False
        Me.fraAptEn.Text = " Incluir  apartados...  "
        '
        'chkEnCancelados
        '
        Me.chkEnCancelados.BackColor = System.Drawing.SystemColors.Control
        Me.chkEnCancelados.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEnCancelados.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEnCancelados.Location = New System.Drawing.Point(201, 19)
        Me.chkEnCancelados.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEnCancelados.Name = "chkEnCancelados"
        Me.chkEnCancelados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEnCancelados.Size = New System.Drawing.Size(95, 26)
        Me.chkEnCancelados.TabIndex = 14
        Me.chkEnCancelados.Text = "Cancelados"
        Me.chkEnCancelados.UseVisualStyleBackColor = False
        '
        'chkEnVigentes
        '
        Me.chkEnVigentes.BackColor = System.Drawing.SystemColors.Control
        Me.chkEnVigentes.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEnVigentes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEnVigentes.Location = New System.Drawing.Point(27, 19)
        Me.chkEnVigentes.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEnVigentes.Name = "chkEnVigentes"
        Me.chkEnVigentes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEnVigentes.Size = New System.Drawing.Size(81, 26)
        Me.chkEnVigentes.TabIndex = 12
        Me.chkEnVigentes.Text = "Vigentes"
        Me.chkEnVigentes.UseVisualStyleBackColor = False
        '
        'chkEnSaldados
        '
        Me.chkEnSaldados.BackColor = System.Drawing.SystemColors.Control
        Me.chkEnSaldados.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEnSaldados.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEnSaldados.Location = New System.Drawing.Point(114, 19)
        Me.chkEnSaldados.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEnSaldados.Name = "chkEnSaldados"
        Me.chkEnSaldados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEnSaldados.Size = New System.Drawing.Size(82, 26)
        Me.chkEnSaldados.TabIndex = 13
        Me.chkEnSaldados.Text = "Saldados"
        Me.chkEnSaldados.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFechaFinal)
        Me.Frame1.Controls.Add(Me.dtpFechaInicial)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(12, 273)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(387, 46)
        Me.Frame1.TabIndex = 15
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo ..."
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(264, 17)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(98, 20)
        Me.dtpFechaFinal.TabIndex = 19
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(73, 17)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(98, 20)
        Me.dtpFechaInicial.TabIndex = 17
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(211, 20)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(58, 17)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Hasta el :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(17, 20)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(62, 17)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Desde el :"
        '
        'dbcCliente
        '
        Me.dbcCliente.Location = New System.Drawing.Point(165, 163)
        Me.dbcCliente.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcCliente.Name = "dbcCliente"
        Me.dbcCliente.Size = New System.Drawing.Size(173, 21)
        Me.dbcCliente.TabIndex = 10
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(165, 57)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(173, 21)
        Me.dbcSucursal.TabIndex = 4
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(6, 125)
        Me.Line1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(305, 1)
        Me.Line1.TabIndex = 16
        '
        'Line3
        '
        Me.Line3.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line3.Location = New System.Drawing.Point(11, 260)
        Me.Line3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Line3.Name = "Line3"
        Me.Line3.Size = New System.Drawing.Size(298, 1)
        Me.Line3.TabIndex = 17
        '
        'Line4
        '
        Me.Line4.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line4.Location = New System.Drawing.Point(11, 260)
        Me.Line4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Line4.Name = "Line4"
        Me.Line4.Size = New System.Drawing.Size(297, 1)
        Me.Line4.TabIndex = 18
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(41, 165)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(50, 17)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Cliente :"
        '
        'Line2
        '
        Me.Line2.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line2.Location = New System.Drawing.Point(7, 125)
        Me.Line2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Line2.Name = "Line2"
        Me.Line2.Size = New System.Drawing.Size(306, 1)
        Me.Line2.TabIndex = 19
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(41, 61)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(64, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Sucursal :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(131, 332)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 100
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(16, 332)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 99
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(246, 333)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 98
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmVtasReportedeApartados
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(409, 382)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.fraAptEn)
        Me.Controls.Add(Me.chkSoloVigentes)
        Me.Controls.Add(Me.chkApartadosSaldados)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.txtCodCliente)
        Me.Controls.Add(Me.optEstadodeCuenta)
        Me.Controls.Add(Me.txtCodSucursal)
        Me.Controls.Add(Me.chkTodaslasSucursales)
        Me.Controls.Add(Me.optReporteGeneral)
        Me.Controls.Add(Me.dbcCliente)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.Line1)
        Me.Controls.Add(Me.Line3)
        Me.Controls.Add(Me.Line4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Line2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(302, 177)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasReportedeApartados"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Apartados"
        Me.fraAptEn.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub


    Sub Imprime()

        Dim RptVtasEstadodeCuentaApartados As New RptVtasEstadodeCuentaApartados
        Dim rptVtasReportedeApartados As New RptVtasReportedeApartados

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        Dim Sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim PeriodoReporte As String
        Dim strWhere As String
        Dim strWhere2 As String
        Dim strHaving As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim strSucursal As String

        Dim strCancel As String
        Dim strCancel2 As String
        Dim strVigente As String
        Dim strVigente2 As String
        Dim strSaldado As String
        Dim strSaldado2 As String
        Dim lHaving As String
        Dim lHaving2 As String

        Dim RsAux As ADODB.Recordset
        On Error GoTo Merr

        If Not ValidaDatos() Then Exit Sub
        Cmd.CommandTimeout = 300
        intCodCliente = CInt(Numerico((txtCodCliente.Text)))
        intCodSucursal = CInt(Numerico((txtCodSucursal.Text)))
        Sql = ""
        strHaving = ""
        strSucursal = ""

        strCancel = ""
        strCancel2 = ""
        strVigente = ""
        strVigente2 = ""
        strSaldado = ""
        strSaldado2 = ""
        lHaving = ""
        lHaving2 = ""

        dtpFechaInicial.Refresh()
        dtpFechaFinal.Refresh()

        'Do While (sglTiempoCambio) <= 2.1
        'Loop

        If dtpFechaInicial.Value > dtpFechaFinal.Value Then
            MsgBox("La Fecha Inicial no Puede ser Mayor que la Fecha Final.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If dtpFechaInicial.Value > Now Then
            MsgBox("la Fecha Inicial no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If dtpFechaFinal.Value > Now Then
            MsgBox("la Fecha Final no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If

        If optReporteGeneral.Checked = True Then
            If CDbl(Numerico(txtCodSucursal.Text)) = 0 And chkTodaslasSucursales.CheckState = 0 Then
                MsgBox("Proporcione un Codigo de Sucursal, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodSucursal.Focus()
                Exit Sub
            End If
            If Trim(dbcSucursal.Text) = "" And chkTodaslasSucursales.CheckState = 0 Then
                MsgBox("Proprcione la Descripción de la Sucursal, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                dbcSucursal.Focus()
                Exit Sub
            End If
            If chkApartadosSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked And chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                MsgBox("Debe seleccionar el tipo de folios a mostrar en el reporte" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                chkApartadosSaldados.Focus()
                Exit Sub
            End If
        ElseIf optEstadodeCuenta.Checked = True Then
            If CInt(Numerico((txtCodCliente.Text))) = 0 Then
                MsgBox("Proporcione el Código de un Cliente" & vbNewLine & "Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                txtCodCliente.Focus()
                Exit Sub
            End If
            If Trim(dbcCliente.Text) = "" Then
                MsgBox("Debe proporcionar el nombre del cliente" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                dbcCliente.Focus()
                Exit Sub
            End If
        End If

        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        'FechaInicial = CStr(DateSerial(Year(dtpFechaInicial.Value), Month(dtpFechaInicial.Value), (Convert.ToInt32(dtpFechaInicial.Value))))
        'FechaFinal = CStr(DateSerial(Year(dtpFechaFinal.Value), Month(dtpFechaFinal.Value), (Convert.ToInt32(dtpFechaFinal.Value))))
        'PeriodoReporte = "Del   " & Format(dtpFechaInicial.Value, "dd/mmm/yyyy") & "   al   " & Format(dtpFechaFinal.Value, "dd/mmm/yyyy")
        FechaInicial = AgregarHoraAFecha(dtpFechaInicial.Value)
        FechaFinal = AgregarHoraAFecha(dtpFechaFinal.Value)
        PeriodoReporte = "Del " & FechaInicial & " al " & FechaFinal

        strWhere = ""

        If optReporteGeneral.Checked = True Then
            If chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Checked Then
                strHaving = "Having (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1)) <> 0 " & "AND (Round(Sum(Case When VtaCab.Moneda = 'D' Then Ing.Total Else (Ing.Total * Ing.TipoCambio) End),1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN ISNULL(DC.TOTALDOCTO,0) ELSE ISNULL(DC.TOTALDOCTO * VTACAB.TIPOCAMBIO,0) END,1)) < " & "(ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1))"
            ElseIf chkApartadosSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
                strHaving = "Having (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1)) <> 0 " & "AND (Round(Sum(Case When VtaCab.Moneda = 'D' Then Ing.Total Else (Ing.Total * Ing.TipoCambio) End),1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN ISNULL(DC.TOTALDOCTO,0) ELSE ISNULL(DC.TOTALDOCTO * VTACAB.TIPOCAMBIO,0) END,1)) >= " & "(ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1)) And ISNULL(Max(Ing.FechaIngreso),'01/01/1900') Between '" & VB6.Format(dtpFechaInicial.Value, C_FORMATFECHAGUARDAR) & "' and '" & VB6.Format(dtpFechaFinal.Value, C_FORMATFECHAGUARDAR) & "' "
            End If
            If chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                strSucursal = " AND VTACAB.CODSUCURSAL = " & CInt(Numerico((txtCodSucursal.Text))) & " "
            Else
                strSucursal = ""
            End If
            Sql = "SELECT SUC.DESCALMACEN,VTACAB.FOLIOVENTA,VTACAB.FECHAVENTA,CATCLI.DESCCLIENTE,(ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1)) AS APARTADO," & "(Round(Sum(Case When VtaCab.Moneda = 'D' Then ISNULL(Ing.Total,0) Else ISNULL(Ing.Total * Ing.TipoCambio,0) End),1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN ISNULL(DC.TOTALDOCTO,0) ELSE ISNULL(DC.TOTALDOCTO * VTACAB.TIPOCAMBIO,0) END,1)) AS ABONOS," & "((ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1))) - ((Round(Sum(Case When VtaCab.Moneda = 'D' Then Ing.Total Else (Ing.Total * Ing.TipoCambio) End),1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN ISNULL(DC.TOTALDOCTO,0) ELSE ISNULL(DC.TOTALDOCTO * VTACAB.TIPOCAMBIO,0) END,1))) as SALDO," & "VtaCab.Moneda,VtaCab.TipoCambio,VtaCab.Estatus,Round(Case When VtaCab.Moneda = 'D' Then (VtaCab.Total+VtaCab.Redondeo) Else ((VtaCab.Total+VtaCab.Redondeo)*VtaCab.TipoCambio) End,1) as VtaReal,Round(Sum(Case When VtaCab.Moneda = 'D' Then Ing.Total Else (Ing.Total*Ing.TipoCambio) End),1) As IngresoReal,Round(IsNull(Case When DF.Moneda = 'D' Then DF.Importe Else (DF.Importe*DF.TipoCambio) End,0),1) As DifCamb,ISNULL(Max(Ing.FechaIngreso),'01/01/1900') as FechaUltIng  " & "FROM MOVIMIENTOSVENTASCAB VTACAB (Nolock) LEFT OUTER JOIN (SELECT * FROM INGRESOS (Nolock) WHERE ESTATUS <> 'C') ING ON VTACAB.FOLIOVENTA = ING.FOLIOMOVTO LEFT OUTER JOIN (SELECT FolioVenta,SUM(TotalDevol) TotalDevol,SUM(TotalDocto) TotalDocto FROM DevolucionesCab WHERE ESTATUS <> 'C' GROUP BY FolioVenta) DC ON VtaCab.FolioVenta = DC.FolioVenta " & "INNER JOIN CATALMACEN SUC (Nolock) ON VTACAB.CODSUCURSAL = SUC.CODALMACEN INNER JOIN CATCLIENTES CATCLI (Nolock) ON VTACAB.CODCLIENTE = CATCLI.CODCLIENTE LEFT OUTER JOIN DiferenciaCambiaria DF (Nolock) ON VtaCab.FolioVenta = DF.FolioVenta " & "WHERE VTACAB.TIPOMOVTO = 'A' " & strSucursal & " AND VTACAB.ESTATUS <> 'C' " & "GROUP BY SUC.DESCALMACEN,VTACAB.FOLIOVENTA,VTACAB.FECHAVENTA,CATCLI.DESCCLIENTE,VTACAB.TOTAL,VTACAB.REDONDEO,VtaCab.Moneda,VtaCab.TipoCambio,VtaCab.Estatus,DF.Moneda,DF.TipoCambio,DF.Importe,DC.TotalDevol,DC.TotalDocto " & strHaving & "ORDER  BY SUC.DESCALMACEN,CATCLI.DESCCLIENTE,VTACAB.FECHAVENTA,VTACAB.FOLIOVENTA"
            BorraCmd()
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd.CommandText = Sql
            frmReportes.rsReport = Cmd.Execute

            If frmReportes.rsReport.RecordCount = 0 Then
                MsgBox("No existen movimientos en este periodo de fechas" & vbNewLine & "Favor de verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                Exit Sub
            Else
                rptVtasReportedeApartados.SetDataSource(frmReportes.rsReport)
            End If

            NombreReporte = UCase("Reporte General de Apartados")
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'frmReportes.rsReport = rsReporte
            'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "PeriodoReporte"}
            'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte}


            'If (NombreEmpresa <> Nothing) Then
            '    pdvNum.Value = NombreEmpresa : pvNum.Add(pdvNum)
            '    RptVtasVEReportedeSalidadeMercancia.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
            'End If

            'If (NombreReporte <> Nothing) Then
            '    pdvNum.Value = NombreReporte : pvNum.Add(pdvNum)
            '    RptVtasVEReportedeSalidadeMercancia.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
            'End If

            'If (PeriodoReporte <> Nothing) Then
            '    pdvNum.Value = PeriodoReporte : pvNum.Add(pdvNum)
            '    RptVtasVEReportedeSalidadeMercancia.DataDefinition.ParameterFields("Periodo").ApplyCurrentValues(pvNum)
            'End If


            frmReportes.Text = "Reporte General de Apartados"
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            frmReportes.reporteActual = rptVtasReportedeApartados
            frmReportes.Show()
            Cursor = System.Windows.Forms.Cursors.Default
            FueraChange = False

        ElseIf optEstadodeCuenta.Checked = True Then

            NombreEmpresa = Trim(gstrCorpoNOMBREEMPRESA)
            NombreReporte = UCase("Estado de Cuenta de Apartados")

            If chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Checked Then
                strCancel = " ( MVC.Estatus = 'C') "
                strCancel2 = " ( DetAbo.Estatus = 'C' ) "
            End If

            If chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked Then
                strVigente = " (  (Round(Sum(Case When MVC.Moneda = 'D' Then Ing.Total Else (Ing.Total*Ing.TipoCambio) End),1) < Round(Max(Case When MVC.Moneda = 'D' Then MVC.Total + MVC.Redondeo Else (MVC.Total + MVC.Redondeo) * MVC.TipoCambio End),1) AND MVC.Estatus <> 'C') " & "    OR  (Round(Max(Case When MVC.Moneda = 'D' Then MVC.Total + MVC.Redondeo Else (MVC.Total + MVC.Redondeo) * MVC.TipoCambio End),1) = 0 And  MVC.Estatus <> 'C')  ) "
                strVigente2 = "(  ( Round(Sum(Case When DETABO.Moneda = 'D' Then Ing.Total    Else (Ing.Total*Ing.TipoCambio)      End),1) < Round(Max(DetAbo.VtaReal),1) ) And DetAbo.Estatus <> 'C' " & "   OR ( Round(Max(DetAbo.VtaReal),1) = 0 And DetAbo.Estatus <> 'C' )  ) "
            End If

            If chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
                strSaldado = " (  (Round(Sum(Case When MVC.Moneda = 'D' Then Ing.Total Else (Ing.Total*Ing.TipoCambio) End),1) >= Round(Max(Case When MVC.Moneda = 'D' Then MVC.Total + MVC.Redondeo Else (MVC.Total + MVC.Redondeo) * MVC.TipoCambio End),1) AND MVC.Estatus <> 'C') " & " AND (Round(Max(Case When MVC.Moneda = 'D' Then MVC.Total + MVC.Redondeo Else (MVC.Total + MVC.Redondeo) * MVC.TipoCambio End),1) > 0 And MVC.Estatus <> 'C')  )  "
                strSaldado2 = " (  ( Round(Sum(Case When DETABO.Moneda = 'D' Then Ing.Total    Else (Ing.Total*Ing.TipoCambio)      End),1)  >=  Round(Max(DetAbo.VtaReal),1) )  And DetAbo.Estatus <> 'C' " & " AND  ( Round(Max(DetAbo.VtaReal),1) > 0 And DetAbo.Estatus <> 'C' )  )  "
            End If

            If chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Checked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
                lHaving = strCancel & " OR " & strVigente & " OR " & strSaldado
                lHaving2 = strCancel2 & " OR " & strVigente2 & " OR " & strSaldado2
            ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Checked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                lHaving = strCancel & " OR " & strVigente
                lHaving2 = strCancel2 & " OR " & strVigente2
            ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Checked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
                lHaving = strCancel & " OR " & strSaldado
                lHaving2 = strCancel2 & " OR " & strSaldado2
            ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
                lHaving = strVigente & " OR " & strSaldado
                lHaving2 = strVigente2 & " OR " & strSaldado2
            ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Checked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                lHaving = strCancel
                lHaving2 = strCancel2
            ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                lHaving = strVigente
                lHaving2 = strVigente2
            ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
                lHaving = strSaldado
                lHaving2 = strSaldado2
            End If

            Sql = "SELECT MVC.FolioVenta,Cte.DescCliente,Cte.Domicilio,Cte.Ciudad,Cte.CP,Cte.Colonia,Cte.Rfc, MVD.NumPartida,MVC.Estatus, MVD.CodArticulo, MVD.DescArticulo, MVD.Cantidad, MVD.PrecioLista AS PrecioLista, MVD.PrecioReal AS PrecioReal, MVD.ImptePromociones + MVD.ImpteDescuentos AS Descuento, MVC.Redondeo, MVC.Moneda, " & "Round(Max(Case When MVC.Moneda = 'D' Then MVC.Total + MVC.Redondeo Else (MVC.Total + MVC.Redondeo) * MVC.TipoCambio End),1) as TotalReal, Round(Sum(Case When MVC.Moneda = 'D' Then Ing.Total Else (Ing.Total*Ing.TipoCambio) End),1)  as IngresosReal, MVC.FechaVenta, Round(IsNull(Case When DF.Moneda = 'D' Then DF.Importe Else (DF.Importe*DF.TipoCambio) End,0),1) as DifCamb " & "FROM  dbo.MovimientosVentasCab MVC INNER JOIN dbo.MovimientosVentasDet MVD (Nolock) ON MVC.FolioVenta = MVD.FolioVenta INNER JOIN dbo.Ingresos Ing ON MVC.FolioVenta = Ing.FolioMovto INNER JOIN dbo.CatClientes Cte ON MVC.CodCliente = Cte.CodCliente INNER JOIN dbo.CatAlmacen Alm ON MVC.CodSucursal = Alm.CodAlmacen Left Outer Join DiferenciaCambiaria DF (Nolock) On MVC.FolioVenta = DF.FolioVenta " & "WHERE MVC.CodCliente = " & intCodCliente & " AND MVC.TIPOMOVTO = 'A' Group by MVC.FolioVenta,Cte.DescCliente,Cte.Domicilio,Cte.Ciudad,Cte.CP,Cte.Colonia,Cte.Rfc, MVD.NumPartida, MVC.Estatus, MVD.CodArticulo, MVD.DescArticulo, MVD.Cantidad, MVD.PrecioLista, MVD.PrecioReal, MVD.ImptePromociones, MVD.ImpteDescuentos, MVC.Redondeo, MVC.Moneda, MVC.FechaVenta, DF.Importe, DF.TipoCambio, DF.Moneda " & "Having " & lHaving & "Order    by MVC.FechaVenta, MVC.FolioVenta "
            BorraCmd()
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd.CommandText = Sql
            frmReportes.rsReport = Cmd.Execute

            'Valores para el SubReporte
            Sql = "SELECT CAST(DETABO.FOLIOINGRESO AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioIngreso,CAST(DETABO.FOLIOVENTA AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioVenta,CAST(DETABO.CONCEPTO AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS Concepto,DETABO.FECHAINGRESO,Round(Case When DETABO.Moneda = 'D' Then DETABO.CARGO Else (DETABO.CARGO*DETABO.TipoCambio) End,1) as CARGO, Round(Case When DETABO.Moneda = 'D' Then DETABO.ABONO Else (DETABO.ABONO*DETABO.TipoCambio) End,1) as Abono,DETABO.TipoMovto,CAST(DetAbo.Estatus AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS Estatus,CAST(DETABO.Moneda AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS Moneda, " & "Round(Max(DetAbo.VtaReal),1) as TotalReal, Round(Sum(Case When DETABO.Moneda = 'D' Then Ing.Total Else (Ing.Total*Ing.TipoCambio) End),1) as IngresosReal FROM ((SELECT CAST(Vta.FOLIOVENTA AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FOLIOINGRESO,CAST(Vta.FOLIOVENTA AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioVenta,'APARTADO       ' AS CONCEPTO,Vta.FECHAVENTA AS FECHAINGRESO,(Vta.TOTAL + Vta.REDONDEO) AS CARGO, 0 AS ABONO, 1 AS TipoMovto, " & "CAST(Vta.Estatus AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS Estatus,CAST(Vta.Moneda AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS Moneda,Vta.TipoCambio,Case When Vta.Moneda = 'D' Then (Vta.TOTAL + Vta.REDONDEO) Else ((Vta.TOTAL + Vta.REDONDEO)*Vta.TipoCambio) End as VtaReal From MovimientosVentasCab Vta (Nolock) WHERE Vta.CODCLIENTE = " & intCodCliente & " AND Vta.TIPOMOVTO = 'A' ) Union (SELECT CAST(I.FolioIngreso AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioIngreso,CAST(MVC.FOLIOVENTA AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioVenta,CASE I.TipoIngreso WHEN 'A' THEN 'ANTICIPO       ' WHEN 'B' THEN 'ABONO          ' END AS Concepto, " & "I.FechaIngreso,0 AS CARGO,I.Total AS ABONO,2 AS TipoMovto,CAST(MVC.Estatus AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS Estatus,CAST(MVC.Moneda AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS Moneda,I.TipoCambio,Case When MVC.Moneda = 'D' Then (MVC.TOTAL + MVC.REDONDEO) Else ((MVC.TOTAL + MVC.REDONDEO)*MVC.TipoCambio) End as VtaReal FROM MovimientosVentasCab MVC INNER JOIN Ingresos I ON MVC.FolioVenta = I.FolioMovto WHERE MVC.CodCliente = " & intCodCliente & " AND MVC.TIPOMOVTO = 'A' ) " & " Union (SELECT CAST(FolioIngreso AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioIngreso,CAST(FolioVenta AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioVenta,CAST(Concepto AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS Concepto,FechaIngreso,Cargo,Abono,TipoMovto,CAST(Estatus AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS Estatus,CAST(MonedaVta AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS Moneda,TipoCambio,VtaReal FROM DetalleDevoluciones_Moneda () WHERE CodCliente = " & intCodCliente & " )) DETABO Inner Join Ingresos Ing (Nolock) On Ing.FolioMovto = DetAbo.FolioVenta " & "Group by DETABO.FOLIOINGRESO,DETABO.FOLIOVENTA,DETABO.CONCEPTO,DETABO.FECHAINGRESO,DETABO.CARGO,DETABO.ABONO,DETABO.TipoMovto,DetAbo.Estatus,DETABO.Moneda,DETABO.TipoCambio " & "Having " & lHaving2 & "ORDER BY DETABO.FolioVenta,DETABO.FechaIngreso,DETABO.TipoMovto "

            ModEstandar.BorraCmd()
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd.CommandText = Sql
            frmReportes.rsReport = Cmd.Execute

            If frmReportes.rsReport.RecordCount = 0 Then
                MsgBox("No existen movimientos que reportar" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                Exit Sub
            Else
                'frmReportes.Report = RptVtasEstadodeCuentaApartados
                'frmReportes.SubReport = RptVtasEstadodeCuentaApartados.Subreport1.OpenSubreport
                RptVtasEstadodeCuentaApartados.SetDataSource(frmReportes.rsReport)
            End If
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'frmReportes.rsReport = rsReporte
            'frmReportes.rsSubReport1 = rsSubReporte
            'frmReportes.aParam_ = New Object() {"NombreEmpresa", "NombreReporte"}
            'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte}
            frmReportes.Text = "Estado de Cuenta de Apartados"
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            frmReportes.reporteActual = RptVtasEstadodeCuentaApartados
            frmReportes.Show()
            Cursor = System.Windows.Forms.Cursors.Default
            FueraChange = False

        End If
        Cmd.CommandTimeout = 90
        Exit Sub

Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error, no es posible generar el reporte" & vbNewLine & Err.Description, MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
        FueraChange = False
    End Sub



    Sub BuscaCliente()
        On Error GoTo Merr
        gStrSql = "SELECT CodCliente,DescCliente,AlmacenVExt FROM CatClientes WHERE CodCliente = " & txtCodCliente.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("CodCliente").Value = 1 Then
                MsgBox("El Cliente Publico en General No Tiene Registrados Apartados, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodCliente.Text = ""
                txtCodCliente.Focus()
                Exit Sub
            End If
            If Not IsDBNull(RsGral.Fields("AlmacenVExt").Value) Then
                MsgBox("Los Clientes Registrados como Vendedores Externo No Tienen Registrados Apartados, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodCliente.Text = ""
                txtCodCliente.Focus()
                Exit Sub
            End If
            'txtCodCliente.Text = Format(txtCodCliente.Text, "00000")
            txtCodCliente.Text = Format(String.Concat(txtCodCliente.Text, "00000"))
            dbcCliente.Text = Trim(RsGral.Fields("DescCliente").Value)
        Else
            MsgBox("Código de Cliente No Existe, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodCliente.Text = ""
            txtCodCliente.Focus()
            Exit Sub
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
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
                txtCodSucursal.Text = Format(String.Concat(txtCodSucursal.Text, "000"))
                dbcSucursal.Text = RsGral.Fields("DescAlmacen").Value
                'dbcSucursal.Text = RsGral.Fields(1).Name
            End If
        Else
            MsgBox("Codigo de Almacen no Existe, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodSucursal.Text = ""
            txtCodSucursal.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
    End Sub

    Sub Limpiar()
        InicializaVariables()
        Nuevo()
        optReporteGeneral.Focus()
    End Sub

    Sub Nuevo()
        FueraChange = True
        optReporteGeneral.Checked = True
        chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        chkApartadosSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtCodSucursal.Text = ""
        dbcSucursal.Text = ""
        txtCodCliente.Text = ""
        dbcCliente.Text = ""
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
        chkEnVigentes.Enabled = False
        chkEnSaldados.Enabled = False
        chkEnCancelados.Enabled = False
        FueraChange = False
    End Sub

    Private Sub chkApartadosSaldados_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkApartadosSaldados.CheckStateChanged
        If chkApartadosSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
            chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked
            dtpFechaInicial.Enabled = True
            dtpFechaFinal.Enabled = True
        End If
    End Sub

    Private Sub chkApartadosSaldados_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkApartadosSaldados.Enter
        Pon_Tool()
    End Sub

    Private Sub chkSoloVigentes_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSoloVigentes.CheckStateChanged
        If chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Checked Then
            chkApartadosSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked
            dtpFechaInicial.Enabled = False
            dtpFechaFinal.Enabled = False
        Else
            Frame1.Enabled = True
            dtpFechaInicial.Enabled = True
            dtpFechaFinal.Enabled = True
        End If
    End Sub

    Private Sub chkTodaslasSucursales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.CheckStateChanged
        If chkTodaslasSucursales.CheckState = 1 Then
            txtCodSucursal.Text = ""
            dbcSucursal.Text = ""
            txtCodSucursal.Enabled = False
            dbcSucursal.Enabled = False
        ElseIf chkTodaslasSucursales.CheckState = 0 Then
            txtCodSucursal.Enabled = True
            dbcSucursal.Enabled = True
        End If
    End Sub

    Private Sub chkTodaslasSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.Enter
        Pon_Tool()
    End Sub

    Private Sub chkTodaslasSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkTodaslasSucursales.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            optReporteGeneral.Focus()
        End If
    End Sub

    Private Sub dbcCliente_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCliente.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        If Trim(dbcCliente.Text) = "" Then txtCodCliente.Text = ""
        gStrSql = "SELECT CodCliente,DescCliente FROM CatClientes (Nolock) WHERE DescCliente LIKE '" & Trim(dbcCliente.Text) & "%' AND ISNULL(AlmacenVExt,0) = 0 AND CodCliente <> 1 ORDER BY DescCliente"
        ModDCombo.DCChange(gStrSql, tecla)
        If Trim(dbcCliente.Text) = "" Then
            intCodCliente = 0
        End If
    End Sub

    Private Sub dbcCliente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCliente.Enter
        'If Screen.ActiveForm.ActiveControl.Name <> dbcCliente.Name Then
        '    Exit Sub
        'End If
        Pon_Tool()
        gStrSql = "SELECT CodCliente, DescCliente FROM CatClientes (Nolock) WHERE ISNULL(AlmacenVExt,0) = 0 AND CodCliente <> 1 ORDER BY DescCliente "
        ModDCombo.DCGotFocus(gStrSql, dbcCliente)
        '''Pon_Tool
        '''FueraChange = False
    End Sub

    Private Sub dbcCliente_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCliente.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then txtCodCliente.Focus()
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then
            chkEnVigentes.Focus()
            '''dbcCliente_LostFocus
        End If
        '   dbcCliente_LostFocus
        '   FueraChange = True
        'End If
    End Sub

    Private Sub dbcCliente_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcCliente.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcCliente_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCliente.KeyUp
        Dim Aux As String
        Aux = dbcCliente.Text
        'If dbcCliente.SelectedItem <> 0 Then
        '    dbcCliente_Leave(dbcCliente, New System.EventArgs())
        'End If
        FueraChange = True
        dbcCliente.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcCliente_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCliente.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        FueraChange = True
        intCodCliente = 0
        gStrSql = "SELECT CodCliente,DescCliente FROM CatClientes (Nolock) WHERE DescCliente LIKE '" & Trim(dbcCliente.Text) & "%' AND ISNULL(AlmacenVExt,0) = 0 AND CodCliente <> 1 ORDER BY DescCliente"
        DCLostFocus(dbcCliente, gStrSql, intCodCliente)
        If intCodCliente <> 0 Then
            'txtCodCliente.Text = Format(intCodCliente, "00000")
            txtCodCliente.Text = Format(String.Concat(intCodCliente, "00000"))
        End If
        FueraChange = False
    End Sub

    Private Sub dbcCliente_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcCliente.MouseUp
        Dim Aux As String
        Aux = dbcCliente.Text
        'If dbcCliente.SelectedItem <> 0 Then dbcCliente_Leave(dbcCliente, New System.EventArgs())
        FueraChange = True
        dbcCliente.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub

        If Trim(dbcSucursal.Text) = "" Then txtCodSucursal.Text = ""
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen "
        DCChange(gStrSql, tecla)
        'dbcSucursal_LostFocus
        '''intCodSucursal = 0
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        '''If Screen.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursal)
        '''FueraChange = False
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then txtCodSucursal.Focus()
        'If KeyCode = vbKeyReturn Then
        '   dbcSucursal_LostFocus
        '   FueraChange = True
        '   Exit Sub
        'End If
    End Sub

    Private Sub dbcSucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcSucursal.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
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

        intCodSucursal = 0
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen like '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
        If intCodSucursal <> 0 Then
            'txtCodSucursal.Text = Format(intCodSucursal, "000")
            txtCodSucursal.Text = Format(String.Concat(intCodSucursal, "000"))
        End If
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        'If dbcSucursal.SelectedItem <> 0 Then dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        FueraChange = True
        dbcSucursal.Text = Aux
        FueraChange = False
    End Sub


    Private Sub frmVtasReportedeApartados_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasReportedeApartados_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasReportedeApartados_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optReporteGeneral" And System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optEstadodeCuenta" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmVtasReportedeApartados_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasReportedeApartados_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        InicializaVariables()
        dtpFechaInicial.MinDate = C_FECHAINICIAL
        dtpFechaInicial.MaxDate = C_FECHAFINAL
        dtpFechaFinal.MinDate = C_FECHAINICIAL
        dtpFechaFinal.MaxDate = C_FECHAFINAL
        Nuevo()
    End Sub

    Private Sub frmVtasReportedeApartados_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasReportedeApartados_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optEstadodeCuenta_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optEstadodeCuenta.CheckedChanged
        If eventSender.Checked Then
            FueraChange = True
            chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkTodaslasSucursales.Enabled = False
            chkEnCancelados.Enabled = True
            chkEnVigentes.Enabled = True
            chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked
            chkEnSaldados.Enabled = True
            txtCodSucursal.Enabled = False
            txtCodSucursal.Text = ""
            txtCodSucursal.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            dbcSucursal.Enabled = False
            dbcSucursal.Text = ""
            dbcSucursal.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            Label1.Enabled = False
            chkApartadosSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkApartadosSaldados.Enabled = False
            chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkSoloVigentes.Enabled = False
            txtCodCliente.Enabled = True
            txtCodCliente.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            dbcCliente.Enabled = True
            dbcCliente.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            Label2.Enabled = True
            Frame1.Enabled = False
            dtpFechaInicial.Enabled = False
            dtpFechaFinal.Enabled = False
            FueraChange = False
        End If
    End Sub

    Private Sub optEstadodeCuenta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optEstadodeCuenta.Enter
        Pon_Tool()
    End Sub

    Private Sub optReporteGeneral_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optReporteGeneral.CheckedChanged
        If eventSender.Checked Then
            FueraChange = True
            txtCodCliente.Enabled = False
            txtCodCliente.Text = ""
            txtCodCliente.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            dbcCliente.Enabled = False
            dbcCliente.Text = ""
            dbcCliente.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            Label2.Enabled = False
            chkTodaslasSucursales.Enabled = True
            chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
            txtCodSucursal.Enabled = False
            txtCodSucursal.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            dbcSucursal.Enabled = False
            dbcSucursal.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            Label1.Enabled = True
            chkApartadosSaldados.Enabled = True
            chkSoloVigentes.Enabled = True
            chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEnCancelados.Enabled = False
            chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEnVigentes.Enabled = False
            chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEnSaldados.Enabled = False
            Frame1.Enabled = True
            dtpFechaInicial.Enabled = True
            dtpFechaFinal.Enabled = True
            FueraChange = False
        End If
    End Sub

    Private Sub optReporteGeneral_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optReporteGeneral.Enter
        Pon_Tool()
    End Sub

    Private Sub txtCodCliente_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodCliente.TextChanged
        If FueraChange Then Exit Sub
        If txtCodCliente.Text = "" Then dbcCliente.Text = ""
    End Sub

    Private Sub txtCodCliente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodCliente.Enter
        Pon_Tool()
    End Sub

    Private Sub txtCodCliente_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodCliente.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodCliente_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodCliente.Leave
        If Trim(txtCodCliente.Text) <> "" Then
            BuscaCliente()
        End If
    End Sub

    Private Sub txtCodSucursal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.TextChanged
        If FueraChange Then Exit Sub
        If txtCodSucursal.Text = "" Then txtCodSucursal.Text = ""
    End Sub

    Private Sub txtCodSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Enter
        Pon_Tool()
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


    Sub ImprimirPV()
        '        Dim RptVtasEstadodeCuentaApartados As Object
        '        'Este Proceso Genera un Reporte para el Estado de Cuenta a los folios de Apartados, para un CLiente Especifico.
        '        'El cual incluye un Reporte Pricipal con los Datos del Encabezado del Reporte y el Detalle del Apartado (Artículos) Cuyo nombre es: rptEstadoCuentaApartado.
        '        'Además se requiere un detalle de Abonos realizados hacia este apartado. Para esto, se Añadió un SubReporte para mostrar los datos de Abonos, llamado:rptEstadoCtaApartadoAbono
        '        'Se debe Enviar un Recordset a cada uno de los Reportes. Para el caso del Recordset del SubReportes, se Está haciendo uso de una función, que hace el calculo del Saldo, y lo almacena en una VAriable tipo Tabla.
        '        'Posteriormente sólo se hace una COnsulta a la FUnción que actúa como Tabla para obtener los DAtos. Esta función requiere los siguientes parametros:
        '        'FOlio de Apartado, Código del Cliente, Tipo de Moneda y TIpo de Cambo.
        '        On Error GoTo Merr
        '        Dim Sql As String
        '        Dim NombreEmpresa As String
        '        Dim NombreReporte As String
        '        Dim strCancel As String
        '        Dim strCancel2 As String
        '        Dim strVigente As String
        '        Dim strVigente2 As String
        '        Dim strSaldado As String
        '        Dim strSaldado2 As String
        '        Dim lHaving As String
        '        Dim lHaving2 As String

        '        Dim RsAux As ADODB.Recordset

        '        strCancel = ""
        '        strCancel2 = ""
        '        strVigente = ""
        '        strVigente2 = ""
        '        strSaldado = ""
        '        strSaldado2 = ""
        '        lHaving = ""
        '        lHaving2 = ""

        '        If ValidaDatos() = False Then Exit Sub

        '        NombreEmpresa = Trim(gstrCorpoNOMBREEMPRESA)
        '        NombreReporte = UCase("Estado de Cuenta de Apartados")

        '        If chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Checked Then
        '            strCancel = " (MVC.Estatus = 'C') "
        '            strCancel2 = " ( DetAbo.Estatus = 'C' ) "
        '        End If

        '        If chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked Then
        '            strVigente = " (Round(Sum(Case When MVC.Moneda = 'D' Then Ing.Total Else (Ing.Total*Ing.TipoCambio) End),1) < Round(Max(Case When MVC.Moneda = 'D' Then MVC.Total + MVC.Redondeo Else (MVC.Total + MVC.Redondeo) * MVC.TipoCambio End),1) And MVC.Estatus <> 'C' ) "
        '            strVigente2 = " ( Round(Sum(Case When DETABO.Moneda = 'D' Then Ing.Total    Else (Ing.Total*Ing.TipoCambio)      End),1) <  Round(Max(Case When DETABO.Moneda = 'D' Then DetAbo.VtaReal Else (DetAbo.VtaReal * DetAbo.TipoCambio) End),1) And DetAbo.Estatus <> 'C' ) "
        '        End If

        '        If chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
        '            strSaldado = " (Round(Sum(Case When MVC.Moneda = 'D' Then Ing.Total Else (Ing.Total*Ing.TipoCambio) End),1) >= Round(Max(Case When MVC.Moneda = 'D' Then MVC.Total + MVC.Redondeo Else (MVC.Total + MVC.Redondeo) * MVC.TipoCambio End),1) And MVC.Estatus <> 'C' ) "
        '            strSaldado2 = " ( Round(Sum(Case When DETABO.Moneda = 'D' Then Ing.Total    Else (Ing.Total*Ing.TipoCambio)      End),1) >= Round(Max(Case When DETABO.Moneda = 'D' Then DetAbo.VtaReal Else (DetAbo.VtaReal * DetAbo.TipoCambio) End),1) And DetAbo.Estatus <> 'C' ) "
        '        End If

        '        If chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Checked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
        '            lHaving = strCancel & " OR " & strVigente & " OR " & strSaldado
        '            lHaving2 = strCancel2 & " OR " & strVigente2 & " OR " & strSaldado2
        '        ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Checked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked Then
        '            lHaving = strCancel & " OR " & strVigente
        '            lHaving2 = strCancel2 & " OR " & strVigente2
        '        ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Checked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
        '            lHaving = strCancel & " OR " & strSaldado
        '            lHaving2 = strCancel2 & " OR " & strSaldado2
        '        ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
        '            lHaving = strVigente & " OR " & strSaldado
        '            lHaving2 = strVigente2 & " OR " & strSaldado2
        '        ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Checked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked Then
        '            lHaving = strCancel
        '            lHaving2 = strCancel2
        '        ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Checked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked Then
        '            lHaving = strVigente
        '            lHaving2 = strVigente2
        '        ElseIf chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Checked Then
        '            lHaving = strSaldado
        '            lHaving2 = strSaldado2
        '        End If

        '        Sql = "SELECT MVC.FolioVenta,Cte.DescCliente,Cte.Domicilio,Cte.Ciudad,Cte.CP,Cte.Colonia,Cte.Rfc, MVD.NumPartida,MVC.Estatus, MVD.CodArticulo, MVD.DescArticulo, MVD.Cantidad, MVD.PrecioLista AS PrecioLista, MVD.PrecioReal AS PrecioReal, MVD.ImptePromociones + MVD.ImpteDescuentos AS Descuento, MVC.Redondeo, MVC.Moneda, " & "Round(Max(Case When MVC.Moneda = 'D' Then MVC.Total + MVC.Redondeo Else (MVC.Total + MVC.Redondeo) * MVC.TipoCambio End),1) as TotalReal, Round(Sum(Case When MVC.Moneda = 'D' Then Ing.Total Else (Ing.Total*Ing.TipoCambio) End),1)  as IngresosReal, MVC.FechaVenta, Round(IsNull(Case When DF.Moneda = 'D' Then DF.Importe Else (DF.Importe*DF.TipoCambio) End,0),1) as DifCamb " & "FROM  dbo.MovimientosVentasCab MVC INNER JOIN dbo.MovimientosVentasDet MVD (Nolock) ON MVC.FolioVenta = MVD.FolioVenta INNER JOIN dbo.Ingresos Ing ON MVC.FolioVenta = Ing.FolioMovto INNER JOIN dbo.CatClientes Cte ON MVC.CodCliente = Cte.CodCliente INNER JOIN dbo.CatAlmacen Alm ON MVC.CodSucursal = Alm.CodAlmacen Left Outer Join DiferenciaCambiaria DF (Nolock) On MVC.FolioVenta = DF.FolioVenta " & "WHERE MVC.CodCliente = " & intCodCliente & " AND MVC.TIPOMOVTO = 'A' Group by MVC.FolioVenta,Cte.DescCliente,Cte.Domicilio,Cte.Ciudad,Cte.CP,Cte.Colonia,Cte.Rfc, MVD.NumPartida, MVC.Estatus, MVD.CodArticulo, MVD.DescArticulo, MVD.Cantidad, MVD.PrecioLista, MVD.PrecioReal, MVD.ImptePromociones, MVD.ImpteDescuentos, MVC.Redondeo, MVC.Moneda, MVC.FechaVenta, DF.Importe, DF.TipoCambio, DF.Moneda " & "Having " & lHaving & "Order    by MVC.FechaVenta, MVC.FolioVenta "
        '        '''         "Having (MVC.Estatus = 'C') " & _
        '        ''''         " OR (Round(Sum(Case When MVC.Moneda = 'D' Then Ing.Total Else (Ing.Total*Ing.TipoCambio) End),1) < Round(Max(Case When MVC.Moneda = 'D' Then MVC.Total + MVC.Redondeo Else (MVC.Total + MVC.Redondeo) * MVC.TipoCambio End),1) And MVC.Estatus <> 'C' ) " & _
        '        ''''         " OR (Round(Sum(Case When MVC.Moneda = 'D' Then Ing.Total Else (Ing.Total*Ing.TipoCambio) End),1) >= Round(Max(Case When MVC.Moneda = 'D' Then MVC.Total + MVC.Redondeo Else (MVC.Total + MVC.Redondeo) * MVC.TipoCambio End),1) And MVC.Estatus <> 'C' ) "
        '        BorraCmd()
        '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        '        Cmd.CommandText = Sql
        '        rsReporte = Cmd.Execute

        '        'Valores para el SubReporte
        '        Sql = "SELECT DETABO.FOLIOINGRESO, DETABO.FOLIOVENTA, DETABO.CONCEPTO, DETABO.FECHAINGRESO, Round(Case When DETABO.Moneda = 'D' Then DETABO.CARGO Else (DETABO.CARGO*DETABO.TipoCambio) End,1) as CARGO, Round(Case When DETABO.Moneda = 'D' Then DETABO.ABONO Else (DETABO.ABONO*DETABO.TipoCambio) End,1) as Abono, DETABO.TipoMovto, DetAbo.Estatus, DETABO.Moneda, " & "Round(Max(Case When DETABO.Moneda = 'D' Then DetAbo.VtaReal Else (DetAbo.VtaReal * DetAbo.TipoCambio) End),1) as TotalReal, Round(Sum(Case When DETABO.Moneda = 'D' Then Ing.Total    Else (Ing.Total*Ing.TipoCambio)      End),1) as IngresosReal " & "FROM ( (SELECT Vta.FOLIOVENTA AS FOLIOINGRESO, Vta.FOLIOVENTA, 'APARTADO       ' AS CONCEPTO, Vta.FECHAVENTA AS FECHAINGRESO, (Vta.TOTAL + Vta.REDONDEO) AS CARGO, 0 AS ABONO, 1 AS TipoMovto, Vta.Estatus, Vta.Moneda, Vta.TipoCambio, (Vta.TOTAL + Vta.REDONDEO) as VtaReal " & "From MovimientosVentasCab Vta (Nolock) WHERE Vta.CODCLIENTE = " & intCodCliente & " AND Vta.TIPOMOVTO = 'A' ) UNION " & "(SELECT I.FolioIngreso,MVC.FOLIOVENTA, CASE I.TipoIngreso WHEN 'A' THEN 'ANTICIPO       ' WHEN 'B' THEN 'ABONO          ' END AS Concepto, I.FechaIngreso, 0 AS CARGO, I.Total AS ABONO, 2 AS TipoMovto, MVC.Estatus, MVC.Moneda, I.TipoCambio, (MVC.TOTAL + MVC.REDONDEO) as VtaReal " & "FROM MovimientosVentasCab MVC INNER JOIN Ingresos I ON MVC.FolioVenta = I.FolioMovto WHERE MVC.CodCliente = " & intCodCliente & " AND MVC.TIPOMOVTO = 'A' ) Union " & "(SELECT FolioIngreso, FolioVenta, Concepto, FechaIngreso, Cargo, Abono, TipoMovto, Estatus, MonedaVta, TipoCambio, VtaReal FROM DetalleDevoluciones_Moneda () WHERE CodCliente = " & intCodCliente & " )  ) DETABO Inner Join Ingresos Ing (Nolock) On Ing.FolioMovto = DetAbo.FolioVenta " & "Group  by DETABO.FOLIOINGRESO, DETABO.FOLIOVENTA, DETABO.CONCEPTO, DETABO.FECHAINGRESO, DETABO.CARGO, DETABO.ABONO, DETABO.TipoMovto, DetAbo.Estatus, DETABO.Moneda, DETABO.TipoCambio " & "Having " & lHaving2 & "ORDER   BY DETABO.FolioVenta, DETABO.FechaIngreso, DETABO.TipoMovto "
        '        '''         "Having  ( DetAbo.Estatus = 'C' ) " & _
        '        ''''         "OR ( Round(Sum(Case When DETABO.Moneda = 'D' Then Ing.Total    Else (Ing.Total*Ing.TipoCambio)      End),1) >= Round(Max(Case When DETABO.Moneda = 'D' Then DetAbo.VtaReal Else (DetAbo.VtaReal * DetAbo.TipoCambio) End),1) And DetAbo.Estatus <> 'C' ) " & _
        '        ''''         "Or ( Round(Sum(Case When DETABO.Moneda = 'D' Then Ing.Total    Else (Ing.Total*Ing.TipoCambio)      End),1) <  Round(Max(Case When DETABO.Moneda = 'D' Then DetAbo.VtaReal Else (DetAbo.VtaReal * DetAbo.TipoCambio) End),1) And DetAbo.Estatus <> 'C' ) " & _
        '        '
        '        ModEstandar.BorraCmd()
        '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        '        Cmd.CommandText = Sql
        '        rsSubReporte = Cmd.Execute
        '        If rsReporte.RecordCount = 0 Then
        '            MsgBox("No Existen Movimientos En Este Periodo de Fechas, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
        '            Exit Sub
        '        Else
        '            frmReportes.Report = RptVtasEstadodeCuentaApartados
        '            frmReportes.SubReport = RptVtasEstadodeCuentaApartados.Subreport1.OpenSubreport
        '        End If
        '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        '        frmReportes.rsReport = rsReporte
        '        frmReportes.rsSubReport1 = rsSubReporte
        '        frmReportes.aParam_ = New Object() {"NombreEmpresa", "NombreReporte"}
        '        frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte}
        '        frmReportes.Text = "Estado de Cuenta de Apartados"
        '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        '        frmReportes.Show()
        '        Cursor = System.Windows.Forms.Cursors.Default
        '        FueraChange = False

        'Merr:
        '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function ValidaDatos() As Boolean
        If optReporteGeneral.Checked Then
            If (chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked) And Trim(dbcSucursal.Text) = "" Then
                MsgBox(C_msgFALTADATO & "Nombre de la sucursal", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                dbcCliente.Focus()
                Exit Function
            End If
            If (chkApartadosSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked And chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked) Then
                MsgBox("No ha sido seleccionado ningún estatus", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                chkApartadosSaldados.Focus()
                Exit Function
            End If
        ElseIf optEstadodeCuenta.Checked Then
            'Validar si la FEcha final es Mayor que la Inicial.
            If Trim(dbcCliente.Text) = "" Then
                MsgBox(C_msgFALTADATO & "Nombre del Cliente.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                dbcCliente.Focus()
                Exit Function
            End If
            If (chkEnVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnSaldados.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEnCancelados.CheckState = System.Windows.Forms.CheckState.Unchecked) Then
                MsgBox("No ha sido seleccionado ningún estatus", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                chkEnVigentes.Focus()
                Exit Function
            End If
        End If
        ValidaDatos = True
        'Return ValidaDatos
    End Function

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

    Private Sub dtpFechaInicial_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub dtpFechaInicial_CursorChanged(sender As Object, e As EventArgs)

    End Sub
End Class