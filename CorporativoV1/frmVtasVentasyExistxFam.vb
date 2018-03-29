Option Strict Off
Option Explicit On
Imports Microsoft.Office.Interop
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasVentasyExistxFam
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''' ****************************************************************************************************************************************************'
    ''' REPORTE DE VENTAS Y EXISTENCIA POR FAMILIA - MUESTRA SOLO LOS ARTICULOS VENDIDOS EN EL PERIODO Y LA EXISTENCIA DE ESTOS
    ''' SE BASO EN EL REPORTE DE VTASYEXIST POR PROVEEDOR
    ''' GENERA 2 REPORTES: 1 POR ARTICULO Y EL OTRO CONCENTRADO POR PESO-CT/COLOR/PUREZA-Q
    ''' 27OCT2010 - MAVF Ver
    '''
    ''' MODIFICACION.- SE AGREGO TOTAL POR GPO Y TOTAL FINAL - SE CREO NUEVO CONCENTRADO POR PUREZA
    ''' 18NOV2010 - MAVF Ver
    '''
    ''' Ver 1.1       Estatus: Aprobado
    '*******************************************************************************************************************************************************'
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkmdsPureza As System.Windows.Forms.CheckBox
    Public WithEvents chkmdsColor As System.Windows.Forms.CheckBox
    Public WithEvents chkmdsPeso As System.Windows.Forms.CheckBox
    Public WithEvents txtMDSPeso2 As System.Windows.Forms.TextBox
    Public WithEvents txtMDSPeso As System.Windows.Forms.TextBox
    Public WithEvents txtMDSColor As System.Windows.Forms.TextBox
    Public WithEvents txtMDSPureza As System.Windows.Forms.TextBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents fraDiamanteSuelto As System.Windows.Forms.GroupBox
    Public WithEvents chkTodasSuc As System.Windows.Forms.CheckBox
    Public WithEvents _fraVtas_0 As System.Windows.Forms.GroupBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents chkConcentrado As System.Windows.Forms.CheckBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents dbcJFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcJLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcJSubLinea As System.Windows.Forms.ComboBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_3 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_4 As System.Windows.Forms.Label
    Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"
    Const C_OPCION As String = "-"

    Const C_ENCABEZADO As Short = 9
    Const C_INICIAL As Short = 3

    Dim msglTiempoCambioI As Single
    Dim msglTiempoCambioF As Single
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean
    Dim mblnFueraChange As Boolean

    Dim mintCodSucursal As Short
    Dim mintJFamilia As Short
    Dim mintJLinea As Short
    Dim mintJSubLinea As Short

    Dim lSucursal As Short
    Dim lFamilia As Short
    Dim lLinea As Short
    Dim lSubLinea As Short
    Dim lPeso As Decimal
    Dim lPeso2 As Decimal
    Dim lColor As String
    Dim lPureza As String
    Dim lFechaIni As String
    Dim lFechaFin As String
    Dim lstrFamilia As String

    Dim tecla As Short
    Dim mblnSalir As Boolean

    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    Dim RsAux As ADODB.Recordset
    Dim RSColor As ADODB.Recordset
    Dim ObjExcel As Object
    Dim objLibro As Excel.Workbook
    Dim objHoja As Excel.Worksheet
    Dim ObjExcelCon As Object
    Dim objLibroCon As Excel.Workbook
    Dim objHojaCon As Excel.Worksheet
    Dim ColumSepar As Integer
    Dim ColumCtoL As Integer
    Dim ColumnaCtoExis As Integer 'Variable para saber donde esta la columna que muestra el importe costeado de la existencia
    Dim RenFinal As Integer 'Variable para saber el renglon final

    Dim Renglon As Short 'Var global para renglon del archivo de EXCEL
    Dim lnumColumnas As Short
    Dim lnumIniE As Short
    Dim lnumFinE As Short
    Dim lnumIniV As Short
    Dim lnumFinV As Short
    Dim lnumColE As Short 'Var global para calculo de celdas en encabezado de reporte
    Dim lArchivo01 As String 'Nombre de archivos
    Dim lArchivo02 As String 'Nombre de archivos
    Dim RsConP As ADODB.Recordset 'RecordSet para concentrado de Pureza
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim RenglonPureza As Short 'Renglon para calculo de formula suma en concentrado pureza
    'Dim cmd As ADODB.Command

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.chkConcentrado = New System.Windows.Forms.CheckBox()
        Me.fraDiamanteSuelto = New System.Windows.Forms.GroupBox()
        Me.chkmdsPureza = New System.Windows.Forms.CheckBox()
        Me.chkmdsColor = New System.Windows.Forms.CheckBox()
        Me.chkmdsPeso = New System.Windows.Forms.CheckBox()
        Me.txtMDSPeso2 = New System.Windows.Forms.TextBox()
        Me.txtMDSPeso = New System.Windows.Forms.TextBox()
        Me.txtMDSColor = New System.Windows.Forms.TextBox()
        Me.txtMDSPureza = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.chkTodasSuc = New System.Windows.Forms.CheckBox()
        Me._fraVtas_0 = New System.Windows.Forms.GroupBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.dbcJFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcJLinea = New System.Windows.Forms.ComboBox()
        Me.dbcJSubLinea = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me._lblVentas_3 = New System.Windows.Forms.Label()
        Me._lblVentas_4 = New System.Windows.Forms.Label()
        Me.fraVtas = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.fraDiamanteSuelto.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame2.SuspendLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(6, 13)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(345, 77)
        Me.txtMensaje.TabIndex = 26
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'chkConcentrado
        '
        Me.chkConcentrado.BackColor = System.Drawing.SystemColors.Control
        Me.chkConcentrado.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkConcentrado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkConcentrado.Location = New System.Drawing.Point(14, 211)
        Me.chkConcentrado.Margin = New System.Windows.Forms.Padding(2)
        Me.chkConcentrado.Name = "chkConcentrado"
        Me.chkConcentrado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkConcentrado.Size = New System.Drawing.Size(136, 20)
        Me.chkConcentrado.TabIndex = 19
        Me.chkConcentrado.Text = "Generar Concentrado"
        Me.ToolTip1.SetToolTip(Me.chkConcentrado, "Indica la generación del concentrado de ventas Peso-Color-Pureza")
        Me.chkConcentrado.UseVisualStyleBackColor = False
        '
        'fraDiamanteSuelto
        '
        Me.fraDiamanteSuelto.BackColor = System.Drawing.SystemColors.Control
        Me.fraDiamanteSuelto.Controls.Add(Me.chkmdsPureza)
        Me.fraDiamanteSuelto.Controls.Add(Me.chkmdsColor)
        Me.fraDiamanteSuelto.Controls.Add(Me.chkmdsPeso)
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSPeso2)
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSPeso)
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSColor)
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSPureza)
        Me.fraDiamanteSuelto.Controls.Add(Me.Label4)
        Me.fraDiamanteSuelto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraDiamanteSuelto.Location = New System.Drawing.Point(87, 128)
        Me.fraDiamanteSuelto.Margin = New System.Windows.Forms.Padding(2)
        Me.fraDiamanteSuelto.Name = "fraDiamanteSuelto"
        Me.fraDiamanteSuelto.Padding = New System.Windows.Forms.Padding(2)
        Me.fraDiamanteSuelto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraDiamanteSuelto.Size = New System.Drawing.Size(217, 78)
        Me.fraDiamanteSuelto.TabIndex = 10
        Me.fraDiamanteSuelto.TabStop = False
        '
        'chkmdsPureza
        '
        Me.chkmdsPureza.BackColor = System.Drawing.SystemColors.Control
        Me.chkmdsPureza.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkmdsPureza.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.chkmdsPureza.Location = New System.Drawing.Point(9, 57)
        Me.chkmdsPureza.Margin = New System.Windows.Forms.Padding(2)
        Me.chkmdsPureza.Name = "chkmdsPureza"
        Me.chkmdsPureza.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkmdsPureza.Size = New System.Drawing.Size(79, 15)
        Me.chkmdsPureza.TabIndex = 17
        Me.chkmdsPureza.Text = "Pureza - Q"
        Me.chkmdsPureza.UseVisualStyleBackColor = False
        '
        'chkmdsColor
        '
        Me.chkmdsColor.BackColor = System.Drawing.SystemColors.Control
        Me.chkmdsColor.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkmdsColor.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.chkmdsColor.Location = New System.Drawing.Point(9, 35)
        Me.chkmdsColor.Margin = New System.Windows.Forms.Padding(2)
        Me.chkmdsColor.Name = "chkmdsColor"
        Me.chkmdsColor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkmdsColor.Size = New System.Drawing.Size(64, 15)
        Me.chkmdsColor.TabIndex = 15
        Me.chkmdsColor.Text = "Color"
        Me.chkmdsColor.UseVisualStyleBackColor = False
        '
        'chkmdsPeso
        '
        Me.chkmdsPeso.BackColor = System.Drawing.SystemColors.Control
        Me.chkmdsPeso.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkmdsPeso.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.chkmdsPeso.Location = New System.Drawing.Point(9, 15)
        Me.chkmdsPeso.Margin = New System.Windows.Forms.Padding(2)
        Me.chkmdsPeso.Name = "chkmdsPeso"
        Me.chkmdsPeso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkmdsPeso.Size = New System.Drawing.Size(71, 15)
        Me.chkmdsPeso.TabIndex = 11
        Me.chkmdsPeso.Text = "Peso - CT"
        Me.chkmdsPeso.UseVisualStyleBackColor = False
        '
        'txtMDSPeso2
        '
        Me.txtMDSPeso2.AcceptsReturn = True
        Me.txtMDSPeso2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtMDSPeso2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSPeso2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSPeso2.Location = New System.Drawing.Point(166, 12)
        Me.txtMDSPeso2.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMDSPeso2.MaxLength = 6
        Me.txtMDSPeso2.Name = "txtMDSPeso2"
        Me.txtMDSPeso2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSPeso2.Size = New System.Drawing.Size(38, 20)
        Me.txtMDSPeso2.TabIndex = 14
        Me.txtMDSPeso2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMDSPeso
        '
        Me.txtMDSPeso.AcceptsReturn = True
        Me.txtMDSPeso.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtMDSPeso.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSPeso.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSPeso.Location = New System.Drawing.Point(92, 11)
        Me.txtMDSPeso.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMDSPeso.MaxLength = 6
        Me.txtMDSPeso.Name = "txtMDSPeso"
        Me.txtMDSPeso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSPeso.Size = New System.Drawing.Size(38, 20)
        Me.txtMDSPeso.TabIndex = 12
        Me.txtMDSPeso.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMDSColor
        '
        Me.txtMDSColor.AcceptsReturn = True
        Me.txtMDSColor.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtMDSColor.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSColor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSColor.Location = New System.Drawing.Point(92, 32)
        Me.txtMDSColor.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMDSColor.MaxLength = 1
        Me.txtMDSColor.Name = "txtMDSColor"
        Me.txtMDSColor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSColor.Size = New System.Drawing.Size(38, 20)
        Me.txtMDSColor.TabIndex = 16
        Me.txtMDSColor.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtMDSPureza
        '
        Me.txtMDSPureza.AcceptsReturn = True
        Me.txtMDSPureza.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtMDSPureza.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSPureza.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSPureza.Location = New System.Drawing.Point(92, 54)
        Me.txtMDSPureza.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMDSPureza.MaxLength = 4
        Me.txtMDSPureza.Name = "txtMDSPureza"
        Me.txtMDSPureza.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSPureza.Size = New System.Drawing.Size(38, 20)
        Me.txtMDSPureza.TabIndex = 18
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(146, 14)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(16, 15)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "al"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'chkTodasSuc
        '
        Me.chkTodasSuc.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodasSuc.Checked = True
        Me.chkTodasSuc.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodasSuc.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodasSuc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodasSuc.Location = New System.Drawing.Point(14, 13)
        Me.chkTodasSuc.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodasSuc.Name = "chkTodasSuc"
        Me.chkTodasSuc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodasSuc.Size = New System.Drawing.Size(128, 15)
        Me.chkTodasSuc.TabIndex = 0
        Me.chkTodasSuc.Text = "Todas las sucursales"
        Me.chkTodasSuc.UseVisualStyleBackColor = False
        '
        '_fraVtas_0
        '
        Me._fraVtas_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraVtas.SetIndex(Me._fraVtas_0, CType(0, Short))
        Me._fraVtas_0.Location = New System.Drawing.Point(7, 32)
        Me._fraVtas_0.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.Name = "_fraVtas_0"
        Me._fraVtas_0.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_0.Size = New System.Drawing.Size(298, 2)
        Me._fraVtas_0.TabIndex = 2
        Me._fraVtas_0.TabStop = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.txtMensaje)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(14, 297)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(362, 105)
        Me.Frame4.TabIndex = 25
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Mensaje Adicional para el reporte ..."
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.dtpFechaInicial)
        Me.Frame2.Controls.Add(Me.dtpFechaFinal)
        Me.Frame2.Controls.Add(Me.Label2)
        Me.Frame2.Controls.Add(Me.Label3)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(14, 235)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(362, 46)
        Me.Frame2.TabIndex = 20
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(62, 17)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(99, 20)
        Me.dtpFechaInicial.TabIndex = 22
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(239, 17)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(95, 20)
        Me.dtpFechaFinal.TabIndex = 24
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(18, 20)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(44, 17)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Desde"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(197, 20)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(38, 17)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "Hasta"
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(146, 8)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(194, 21)
        Me.dbcSucursal.TabIndex = 1
        '
        'dbcJFamilia
        '
        Me.dbcJFamilia.Location = New System.Drawing.Point(87, 62)
        Me.dbcJFamilia.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcJFamilia.Name = "dbcJFamilia"
        Me.dbcJFamilia.Size = New System.Drawing.Size(182, 21)
        Me.dbcJFamilia.TabIndex = 7
        '
        'dbcJLinea
        '
        Me.dbcJLinea.Location = New System.Drawing.Point(87, 82)
        Me.dbcJLinea.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcJLinea.Name = "dbcJLinea"
        Me.dbcJLinea.Size = New System.Drawing.Size(182, 21)
        Me.dbcJLinea.TabIndex = 8
        '
        'dbcJSubLinea
        '
        Me.dbcJSubLinea.Location = New System.Drawing.Point(87, 102)
        Me.dbcJSubLinea.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcJSubLinea.Name = "dbcJSubLinea"
        Me.dbcJSubLinea.Size = New System.Drawing.Size(182, 21)
        Me.dbcJSubLinea.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(14, 41)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(55, 18)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "JOYERIA"
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_0, CType(0, Short))
        Me._lblVentas_0.Location = New System.Drawing.Point(44, 64)
        Me._lblVentas_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(39, 13)
        Me._lblVentas_0.TabIndex = 4
        Me._lblVentas_0.Text = "Familia"
        '
        '_lblVentas_3
        '
        Me._lblVentas_3.AutoSize = True
        Me._lblVentas_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_3, CType(3, Short))
        Me._lblVentas_3.Location = New System.Drawing.Point(45, 85)
        Me._lblVentas_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_3.Name = "_lblVentas_3"
        Me._lblVentas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_3.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_3.TabIndex = 5
        Me._lblVentas_3.Text = "Línea"
        '
        '_lblVentas_4
        '
        Me._lblVentas_4.AutoSize = True
        Me._lblVentas_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_4, CType(4, Short))
        Me._lblVentas_4.Location = New System.Drawing.Point(32, 105)
        Me._lblVentas_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_4.Name = "_lblVentas_4"
        Me._lblVentas_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_4.Size = New System.Drawing.Size(54, 13)
        Me._lblVentas_4.TabIndex = 6
        Me._lblVentas_4.Text = "SubLínea"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(138, 415)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 79
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(23, 415)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 78
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(253, 416)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 77
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmVtasVentasyExistxFam
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(395, 471)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.fraDiamanteSuelto)
        Me.Controls.Add(Me.chkTodasSuc)
        Me.Controls.Add(Me._fraVtas_0)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.chkConcentrado)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.dbcJFamilia)
        Me.Controls.Add(Me.dbcJLinea)
        Me.Controls.Add(Me.dbcJSubLinea)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me._lblVentas_0)
        Me.Controls.Add(Me._lblVentas_3)
        Me.Controls.Add(Me._lblVentas_4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(371, 169)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmVtasVentasyExistxFam"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Ventas y Existencias por Familia"
        Me.fraDiamanteSuelto.ResumeLayout(False)
        Me.fraDiamanteSuelto.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub



    Public Sub Limpiar()
        On Error Resume Next
        Nuevo()
        chkTodasSuc.Focus()
    End Sub

    Public Sub Nuevo()
        chkTodasSuc.CheckState = System.Windows.Forms.CheckState.Checked
        chkTodasSuc_CheckStateChanged(chkTodasSuc, New System.EventArgs())

        mintCodSucursal = 0
        mintJFamilia = 0
        mintJLinea = 0
        mintJSubLinea = 0

        lSucursal = 0
        lFamilia = 0
        lLinea = 0
        lSubLinea = 0
        lPeso = 0
        lPeso2 = 0
        lColor = ""
        lPureza = ""
        lFechaIni = ""
        lFechaFin = ""
        lstrFamilia = ""

        mblnFueraChange = True
        dbcSucursal.Text = ""
        dbcJFamilia.Text = ""
        dbcJLinea.Text = ""
        dbcJSubLinea.Text = ""
        mblnFueraChange = False

        lnumColumnas = 0
        lnumIniE = 0
        lnumFinE = 0
        lnumIniV = 0
        lnumFinV = 0
        lnumColE = 0
        lArchivo01 = ""
        lArchivo02 = ""

        chkmdsPeso.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtMDSPeso.Enabled = False
        txtMDSPeso2.Enabled = False
        chkmdsColor.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtMDSColor.Enabled = False
        chkmdsPureza.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtMDSPureza.Enabled = False

        chkConcentrado.CheckState = System.Windows.Forms.CheckState.Unchecked

        dtpFechaInicial.Value = Format(Today, "dd/MMM/yyyy")
        dtpFechaFinal.Value = Format(Today, "dd/MMM/yyyy")
        txtMensaje.Text = ""
        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
    End Sub

    Function DevuelveQuery() As String
        On Error GoTo Err_Renamed
        Dim lStrSql As String

        lStrSql = ""
        '''lStrSql = lStrSql & "Select   CodProveedor, CodArticulo, DescArticulo, CodigoArticuloProv, OrigenAnt, CodigoAnt, PrecioPubDolar, CostoReal, mdsPeso, mdsColor, mdsPureza, mdsCertificado, CodFamilia, CodLinea, CodSubLinea, Articulo, CodigoArticuloP, PrecioReal, Exist, Vta "
        '''lStrSql = lStrSql & "Select   CodProveedor, CodArticulo, DescArticulo, CodigoArticuloProv, OrigenAnt, CodigoAnt, PrecioPubDolar, CostoReal, mdsPeso, mdsColor, mdsPureza, mdsCertificado, CodFamilia, CodLinea, CodSubLinea, Articulo, CodigoArticuloP, PrecioReal, Vta, Exist "
        lStrSql = lStrSql & "Select   CodArticulo, DescArticulo, mdsPeso, mdsColor, mdsPureza, mdsCertificado, PrecioPCat, CostoCat, CodFamilia, CodLinea, CodSubLinea, Articulo, Vta, Exist "
        lStrSql = lStrSql & "From     dbo.fnRptVtasyExistxFamilia ('" & lFechaIni & "','" & lFechaFin & "',  " & lSucursal & ",  " & lFamilia & "," & lLinea & "," & lSubLinea & ",  " & VB6.Format(lPeso, "#####0.00") & "," & VB6.Format(lPeso2, "#####0.00") & ",'" & lColor & "','" & lPureza & "')"
        lStrSql = lStrSql & "ORDER    BY CodArticulo "
        DevuelveQuery = lStrSql

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function DevuelveQuery_Concentrado() As String
        On Error GoTo Err_Renamed
        Dim RsAux As ADODB.Recordset
        Dim lStrSql As String
        Dim lstrVta As String
        Dim lstrExs As String
        Dim cSelect As String
        Dim cWHERE As String
        Dim cGroup As String
        Dim cOrder As String
        Dim lVtaCmp As String
        Dim lExsCmp As String

        lStrSql = ""
        lstrVta = ""
        lstrExs = ""
        cSelect = ""
        cWHERE = ""
        cGroup = ""
        cOrder = ""
        lVtaCmp = ""
        lExsCmp = ""

        cSelect = "Select mdsPeso, mdsPureza, "

        lStrSql = "Select Color From catColores (Nolock) Where lTrim(rTrim(Descripcion)) <> '' Order by codigo "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, lStrSql))
        RsAux = Cmd.Execute
        RSColor = Cmd.Execute

        If RsAux.RecordCount > 0 Then
            Do While Not RsAux.EOF
                lVtaCmp = lVtaCmp & "C_" & Trim(RsAux.Fields("Color").Value) & ","
                lstrVta = lstrVta & "Sum(Case When mdsColor = '" & RsAux.Fields("Color").Value & "' Then Vta   Else 0 End) as C_" & Trim(RsAux.Fields("Color").Value) & ","
                lExsCmp = lExsCmp & "E_" & Trim(RsAux.Fields("Color").Value) & ","
                lstrExs = lstrExs & "Sum(Case When mdsColor = '" & RsAux.Fields("Color").Value & "' Then Exist Else 0 End) as E_" & Trim(RsAux.Fields("Color").Value) & ","
                RsAux.MoveNext()
            Loop
            lstrExs = Mid(lstrExs, 1, Len(lstrExs) - 1) & " "
            lExsCmp = Mid(lExsCmp, 1, Len(lExsCmp) - 1) & " "

            cWHERE = " From dbo.fnRptVtasyExistxFamilia ('" & lFechaIni & "','" & lFechaFin & "',  " & lSucursal & ",  " & lFamilia & "," & lLinea & "," & lSubLinea & ",  " & lPeso & "," & lPeso2 & ",'" & lColor & "','" & lPureza & "') "
            cGroup = " Group by mdsPeso, mdsPureza "
            cOrder = " Order by mdsPeso, mdsPureza "

            lStrSql = cSelect & lstrVta & lstrExs & cWHERE & cGroup & cOrder
            DevuelveQuery_Concentrado = lStrSql
        Else
            DevuelveQuery_Concentrado = ""
        End If

        ''' PARAMETROS PARA DETERMINAR COLUMNAS ( COLORES ) PARA EL REPORTE EN EXCEL
        If Not RSColor.EOF Then
            RSColor.MoveFirst()
            lnumColumnas = RSColor.RecordCount
            lnumIniE = C_INICIAL + 2 '''SIEMPRE ES COLUMNA 5
            lnumFinE = lnumIniE + (lnumColumnas - 1)
            lnumIniV = lnumFinE + 4
            lnumFinV = lnumIniV + (lnumColumnas - 1)
            lnumColE = 4 + 1 + (lnumColumnas * 2)
        Else
        End If
        ''' **************************************************************************

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function DevuelveQuery_ConcentradoPureza() As String
        On Error GoTo Err_Renamed
        Dim RsAux As ADODB.Recordset
        Dim lStrSql As String
        Dim lstrVta As String
        Dim lstrExs As String
        Dim cSelect As String
        Dim cWHERE As String
        Dim cGroup As String
        Dim cOrder As String
        Dim lVtaCmp As String
        Dim lExsCmp As String

        lStrSql = ""
        lstrVta = ""
        lstrExs = ""
        cSelect = ""
        cWHERE = ""
        cGroup = ""
        cOrder = ""
        lVtaCmp = ""
        lExsCmp = ""

        cSelect = "Select mdsPureza, "

        lStrSql = "Select Color From catColores (Nolock) Where lTrim(rTrim(Descripcion)) <> '' Order by codigo "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, lStrSql))
        RsAux = Cmd.Execute
        RSColor = Cmd.Execute

        If RsAux.RecordCount > 0 Then
            Do While Not RsAux.EOF
                lVtaCmp = lVtaCmp & "C_" & Trim(RsAux.Fields("Color").Value) & ","
                lstrVta = lstrVta & "Sum(Case When mdsColor = '" & RsAux.Fields("Color").Value & "' Then Vta   Else 0 End) as C_" & Trim(RsAux.Fields("Color").Value) & ","
                lExsCmp = lExsCmp & "E_" & Trim(RsAux.Fields("Color").Value) & ","
                lstrExs = lstrExs & "Sum(Case When mdsColor = '" & RsAux.Fields("Color").Value & "' Then Exist Else 0 End) as E_" & Trim(RsAux.Fields("Color").Value) & ","
                RsAux.MoveNext()
            Loop
            lstrExs = Mid(lstrExs, 1, Len(lstrExs) - 1) & " "
            lExsCmp = Mid(lExsCmp, 1, Len(lExsCmp) - 1) & " "

            cWHERE = " From dbo.fnRptVtasyExistxFamilia ('" & lFechaIni & "','" & lFechaFin & "',  " & lSucursal & ",  " & lFamilia & "," & lLinea & "," & lSubLinea & ",  " & lPeso & "," & lPeso2 & ",'" & lColor & "','" & lPureza & "') "
            cGroup = " Group by mdsPureza "
            cOrder = " Order by mdsPureza "

            lStrSql = cSelect & lstrVta & lstrExs & cWHERE & cGroup & cOrder
            DevuelveQuery_ConcentradoPureza = lStrSql
        Else
            DevuelveQuery_ConcentradoPureza = ""
        End If

        ''' PARAMETROS PARA DETERMINAR COLUMNAS ( COLORES ) PARA EL REPORTE EN EXCEL
        If Not RSColor.EOF Then
            RSColor.MoveFirst()
            lnumColumnas = RSColor.RecordCount
            lnumIniE = C_INICIAL + 2 '''SIEMPRE ES COLUMNA 5
            lnumFinE = lnumIniE + (lnumColumnas - 1)
            lnumIniV = lnumFinE + 4
            lnumFinV = lnumIniV + (lnumColumnas - 1)
            lnumColE = 4 + 1 + (lnumColumnas * 2)
        Else
        End If
        ''' **************************************************************************

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Encabezado()
        On Error GoTo Err_Renamed
        Dim Columna As Integer

        With objHoja
            .Range("C1").FormulaR1C1 = Trim(gstrCorpoNOMBREEMPRESA)
            .Range("C1:k1").Select()
            .Range("C1:K1").MergeCells = True
            .Range("C1:K1").HorizontalAlignment = Excel.Constants.xlCenter
            .Range("C1:K1").NumberFormat = "@"
            With .Range("C1:K1").Font
                .Bold = True
                .Size = 12
                .Name = "Arial"
            End With
            .Range("C2").FormulaR1C1 = "Ventas y Existencias por Familia"
            .Range("C2:K2").Select()
            .Range("C2:K2").MergeCells = True
            .Range("C2:K2").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C2:K2").Font
                .Bold = False
                .Size = 11
                .Name = "Arial"
            End With
            .Range("C3").FormulaR1C1 = "Desde el " & Format(dtpFechaInicial.Value, "dd/mmm/yyyy") & " Hasta el " & Format(dtpFechaFinal.Value, "dd/mmm/yyyy")
            .Range("C3:K3").Select()
            .Range("C3:K3").MergeCells = True
            .Range("C3:K3").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C3:K3").Font
                .Bold = False
                .Size = 10
                .Name = "Arial"
            End With

            .Range("A4").FormulaR1C1 = "Fecha: " & Format(Today, "dd/mmm/yyyy")
            .Range("A4:B4").Select()
            .Range("A4:B4").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A4:B4").Font
                .Bold = False
                .Size = 9
                .Name = "Arial"
            End With
            .Range("A5").FormulaR1C1 = "Familia: " & lstrFamilia
            .Range("A5:B5").Select()
            .Range("A5:B5").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A5:B5").Font
                .Bold = False
                .Size = 9
                .Name = "Arial"
            End With
            .Range("A6").FormulaR1C1 = "Mensaje: "
            .Range("A6").Select()
            .Range("A6").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A6").Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            If Trim(txtMensaje.Text) <> "" Then
                .Range("B6").FormulaR1C1 = Trim(QuitaEnter(txtMensaje.Text))
                .Range("B6:J6").Select()
                .Range("B6:J6").MergeCells = True
                .Range("B6:J6").HorizontalAlignment = Excel.Constants.xlLeft
                With .Range("B6:J6").Font
                    .Bold = False
                    .Size = 9
                    .Name = "Arial"
                End With
            End If

            Columna = 2
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Código Art"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 9

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Descripción"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 40
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With

            Columna = Columna + 1 '''RENGLON EN BLANCO
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = ""
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 1

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Peso - CT"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 8.29

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Color"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 8.29

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Pureza - Q"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, 6), .Cells._Default(C_ENCABEZADO, 6)).ColumnWidth = 8.29

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Certificado"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 14

            Columna = Columna + 1 '''RENGLON EN BLANCO
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = ""
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 1

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Precio Pub"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 10

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Costo L"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 10
            ColumCtoL = Columna

            Columna = Columna + 1 '''RENGLON EN BLANCO
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = ""
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 1

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "VENTAS"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 9.14
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).MergeCells = True
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "EXISTENCIA"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 9.14
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous

        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Encabezado_Concentrado()
        On Error GoTo Err_Renamed
        Dim I As Short
        Dim Columna As Short

        With objHojaCon

            ''' ENCABEZADO REPORTE **********************************************************************************************************
            '''       C + TotalCols + ETIQS
            Columna = (3 + lnumColumnas + 2) '''NOMBRE DE LA EMPRESA
            Renglon = 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Trim(gstrCorpoNOMBREEMPRESA)
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 12
                .Name = "Arial"
            End With
            Columna = (3 + lnumColumnas + 2) '''NOMBRE DEL REPORTE
            Renglon = Renglon + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Ventas y Existencias por Familia"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = False
                .Size = 11
                .Name = "Arial"
            End With
            Columna = (3 + lnumColumnas + 2) '''TITULO REPORTE
            Renglon = Renglon + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Concentrado por Peso-Color-Pureza"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = False
                .Size = 11
                .Name = "Arial"
            End With
            Columna = (3 + lnumColumnas + 2) '''FECHA
            Renglon = Renglon + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Desde el " & Format(dtpFechaInicial.Value, "dd/mmm/yyyy") & " Hasta el " & Format(dtpFechaFinal.Value, "dd/mmm/yyyy")
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = False
                .Size = 11
                .Name = "Arial"
            End With

            Columna = 3 '''FECHA DE GENERACION
            Renglon = Renglon + 2
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Fecha: " & VB6.Format(Today, "dd/mmm/yyyy")
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 7.57
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = False
                .Size = 9
                .Name = "Arial"
            End With
            Columna = 3 '''FAMILIA
            Renglon = Renglon + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Familia: " & lstrFamilia
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = False
                .Size = 9
                .Name = "Arial"
            End With
            Columna = 3 '''MENSAJE
            Renglon = Renglon + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Mensaje: " & Trim(QuitaEnter(txtMensaje.Text))
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With

            ''' ENCABEZADO X GRUPO  **********************************************************************************************************
            Columna = C_INICIAL
            Renglon = C_ENCABEZADO + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "VENTAS"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + lnumColumnas + 1)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 9

            Columna = lnumIniV - 2
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "EXISTENCIAS"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + lnumColumnas + 1)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 9

        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    '''18NOV2010 - MAVF
    Sub Encabezado_ConcentradoPureza()
        On Error GoTo Err_Renamed
        Dim I As Short
        Dim Columna As Short
        Dim ColVtas As Short

        With objHojaCon

            ''' ENCABEZADO VENTAS Y EXISTENCIAS - CONCENTRADO PUREZA ***********************************************************************************
            Columna = C_INICIAL

            Renglon = Renglon + 3 '''ENCABEZADO DEL GRUPO
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "CONCENTRADO POR PUREZA"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With

            Renglon = Renglon + 2
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "VENTAS"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + lnumColumnas + 1)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With

            Columna = lnumIniV - 2
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "EXISTENCIAS"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + lnumColumnas + 1)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With

            Columna = lnumIniE
            Renglon = Renglon + 2
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).MergeCells = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1)))._Default = "COLOR"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With

            Columna = lnumIniV
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).MergeCells = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1)))._Default = "COLOR"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With

            ''' *************************************************************************************************************
            ''' COLUMNAS DE ENCABEZADOS Y COLORES X GRUPO - PUREZA
            Columna = lnumIniE - 2
            Renglon = Renglon + 1

            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "PUREZA"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 8

            Columna = lnumIniV - 2
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 7

            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "PUREZA"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 8

            Columna = lnumIniE
            ColVtas = lnumIniV
            RenglonPureza = Renglon
            RSColor.MoveFirst()

            For I = 1 To lnumColumnas
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Select()
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1)))._Default = Trim(RSColor.Fields("Color").Value)
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).ColumnWidth = 5

                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Select()
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1)))._Default = Trim(RSColor.Fields("Color").Value)
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).ColumnWidth = 5
                RSColor.MoveNext()

            Next I

            ''' *****************************************************************************************

            Columna = ColVtas - 3 '''RENGLON EN BLANCO
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = ""
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 1

        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Encabezado_ConcentradoGpo()
        On Error GoTo Err_Renamed
        Dim I As Short
        Dim Columna As Short
        Dim ColVtas As Short

        With objHojaCon

            Columna = lnumIniE
            Renglon = Renglon + 2
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).MergeCells = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1)))._Default = "COLOR"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With

            Columna = lnumIniV
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).MergeCells = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1)))._Default = "COLOR"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + (lnumColumnas - 1))).Font
                .Bold = True
                .Size = 10
                .Name = "Arial"
            End With

            ''' COLUMNAS DE ENCABEZADOS Y COLORES
            Columna = lnumIniE - 2
            Renglon = Renglon + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "CT"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 7

            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "PUREZA"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 8

            Columna = lnumIniV - 2
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "CT"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 7

            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "PUREZA"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 8

            Columna = lnumIniE
            ColVtas = lnumIniV
            RSColor.MoveFirst()

            For I = 1 To lnumColumnas
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Select()
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1)))._Default = Trim(RSColor.Fields("Color").Value)
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
                .Range(.Cells._Default(Renglon, Columna + (I - 1)), .Cells._Default(Renglon, Columna + (I - 1))).ColumnWidth = 5

                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Select()
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1)))._Default = Trim(RSColor.Fields("Color").Value)
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
                .Range(.Cells._Default(Renglon, ColVtas + (I - 1)), .Cells._Default(Renglon, ColVtas + (I - 1))).ColumnWidth = 5
                RSColor.MoveNext()

            Next I

            ''' *****************************************************************************************

            Columna = ColVtas - 3 '''RENGLON EN BLANCO
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = ""
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 1

        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Function ArchivoAbierto() As Boolean
        On Error GoTo Err_Renamed
        Dim Archivo As String
        If Dir(gstrCorpoDriveLocal & "\Sistema\", FileAttribute.Directory + FileAttribute.Hidden) = "" Then
            MsgBox("No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            ArchivoAbierto = True
            Exit Function
        End If
        Archivo = "VE" & CStr(Format(Month(Today), "00")) & CStr(Format((Today), "00")) & (CStr(Format(Year(Today), "00"))) & ".xls"
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\", FileAttribute.Directory) = "" Then
            MkDir(gstrCorpoDriveLocal & "\Sistema\Informes\")
        End If
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo, FileAttribute.Archive) <> "" Then
            Kill(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo)
        End If
        ObjExcel = CreateObject("Excel.Application")
        objLibro = ObjExcel.Workbooks.Add
        objHoja = objLibro.ActiveSheet
        ObjExcel.Visible = False
        objLibro.Sheets(1).Select()
        objHoja = objLibro.ActiveSheet
        objLibro.ActiveSheet.Name = "Vtas. y Exist. por Prov."
        objLibro.SaveAs(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)
        CierraInstanciasdeExcel(1)
        ArchivoAbierto = False
Err_Renamed:
        If Err.Number = 70 Then
            MsgBox("No se puede generar un nuevo archivo hasta que el anterior este cerrado.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            CierraInstanciasdeExcel(2)
            ArchivoAbierto = True
        ElseIf Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            ArchivoAbierto = True
        End If
    End Function

    Function EnviaExcel() As Boolean
        On Error GoTo Err_Renamed
        Dim Archivo As String
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        System.Windows.Forms.Application.DoEvents()
        If Dir(gstrCorpoDriveLocal & "\Sistema\", FileAttribute.Directory + FileAttribute.Hidden) = "" Then
            MsgBox("No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Function
        End If
        Archivo = "VF" & CStr(Format(Month(Today), "00")) & CStr(Format((Today), "00")) & (CStr(Format(Year(Today), "00"))) & ".xls"
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\", FileAttribute.Directory) = "" Then
            MkDir(gstrCorpoDriveLocal & "\Sistema\Informes\")
        End If
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo, FileAttribute.Archive) <> "" Then
            Kill(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo)
        End If

        ObjExcel = CreateObject("Excel.Application")
        objLibro = ObjExcel.Workbooks.Add
        objHoja = objLibro.ActiveSheet
        ObjExcel.Visible = False
        objLibro.Sheets(1).Select()
        objHoja = objLibro.ActiveSheet
        objLibro.ActiveSheet.Name = "Vtas. y Exist. por Familia"
        Encabezado()
        LlenaDatos()

        objLibro.SaveAs(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)

        lArchivo01 = Archivo
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        System.Windows.Forms.Application.DoEvents()
        EnviaExcel = True

Err_Renamed:
        If Err.Number = 70 Then
            MsgBox("No se puede generar un nuevo archivo hasta que el anterior este cerrado.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            CierraInstanciasdeExcel(2)
            EnviaExcel = False
        ElseIf Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            EnviaExcel = False
        End If
    End Function

    Function EnviaExcel_Concentrado() As Boolean
        On Error GoTo Err_Renamed
        Dim Archivo As String
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        System.Windows.Forms.Application.DoEvents()
        If Dir(gstrCorpoDriveLocal & "\Sistema\", FileAttribute.Directory + FileAttribute.Hidden) = "" Then
            MsgBox("No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Function
        End If
        Archivo = "VFC" & CStr(Format(Month(Today), "00")) & CStr(Format((Today), "00")) & (CStr(Format(Year(Today), "00"))) & ".xls"
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\", FileAttribute.Directory) = "" Then
            MkDir(gstrCorpoDriveLocal & "\Sistema\Informes\")
        End If
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo, FileAttribute.Archive) <> "" Then
            Kill(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo)
        End If

        ObjExcelCon = CreateObject("Excel.Application")
        objLibroCon = ObjExcelCon.Workbooks.Add
        objHojaCon = objLibroCon.ActiveSheet
        ObjExcelCon.Visible = False
        objLibroCon.Sheets(1).Select()
        objHojaCon = objLibroCon.ActiveSheet
        objLibroCon.ActiveSheet.Name = "VyE x Familia - Concentrado"
        Encabezado_Concentrado()
        LlenaDatos_Concentrado()

        '''18NOV2010 - MAVF - CONCENTRADO DE PUREZA
        If RsConP.RecordCount > 0 Then
            Encabezado_ConcentradoPureza()
            LlenaDatos_ConcentradoPureza()
        End If

        objLibroCon.SaveAs(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)

        lArchivo02 = Archivo
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        System.Windows.Forms.Application.DoEvents()
        EnviaExcel_Concentrado = True

Err_Renamed:
        If Err.Number = 70 Then
            MsgBox("No se puede generar un nuevo archivo hasta que el anterior este cerrado.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            CierraInstanciasdeExcel_Concentrado(2)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            EnviaExcel_Concentrado = False
        ElseIf Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel_Concentrado(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            EnviaExcel_Concentrado = False
        End If
    End Function

    Sub CierraInstanciasdeExcel(ByRef Tipo As Short)
        If Tipo = 1 Then
            objLibro.Close()
            ObjExcel.Quit()
        End If

        If ObjExcel Is Nothing Then ObjExcel = Nothing
        If objLibro Is Nothing Then objLibro = Nothing
        If objHoja Is Nothing Then objHoja = Nothing
    End Sub

    Sub CierraInstanciasdeExcel_Concentrado(ByRef Tipo As Short)
        If Tipo = 1 Then
            objLibroCon.Close()
            ObjExcelCon.Quit()
        End If

        If ObjExcelCon Is Nothing Then ObjExcelCon = Nothing
        If objLibroCon Is Nothing Then objLibroCon = Nothing
        If objHojaCon Is Nothing Then objHojaCon = Nothing
    End Sub

    Sub LlenaDatos()
        On Error GoTo Err_Renamed
        Dim Renglon As Integer
        Dim Columna As Integer
        Dim RenRecorridos As Integer
        Dim CodProveedor As Short
        Dim I As Short
        Dim Rango As String
        Dim Formula As String
        Dim Totales As String
        Dim Cantidad As String
        Dim blnCtoL As Boolean

        Renglon = 9
        CodProveedor = 0
        blnCtoL = True

        With objHoja
            RsGral.MoveFirst()
            Do While Not RsGral.EOF
                Renglon = Renglon + 1

                Rango = "B" & Renglon '''CODIGO DEL ARTICULO
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("CodArticulo").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "C" & Renglon '''DESCRIPCION DEL ARTICULO
                .Range(Rango).Select()
                .Range(Rango)._Default = Trim(RsGral.Fields("DescArticulo").Value)
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlLeft
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "E" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("mdsPeso").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
                .Range(Rango).NumberFormat = "###,##0.00"
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "F" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("mdsColor").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlCenter
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "G" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("mdsPureza").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlLeft
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "H" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("mdsCertificado").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlLeft
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "J" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("PrecioPCat").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
                .Range(Rango).NumberFormat = "###,##0.00"
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "K" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("CostoCat").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
                .Range(Rango).NumberFormat = "###,##0.00"
                .Range(Rango).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = 13
                For I = 1 To 2
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields(I + 11).Value
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                    If RenRecorridos = 0 Then
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    End If
                    If (Columna - 1) = ColumSepar Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                    If (Columna - 1) = ColumCtoL Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous

                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    Columna = Columna + 1
                Next
                RenRecorridos = RenRecorridos + 1

                RsGral.MoveNext()
            Loop

            RsGral.MoveFirst()
            Renglon = Renglon + 1
            Rango = "J" & Renglon
            .Range(Rango).Select()
            .Range(Rango)._Default = "GRAN TOTAL"
            .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
            With .Range(Rango).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With

            RenFinal = Renglon
            Columna = 13
            For I = 1 To 2
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).FormulaR1C1 = "=SUM(R[-" & RenRecorridos & "]C:R[-1]C)"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"

                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
            Next

            .Range("A1").Select()
            .Range("A1").Activate()
        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub LlenaDatos_Concentrado()
        On Error GoTo Err_Renamed
        Dim Columna As Short
        Dim ColVtas As Short
        Dim RenRecorridos As Short
        Dim I As Short
        Dim Rango As String
        Dim Formula As String
        Dim Totales As String
        Dim Cantidad As String
        Dim totalRenxGpo As Short
        Dim lmdsPeso As Decimal
        Dim RenGrupo As Short
        Dim valorC As Short '''Valores de columnas para Color-Pureza en secciones Exist-Vtas
        Dim valorD As Short
        Dim valorM As Short
        Dim valorN As Short
        Dim pzasVta As Decimal '''Valores de Ventas y Existencias X Gpo - 18NOV2010 - MAVF
        Dim pzasExist As Decimal '''Valores de Ventas y Existencias X Gpo - 18NOV2010 - MAVF
        Dim pzasVtaT As Decimal '''Valores Totales de Ventas y Existencias - 18NOV2010 - MAVF
        Dim pzasExistT As Decimal '''Valores Totales de Ventas y Existencias - 18NOV2010 - MAVF

        With objHojaCon
            RsGral.MoveFirst()
            lmdsPeso = RsGral.Fields("mdsPeso").Value '''Tomar valor del primer registro
            pzasExistT = 0 '''18NOV2010 - MAVF
            pzasVtaT = 0 '''18NOV2010 - MAVF

            Do While Not RsGral.EOF

                Encabezado_ConcentradoGpo()

                pzasExist = 0 '''18NOV2010 - MAVF
                pzasVta = 0 '''18NOV2010 - MAVF
                Do While (lmdsPeso = RsGral.Fields("mdsPeso").Value) '''Cambio de grupo - PesoCT
                    Renglon = Renglon + 1

                    If totalRenxGpo = 0 Then
                        valorC = Asc("C") '''Columna fija en la que siempre comienza
                        valorD = valorC + 1
                        valorM = valorC + lnumColumnas + 1 + 2
                        valorN = valorM + 1

                        Rango = Chr(valorC) & Renglon '''Peso - CT
                        .Range(Rango).Select()
                        .Range(Rango)._Default = RsGral.Fields("mdsPeso").Value
                        .Range(Rango).HorizontalAlignment = Excel.Constants.xlCenter
                        .Range(Rango).NumberFormat = "###,##0.00"
                        With .Range(Rango).Font
                            .Size = 9
                            .Name = "Arial"
                        End With
                    End If

                    Rango = Chr(valorD) & Renglon '''Pureza
                    .Range(Rango).Select()
                    .Range(Rango)._Default = RsGral.Fields("mdsPureza").Value
                    .Range(Rango).HorizontalAlignment = Excel.Constants.xlCenter
                    With .Range(Rango).Font
                        .Size = 9
                        .Name = "Arial"
                    End With

                    If totalRenxGpo = 0 Then
                        Rango = Chr(valorM) & Renglon '''Peso - CT
                        .Range(Rango).Select()
                        .Range(Rango)._Default = RsGral.Fields("mdsPeso").Value
                        .Range(Rango).HorizontalAlignment = Excel.Constants.xlCenter
                        .Range(Rango).NumberFormat = "###,##0.00"
                        With .Range(Rango).Font
                            .Size = 9
                            .Name = "Arial"
                        End With
                    End If

                    Rango = Chr(valorN) & Renglon '''Pureza
                    .Range(Rango).Select()
                    .Range(Rango)._Default = RsGral.Fields("mdsPureza").Value
                    .Range(Rango).HorizontalAlignment = Excel.Constants.xlCenter
                    With .Range(Rango).Font
                        .Size = 9
                        .Name = "Arial"
                    End With

                    totalRenxGpo = totalRenxGpo + 1

                    Columna = lnumIniE
                    ColVtas = lnumIniV
                    For I = 1 To lnumColumnas
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields(I + 1).Value
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                        With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                            .Size = 9
                            .Name = "Arial"
                        End With
                        pzasVta = pzasVta + RsGral.Fields(I + 1).Value '''18NOV2010 - MAVF

                        .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Select()
                        .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas))._Default = RsGral.Fields((I + 1) + lnumColumnas).Value
                        .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).NumberFormat = "###,##0"
                        With .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Font
                            .Size = 9
                            .Name = "Arial"
                        End With
                        pzasExist = pzasExist + RsGral.Fields(I + 1 + lnumColumnas).Value '''18NOV2010 - MAVF

                        Columna = Columna + 1
                        ColVtas = ColVtas + 1
                    Next
                    RenRecorridos = RenRecorridos + 1
                    RsGral.MoveNext()

                    If RsGral.EOF Then Exit Do

                Loop  '''Ciclo por grupo - PesoCT

                If Not RsGral.EOF Then lmdsPeso = RsGral.Fields("mdsPeso").Value
                totalRenxGpo = 0

                'PONER RAYITA POR GRUPO
                RenGrupo = Renglon
                Columna = lnumIniE
                ColVtas = lnumIniV
                For I = 1 To lnumColumnas
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous

                    .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Select()
                    .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous

                    Columna = Columna + 1
                    ColVtas = ColVtas + 1
                Next

                '''18NOV2010 - MAVF
                Renglon = Renglon + 1
                .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).Select() '''Total VENTAS x Gpo
                .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).NumberFormat = "###,##0"
                .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).HorizontalAlignment = Excel.Constants.xlCenter
                With .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).Font
                    .Size = 9
                    .Name = "Arial"
                End With
                .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2))._Default = pzasVta

                .Range(.Cells._Default(Renglon, (lnumIniV - 2)), .Cells._Default(Renglon, lnumIniV - 2)).Select() '''Total EXISTENCIAS x Gpo
                .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2)).NumberFormat = "###,##0"
                .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2)).HorizontalAlignment = Excel.Constants.xlCenter
                With .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2)).Font
                    .Size = 9
                    .Name = "Arial"
                End With
                .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2))._Default = pzasExist

                pzasVtaT = pzasVtaT + pzasVta '''Acumulado de Ventas
                pzasExistT = pzasExistT + pzasExist '''Acumulado de Existencias
                ''' ***********************************************************************************************************

            Loop  '''Ciclo General

            RsGral.MoveFirst()
            Renglon = Renglon + 3

            '''LEYENDAS - 18NOV2010 - MAVF
            .Range(.Cells._Default(Renglon, lnumIniE - 1), .Cells._Default(Renglon, lnumIniE - 1)).Select() '''Leyenda Total VENTAS
            .Range(.Cells._Default(Renglon, lnumIniE - 1), .Cells._Default(Renglon, lnumIniE - 1)).HorizontalAlignment = Excel.Constants.xlRight
            With .Range(.Cells._Default(Renglon, lnumIniE - 1), .Cells._Default(Renglon, lnumIniE - 1)).Font
                .Size = 9
                .Bold = True
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, lnumIniE - 1), .Cells._Default(Renglon, lnumIniE - 1))._Default = "TOTAL"

            .Range(.Cells._Default(Renglon, (lnumIniV - 1)), .Cells._Default(Renglon, lnumIniV - 1)).Select() '''Leyenda Total EXISTENCIAS
            .Range(.Cells._Default(Renglon, lnumIniV - 1), .Cells._Default(Renglon, lnumIniV - 1)).HorizontalAlignment = Excel.Constants.xlRight
            With .Range(.Cells._Default(Renglon, lnumIniV - 1), .Cells._Default(Renglon, lnumIniV - 1)).Font
                .Size = 9
                .Bold = True
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, lnumIniV - 1), .Cells._Default(Renglon, lnumIniV - 1))._Default = "TOTAL"
            ''' ********************************************************************************

            'Almacenamos el renglon final para saber hasta donde vamos a borrar cuando no mostremos el Costo L
            RenFinal = Renglon
            Columna = lnumIniE
            ColVtas = lnumIniV
            For I = 1 To lnumColumnas
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).FormulaR1C1 = "=SUM(R[-" & (RenFinal - 14) & "]C:R[-1]C)"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"

                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 10
                    .Name = "Arial"
                End With

                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Select()
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).FormulaR1C1 = "=SUM(R[-" & (RenFinal - 14) & "]C:R[-1]C)"
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).NumberFormat = "###,##0"

                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Font
                    .Bold = True
                    .Size = 10
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                ColVtas = ColVtas + 1
            Next

            '''PIEZAS TOTALES - 18NOV2010
            .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).Select() '''VENTAS
            .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).NumberFormat = "###,##0"
            .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).Font
                .Size = 9
                .Bold = True
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2))._Default = pzasVtaT

            .Range(.Cells._Default(Renglon, (lnumIniV - 2)), .Cells._Default(Renglon, lnumIniV - 2)).Select() '''EXISTENCIAS
            .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2)).NumberFormat = "###,##0"
            .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2)).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2)).Font
                .Size = 9
                .Bold = True
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2))._Default = pzasExistT
            ''' ************************************************************************************

            .Range("A1").Select()
            .Range("A1").Activate()
        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub LlenaDatos_ConcentradoPureza()
        On Error GoTo Err_Renamed
        Dim Columna As Short
        Dim ColVtas As Short
        Dim RenRecorridos As Short
        Dim I As Short
        Dim Rango As String
        Dim Formula As String
        Dim Totales As String
        Dim Cantidad As String
        Dim totalRenxGpo As Short
        Dim lmdsPureza As String
        Dim RenGrupo As Short
        Dim valorC As Short '''Valores de columnas para Color-Pureza en secciones Exist-Vtas
        Dim valorD As Short
        Dim valorM As Short
        Dim valorN As Short
        Dim pzasVta As Decimal '''Valores de Ventas y Existencias X Gpo
        Dim pzasExist As Decimal '''Valores de Ventas y Existencias X Gpo

        With objHojaCon
            RsConP.MoveFirst()
            lmdsPureza = RsConP.Fields("mdsPureza").Value '''Tomar valor del primer registro
            pzasExist = 0
            pzasVta = 0

            Do While Not RsConP.EOF

                Do While (lmdsPureza = RsConP.Fields("mdsPureza").Value) '''Cambio de grupo - PesoCT
                    Renglon = Renglon + 1

                    If totalRenxGpo = 0 Then
                        valorC = Asc("C") '''Columna fija en la que siempre comienza
                        valorD = valorC + 1
                        valorM = valorC + lnumColumnas + 1 + 2
                        valorN = valorM + 1
                    End If

                    Rango = Chr(valorD) & Renglon '''Pureza
                    .Range(Rango).Select()
                    .Range(Rango)._Default = RsConP.Fields("mdsPureza").Value
                    .Range(Rango).HorizontalAlignment = Excel.Constants.xlCenter
                    With .Range(Rango).Font
                        .Size = 9
                        .Name = "Arial"
                    End With

                    Rango = Chr(valorN) & Renglon '''Pureza
                    .Range(Rango).Select()
                    .Range(Rango)._Default = RsConP.Fields("mdsPureza").Value
                    .Range(Rango).HorizontalAlignment = Excel.Constants.xlCenter
                    With .Range(Rango).Font
                        .Size = 9
                        .Name = "Arial"
                    End With

                    totalRenxGpo = totalRenxGpo + 1

                    Columna = lnumIniE
                    ColVtas = lnumIniV
                    For I = 1 To lnumColumnas
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsConP.Fields(I).Value '''(I+1)
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                        With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                            .Size = 9
                            .Name = "Arial"
                        End With
                        pzasVta = pzasVta + RsConP.Fields(I).Value

                        .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Select()
                        .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas))._Default = RsConP.Fields((I) + lnumColumnas).Value ''''''(I+1) + lnumcolumnas
                        .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).NumberFormat = "###,##0"
                        With .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Font
                            .Size = 9
                            .Name = "Arial"
                        End With
                        pzasExist = pzasExist + RsConP.Fields(I + lnumColumnas).Value

                        Columna = Columna + 1
                        ColVtas = ColVtas + 1
                    Next
                    RenRecorridos = RenRecorridos + 1
                    RsConP.MoveNext()

                    If RsConP.EOF Then Exit Do

                Loop  '''Ciclo por grupo - PesoCT

                If Not RsConP.EOF Then lmdsPureza = RsConP.Fields("mdsPureza").Value
                totalRenxGpo = 0

            Loop  '''Ciclo General

            'PONER RAYITA POR GRUPO
            RenGrupo = Renglon
            Columna = lnumIniE
            ColVtas = lnumIniV
            For I = 1 To lnumColumnas
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous

                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Select()
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous

                Columna = Columna + 1
                ColVtas = ColVtas + 1
            Next

            RsConP.MoveFirst()
            Renglon = Renglon + 3

            '''LEYENDAS
            ''' ***********************************************************************************************************
            .Range(.Cells._Default(Renglon, lnumIniE - 1), .Cells._Default(Renglon, lnumIniE - 1)).Select() '''Leyenda Total VENTAS
            .Range(.Cells._Default(Renglon, lnumIniE - 1), .Cells._Default(Renglon, lnumIniE - 1)).HorizontalAlignment = Excel.Constants.xlRight
            With .Range(.Cells._Default(Renglon, lnumIniE - 1), .Cells._Default(Renglon, lnumIniE - 1)).Font
                .Size = 9
                .Bold = True
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, lnumIniE - 1), .Cells._Default(Renglon, lnumIniE - 1))._Default = "TOTAL"

            .Range(.Cells._Default(Renglon, (lnumIniV - 1)), .Cells._Default(Renglon, lnumIniV - 1)).Select() '''Leyenda Total EXISTENCIAS
            .Range(.Cells._Default(Renglon, lnumIniV - 1), .Cells._Default(Renglon, lnumIniV - 1)).HorizontalAlignment = Excel.Constants.xlRight
            With .Range(.Cells._Default(Renglon, lnumIniV - 1), .Cells._Default(Renglon, lnumIniV - 1)).Font
                .Size = 9
                .Bold = True
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, lnumIniV - 1), .Cells._Default(Renglon, lnumIniV - 1))._Default = "TOTAL"

            '''VALORES
            .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).Select() '''Total VENTAS x Gpo
            .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).NumberFormat = "###,##0"
            .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2)).Font
                .Size = 9
                .Bold = True
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, lnumIniE - 2), .Cells._Default(Renglon, lnumIniE - 2))._Default = pzasVta

            .Range(.Cells._Default(Renglon, (lnumIniV - 2)), .Cells._Default(Renglon, lnumIniV - 2)).Select() '''Total EXISTENCIAS x Gpo
            .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2)).NumberFormat = "###,##0"
            .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2)).HorizontalAlignment = Excel.Constants.xlCenter
            With .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2)).Font
                .Size = 9
                .Bold = True
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, lnumIniV - 2), .Cells._Default(Renglon, lnumIniV - 2))._Default = pzasExist
            ''' ***********************************************************************************************************

            RenFinal = Renglon
            Columna = lnumIniE
            ColVtas = lnumIniV
            For I = 1 To lnumColumnas
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).FormulaR1C1 = "=SUM(R[-" & (Renglon - RenglonPureza) & "]C:R[-1]C)"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"

                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 10
                    .Name = "Arial"
                End With

                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Select()
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).FormulaR1C1 = "=SUM(R[-" & (Renglon - RenglonPureza) & "]C:R[-1]C)"
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).NumberFormat = "###,##0"

                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, ColVtas), .Cells._Default(Renglon, ColVtas)).Font
                    .Bold = True
                    .Size = 10
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                ColVtas = ColVtas + 1
            Next

            .Range("A1").Select()
            .Range("A1").Activate()
        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Public Sub Imprime()
        On Error GoTo Merr
        Dim lStrSql As String
        Dim lRep01 As Boolean
        Dim lRep02 As Boolean
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Not ValidaDatos() Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Sub
        End If

        ''' Asignar valor a Variables ****************************************
        lSucursal = IIf(chkTodasSuc.CheckState = System.Windows.Forms.CheckState.Checked, 0, mintCodSucursal)
        lFamilia = IIf(Trim(dbcJFamilia.Text) = C_TODAS, 0, mintJFamilia)
        lLinea = IIf(Trim(dbcJLinea.Text) = C_TODAS, 0, mintJLinea)
        lSubLinea = IIf(Trim(dbcJSubLinea.Text) = C_TODAS, 0, mintJSubLinea)
        lPeso = IIf(chkmdsPeso.CheckState = System.Windows.Forms.CheckState.Checked, txtMDSPeso.Text, 0)
        lPeso2 = IIf(chkmdsPeso.CheckState = System.Windows.Forms.CheckState.Checked, txtMDSPeso2.Text, 0)
        lColor = IIf(chkmdsColor.CheckState = System.Windows.Forms.CheckState.Checked, IIf(Trim(txtMDSColor.Text) = C_OPCION, "0", Trim(txtMDSColor.Text)), "0")
        lPureza = IIf(chkmdsPureza.CheckState = System.Windows.Forms.CheckState.Checked, IIf(Trim(txtMDSPureza.Text) = C_OPCION, "0", Trim(txtMDSPureza.Text)), "0")
        lFechaIni = Format(dtpFechaInicial.Value, "yyyymmdd")
        lFechaFin = Format(dtpFechaFinal.Value, "yyyymmdd")
        lstrFamilia = Trim(dbcJFamilia.Text) & IIf(Trim(dbcJLinea.Text) = C_TODAS, "", " - " & Trim(dbcJLinea.Text)) & IIf(Trim(dbcJSubLinea.Text) = C_TODAS, "", " - " & Trim(dbcJSubLinea.Text))

        ''' Genera_Reporte01 **************************************************************************
        lStrSql = DevuelveQuery()
        ModEstandar.BorraCmd()
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, lStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            If EnviaExcel() Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                lRep01 = True
            Else
                lRep01 = False
            End If
        Else
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            MsgBox("No existe información por mostrar en este periodo de fechas" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
        Cmd.CommandTimeout = 90
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText

        If Not lRep01 Then Exit Sub '''SI EL PRIMER REPORTE FALLA, SE SALE DEL PROCESO

        ''' Genera_Reporte02 ES OPCIONAL **************************************************************
        If chkConcentrado.CheckState = System.Windows.Forms.CheckState.Checked Then
            lStrSql = DevuelveQuery_Concentrado()
            ModEstandar.BorraCmd()
            Cmd.CommandTimeout = 300
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, lStrSql))
            RsGral = Cmd.Execute

            ''' 18NOV2010 - MAVF - CONCENTRADO DE PUREZA
            lStrSql = DevuelveQuery_ConcentradoPureza()
            ModEstandar.BorraCmd()
            Cmd.CommandTimeout = 300
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, lStrSql))
            RsConP = Cmd.Execute

            If RsGral.RecordCount > 0 Then
                If EnviaExcel_Concentrado() Then
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    lRep02 = True
                Else
                    lRep02 = False
                End If
            Else
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                MsgBox("No existe información por mostrar en este periodo de fechas PARA EL CONCENTRADO " & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            End If
            Cmd.CommandTimeout = 90
        End If

        '''MOSTRAR ARCHIVOS EXCEL *********************************************
        If chkConcentrado.CheckState = System.Windows.Forms.CheckState.Checked Then
            ''' SE INDICO LA GENERACION DE AMBOS REPORTES - FAMILIAS/CONCENTRADO
            If lRep01 And lRep02 Then
                Select Case MsgBox("Se han creado los archivos " & lArchivo01 & " y " & lArchivo02 & " ¿Desea abrirlos?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        ObjExcel.Visible = True
                        ObjExcel = Nothing
                        objLibro = Nothing
                        objHoja = Nothing

                        ObjExcelCon.Visible = True
                        ObjExcelCon = Nothing
                        objLibroCon = Nothing
                        objHojaCon = Nothing

                    Case MsgBoxResult.No Or MsgBoxResult.Cancel
                        CierraInstanciasdeExcel(1)
                        CierraInstanciasdeExcel_Concentrado(1)
                End Select

            ElseIf lRep01 And Not lRep02 Then

                '''SOLO SE GENERO CORRECTAMENTE EL REPORTE DE FAMILIAS
                Select Case MsgBox("El archivo Concentrado no fue creado" & vbNewLine & "Se ha creado el archivo " & lArchivo01 & " ¿Desea abrirlo?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        ObjExcel.Visible = True
                        ObjExcel = Nothing
                        objLibro = Nothing
                        objHoja = Nothing

                    Case MsgBoxResult.No Or MsgBoxResult.Cancel
                        CierraInstanciasdeExcel(1)
                End Select

            ElseIf Not lRep01 And lRep02 Then

                '''SOLO SE GENERO CORRECTAMENTE EL REPORTE CONCENTRADO
                Select Case MsgBox("El archivo por Familia no fue creado" & vbNewLine & "Se ha creado el archivo " & lArchivo02 & " ¿Desea abrirlo?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        ObjExcelCon.Visible = True
                        ObjExcelCon = Nothing
                        objLibroCon = Nothing
                        objHojaCon = Nothing

                    Case MsgBoxResult.No Or MsgBoxResult.Cancel
                        CierraInstanciasdeExcel_Concentrado(1)
                End Select

            End If
        Else
            '''SOLO SE GENERA EL REPORTE X FAMILIA
            If lRep01 Then

                Select Case MsgBox("Se ha creado el archivo " & lArchivo01 & " ¿Desea abrirlo?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        ObjExcel.Visible = True
                        ObjExcel = Nothing
                        objLibro = Nothing
                        objHoja = Nothing

                    Case MsgBoxResult.No Or MsgBoxResult.Cancel
                        CierraInstanciasdeExcel(1)
                End Select

            End If
        End If

Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function ValidaDatos() As Boolean

        If chkTodasSuc.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            If dbcSucursal.Text = "" Then
                MsgBox("Debe seleccionar una sucursal", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                dbcSucursal.Focus()
                Exit Function
            End If
        End If

        If Trim(dbcJFamilia.Text) = "" And Trim(dbcJLinea.Text) = "" And Trim(dbcJSubLinea.Text) = "" Then
            MsgBox("Debe indicar al menos 1 parámetro para generar información (Familia-Línea-SubLínea)", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            ValidaDatos = False
            dbcJFamilia.Focus()
            Exit Function
        Else
            If Trim(dbcJFamilia.Text) = "" Then
                MsgBox("Debe indicar la familia para generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                dbcJFamilia.Focus()
                Exit Function
            End If
            If Trim(dbcJLinea.Text) = "" Then
                MsgBox("Debe indicar una línea o seleccionar [TODAS] para generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                dbcJLinea.Focus()
                Exit Function
            End If
            If Trim(dbcJSubLinea.Text) = "" Then
                MsgBox("Debe indicar una sublínea o seleccionar [TODAS] para el reporte ", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                dbcJSubLinea.Focus()
                Exit Function
            End If
        End If

        If (chkmdsPeso.CheckState = System.Windows.Forms.CheckState.Unchecked And chkmdsColor.CheckState = System.Windows.Forms.CheckState.Unchecked And chkmdsPureza.CheckState = System.Windows.Forms.CheckState.Unchecked) Then
            MsgBox("Debe indicar al menos 1 característica del dimante (Peso-Color-Pureza)", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            ValidaDatos = False
            chkmdsPeso.Focus()
            Exit Function
        End If

        If chkmdsPeso.CheckState = System.Windows.Forms.CheckState.Checked Then
            If Trim(txtMDSPeso.Text) = "" Or Trim(txtMDSPeso2.Text) = "" Then
                MsgBox("Debe indicar el rango del peso del diamante ... (Peso-CT)", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                txtMDSPeso.Focus()
                Exit Function
            End If
            If CDec(ModEstandar.Numerico((txtMDSPeso.Text))) > CDec(ModEstandar.Numerico((txtMDSPeso2.Text))) Then
                MsgBox("El peso inicial no debe ser mayor al final, verifique por favor ... (Peso-CT)", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                txtMDSPeso.Focus()
                Exit Function
            End If
        End If
        If chkmdsColor.CheckState = System.Windows.Forms.CheckState.Checked Then
            If Trim(txtMDSColor.Text) = "" Then
                MsgBox("Debe indicar el color del diamante ... (Color)", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                txtMDSColor.Focus()
                Exit Function
            End If
        End If
        If chkmdsPureza.CheckState = System.Windows.Forms.CheckState.Checked Then
            If Trim(txtMDSPureza.Text) = "" Then
                MsgBox("Debe indicar la pureza del diamante ... (Pureza-Q)", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                txtMDSPureza.Focus()
                Exit Function
            End If
        End If

        If mblnTecleoFechaI Then
            Do While (msglTiempoCambioI) <= CDec(2.1)
            Loop
            mblnTecleoFechaI = False
        End If
        If mblnTecleoFechaF Then
            Do While (msglTiempoCambioF) <= CDec(2.1)
            Loop
            mblnTecleoFechaF = False
        End If
        System.Windows.Forms.Application.DoEvents()
        Select Case True
            Case dtpFechaInicial.Value > Me.dtpFechaFinal.Value
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Final", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                dtpFechaInicial.Focus()
                Exit Function
        End Select

        ValidaDatos = True
    End Function

    Private Sub chkmdsColor_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkmdsColor.CheckStateChanged
        If chkmdsColor.CheckState = System.Windows.Forms.CheckState.Checked Then
            txtMDSColor.Enabled = True
        Else
            txtMDSColor.Text = ""
            txtMDSColor.Enabled = False
        End If
    End Sub

    Private Sub chkmdsPeso_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkmdsPeso.CheckStateChanged
        If chkmdsPeso.CheckState = System.Windows.Forms.CheckState.Checked Then
            txtMDSPeso.Enabled = True
            txtMDSPeso2.Enabled = True
        Else
            txtMDSPeso.Text = ""
            txtMDSPeso2.Text = ""
            txtMDSPeso.Enabled = False
            txtMDSPeso2.Enabled = False
        End If
    End Sub

    Private Sub chkmdsPureza_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkmdsPureza.CheckStateChanged
        If chkmdsPureza.CheckState = System.Windows.Forms.CheckState.Checked Then
            txtMDSPureza.Enabled = True
        Else
            txtMDSPureza.Text = ""
            txtMDSPureza.Enabled = False
        End If
    End Sub

    Private Sub chkTodasSuc_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodasSuc.CheckStateChanged
        Select Case chkTodasSuc.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                dbcSucursal.Text = C_TODAS
                dbcSucursal.Tag = ""
                mintCodSucursal = 0
                dbcSucursal.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                dbcSucursal.Text = ""
                dbcSucursal.Tag = ""
                mintCodSucursal = 0
                dbcSucursal.Enabled = True
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub dbcJFAmilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJFamilia.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        lStrSql = "SELECT codFamilia, LTrim(RTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia LIKE '" & Trim(Me.dbcJFamilia.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcJFamilia))
        If Trim(Me.dbcJFamilia.Text) = "" Then
            mintJFamilia = 0
            mblnFueraChange = True
            Me.dbcJFamilia.Enabled = True
            mintJLinea = 0
            Me.dbcJLinea.Text = C_TODAS
            Me.dbcJLinea.Enabled = False
            mintJSubLinea = 0
            Me.dbcJSubLinea.Text = C_TODAS
            Me.dbcJSubLinea.Enabled = False
            mblnFueraChange = False
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcjFAmilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.Enter
        Pon_Tool()
        gStrSql = "SELECT codFamilia, LTrim(RTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcJFamilia))
    End Sub

    Private Sub dbcJFAmilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJFamilia.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            eventSender.KeyCode = 0
            If chkTodasSuc.CheckState = System.Windows.Forms.CheckState.Checked Then chkTodasSuc.Focus() Else dbcSucursal.Focus()
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcJFamilia_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcJFamilia.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
        ModEstandar.gp_CampoAlfanumerico(eventSender.keyAscii, ".,:;()[]#$%&/\-_+*<>")
    End Sub

    Private Sub dbcJFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.Leave
        Dim I As Short
        Dim Aux As Short
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codFamilia, LTrim(RTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia LIKE '" & Trim(dbcJFamilia.Text) & "%'"
        Aux = mintJFamilia
        mintJFamilia = 0
        If Trim(dbcJFamilia.Text) <> Trim(C_TODAS) Or Trim(dbcJFamilia.Text) = "" Then
            ModDCombo.DCLostFocus(dbcJFamilia, gStrSql, mintJFamilia)
        End If
        If Aux <> mintJFamilia Then
            If mintJFamilia = 0 Then
                mblnFueraChange = True
                '''dbcJFamilia.text = C_TODAS
                dbcJFamilia.Text = ""
                dbcJFamilia.Enabled = True
                mintJLinea = 0
                '''dbcJLinea.text = C_TODAS
                dbcJLinea.Text = ""
                dbcJLinea.Enabled = False
                mintJSubLinea = 0
                '''dbcJSubLinea.text = C_TODAS
                dbcJSubLinea.Text = ""
                dbcJSubLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintJLinea = 0
                dbcJLinea.Text = C_TODAS
                dbcJLinea.Enabled = True
                mintJSubLinea = 0
                dbcJSubLinea.Text = C_TODAS
                dbcJSubLinea.Enabled = False
                mblnFueraChange = False
                dbcJLinea.Focus()
            End If
        End If
        '''If Trim(dbcJFamilia.text) = "" Then dbcJFamilia.text = C_TODAS
    End Sub

    Private Sub dbcJLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJLinea.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = " SELECT codLinea, LTrim(RTrim(descLinea)) as DescLinea FROM CatLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFamilia & " and descLinea LIKE '" & Trim(dbcJLinea.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcJLinea)

        If Trim(dbcJLinea.Text) = "" Then
            mintJLinea = 0
            mblnFueraChange = True
            dbcJLinea.Enabled = True
            mintJSubLinea = 0
            dbcJSubLinea.Text = C_TODAS
            dbcJSubLinea.Enabled = False
            mblnFueraChange = False
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
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

    Private Sub dbcJLinea_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcJLinea.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
        ModEstandar.gp_CampoAlfanumerico(eventSender.keyAscii, ".,:;()[]#$%&/\-_+*<>")
    End Sub

    Private Sub dbcJLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.Leave
        Dim Aux As Integer

        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub

        gStrSql = " SELECT codLinea, LTrim(RTrim(descLinea)) as DescLinea FROM CatLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFamilia & " and descLinea LIKE '" & Trim(dbcJLinea.Text) & "%'"
        Aux = mintJLinea
        mintJLinea = 0
        If Trim(dbcJLinea.Text) <> Trim(C_TODAS) Or Trim(dbcJLinea.Text) = "" Then
            ModDCombo.DCLostFocus(dbcJLinea, gStrSql, mintJLinea)
        End If

        If Aux <> mintJLinea Then
            If mintJLinea = 0 Then
                mblnFueraChange = True
                dbcJLinea.Text = C_TODAS
                dbcJLinea.Enabled = True
                mintJSubLinea = 0
                dbcJSubLinea.Text = C_TODAS
                dbcJSubLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintJSubLinea = 0
                dbcJSubLinea.Text = C_TODAS
                dbcJSubLinea.Enabled = True
                mblnFueraChange = False
                dbcJSubLinea.Focus()
            End If
        End If
        If Trim(dbcJLinea.Text) = "" Then dbcJLinea.Text = C_TODAS

    End Sub

    Private Sub dbcJSubLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJSubLinea.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        lStrSql = " SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as DescSubLinea FROM CatSubLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFamilia & " and CodLinea = " & mintJLinea & " and descSubLinea LIKE '" & Trim(dbcJSubLinea.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcJSubLinea)

        If Trim(dbcJSubLinea.Text) = "" Then
            mintJSubLinea = 0
            mblnFueraChange = True
            dbcJSubLinea.Enabled = True
            mblnFueraChange = False
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
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

    Private Sub dbcJSubLinea_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcJSubLinea.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
        ModEstandar.gp_CampoAlfanumerico(eventSender.keyAscii, ".,:;()[]#$%&/\-_+*<>")
    End Sub

    Private Sub dbcJSubLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.Leave
        Dim Aux As Integer

        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub

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

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcSucursal)
        If Trim(dbcSucursal.Text) = "" Then
            mintCodSucursal = 0
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen WHERE TipoAlmacen = 'P'"
        ModDCombo.DCGotFocus(gStrSql, dbcSucursal)
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodasSuc.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcSucursal.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
        ModEstandar.gp_CampoAlfanumerico(eventSender.keyAscii, ".,:;()[]#$%&/\-_+*<>")
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        Dim I As Short
        Dim Aux As Short

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

    Private Sub dtpFechaFinal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaFinal.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaFinal.Click
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaFinal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaFinal.KeyPress
        ' sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInicial.CursorChanged
        '  sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInicial.Click
        ' sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaInicial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaInicial.KeyPress
        ' sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub frmVtasVentasyExistxFam_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasVentasyExistxFam_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasVentasyExistxFam_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
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

    Private Sub frmVtasVentasyExistxFam_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasVentasyExistxFam_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        dtpFechaInicial.MinDate = C_FECHAINICIAL
        dtpFechaInicial.MaxDate = C_FECHAFINAL
        dtpFechaFinal.MinDate = C_FECHAINICIAL
        dtpFechaFinal.MaxDate = C_FECHAFINAL
        Call Me.Nuevo()
    End Sub

    Private Sub frmVtasVentasyExistxFam_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmVtasVentasyExistxFam_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'cmd.CommandTimeout = 90
        frmVtasRPTVentasSalidadeMercanciaClasifArtic = Nothing
    End Sub

    Private Sub txtMDSColor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSColor.Enter
        SelTextoTxt(txtMDSColor)
    End Sub

    Private Sub txtMDSColor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtMDSColor.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If Trim(txtMDSColor.Text) = "" Then txtMDSColor.Text = C_OPCION
        End If
    End Sub

    Private Sub txtMDSColor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSColor.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        ModEstandar.gp_CampoLetras(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMDSPeso_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPeso.TextChanged
        If CDec(ModEstandar.Numerico((txtMDSPeso.Text))) > 100 Then
            MsgBox("Valor incorrecto" & vbNewLine & "El peso no debe pasar de 100.00" & vbNewLine & vbNewLine & "Vefifique por favor...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            txtMDSPeso.Focus()
        End If
    End Sub

    Private Sub txtMDSPeso_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPeso.Enter
        SelTextoTxt(txtMDSPeso)
    End Sub

    Private Sub txtMDSPeso_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtMDSPeso.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 13 Then txtMDSPeso.Text = Format(ModEstandar.Numerico((txtMDSPeso.Text)), "##0.00")
    End Sub

    Private Sub txtMDSPeso_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSPeso.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, ".")
        KeyAscii = ModEstandar.MskCantidad((txtMDSPeso.Text), KeyAscii, 3, 2, (txtMDSPeso.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMDSPeso_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPeso.Leave
        txtMDSPeso.Text = Format(ModEstandar.Numerico((txtMDSPeso.Text)), "##0.00")
    End Sub

    Private Sub txtMDSPeso2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPeso2.TextChanged
        If CDec(ModEstandar.Numerico((txtMDSPeso2.Text))) > 100 Then
            MsgBox("Valor incorrecto" & vbNewLine & "El peso no debe pasar de 100.00" & vbNewLine & vbNewLine & "Vefifique por favor...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            txtMDSPeso2.Focus()
        End If
    End Sub

    Private Sub txtMDSPeso2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPeso2.Enter
        SelTextoTxt(txtMDSPeso2)
    End Sub

    Private Sub txtMDSPeso2_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtMDSPeso2.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 13 Then txtMDSPeso2.Text = Format(ModEstandar.Numerico((txtMDSPeso2.Text)), "##0.00")
    End Sub

    Private Sub txtMDSPeso2_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSPeso2.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, ".")
        KeyAscii = ModEstandar.MskCantidad((txtMDSPeso2.Text), KeyAscii, 3, 2, (txtMDSPeso2.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMDSPeso2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPeso2.Leave
        txtMDSPeso2.Text = Format(ModEstandar.Numerico((txtMDSPeso2.Text)), "##0.00")
    End Sub

    Private Sub txtMDSPureza_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPureza.Enter
        SelTextoTxt(txtMDSPureza)
    End Sub

    Private Sub txtMDSPureza_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtMDSPureza.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If Trim(txtMDSPureza.Text) = "" Then txtMDSPureza.Text = C_OPCION
        End If
    End Sub

    Private Sub txtMDSPureza_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSPureza.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class