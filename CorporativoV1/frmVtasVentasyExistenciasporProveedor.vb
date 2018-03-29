Option Strict Off
Option Explicit On
Imports Microsoft.Office.Interop
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasVentasyExistenciasporProveedor
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE VENTAS Y EXISTENCIAS POR PROVEEDOR                                                *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      MARTES 18 DE MAYO DE 2004                                                                    *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkMostrarCL As System.Windows.Forms.CheckBox
    Public WithEvents flexGrid As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents chkArticulosMovimiento As System.Windows.Forms.CheckBox
    Public WithEvents chkDescendente2 As System.Windows.Forms.CheckBox
    Public WithEvents chkDescendente1 As System.Windows.Forms.CheckBox
    Public WithEvents cboOrdenado As System.Windows.Forms.ComboBox
    Public WithEvents optPiezasVenta As System.Windows.Forms.RadioButton
    Public WithEvents optProveedor As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dbcProveedores As System.Windows.Forms.ComboBox
    Public WithEvents chkTodosProveedores As System.Windows.Forms.CheckBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox



    Dim mblnSalir As Boolean
    Dim mblnFueraChange As Boolean
    Dim mintCodProveedor As Integer
    Dim tecla As Integer
    Dim RsAux As ADODB.Recordset
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    Dim ObjExcel As Object
    Dim objLibro As Excel.Workbook
    Dim objHoja As Excel.Worksheet
    Dim ColumSepar As Integer
    Dim ColumCtoL As Integer
    Dim MostrarCostoL As Boolean

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Juan Carlos Osuna Corrales 10/Noviembre/2006                                              '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Variable para saber donde esta la columna que muestra el importe costeado de la existencia
    Dim ColumnaCtoExis As Integer
    'Variable para saber el renglon final
    Dim RenFinal As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Const C_ENCABEZADO As Integer = 9

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVtasVentasyExistenciasporProveedor))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkMostrarCL = New System.Windows.Forms.CheckBox()
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.chkArticulosMovimiento = New System.Windows.Forms.CheckBox()
        Me.cboOrdenado = New System.Windows.Forms.ComboBox()
        Me.optPiezasVenta = New System.Windows.Forms.RadioButton()
        Me.optProveedor = New System.Windows.Forms.RadioButton()
        Me.chkTodosProveedores = New System.Windows.Forms.CheckBox()
        Me.flexGrid = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.chkDescendente2 = New System.Windows.Forms.CheckBox()
        Me.chkDescendente1 = New System.Windows.Forms.CheckBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dbcProveedores = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        CType(Me.flexGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkMostrarCL
        '
        Me.chkMostrarCL.BackColor = System.Drawing.SystemColors.Control
        Me.chkMostrarCL.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMostrarCL.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkMostrarCL.Location = New System.Drawing.Point(215, 230)
        Me.chkMostrarCL.Margin = New System.Windows.Forms.Padding(2)
        Me.chkMostrarCL.Name = "chkMostrarCL"
        Me.chkMostrarCL.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMostrarCL.Size = New System.Drawing.Size(82, 19)
        Me.chkMostrarCL.TabIndex = 19
        Me.chkMostrarCL.Text = "Costo L"
        Me.ToolTip1.SetToolTip(Me.chkMostrarCL, "Muestra Cto L")
        Me.chkMostrarCL.UseVisualStyleBackColor = False
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
        Me.txtMensaje.Size = New System.Drawing.Size(319, 71)
        Me.txtMensaje.TabIndex = 17
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'chkArticulosMovimiento
        '
        Me.chkArticulosMovimiento.BackColor = System.Drawing.SystemColors.Control
        Me.chkArticulosMovimiento.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkArticulosMovimiento.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkArticulosMovimiento.Location = New System.Drawing.Point(6, 228)
        Me.chkArticulosMovimiento.Margin = New System.Windows.Forms.Padding(2)
        Me.chkArticulosMovimiento.Name = "chkArticulosMovimiento"
        Me.chkArticulosMovimiento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkArticulosMovimiento.Size = New System.Drawing.Size(195, 20)
        Me.chkArticulosMovimiento.TabIndex = 16
        Me.chkArticulosMovimiento.Text = "Mostrar Articulos Sin Movimientos"
        Me.ToolTip1.SetToolTip(Me.chkArticulosMovimiento, "Muestra Articulos que no Tienen Movimientos en el Periodo Especificado")
        Me.chkArticulosMovimiento.UseVisualStyleBackColor = False
        '
        'cboOrdenado
        '
        Me.cboOrdenado.BackColor = System.Drawing.SystemColors.Window
        Me.cboOrdenado.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboOrdenado.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboOrdenado.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboOrdenado.Items.AddRange(New Object() {"Código", "Descripción"})
        Me.cboOrdenado.Location = New System.Drawing.Point(89, 24)
        Me.cboOrdenado.Margin = New System.Windows.Forms.Padding(2)
        Me.cboOrdenado.Name = "cboOrdenado"
        Me.cboOrdenado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboOrdenado.Size = New System.Drawing.Size(96, 21)
        Me.cboOrdenado.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.cboOrdenado, "Por Código/Proveedor")
        '
        'optPiezasVenta
        '
        Me.optPiezasVenta.BackColor = System.Drawing.SystemColors.Control
        Me.optPiezasVenta.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPiezasVenta.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPiezasVenta.Location = New System.Drawing.Point(8, 58)
        Me.optPiezasVenta.Margin = New System.Windows.Forms.Padding(2)
        Me.optPiezasVenta.Name = "optPiezasVenta"
        Me.optPiezasVenta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPiezasVenta.Size = New System.Drawing.Size(110, 17)
        Me.optPiezasVenta.TabIndex = 14
        Me.optPiezasVenta.TabStop = True
        Me.optPiezasVenta.Text = "Piezas de Venta"
        Me.ToolTip1.SetToolTip(Me.optPiezasVenta, "Ordenar por Numero de Piezas Vendidas")
        Me.optPiezasVenta.UseVisualStyleBackColor = False
        '
        'optProveedor
        '
        Me.optProveedor.BackColor = System.Drawing.SystemColors.Control
        Me.optProveedor.Cursor = System.Windows.Forms.Cursors.Default
        Me.optProveedor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optProveedor.Location = New System.Drawing.Point(8, 26)
        Me.optProveedor.Margin = New System.Windows.Forms.Padding(2)
        Me.optProveedor.Name = "optProveedor"
        Me.optProveedor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optProveedor.Size = New System.Drawing.Size(79, 17)
        Me.optProveedor.TabIndex = 11
        Me.optProveedor.TabStop = True
        Me.optProveedor.Text = "Proveedor"
        Me.ToolTip1.SetToolTip(Me.optProveedor, "Ordena por Proveedor")
        Me.optProveedor.UseVisualStyleBackColor = False
        '
        'chkTodosProveedores
        '
        Me.chkTodosProveedores.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodosProveedores.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodosProveedores.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTodosProveedores.Location = New System.Drawing.Point(6, 13)
        Me.chkTodosProveedores.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodosProveedores.Name = "chkTodosProveedores"
        Me.chkTodosProveedores.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodosProveedores.Size = New System.Drawing.Size(151, 17)
        Me.chkTodosProveedores.TabIndex = 4
        Me.chkTodosProveedores.Text = "Todos los Proveedores"
        Me.ToolTip1.SetToolTip(Me.chkTodosProveedores, "Muestra Todos los Proveedores")
        Me.chkTodosProveedores.UseVisualStyleBackColor = False
        '
        'flexGrid
        '
        Me.flexGrid.DataSource = Nothing
        Me.flexGrid.Location = New System.Drawing.Point(11, 361)
        Me.flexGrid.Margin = New System.Windows.Forms.Padding(2)
        Me.flexGrid.Name = "flexGrid"
        Me.flexGrid.OcxState = CType(resources.GetObject("flexGrid.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexGrid.Size = New System.Drawing.Size(336, 96)
        Me.flexGrid.TabIndex = 18
        Me.flexGrid.Visible = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.txtMensaje)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(6, 254)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(341, 98)
        Me.Frame4.TabIndex = 3
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Texto Adicional"
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.chkDescendente2)
        Me.Frame3.Controls.Add(Me.chkDescendente1)
        Me.Frame3.Controls.Add(Me.cboOrdenado)
        Me.Frame3.Controls.Add(Me.optPiezasVenta)
        Me.Frame3.Controls.Add(Me.optProveedor)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(6, 130)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(308, 92)
        Me.Frame3.TabIndex = 2
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Ordenado Por"
        '
        'chkDescendente2
        '
        Me.chkDescendente2.BackColor = System.Drawing.SystemColors.Control
        Me.chkDescendente2.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDescendente2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDescendente2.Location = New System.Drawing.Point(194, 55)
        Me.chkDescendente2.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDescendente2.Name = "chkDescendente2"
        Me.chkDescendente2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDescendente2.Size = New System.Drawing.Size(98, 20)
        Me.chkDescendente2.TabIndex = 15
        Me.chkDescendente2.Text = "Descendente"
        Me.chkDescendente2.UseVisualStyleBackColor = False
        '
        'chkDescendente1
        '
        Me.chkDescendente1.BackColor = System.Drawing.SystemColors.Control
        Me.chkDescendente1.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDescendente1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDescendente1.Location = New System.Drawing.Point(194, 23)
        Me.chkDescendente1.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDescendente1.Name = "chkDescendente1"
        Me.chkDescendente1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDescendente1.Size = New System.Drawing.Size(98, 20)
        Me.chkDescendente1.TabIndex = 13
        Me.chkDescendente1.Text = "Descendente"
        Me.chkDescendente1.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.dtpFechaInicial)
        Me.Frame2.Controls.Add(Me.dtpFechaFinal)
        Me.Frame2.Controls.Add(Me.Label3)
        Me.Frame2.Controls.Add(Me.Label2)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(6, 78)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(341, 46)
        Me.Frame2.TabIndex = 1
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(67, 19)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(99, 20)
        Me.dtpFechaInicial.TabIndex = 8
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(229, 19)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(96, 20)
        Me.dtpFechaFinal.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(182, 21)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(42, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Hasta"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(22, 22)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(47, 17)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Desde"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dbcProveedores)
        Me.Frame1.Controls.Add(Me.chkTodosProveedores)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(6, 6)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(308, 66)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'dbcProveedores
        '
        Me.dbcProveedores.Location = New System.Drawing.Point(101, 34)
        Me.dbcProveedores.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcProveedores.Name = "dbcProveedores"
        Me.dbcProveedores.Size = New System.Drawing.Size(191, 21)
        Me.dbcProveedores.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(30, 40)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(67, 14)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Proveedor :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(130, 472)
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
        Me.btnImprimir.Location = New System.Drawing.Point(15, 472)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 78
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(245, 473)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 77
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmVtasVentasyExistenciasporProveedor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(367, 520)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.chkMostrarCL)
        Me.Controls.Add(Me.flexGrid)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.chkArticulosMovimiento)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(344, 160)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasVentasyExistenciasporProveedor"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Ventas y Existencias por Proveedor"
        CType(Me.flexGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub


    Sub CierraInstanciasdeExcel(ByRef Tipo As Integer)
        If Tipo = 1 Then
            objLibro.Close()
            ObjExcel.Quit()
        End If
        If ObjExcel Is Nothing Then ObjExcel = Nothing
        If objLibro Is Nothing Then objLibro = Nothing
        If objHoja Is Nothing Then objHoja = Nothing
    End Sub

    Function DevuelveQuery() As String
        On Error GoTo Err_Renamed
        Dim Sql As String
        Dim Inv As String
        Dim GroupBy As String
        Dim cSELECT As String
        Dim cSubSelectVtas As String
        Dim cSubSelectExis As String

        cSELECT = "Select  V.CodProveedor,V.DescProvAcreed,V.CodArticulo,V.DescArticulo,V.CodigoArticuloProv,V.CodigoAnt, " & "A.CodProveedor AS CodProv,A.DescProvAcreed AS DescProv,A.CodArticulo AS CodArt,A.DescArticulo AS DescArt, " & "A.CodigoArticuloProv AS CodArtProv, " & "Convert(char(1),A.OrigenAnt) + '-' + right('00000' + ltrim(rtrim(convert(char(5),A.CodigoAnt))),5) as CodAnt, " & "A.PrecioPubDolar, Round(A.CostoReal,2) as CostoReal "

        Sql = "(SELECT Vta.CodProveedor,IsNull(P.DescProvAcreed,'') as DescProvAcreed,Vta.CodArticulo,Convert(char(1),Vta.OrigenAnt) + '-' + right('00000' + ltrim(rtrim(convert(char(5),Vta.CodigoAnt))),5) as CodigoAnt," & "SubString(Vta.DescArticulo,1,75) as DescArticulo,Vta.CodigoArticuloProv,PrecioPubDolar,"
        Inv = "(SELECT Inv.CodArticulo,SUM((Inv.ExistenciaInicial+Inv.Entradas) - (Inv.Salidas+Inv.Apartados)) as ExistTotal"

        gStrSql = "Select CodAlmacen,DescAlmacen From CatAlmacen Where TipoAlmacen = 'P' Order By CodAlmacen"

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsAux = Cmd.Execute
        If RsAux.RecordCount > 0 Then
            Do While Not RsAux.EOF
                Sql = Sql & "Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then Vta.Cantidad - Vta.CantidadDev Else 0 End) as CantidadVta" & RsAux.Fields("CodAlmacen").Value & ","
                Inv = Inv & ",SUM(Case When Inv.CodAlmacen = " & RsAux.Fields("CodAlmacen").Value & " then (ExistenciaInicial+Entradas) - (Salidas+Apartados) else 0 end) as Existencia" & RsAux.Fields("CodAlmacen").Value
                cSubSelectVtas = cSubSelectVtas & ",ISNULL(V.CantidadVta" & RsAux.Fields("CodAlmacen").Value & ",0) AS CantidadVta" & RsAux.Fields("CodAlmacen").Value & ""
                cSubSelectExis = cSubSelectExis & ",ISNULL(Inv.Existencia" & RsAux.Fields("CodAlmacen").Value & ",0) AS Existencia" & RsAux.Fields("CodAlmacen").Value & ""
                RsAux.MoveNext()
            Loop
        End If

        cSubSelectVtas = cSubSelectVtas & ",ISNULL(V.VentaTotal,0) AS VentaTotal"

        cSubSelectExis = cSubSelectExis & ", ISNULL(Inv.ExistTotal,0) AS ExistenciaTotal, (ISNULL(Inv.ExistTotal,0) * Round(A.CostoReal,2)) as ImpteCtoLetras " & "FROM (SELECT CA.CodProveedor,CP.DescProvAcreed,CA.CodArticulo,CA.DescArticulo,CA.CodigoArticuloProv,CA.OrigenAnt," & "CA.CodigoAnt,CA.PrecioPubDolar, CA.CostoReal FROM CatArticulos CA Left Outer Join CatProvAcreed CP ON CA.CodProveedor = CP.CodProvAcreed " & IIf(mintCodProveedor <> 0, "WHERE CodProveedor = " & mintCodProveedor & "", "") & ") A " & IIf(chkArticulosMovimiento.CheckState = System.Windows.Forms.CheckState.Checked, "LEFT OUTER JOIN ", "INNER JOIN ")
        Sql = Sql & "Sum(Vta.Cantidad - Vta.CantidadDev) as VentaTotal "

        Inv = Inv & " From Inventario Inv (Nolock) Inner Join CatAlmacen A On Inv.CodAlmacen = A.CodAlmacen And A.TipoAlmacen = 'P' " & "Inner Join CatArticulos C On Inv.CodArticulo = C.CodArticulo " & IIf(chkTodosProveedores.CheckState = System.Windows.Forms.CheckState.Checked, "", "where c.codproveedor = " & mintCodProveedor & " ") & "Group by Inv.CodArticulo) Inv On A.CodArticulo = Inv.CodArticulo "
        Sql = Sql & "From DBO.VENTAS_SALIDAMCIA('" & Format(dtpFechaInicial.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpFechaFinal.Value, C_FORMATFECHAGUARDAR) & "') Vta " & "Left Outer Join CatProvAcreed P On Vta.CodProveedor = P.CodProvAcreed " & "Where Vta.Tipo <> 'R' " & IIf(chkTodosProveedores.CheckState = System.Windows.Forms.CheckState.Checked, "", "and vta.codproveedor = " & mintCodProveedor & " ") & "Group by Vta.CodArticulo,Vta.DescArticulo,Vta.CodProveedor,Vta.CodigoArticuloProv,Vta.OrigenAnt,Vta.CodigoAnt,IsNull(P.DescProvAcreed,''),PrecioPubDolar " & IIf(chkArticulosMovimiento.CheckState = System.Windows.Forms.CheckState.Unchecked, "HAVING Sum(Vta.Cantidad - Vta.CantidadDev) > 0", "") & ") V " & "ON A.CodArticulo = V.CodArticulo And A.CodProveedor = V.CodProveedor " & "Left Outer Join " & Inv
        DevuelveQuery = cSELECT & cSubSelectVtas & cSubSelectExis & Sql & "ORDER BY " & IIf(optProveedor.Checked = True, IIf(cboOrdenado.SelectedIndex = 0, IIf(chkDescendente1.CheckState = System.Windows.Forms.CheckState.Checked, "A.CodProveedor Desc,A.CodArticulo ASC", "A.CodProveedor,A.CodArticulo"), IIf(chkDescendente1.CheckState = System.Windows.Forms.CheckState.Checked, "IsNull(A.DescProvAcreed,'') Desc,A.CodArticulo ASC", "IsNull(A.DescProvAcreed,''),A.CodArticulo ")), IIf(chkDescendente2.CheckState = System.Windows.Forms.CheckState.Checked, "A.CodProveedor ASC,V.VentaTotal Desc", "A.CodProveedor,V.VentaTotal "))

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Encabezado()
        On Error GoTo Err_Renamed
        Dim Columna As Integer
        With objHoja
            .Range("C1").FormulaR1C1 = Trim(gstrCorpoNOMBREEMPRESA)
            .Range("C1:G1").Select()
            .Range("C1:G1").MergeCells = True
            .Range("C1:G1").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C1:G1").Font
                .Bold = True
                .Size = 12
                .Name = "Arial"
            End With
            .Range("C2").FormulaR1C1 = "Ventas y Existencias por Proveedor"
            .Range("C2:G2").Select()
            .Range("C2:G2").MergeCells = True
            .Range("C2:G2").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C2:G2").Font
                .Bold = False
                .Size = 11
                .Name = "Arial"
            End With
            .Range("C3").FormulaR1C1 = "Desde el " & Format(dtpFechaInicial.Value, "dd/mmm/yyyy") & " Hasta el " & Format(dtpFechaFinal.Value, "dd/mmm/yyyy")
            .Range("C3:G3").Select()
            .Range("C3:G3").MergeCells = True
            .Range("C3:G3").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C3:G3").Font
                .Bold = False
                .Size = 10
                .Name = "Arial"
            End With
            .Range("A4").FormulaR1C1 = "Fecha: " & Format(Today, "dd/mmm/yyyy")
            .Range("A4:B4").Select()
            '.Range("A4:B4").MergeCells = True
            .Range("A4:B4").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A4:B4").Font
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
                .Range("B7").FormulaR1C1 = Trim(QuitaEnter(txtMensaje.Text))
                .Range("B7:J7").Select()
                .Range("B7:J7").MergeCells = True
                .Range("B7:J7").HorizontalAlignment = Excel.Constants.xlLeft
                With .Range("B7:J7").Font
                    .Bold = False
                    .Size = 9
                    .Name = "Arial"
                End With
            End If
            .Range(.Cells._Default(C_ENCABEZADO, 1), .Cells._Default(C_ENCABEZADO, 1)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, 1), .Cells._Default(C_ENCABEZADO, 1))._Default = "Proveedor"
            .Range(.Cells._Default(C_ENCABEZADO, 1), .Cells._Default(C_ENCABEZADO, 1)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, 1), .Cells._Default(C_ENCABEZADO, 1)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, 1), .Cells._Default(C_ENCABEZADO, 1)).WrapText = True
            .Range(.Cells._Default(C_ENCABEZADO, 1), .Cells._Default(C_ENCABEZADO, 1)).RowHeight = 90
            .Range(.Cells._Default(C_ENCABEZADO, 1), .Cells._Default(C_ENCABEZADO, 1)).ColumnWidth = 7
            .Range(.Cells._Default(C_ENCABEZADO, 1), .Cells._Default(C_ENCABEZADO, 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(C_ENCABEZADO, 1), .Cells._Default(C_ENCABEZADO, 1)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, 1), .Cells._Default(C_ENCABEZADO, 1)).ColumnWidth = 11.5
            '.Range(.Cells(C_ENCABEZADO, 1), .Cells(C_ENCABEZADO, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous

            .Range(.Cells._Default(C_ENCABEZADO, 2), .Cells._Default(C_ENCABEZADO, 2)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, 2), .Cells._Default(C_ENCABEZADO, 2))._Default = "Código Art"
            .Range(.Cells._Default(C_ENCABEZADO, 2), .Cells._Default(C_ENCABEZADO, 2)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, 2), .Cells._Default(C_ENCABEZADO, 2)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, 2), .Cells._Default(C_ENCABEZADO, 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, 2), .Cells._Default(C_ENCABEZADO, 2)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, 2), .Cells._Default(C_ENCABEZADO, 2)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, 2), .Cells._Default(C_ENCABEZADO, 2)).ColumnWidth = 7

            .Range(.Cells._Default(C_ENCABEZADO, 3), .Cells._Default(C_ENCABEZADO, 3)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, 3), .Cells._Default(C_ENCABEZADO, 3))._Default = "Descripción"
            .Range(.Cells._Default(C_ENCABEZADO, 3), .Cells._Default(C_ENCABEZADO, 3)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, 3), .Cells._Default(C_ENCABEZADO, 3)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, 3), .Cells._Default(C_ENCABEZADO, 3)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, 3), .Cells._Default(C_ENCABEZADO, 3)).WrapText = True
            .Range(.Cells._Default(C_ENCABEZADO, 3), .Cells._Default(C_ENCABEZADO, 3)).ColumnWidth = 40
            With .Range(.Cells._Default(C_ENCABEZADO, 3), .Cells._Default(C_ENCABEZADO, 3)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With

            .Range(.Cells._Default(C_ENCABEZADO, 4), .Cells._Default(C_ENCABEZADO, 4)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, 4), .Cells._Default(C_ENCABEZADO, 4))._Default = "Código Articulo Prov"
            .Range(.Cells._Default(C_ENCABEZADO, 4), .Cells._Default(C_ENCABEZADO, 4)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, 4), .Cells._Default(C_ENCABEZADO, 4)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, 4), .Cells._Default(C_ENCABEZADO, 4)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, 4), .Cells._Default(C_ENCABEZADO, 4)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, 4), .Cells._Default(C_ENCABEZADO, 4)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, 4), .Cells._Default(C_ENCABEZADO, 4)).ColumnWidth = 12.5

            .Range(.Cells._Default(C_ENCABEZADO, 5), .Cells._Default(C_ENCABEZADO, 5)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, 5), .Cells._Default(C_ENCABEZADO, 5))._Default = "Código Ant"
            .Range(.Cells._Default(C_ENCABEZADO, 5), .Cells._Default(C_ENCABEZADO, 5)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, 5), .Cells._Default(C_ENCABEZADO, 5)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, 5), .Cells._Default(C_ENCABEZADO, 5)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, 5), .Cells._Default(C_ENCABEZADO, 5)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, 5), .Cells._Default(C_ENCABEZADO, 5)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, 5), .Cells._Default(C_ENCABEZADO, 5)).ColumnWidth = 7

            .Range(.Cells._Default(C_ENCABEZADO, 6), .Cells._Default(C_ENCABEZADO, 6)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, 6), .Cells._Default(C_ENCABEZADO, 6))._Default = "Precio Pub"
            .Range(.Cells._Default(C_ENCABEZADO, 6), .Cells._Default(C_ENCABEZADO, 6)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, 6), .Cells._Default(C_ENCABEZADO, 6)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, 6), .Cells._Default(C_ENCABEZADO, 6)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, 6), .Cells._Default(C_ENCABEZADO, 6)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, 6), .Cells._Default(C_ENCABEZADO, 6)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, 6), .Cells._Default(C_ENCABEZADO, 6)).ColumnWidth = 10

            .Range(.Cells._Default(C_ENCABEZADO, 7), .Cells._Default(C_ENCABEZADO, 7)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, 7), .Cells._Default(C_ENCABEZADO, 7))._Default = "Costo L"
            .Range(.Cells._Default(C_ENCABEZADO, 7), .Cells._Default(C_ENCABEZADO, 7)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, 7), .Cells._Default(C_ENCABEZADO, 7)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(C_ENCABEZADO, 7), .Cells._Default(C_ENCABEZADO, 7)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, 7), .Cells._Default(C_ENCABEZADO, 7)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, 7), .Cells._Default(C_ENCABEZADO, 7)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, 7), .Cells._Default(C_ENCABEZADO, 7)).ColumnWidth = 10


            Columna = 8
            gStrSql = "Select CodAlmacen,DescAlmacen From CatAlmacen Where TipoAlmacen = 'P' Order By CodAlmacen"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsAux = Cmd.Execute
            If RsAux.RecordCount > 0 Then
                Do While Not RsAux.EOF
                    .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Select()
                    .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).VerticalAlignment = Excel.Constants.xlTop
                    .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).WrapText = True
                    .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).RowHeight = 26
                    If Columna = 8 Then
                        .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3)).Select()
                        .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3)).MergeCells = True
                        .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3))._Default = "Ventas"
                    End If
                    With .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                    .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Interior.ColorIndex = 15
                    .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).ColumnWidth = 3.6

                    .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
                    .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = Trim(UCase((RsAux.Fields("DescAlmacen").Value)) & LCase((RsAux.Fields("DescAlmacen").Value)))
                    .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
                    .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                    .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Orientation = 90
                    .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
                    With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                    RsAux.MoveNext()
                    .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 15
                    .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 3.6
                    .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    Columna = Columna + 1
                Loop
            End If

            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Select()
            With .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).ColumnWidth = 5
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "TOTAL VTA"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Orientation = 90
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 5
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).ColumnWidth = 3.6
            ColumSepar = Columna
            Columna = Columna + 1
            RsAux.MoveFirst()

            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3)).Select()
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3)).MergeCells = True
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3)).VerticalAlignment = Excel.Constants.xlTop
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3)).WrapText = True
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3))._Default = "Existencias"
            With .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3)).Interior.ColorIndex = 15
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3)).ColumnWidth = 3.6
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna + 3)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous

            Do While Not RsAux.EOF
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = Trim(UCase((RsAux.Fields("DescAlmacen").Value)) & LCase((RsAux.Fields("DescAlmacen").Value)))
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Orientation = 90
                If Columna - 1 = ColumSepar Then
                    .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                End If
                With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                RsAux.MoveNext()
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 15
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 3.6
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                Columna = Columna + 1

                If Not RsAux.EOF Then
                    .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Select()
                    With .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                    .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Interior.ColorIndex = 15
                    .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).ColumnWidth = 3.6
                End If
            Loop

            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).MergeCells = True
            With .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(C_ENCABEZADO - 1, Columna), .Cells._Default(C_ENCABEZADO - 1, Columna)).ColumnWidth = 5

            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).MergeCells = True
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "TOTAL EXIST"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Orientation = 90
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 5
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 2
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 3.6
            ColumCtoL = Columna

            Columna = Columna + 1
            'Juan Carlos Osuna Corrales 10/Noviembre/2006
            'guardamos el numero de columna donde mostramos el impoirte costeado de la existencia
            ColumnaCtoExis = Columna
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).MergeCells = True
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Impte Existencia"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Orientation = 90
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = True
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 9
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous

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

    Sub EnviaExcel()
        Dim Archivo As String
        On Error GoTo Err_Renamed
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        System.Windows.Forms.Application.DoEvents()
        If Dir(gstrCorpoDriveLocal & "\Sistema\", FileAttribute.Directory + FileAttribute.Hidden) = "" Then
            MsgBox("No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        Archivo = "VE" & CStr(Format(Month(Today), "00")) & CStr(Format((Today) + "00")) & (CStr(Format(Year(Today) + "00"))) & ".xls"
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
        Encabezado()
        LlenaDatos()
        objLibro.SaveAs(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        System.Windows.Forms.Application.DoEvents()
        Select Case MsgBox("Se ha creado el archivo " & Archivo & " ¿Desea abrirlo?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
            Case MsgBoxResult.Yes
                ObjExcel.Visible = True
                ObjExcel = Nothing
                objLibro = Nothing
                objHoja = Nothing
            Case MsgBoxResult.No Or MsgBoxResult.Cancel
                CierraInstanciasdeExcel(1)
        End Select
Err_Renamed:
        If Err.Number = 70 Then
            MsgBox("No se puede generar un nuevo archivo hasta que el anterior este cerrado.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            CierraInstanciasdeExcel(2)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        ElseIf Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub LlenaDatos()
        On Error GoTo Err_Renamed
        Dim Renglon As Integer
        Dim Columna As Integer
        Dim RenRecorridos As Integer
        Dim CodProveedor As Integer
        Dim I As Integer
        Dim Rango As String
        Dim Formula As String
        Dim Totales As String
        Dim Cantidad As String
        Dim blnCtoL As Boolean

        Renglon = 10
        CodProveedor = 0
        flexGrid.Clear()
        flexGrid.Rows = 2
        flexGrid.set_Cols(0, 1)
        blnCtoL = True

        With objHoja
            RsGral.MoveFirst()
            Do While Not RsGral.EOF
                If CodProveedor = 0 Then
                    CodProveedor = RsGral.Fields("CodProv").Value
                    Rango = "A" & Renglon
                    .Range(Rango)._Default = "Prov: " & ("000" & CStr(RsGral.Fields("CodProv").Value))
                    .Range(Rango).HorizontalAlignment = Excel.Constants.xlLeft
                    With .Range(Rango).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                    Rango = "C" & Renglon
                    .Range(Rango).Select()
                    .Range(Rango)._Default = Trim(RsGral.Fields("DescProv").Value)
                    .Range(Rango).HorizontalAlignment = Excel.Constants.xlLeft
                    With .Range(Rango).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                    RenRecorridos = 0
                ElseIf CodProveedor <> RsGral.Fields("CodProv").Value Then
                    'hacemos corte
                    CodProveedor = RsGral.Fields("CodProv").Value
                    Renglon = Renglon + 1
                    Rango = "F" & Renglon
                    .Range(Rango).Select()
                    .Range(Rango)._Default = "Total Prov"
                    .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
                    With .Range(Rango).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                    Columna = 8
                    For I = 14 To RsGral.Fields.Count - 1
                        If Columna = ColumSepar Then
                            Columna = Columna + 1
                        ElseIf Columna = ColumCtoL Then
                            Columna = Columna + 1
                        End If
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).FormulaR1C1 = "=SUM(R[-" & RenRecorridos & "]C:R[-1]C)"
                        If I = (RsGral.Fields.Count - 1) Then
                            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                        Else
                            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                        End If
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        If (Columna - 1) = ColumSepar Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        If (Columna - 1) = ColumCtoL Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous

                        With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                            .Size = 8
                            .Name = "Arial"
                        End With
                        Columna = Columna + 1
                    Next
                    Renglon = Renglon + 2
                    Rango = "A" & Renglon
                    .Range(Rango)._Default = "Prov: " & ("000" & CStr(RsGral.Fields("CodProv").Value))
                    .Range(Rango).HorizontalAlignment = Excel.Constants.xlLeft
                    With .Range(Rango).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                    Rango = "C" & Renglon
                    .Range(Rango).Select()
                    .Range(Rango)._Default = Trim(RsGral.Fields("DescProv").Value)
                    .Range(Rango).HorizontalAlignment = Excel.Constants.xlLeft
                    With .Range(Rango).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                    RenRecorridos = 0
                End If

                Renglon = Renglon + 1
                Rango = "C" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = Trim(RsGral.Fields("DescArt").Value)
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlLeft
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                Rango = "B" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("CodArt").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "D" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("CodArtProv").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlLeft
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "E" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("CodAnt").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "F" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("PrecioPubDolar").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
                .Range(Rango).NumberFormat = "###,##0.00"
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Rango = "G" & Renglon
                .Range(Rango).Select()
                .Range(Rango)._Default = RsGral.Fields("CostoReal").Value
                .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
                .Range(Rango).NumberFormat = "###,##0.00"
                .Range(Rango).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(Rango).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = 8
                For I = 14 To RsGral.Fields.Count - 1
                    If I = 14 Then
                        flexGrid.Row = 1
                        flexGrid.Col = 0
                        Cantidad = Numerico(flexGrid.get_TextMatrix(flexGrid.Row, flexGrid.Col))
                        Cantidad = Cantidad + RsGral.Fields(I).Value
                        flexGrid.set_TextMatrix(flexGrid.Row, flexGrid.Col, Cantidad)
                    Else
                        If flexGrid.Col = flexGrid.get_Cols() - 1 Then
                            flexGrid.set_Cols(0, flexGrid.get_Cols() + 1)
                        End If
                        flexGrid.Col = flexGrid.Col + 1
                        Cantidad = Numerico(flexGrid.get_TextMatrix(flexGrid.Row, flexGrid.Col))
                        Cantidad = Cantidad + RsGral.Fields(I).Value
                        flexGrid.set_TextMatrix(flexGrid.Row, flexGrid.Col, Cantidad)
                    End If
                    If Columna = ColumSepar Then
                        Columna = Columna + 1
                    ElseIf Columna = ColumCtoL Then
                        Columna = Columna + 1
                    End If
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields(I).Value
                    If I = (RsGral.Fields.Count - 1) Then
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                    Else
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                    End If
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
                If RsGral.EOF Then
                    Renglon = Renglon + 1
                    Rango = "F" & Renglon
                    .Range(Rango).Select()
                    .Range(Rango)._Default = "Total Prov"
                    .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
                    With .Range(Rango).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                    Columna = 8
                    For I = 14 To RsGral.Fields.Count - 1
                        If Columna = ColumSepar Then
                            Columna = Columna + 1
                        ElseIf Columna = ColumCtoL Then
                            Columna = Columna + 1
                        End If
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).MergeCells = True
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).FormulaR1C1 = "=SUM(R[-" & RenRecorridos & "]C:R[-1]C)"
                        If I = (RsGral.Fields.Count - 1) Then
                            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                        Else
                            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                        End If
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        If (Columna - 1) = ColumSepar Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        If (Columna - 1) = ColumCtoL Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous

                        With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                            .Size = 8
                            .Name = "Arial"
                        End With
                        Columna = Columna + 1
                    Next
                End If
            Loop
            RsGral.MoveFirst()
            Renglon = Renglon + 2
            Rango = "F" & Renglon
            .Range(Rango).Select()
            .Range(Rango)._Default = "GRAN TOTAL"
            .Range(Rango).HorizontalAlignment = Excel.Constants.xlRight
            With .Range(Rango).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            'Juan Carlos Osuna Corrales 10/Noviembre/2006
            'Almacenamos el renglon final para saber hasta donde vamos a borrar cuando no mostremos el Costo L
            RenFinal = Renglon
            Columna = 8
            flexGrid.Col = 0
            flexGrid.Row = 1
            For I = 14 To RsGral.Fields.Count - 1
                If Columna = ColumSepar Then
                    Columna = Columna + 1
                ElseIf Columna = ColumCtoL Then
                    Columna = Columna + 1
                End If
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).FormulaR1C1 = flexGrid.get_TextMatrix(flexGrid.Row, flexGrid.Col)
                If I = (RsGral.Fields.Count - 1) Then
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                Else
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                End If
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                If (Columna - 1) = ColumSepar Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                If (Columna - 1) = ColumCtoL Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous

                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                If flexGrid.Col < flexGrid.get_Cols() - 1 Then
                    flexGrid.Col = flexGrid.Col + 1
                End If
            Next

            '''OCULTAR COLUMNAS
            '''.Range("G:G").EntireColumn.Hidden = True
            '''.Range("AE:AF").EntireColumn.Hidden = True
            If Not MostrarCostoL Then
                .Range("G:G").Delete(Excel.XlDirection.xlToLeft)
                'esto ya no sirve
                '.Range("AD:EA").Delete (xlToLeft)
                'Juan Carlos Osuna Corrales 10/Noviembre/2006
                'Eliminamos la columna del importe costeado
                'desde el renglon 1 hasta el renglon final(RenFinal) que es el renglon hasta donde termina el reporte
                'a ColumnaCtoExis - 1, le restamos 1 porque ya eliminamos la columna del costo
                'y ya se recorrieron en 1 hacia atras todas las columnas que estan despues del costo
                .Range(.Cells._Default(1, ColumnaCtoExis - 1), .Cells._Default(RenFinal, ColumnaCtoExis - 1)).Delete(Excel.XlDirection.xlToLeft)
            Else
                If chkMostrarCL.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                    .Range("G:G").Delete(Excel.XlDirection.xlToLeft)
                    'esto ya no sirve
                    '.Range("AD:EA").Delete (xlToLeft)
                    'Juan Carlos Osuna Corrales 10/Noviembre/2006
                    'Eliminamos la columna del importe costeado
                    'desde el renglon 1 hasta el renglon final(RenFinal) que es el renglon hasta donde termina el reporte
                    'a ColumnaCtoExis - 1, le restamos 1 porque ya eliminamos la columna del costo
                    'y ya se recorrieron en 1 hacia atras todas las columnas que estan despues del costo
                    .Range(.Cells._Default(1, ColumnaCtoExis - 1), .Cells._Default(RenFinal, ColumnaCtoExis - 1)).Delete(Excel.XlDirection.xlToLeft)
                End If
            End If
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

    Sub Imprime()
        On Error GoTo ImprimeErr
        If Not ValidaDatos() Then Exit Sub
        gStrSql = DevuelveQuery()
        ModEstandar.BorraCmd()
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            EnviaExcel()
        Else
            MsgBox("No existe información por mostrar en este periodo de fechas, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        End If
        Cmd.CommandTimeout = 90

ImprimeErr:
        If Err.Number <> 0 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
            mblnFueraChange = False
        End If
    End Sub

    Sub Limpiar()
        Nuevo()
        chkTodosProveedores.Focus()
    End Sub

    Sub Nuevo()
        chkTodosProveedores.CheckState = System.Windows.Forms.CheckState.Checked
        dbcProveedores.Enabled = False
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
        optProveedor.Checked = True
        optPiezasVenta.Checked = False
        chkArticulosMovimiento.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtMensaje.Text = ""
        cboOrdenado.SelectedIndex = 0
        chkDescendente1.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkDescendente2.CheckState = System.Windows.Forms.CheckState.Unchecked

        If MostrarCostoL Then
            chkMostrarCL.Visible = True
            chkMostrarCL.CheckState = System.Windows.Forms.CheckState.Checked
        Else
            chkMostrarCL.CheckState = System.Windows.Forms.CheckState.Checked
            chkMostrarCL.Visible = False
        End If

        mblnSalir = False
    End Sub

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If chkTodosProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            If mintCodProveedor = 0 Then
                MsgBox("Proporcione un Proveedor, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                dbcProveedores.Focus()
                Exit Function
            End If
        End If
        Do While (sglTiempoCambio) <= 2.1
        Loop
        System.Windows.Forms.Application.DoEvents()
        If dtpFechaInicial.Value > dtpFechaFinal.Value Then
            MsgBox("La Fecha Inicial no Puede ser Mayor que la Fecha Final.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaInicial.Focus()
            Exit Function
        End If
        If dtpFechaInicial.Value > Now Then
            MsgBox("la Fecha Inicial no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaInicial.Focus()
            Exit Function
        End If
        If dtpFechaFinal.Value > Now Then
            MsgBox("la Fecha Final no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaFinal.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub cboOrdenado_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboOrdenado.Enter
        Pon_Tool()
    End Sub

    Private Sub chkArticulosMovimiento_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkArticulosMovimiento.Enter
        Pon_Tool()
    End Sub

    Private Sub chkDescendente1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDescendente1.Enter
        Pon_Tool()
    End Sub

    Private Sub chkDescendente2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDescendente2.Enter
        Pon_Tool()
    End Sub

    Private Sub chkMostrarCL_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkMostrarCL.Enter
        Pon_Tool()
    End Sub

    Private Sub chkTodosProveedores_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodosProveedores.CheckStateChanged
        If chkTodosProveedores.CheckState = System.Windows.Forms.CheckState.Checked Then
            mblnFueraChange = True
            dbcProveedores.Text = "[TODOS LOS PROVEEDORES...]"
            dbcProveedores.Enabled = False
            mblnFueraChange = False
            mintCodProveedor = 0
        ElseIf chkTodosProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            mblnFueraChange = True
            dbcProveedores.Enabled = True
            dbcProveedores.Text = ""
            mblnFueraChange = False
        End If
    End Sub

    Private Sub chkTodosProveedores_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodosProveedores.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcProveedores_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedores.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String
        If mblnFueraChange Then Exit Sub
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcProveedores.Name Then Exit Sub
        lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedores.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcProveedores)
        If Trim(Me.dbcProveedores.Text) = "" Then
            mintCodProveedor = 0
        End If
        If dbcProveedores.SelectedItem <> "" Then
            Call dbcProveedores_Leave(dbcProveedores, New System.EventArgs())
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedores_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedores.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcProveedores.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' ORDER BY descProvAcreed"
        ModDCombo.DCGotFocus(gStrSql, dbcProveedores)
    End Sub

    Private Sub dbcProveedores_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcProveedores.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodosProveedores.Focus()
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcProveedores_KeyPresss(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcProveedores.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcProveedores_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedores.Leave
        Dim I As Integer
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedores.Text) & "%'"
        Aux = mintCodProveedor
        mintCodProveedor = 0
        ModDCombo.DCLostFocus(dbcProveedores, gStrSql, mintCodProveedor)
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
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Click
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaInicial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaInicial.KeyPress
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub frmVtasVentasyExistenciasporProveedor_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasVentasyExistenciasporProveedor_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasVentasyExistenciasporProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "chkTodosProveedores" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmVtasVentasyExistenciasporProveedor_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        gStrSql = "Select * From CatUsuarios (Nolock) Where CodUsuario = " & gIntCodUsuario
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Tipo").Value = C_TADMIN Then
                MostrarCostoL = True
            Else
                MostrarCostoL = False
            End If
        End If
        Nuevo()

    End Sub

    Private Sub frmVtasVentasyExistenciasporProveedor_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        'If mblnSalir Then
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

    Private Sub frmVtasVentasyExistenciasporProveedor_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        'Me = Nothing
    End Sub

    Private Sub optPiezasVenta_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPiezasVenta.CheckedChanged
        If eventSender.Checked Then
            cboOrdenado.Enabled = False
            chkDescendente1.Enabled = False
            chkDescendente2.Enabled = True
            cboOrdenado.SelectedIndex = 0
            chkDescendente1.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkDescendente2.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If
    End Sub

    Private Sub optPiezasVenta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPiezasVenta.Enter
        Pon_Tool()
    End Sub

    Private Sub optProveedor_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optProveedor.CheckedChanged
        If eventSender.Checked Then
            cboOrdenado.Enabled = True
            chkDescendente1.Enabled = True
            chkDescendente2.Enabled = False
            cboOrdenado.SelectedIndex = 0
            chkDescendente1.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkDescendente2.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If
    End Sub

    Private Sub optProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optProveedor.Enter
        Pon_Tool()
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class