Option Strict Off
Option Explicit On
Imports Microsoft.Office.Interop
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVentasPorGrupo
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE VENTAS POR GRUPO                                                                  *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      LUNES 24 DE MAYO DE 2004                                                                     *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents flexVentas As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents flexUtilidad As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents chkDescendente As System.Windows.Forms.CheckBox
    Public WithEvents optTotalPiezas As System.Windows.Forms.RadioButton
    Public WithEvents optTotalGlobal As System.Windows.Forms.RadioButton
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents optVarios As System.Windows.Forms.RadioButton
    Public WithEvents optRelojeria As System.Windows.Forms.RadioButton
    Public WithEvents optJoyeria As System.Windows.Forms.RadioButton
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox

    Dim mblnSalir As Boolean
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    Dim RsAux As ADODB.Recordset
    Dim rsVentas As ADODB.Recordset
    Dim ObjExcel As Object
    Dim objLibro As Excel.Workbook
    Dim objHoja As Excel.Worksheet
    Dim I As Integer
    Dim Renglon As Integer
    Dim Columna As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim cmd As ADODB.Command


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.chkDescendente = New System.Windows.Forms.CheckBox()
        Me.optTotalPiezas = New System.Windows.Forms.RadioButton()
        Me.optTotalGlobal = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optVarios = New System.Windows.Forms.RadioButton()
        Me.optRelojeria = New System.Windows.Forms.RadioButton()
        Me.optJoyeria = New System.Windows.Forms.RadioButton()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame5.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
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
        Me.txtMensaje.Size = New System.Drawing.Size(298, 82)
        Me.txtMensaje.TabIndex = 17
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'chkDescendente
        '
        Me.chkDescendente.BackColor = System.Drawing.SystemColors.Control
        Me.chkDescendente.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDescendente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDescendente.Location = New System.Drawing.Point(166, 28)
        Me.chkDescendente.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDescendente.Name = "chkDescendente"
        Me.chkDescendente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDescendente.Size = New System.Drawing.Size(94, 28)
        Me.chkDescendente.TabIndex = 16
        Me.chkDescendente.Text = "Descendente"
        Me.ToolTip1.SetToolTip(Me.chkDescendente, "Ordenamiento Descendente")
        Me.chkDescendente.UseVisualStyleBackColor = False
        '
        'optTotalPiezas
        '
        Me.optTotalPiezas.BackColor = System.Drawing.SystemColors.Control
        Me.optTotalPiezas.Cursor = System.Windows.Forms.Cursors.Default
        Me.optTotalPiezas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTotalPiezas.Location = New System.Drawing.Point(56, 38)
        Me.optTotalPiezas.Margin = New System.Windows.Forms.Padding(2)
        Me.optTotalPiezas.Name = "optTotalPiezas"
        Me.optTotalPiezas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optTotalPiezas.Size = New System.Drawing.Size(97, 17)
        Me.optTotalPiezas.TabIndex = 15
        Me.optTotalPiezas.TabStop = True
        Me.optTotalPiezas.Text = "Piezas"
        Me.ToolTip1.SetToolTip(Me.optTotalPiezas, "Ordenado por el Total Global de las Piezas")
        Me.optTotalPiezas.UseVisualStyleBackColor = False
        '
        'optTotalGlobal
        '
        Me.optTotalGlobal.BackColor = System.Drawing.SystemColors.Control
        Me.optTotalGlobal.Checked = True
        Me.optTotalGlobal.Cursor = System.Windows.Forms.Cursors.Default
        Me.optTotalGlobal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTotalGlobal.Location = New System.Drawing.Point(56, 18)
        Me.optTotalGlobal.Margin = New System.Windows.Forms.Padding(2)
        Me.optTotalGlobal.Name = "optTotalGlobal"
        Me.optTotalGlobal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optTotalGlobal.Size = New System.Drawing.Size(89, 24)
        Me.optTotalGlobal.TabIndex = 14
        Me.optTotalGlobal.TabStop = True
        Me.optTotalGlobal.Text = "Total Global"
        Me.ToolTip1.SetToolTip(Me.optTotalGlobal, "Ordenado por el Total Global de las Ventas")
        Me.optTotalGlobal.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Checked = True
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(56, 18)
        Me.optPesos.Margin = New System.Windows.Forms.Padding(2)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(73, 17)
        Me.optPesos.TabIndex = 12
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me.optPesos, "Muestra los Importes en Pesos")
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(166, 18)
        Me.optDolares.Margin = New System.Windows.Forms.Padding(2)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(73, 17)
        Me.optDolares.TabIndex = 13
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me.optDolares, "Muestra los Importes en Dólares")
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optVarios
        '
        Me.optVarios.BackColor = System.Drawing.SystemColors.Control
        Me.optVarios.Cursor = System.Windows.Forms.Cursors.Default
        Me.optVarios.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optVarios.Location = New System.Drawing.Point(198, 20)
        Me.optVarios.Margin = New System.Windows.Forms.Padding(2)
        Me.optVarios.Name = "optVarios"
        Me.optVarios.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optVarios.Size = New System.Drawing.Size(67, 22)
        Me.optVarios.TabIndex = 7
        Me.optVarios.TabStop = True
        Me.optVarios.Text = "Varios"
        Me.ToolTip1.SetToolTip(Me.optVarios, "Muestra las Ventas de Varios")
        Me.optVarios.UseVisualStyleBackColor = False
        '
        'optRelojeria
        '
        Me.optRelojeria.BackColor = System.Drawing.SystemColors.Control
        Me.optRelojeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.optRelojeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optRelojeria.Location = New System.Drawing.Point(108, 20)
        Me.optRelojeria.Margin = New System.Windows.Forms.Padding(2)
        Me.optRelojeria.Name = "optRelojeria"
        Me.optRelojeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optRelojeria.Size = New System.Drawing.Size(78, 22)
        Me.optRelojeria.TabIndex = 6
        Me.optRelojeria.TabStop = True
        Me.optRelojeria.Text = "Relojeria"
        Me.ToolTip1.SetToolTip(Me.optRelojeria, "Muestra las Ventas de Relojeria")
        Me.optRelojeria.UseVisualStyleBackColor = False
        '
        'optJoyeria
        '
        Me.optJoyeria.BackColor = System.Drawing.SystemColors.Control
        Me.optJoyeria.Checked = True
        Me.optJoyeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.optJoyeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optJoyeria.Location = New System.Drawing.Point(24, 20)
        Me.optJoyeria.Margin = New System.Windows.Forms.Padding(2)
        Me.optJoyeria.Name = "optJoyeria"
        Me.optJoyeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optJoyeria.Size = New System.Drawing.Size(89, 22)
        Me.optJoyeria.TabIndex = 5
        Me.optJoyeria.TabStop = True
        Me.optJoyeria.Text = "Joyería"
        Me.ToolTip1.SetToolTip(Me.optJoyeria, "Muestra las Ventas de Joyeria")
        Me.optJoyeria.UseVisualStyleBackColor = False
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.txtMensaje)
        Me.Frame5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame5.Location = New System.Drawing.Point(9, 262)
        Me.Frame5.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(319, 99)
        Me.Frame5.TabIndex = 4
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Texto Adicional"
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.chkDescendente)
        Me.Frame4.Controls.Add(Me.optTotalPiezas)
        Me.Frame4.Controls.Add(Me.optTotalGlobal)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(6, 162)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(265, 68)
        Me.Frame4.TabIndex = 3
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Ordenado Por"
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.optPesos)
        Me.Frame3.Controls.Add(Me.optDolares)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(6, 110)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(265, 46)
        Me.Frame3.TabIndex = 2
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Moneda"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.dtpFechaInicial)
        Me.Frame2.Controls.Add(Me.dtpFechaFinal)
        Me.Frame2.Controls.Add(Me.Label2)
        Me.Frame2.Controls.Add(Me.Label3)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(6, 58)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(322, 46)
        Me.Frame2.TabIndex = 1
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(56, 17)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaInicial.TabIndex = 9
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(211, 17)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(96, 20)
        Me.dtpFechaFinal.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(12, 20)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(47, 18)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Desde"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(173, 19)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(45, 18)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Hasta"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optVarios)
        Me.Frame1.Controls.Add(Me.optRelojeria)
        Me.Frame1.Controls.Add(Me.optJoyeria)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(6, 6)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(274, 46)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Grupo"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(125, 381)
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
        Me.btnImprimir.Location = New System.Drawing.Point(10, 381)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 78
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(240, 382)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 77
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmVentasPorGrupo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(371, 429)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame5)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(374, 125)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVentasPorGrupo"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Ventas Salida de Mercancia por Grupo"
        Me.Frame5.ResumeLayout(False)
        Me.Frame5.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Sub CalculaTotales()
        Dim I As Integer
        Dim TotalVentas As Decimal
        Dim TotalPiezas As Decimal
        If rsVentas.RecordCount > 0 Then
            rsVentas.MoveFirst()
        End If
        flexVentas.Clear()
        flexVentas.Rows = 2
        flexVentas.set_Cols(0, rsVentas.Fields.Count - 3)
        flexVentas.Col = 0
        flexVentas.Row = 1
        TotalVentas = 0
        TotalPiezas = 0
        Do While Not rsVentas.EOF
            For I = 3 To rsVentas.Fields.Count - 1 Step 2
                flexVentas.set_TextMatrix(flexVentas.Row, I - 3, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, I - 3))) + System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO))
                flexVentas.set_TextMatrix(flexVentas.Row, I - 2, CDec(Numerico(flexVentas.get_TextMatrix(flexVentas.Row, I - 2))) + System.Math.Round(rsVentas.Fields(I + 1).Value, C_REDONDEO))
            Next
            rsVentas.MoveNext()
        Loop
        For I = 0 To flexVentas.get_Cols() - 3 Step 2
            TotalVentas = TotalVentas + CDec(Numerico(flexVentas.get_TextMatrix(1, I)))
        Next
        flexVentas.set_TextMatrix(1, flexVentas.get_Cols() - 2, TotalVentas)
        If rsVentas.RecordCount > 0 Then
            rsVentas.MoveFirst()
        End If
        flexVentas.Row = 1
        flexVentas.Col = 0
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
        Dim Sql As String
        Dim Total As String
        Dim TotalPiezas As String
        Total = ""
        TotalPiezas = ""
        Sql = "Select Vta.CodGrupo,G.DescGrupo,Vta.DescFamilia"
        gStrSql = "Select CodAlmacen,DescAlmacen From CatAlmacen Where TipoAlmacen = 'P' Order By CodAlmacen"
        ModEstandar.BorraCmd()
        cmd.CommandText = "dbo.UP_Select_Datos"
        cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        cmd.Parameters.Append(cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        cmd.Parameters.Append(cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsAux = cmd.Execute
        If RsAux.RecordCount > 0 Then
            Do While Not RsAux.EOF
                If optPesos.Checked = True Then
                    Sql = Sql & ",Round(Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then (Case When Tipo = 'R' Then Total Else (PrecioReal*(Cantidad-CantidadDev)) + (Case When NumPartida = 1 then Redondeo Else 0 End) End) * TipoCambio Else 0 End),1) as Vta" & RsAux.Fields("CodAlmacen").Value & ""
                    If Trim(Total) = "" Then
                        Total = Total & ",Round(Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then (Case When Tipo = 'R' Then Total Else (PrecioReal*(Cantidad-CantidadDev)) + (Case When NumPartida = 1 then Redondeo Else 0 End) End) * TipoCambio Else 0 End),1) "
                    Else
                        Total = Total & "+ Round(Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then (Case When Tipo = 'R' Then Total Else (PrecioReal*(Cantidad-CantidadDev)) + (Case When NumPartida = 1 then Redondeo Else 0 End) End) * TipoCambio Else 0 End),1) "
                    End If
                ElseIf optDolares.Checked = True Then
                    Sql = Sql & ",Round(Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then Case When Tipo = 'R' Then Total Else (PrecioReal*(Cantidad-CantidadDev)) + (Case When NumPartida = 1 then Redondeo Else 0 End) End Else 0 End),2) as Vta" & RsAux.Fields("CodAlmacen").Value & ""
                    If Trim(Total) = "" Then
                        Total = Total & ",Round(Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then Case When Tipo = 'R' Then Total Else (PrecioReal*(Cantidad-CantidadDev)) + (Case When NumPartida = 1 then Redondeo Else 0 End) End Else 0 End),2) "
                    Else
                        Total = Total & "+ Round(Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then Case When Tipo = 'R' Then Total Else (PrecioReal*(Cantidad-CantidadDev)) + (Case When NumPartida = 1 then Redondeo Else 0 End) End Else 0 End),2) "
                    End If
                End If
                Sql = Sql & ",Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then Case When Tipo = 'R' Then 1 Else (Cantidad-CantidadDev) End Else 0 End) as Piezas" & RsAux.Fields("CodAlmacen").Value & ""
                If Trim(TotalPiezas) = "" Then
                    TotalPiezas = TotalPiezas & ",Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then Case When Tipo = 'R' Then 1 Else (Cantidad-CantidadDev) End Else 0 End) "
                Else
                    TotalPiezas = TotalPiezas & "+ Sum(Case When Vta.CodSucursal = " & RsAux.Fields("CodAlmacen").Value & " then Case When Tipo = 'R' Then 1 Else (Cantidad-CantidadDev) End Else 0 End) "
                End If
                RsAux.MoveNext()
            Loop
            Total = Total & " as Total"
            TotalPiezas = TotalPiezas & " as TotalPiezas "
            Sql = Sql & Total & TotalPiezas
            Sql = Sql & "FROM VENTAS_SALIDAMCIA('" & Format(dtpFechaInicial.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpFechaFinal.Value, C_FORMATFECHAGUARDAR) & "') Vta " & "Inner Join CatGrupos G On Vta.CodGrupo = G.CodGrupo Where Vta.Tipo <> 'R' AND Vta.CodGrupo = " & IIf(optJoyeria.Checked = True, gCODJOYERIA, IIf(optRelojeria.Checked = True, gCODRELOJERIA, gCODVARIOS)) & " " & "Group By Vta.CodGrupo, G.DescGrupo, DescFamilia " & "Order By " & IIf(optTotalGlobal.Checked = True, "Total", "TotalPiezas") & " " & IIf(chkDescendente.CheckState = System.Windows.Forms.CheckState.Checked, "Desc", "")
        End If
        DevuelveQuery = Sql
    End Function

    Sub Encabezado()
        On Error GoTo Err_Renamed
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
            If optJoyeria.Checked = True Then
                .Range("C2").FormulaR1C1 = "Ventas de Joyeria por Linea"
            ElseIf optRelojeria.Checked = True Then
                .Range("C2").FormulaR1C1 = "Ventas de Relojeria por Marca"
            ElseIf optVarios.Checked = True Then
                .Range("C2").FormulaR1C1 = "Ventas de Varios por Familia"
            End If
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
            .Range("I4").FormulaR1C1 = "Fecha: " & Format(Today, "dd/mmm/yyyy")
            .Range("I4:J4").Select()
            .Range("I4:J4").MergeCells = True
            .Range("I4:J4").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("I4:J4").Font
                .Bold = False
                .Size = 9
                .Name = "Arial"
            End With
            .Range("A5").FormulaR1C1 = "Mensaje: "
            .Range("A5").Select()
            .Range("A5").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A5").Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            If Trim(txtMensaje.Text) <> "" Then
                .Range("A5").FormulaR1C1 = .Range("A5").FormulaR1C1 & Trim(QuitaEnter(txtMensaje.Text))
                .Range("A5:J5").Select()
                .Range("A5:J5").MergeCells = True
                .Range("A5:J5").HorizontalAlignment = Excel.Constants.xlLeft
                With .Range("A5:J5").Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
            End If
            .Range("A6:C6").Select()
            .Range("A6:C6").MergeCells = True
            If optPesos.Checked = True Then
                .Range("A6:C6").FormulaR1C1 = "**Los importes estan expresados en pesos"
            ElseIf optDolares.Checked = True Then
                .Range("A6:C6").FormulaR1C1 = "**Los importes estan expresados en dólares"
            End If
            .Range("A6:C6").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A6:C6").Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Renglon = 8
            Columna = 1
            If optJoyeria.Checked = True Then
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Linea"
            ElseIf optRelojeria.Checked = True Then
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Marca"
            ElseIf optVarios.Checked = True Then
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Familia"
            End If
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlTop
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 25
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            Columna = Columna + 1
            gStrSql = "Select CodAlmacen,DescAlmacen From CatAlmacen (Nolock) Where TipoAlmacen = 'P' Order By CodAlmacen"
            ModEstandar.BorraCmd()
            cmd.CommandText = "dbo.UP_Select_Datos"
            cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            cmd.Parameters.Append(cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            cmd.Parameters.Append(cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsAux = cmd.Execute
            If RsAux.RecordCount > 0 Then
                Do While Not RsAux.EOF
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).MergeCells = True
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1))._Default = UCase((RsAux.Fields("DescAlmacen").Value)) & LCase(Mid(RsAux.Fields("DescAlmacen").Value, 2, 39))
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).WrapText = True
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlCenter
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).VerticalAlignment = Excel.Constants.xlTop
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Interior.ColorIndex = 15
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Font
                        .Bold = True
                        .Size = 9
                        .Name = "Arial"
                    End With
                    Columna = Columna + 2
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Piezas"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlTop
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 6
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Bold = True
                        .Size = 9
                        .Name = "Arial"
                    End With
                    RsAux.MoveNext()
                    Columna = Columna + 1
                Loop
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).MergeCells = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1))._Default = "Global"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).VerticalAlignment = Excel.Constants.xlTop
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Interior.ColorIndex = 15
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
                Columna = Columna + 2
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Piezas"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlTop
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
                Columna = Columna + 2
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Precio Prom."
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlTop
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 9
                    .Name = "Arial"
                End With
            End If
        End With
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub EnviaExcel()
        Dim Archivo As String
        On Error GoTo Err_Renamed
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        System.Windows.Forms.Application.DoEvents()
        If Dir(gstrCorpoDriveLocal & "\Sistema\", FileAttribute.Directory + FileAttribute.Hidden) = "" Then
            MsgBox("No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        Archivo = IIf(optJoyeria.Checked = True, "VJ", IIf(optRelojeria.Checked = True, "VR", "VV")) & CStr(Format(Month(Today), "00")) & CStr(Format((Today), "00")) & (CStr(Format(Year(Today), "00"))) & ".xls"
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
        If optJoyeria.Checked = True Then
            objLibro.ActiveSheet.Name = "Ventas de Joyeria por Linea"
        ElseIf optRelojeria.Checked = True Then
            objLibro.ActiveSheet.Name = "Ventas de Relojeria por Marca"
        ElseIf optVarios.Checked = True Then
            objLibro.ActiveSheet.Name = "Ventas de Varios por Familia"
        End If
        Encabezado()
        LlenaDatos()
        objLibro.SaveAs(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        System.Windows.Forms.Application.DoEvents()
        MsgBox("Se ha creado el archivo " & gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        '    Case vbYes:
        '        ObjExcel.Visible = True
        '        Set ObjExcel = Nothing
        '        Set objLibro = Nothing
        '        Set objHoja = Nothing
        '    Case vbNo Or vbCancel:
        CierraInstanciasdeExcel(1)
        'End Select
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

    Sub Imprime()
        On Error GoTo Err_Renamed
        Dim Query As String

        If Not ValidaDatos() Then Exit Sub
        Query = DevuelveQuery()
        ModEstandar.BorraCmd()
        cmd.CommandTimeout = 300
        cmd.CommandText = "dbo.Up_Select_Datos"
        cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        cmd.Parameters.Append(cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        cmd.Parameters.Append(cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Query))
        rsVentas = cmd.Execute
        If rsVentas.RecordCount > 0 Then
            CalculaTotales()
            EnviaExcel()
        Else
            MsgBox("No existe información por mostrar en este periodo, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        End If
        cmd.CommandTimeout = 90

Err_Renamed:
        If Err.Number <> 0 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Sub

    Sub Limpiar()
        Nuevo()
        optJoyeria.Focus()
    End Sub

    Sub LlenaDatos()
        On Error GoTo Err_Renamed
        Dim Total As Decimal
        Dim TotPiezas As Integer
        Dim Porcentaje As Decimal
        Dim PrecioProm As Decimal
        With objHoja
            If rsVentas.RecordCount > 0 Then
                rsVentas.MoveFirst()
            End If
            Renglon = Renglon + 1
            Do While Not rsVentas.EOF
                Columna = 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Trim(rsVentas.Fields("DescFamilia").Value)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous

                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                Total = 0
                TotPiezas = 0
                For I = 3 To rsVentas.Fields.Count - 3 Step 2
                    Total = Total + System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO)
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO)
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 12.29
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    If CDec(Numerico(flexVentas.get_TextMatrix(1, I - 3))) <> 0 Then
                        Porcentaje = System.Math.Round((System.Math.Round(rsVentas.Fields(I).Value, C_REDONDEO) / CDec(Numerico(flexVentas.get_TextMatrix(1, I - 3)))) * 100, 2)
                    Else
                        Porcentaje = 0
                    End If
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).FormulaR1C1 = VB6.Format(Porcentaje, "###,##0.00") & "%"
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).ColumnWidth = 6.71
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    '.Range(.Cells(Renglon, Columna + 1), .Cells(Renglon, Columna + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    Columna = Columna + 2
                    TotPiezas = TotPiezas + System.Math.Round(rsVentas.Fields(I + 1).Value, C_REDONDEO)
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = rsVentas.Fields(I + 1).Value
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 6
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    Columna = Columna + 1
                Next
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Total
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 12.29
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                If CDec(Numerico(flexVentas.get_TextMatrix(1, flexVentas.get_Cols() - 2))) <> 0 Then
                    Porcentaje = System.Math.Round((Total / CDec(Numerico(flexVentas.get_TextMatrix(1, flexVentas.get_Cols() - 2)))) * 100, 2)
                Else
                    Porcentaje = 0
                End If
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).FormulaR1C1 = VB6.Format(Porcentaje, "###,##0.00") & "%"
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).ColumnWidth = 6.71
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                '.Range(.Cells(Renglon, Columna + 1), .Cells(Renglon, Columna + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 2
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = TotPiezas
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 6
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                If TotPiezas <> 0 Then
                    PrecioProm = System.Math.Round(Total / TotPiezas, C_REDONDEO)
                Else
                    PrecioProm = 0
                End If
                Columna = Columna + 2
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = PrecioProm
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 9
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                Renglon = Renglon + 1
                rsVentas.MoveNext()
            Loop
            Columna = 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Total"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Columna = Columna + 1
            For I = 0 To flexVentas.get_Cols() - 3 Step 2
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = flexVentas.get_TextMatrix(1, I)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Interior.ColorIndex = 15
                .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2))._Default = flexVentas.get_TextMatrix(1, I + 1)
                .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).Interior.ColorIndex = 15
                .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 3
            Next
            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = flexVentas.get_TextMatrix(1, flexVentas.get_Cols() - 2)
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(Renglon, Columna + 1), .Cells._Default(Renglon, Columna + 1)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2))._Default = flexVentas.get_TextMatrix(1, flexVentas.get_Cols() - 1)
            .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(Renglon, Columna + 2), .Cells._Default(Renglon, Columna + 2)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Columna = Columna + 4
            If CDec(Numerico(flexVentas.get_TextMatrix(1, flexVentas.get_Cols() - 1))) <> 0 Then
                PrecioProm = System.Math.Round(CDec(Numerico(flexVentas.get_TextMatrix(1, flexVentas.get_Cols() - 2))) / CDec(Numerico(flexVentas.get_TextMatrix(1, flexVentas.get_Cols() - 1))), C_REDONDEO)
            Else
                PrecioProm = 0
            End If
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = PrecioProm
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Renglon = Renglon + 5
            'Seguimos con la Grafica
            If rsVentas.RecordCount > 0 Then
                rsVentas.MoveFirst()
            End If
            I = 0
            Do While Not rsVentas.EOF
                I = I + 1
                .Range("A" & 1000 + I)._Default = Trim(rsVentas.Fields("DescFamilia").Value)
                .Range("B" & 1000 + I).FormulaR1C1 = System.Math.Round(rsVentas.Fields("Total").Value, C_REDONDEO)
                rsVentas.MoveNext()
            Loop
            .Range("B" & (1000 + (I + 1))).Select()
            .Range("B" & (1000 + (I + 1))).FormulaR1C1 = "SUM(R[-" & I & "]C:R[-1]C)"
            .Range("D" & Renglon + 15).Select()
            .Application.Charts.Add()
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPie
            .Application.ActiveChart.ChartType = Excel.XlChartType.xl3DPieExploded
            .Application.ActiveChart.SetSourceData(Source:= .Range("A1001:B" & 1000 + I), PlotBy:=Excel.XlRowCol.xlColumns)
            If optJoyeria.Checked = True Then
                .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Ventas de Joyeria por Linea")
            ElseIf optRelojeria.Checked = True Then
                .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Ventas de Relojeria por Marca")
            ElseIf optVarios.Checked = True Then
                .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Ventas de Varios por Familia")
            End If
            With .Application.ActiveChart
                .HasTitle = True
                If optJoyeria.Checked = True Then
                    .ChartTitle.Characters.Text = "Venta Global Joyeria"
                ElseIf optRelojeria.Checked = True Then
                    .ChartTitle.Characters.Text = "Venta Global Relojeria"
                ElseIf optVarios.Checked = True Then
                    .ChartTitle.Characters.Text = "Venta Global Varios"
                End If
            End With
            .Application.ActiveSheet.Shapes("Chart 1").ScaleHeight(1.46, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft)
            .Application.ActiveSheet.Shapes("Chart 1").ScaleWidth(2.15, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromBottomRight)
            '''        If optJoyeria.Value = True Then
            '''            .Application.ActiveSheet.Shapes("Chart 1").ScaleHeight 1.46, msoFalse, msoScaleFromTopLeft
            '''            .Application.ActiveSheet.Shapes("Chart 1").ScaleWidth 2.15, msoFalse, msoScaleFromBottomRight
            '''        End If
            .Application.ActiveChart.PlotArea.Interior.ColorIndex = 2
            .Application.ActiveChart.PlotArea.Border.LineStyle = Excel.Constants.xlNone
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            With .Application.ActiveChart.SeriesCollection(1).DataLabels.Font
                .Size = 8
                .Name = "Arial"
            End With
            .Application.ActiveChart.ChartArea.Select()
            .Application.ActiveChart.HasLegend = True
            .Application.ActiveChart.Legend.Select()
            .Application.Selection.Position = Excel.Constants.xlRight
            .Application.ActiveChart.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, LegendKey:=False, HasLeaderLines:=True)
            .Application.ActiveChart.ShowWindow = True
            .Application.ActiveChart.ShowWindow = False
            .Application.Selection.AutoScaleFont = True
            With .Application.Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .Underline = False
            End With
            If optJoyeria.Checked = True Then
                .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Ventas de Joyeria por Linea")
            ElseIf optRelojeria.Checked = True Then
                .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Ventas de Relojeria por Marca")
            ElseIf optVarios.Checked = True Then
                .Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:="Ventas de Varios por Familia")
            End If
            .Application.Selection.Top = 1
            .Application.Selection.Left = 620
            .Application.Selection.Height = 350
            .Application.Selection.Width = 160
            .Application.ActiveWindow.Visible = False
            .Application.ActiveWindow.Zoom = 80
            .Range("A1001:B" & 1000 + ((I + 1))).Font.ColorIndex = 2
            .Range("A1").Select()
        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Nuevo()
        optJoyeria.Checked = True
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
        optPesos.Checked = True
        optDolares.Checked = False
        optTotalGlobal.Checked = True
        chkDescendente.CheckState = System.Windows.Forms.CheckState.Unchecked
        txtMensaje.Text = ""
        mblnSalir = False
    End Sub

    Function ValidaDatos() As Boolean
        ValidaDatos = False
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

    Private Sub chkDescendente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDescendente.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaFinal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaFinal.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaFinal.Click
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaFinal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaFinal.KeyPress
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInicial.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInicial.Click
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaInicial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaInicial.KeyPress
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub frmVentasPorGrupo_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub
    Private Sub frmVentasPorGrupo_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVentasPorGrupo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optJoyeria" And System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optRelojeria" And System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optVarios" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmVentasPorGrupo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVentasPorGrupo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Nuevo()
    End Sub

    Private Sub frmVentasPorGrupo_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        If mblnSalir Then
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSalir = False
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVentasPorGrupo_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub optJoyeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optJoyeria.Enter
        Pon_Tool()
    End Sub

    Private Sub optPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPesos.Enter
        Pon_Tool()
    End Sub

    Private Sub optRelojeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRelojeria.Enter
        Pon_Tool()
    End Sub

    Private Sub optTotalGlobal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTotalGlobal.Enter
        Pon_Tool()
    End Sub

    Private Sub optTotalPiezas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTotalPiezas.Enter
        Pon_Tool()
    End Sub

    Private Sub optVarios_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optVarios.Enter
        Pon_Tool()
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