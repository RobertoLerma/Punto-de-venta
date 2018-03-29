Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmVtasRPTVentasSalidadeMercanciaPorCliente
    Inherits System.Windows.Forms.Form
    ''' ****************************************************************************************************************************************************'
    ''' MODIFIC.- PROBLEMA DE MANEJO CON COMBOS PARA DESCRIPCIONES IGUALES - SE CAMBIO COMBO POR TXT
    ''' 03MAR2008 - MAVF
    '*******************************************************************************************************************************************************'


    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"
    Const C_NINGUNA As String = "[ Vacío ... ]"

    Const ColorGris As Integer = &H8000000F '''03MAR2008 - MAVF
    Const ColorAmarillo As Integer = &HC0FFFF '''03MAR2008 - MAVF
    Const ColorBlanco As Integer = &HFFFFFF '''03MAR2008 - MAVF

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean
    Dim mblnFueraChange As Boolean
    Dim mintCodSucursal As Integer
    Dim mintCodCliente As Integer
    Dim tecla As Integer
    Dim cTablaTmp As String
    Dim mblnSalir As Boolean

    Public gBlnFueraChange As Boolean '''03MAR2008 - MAVF
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkClientes As System.Windows.Forms.CheckBox
    Public WithEvents txtCodCliente As System.Windows.Forms.TextBox
    Public WithEvents txtNombre As System.Windows.Forms.TextBox
    Public WithEvents chkTodas As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_0 As System.Windows.Forms.GroupBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_2 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_1 As System.Windows.Forms.GroupBox
    Public WithEvents chkImpuesto As System.Windows.Forms.CheckBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents _lblCliente_5 As System.Windows.Forms.Label
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblCliente As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Friend WithEvents btnBuscar As Button
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public strControlActual As String 'Nombre del control actual
    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendrá el string del tag que se le mandara al fromulario de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas


        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        If strControlActual = "TXTNOMBRE" Then
            strCaptionForm = "Consulta de Clientes"
            gStrSql = "Select Right('00000' + ltrim(rtrim(CodCliente)),5) as Codigo, DescCliente as Nombre From CatClientes (Nolock) Where DescCliente Like '" & Trim(txtNombre.Text) & "%' Order by DescCliente "
        End If

        strSQL = gStrSql 'Se hace uso de una variable temporal para el query
        gStrSql = strSQL 'Se regresa el valor de la variable temporal a la variable original

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsgBox(C_msgSINDATOS & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta
        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTNOMBRE"
                    .set_ColWidth(0, 0, 1000) 'Columna del Código
                    .set_ColWidth(1, 0, 6000) 'Columna de la Descripción
                    .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            End Select
        End With
        FrmConsultas.ShowDialog()

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub
    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCodCliente = New System.Windows.Forms.TextBox()
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.chkClientes = New System.Windows.Forms.CheckBox()
        Me._fraVtas_0 = New System.Windows.Forms.GroupBox()
        Me.chkTodas = New System.Windows.Forms.CheckBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me._fraVtas_1 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me.chkImpuesto = New System.Windows.Forms.CheckBox()
        Me._lblCliente_5 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.fraVtas = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblCliente = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me._fraVtas_0.SuspendLayout()
        Me._fraVtas_1.SuspendLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCliente, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCodCliente
        '
        Me.txtCodCliente.AcceptsReturn = True
        Me.txtCodCliente.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodCliente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodCliente.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodCliente.Location = New System.Drawing.Point(64, 88)
        Me.txtCodCliente.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodCliente.MaxLength = 5
        Me.txtCodCliente.Name = "txtCodCliente"
        Me.txtCodCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodCliente.Size = New System.Drawing.Size(51, 20)
        Me.txtCodCliente.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtCodCliente, "Código del Cliente")
        '
        'txtNombre
        '
        Me.txtNombre.AcceptsReturn = True
        Me.txtNombre.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtNombre.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombre.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombre.Location = New System.Drawing.Point(119, 88)
        Me.txtNombre.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNombre.MaxLength = 0
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombre.Size = New System.Drawing.Size(198, 20)
        Me.txtNombre.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtNombre, "F3 - Nombre Cliente")
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(9, 230)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(325, 80)
        Me.txtMensaje.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'chkClientes
        '
        Me.chkClientes.BackColor = System.Drawing.SystemColors.Control
        Me.chkClientes.Checked = True
        Me.chkClientes.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkClientes.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkClientes.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkClientes.Location = New System.Drawing.Point(14, 68)
        Me.chkClientes.Margin = New System.Windows.Forms.Padding(2)
        Me.chkClientes.Name = "chkClientes"
        Me.chkClientes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkClientes.Size = New System.Drawing.Size(172, 20)
        Me.chkClientes.TabIndex = 4
        Me.chkClientes.Text = "Todos los clientes"
        Me.chkClientes.UseVisualStyleBackColor = False
        '
        '_fraVtas_0
        '
        Me._fraVtas_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_0.Controls.Add(Me.chkTodas)
        Me._fraVtas_0.Controls.Add(Me.dbcSucursal)
        Me._fraVtas_0.Controls.Add(Me._lblVentas_0)
        Me._fraVtas_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_0.Location = New System.Drawing.Point(6, 6)
        Me._fraVtas_0.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.Name = "_fraVtas_0"
        Me._fraVtas_0.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_0.Size = New System.Drawing.Size(328, 58)
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
        Me.chkTodas.Size = New System.Drawing.Size(137, 17)
        Me.chkTodas.TabIndex = 1
        Me.chkTodas.Text = "Todas las sucursales"
        Me.chkTodas.UseVisualStyleBackColor = False
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(75, 20)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(218, 21)
        Me.dbcSucursal.TabIndex = 3
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_0.Location = New System.Drawing.Point(20, 22)
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
        Me._fraVtas_1.Location = New System.Drawing.Point(6, 115)
        Me._fraVtas_1.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.Name = "_fraVtas_1"
        Me._fraVtas_1.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_1.Size = New System.Drawing.Size(328, 46)
        Me._fraVtas_1.TabIndex = 8
        Me._fraVtas_1.TabStop = False
        Me._fraVtas_1.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(64, 17)
        Me.dtpDesde.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(96, 20)
        Me.dtpDesde.TabIndex = 10
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(225, 17)
        Me.dtpHasta.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(95, 20)
        Me.dtpHasta.TabIndex = 12
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_1.Location = New System.Drawing.Point(9, 21)
        Me._lblVentas_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(52, 13)
        Me._lblVentas_1.TabIndex = 9
        Me._lblVentas_1.Text = "Desde el "
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_2.Location = New System.Drawing.Point(175, 21)
        Me._lblVentas_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(46, 13)
        Me._lblVentas_2.TabIndex = 11
        Me._lblVentas_2.Text = "Hasta el"
        '
        'chkImpuesto
        '
        Me.chkImpuesto.BackColor = System.Drawing.SystemColors.Control
        Me.chkImpuesto.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkImpuesto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkImpuesto.Location = New System.Drawing.Point(8, 171)
        Me.chkImpuesto.Margin = New System.Windows.Forms.Padding(2)
        Me.chkImpuesto.Name = "chkImpuesto"
        Me.chkImpuesto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkImpuesto.Size = New System.Drawing.Size(116, 24)
        Me.chkImpuesto.TabIndex = 13
        Me.chkImpuesto.Text = "Incluir Impuesto"
        Me.chkImpuesto.UseVisualStyleBackColor = False
        '
        '_lblCliente_5
        '
        Me._lblCliente_5.AutoSize = True
        Me._lblCliente_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblCliente_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblCliente_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblCliente_5.Location = New System.Drawing.Point(15, 90)
        Me._lblCliente_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblCliente_5.Name = "_lblCliente_5"
        Me._lblCliente_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblCliente_5.Size = New System.Drawing.Size(45, 13)
        Me._lblCliente_5.TabIndex = 5
        Me._lblCliente_5.Text = "Cliente :"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblRpt_2.Location = New System.Drawing.Point(6, 206)
        Me._lblRpt_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 14
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(241, 333)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 68
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(126, 332)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 70
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(11, 332)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 69
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasRPTVentasSalidadeMercanciaPorCliente
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(361, 383)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.chkClientes)
        Me.Controls.Add(Me.txtCodCliente)
        Me.Controls.Add(Me.txtNombre)
        Me.Controls.Add(Me._fraVtas_0)
        Me.Controls.Add(Me._fraVtas_1)
        Me.Controls.Add(Me.chkImpuesto)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me._lblCliente_5)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(194, 181)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasRPTVentasSalidadeMercanciaPorCliente"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Ventas por Cliente"
        Me._fraVtas_0.ResumeLayout(False)
        Me._fraVtas_0.PerformLayout()
        Me._fraVtas_1.ResumeLayout(False)
        Me._fraVtas_1.PerformLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCliente, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Public Sub Limpiar()
        On Error Resume Next
        Call Nuevo()
        chkTodas.Focus()
    End Sub

    Public Sub Nuevo()
        chkTodas.CheckState = System.Windows.Forms.CheckState.Checked
        chkTodas_CheckStateChanged(chkTodas, New System.EventArgs())
        chkClientes.CheckState = System.Windows.Forms.CheckState.Checked
        chkClientes_CheckStateChanged(chkClientes, New System.EventArgs())
        dtpDesde.Value = Format(Today, "dd/MMM/yyyy")
        dtpHasta.Value = Format(Today, "dd/MMM/yyyy")
        chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked
        txtMensaje.Text = ""
        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
        gBlnFueraChange = False
    End Sub

    ''' SE AGREGO EL CAMPO CODIGO DEL ARTICULO (Abc de articulos) AL QUERY Y AL REPORTE
    ''' 10OCT2006
    Function DevuelveQuery() As String
        On Error GoTo Err_Renamed
        Dim Sql As String

        Sql = "SELECT VTA.CodSucursal, CA.DescAlmacen, VTA.CodCliente, VTA.Nombre, VTA.FolioVenta, VTA.FechaVenta, Vta.CodArticulo, Vta.DescArticulo, IMPFOLIO.importe, ImpVta.ImpVta " & "FROM DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') VTA " & "INNER JOIN (SELECT * FROM CatAlmacen WHERE TipoAlmacen = 'P') CA ON VTA.CodSucursal = CA.CodAlmacen " & "INNER JOIN (SELECT CodSucursal,CodCliente,FolioVenta," & IIf(chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "SUM(ROUND(PrecioReal * (Cantidad - CantidadDev) + CASE WHEN NumPartida = 1 THEN Redondeo ELSE 0 END,2)) AS Importe ", "SUM(ROUND((PrecioListaSinIva - Descuento) * (Cantidad - CantidadDev) + CASE WHEN NumPartida = 1 THEN Redondeo ELSE 0 END,2)) AS Importe ") & "FROM DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') " & "GROUP BY CodSucursal,CodCliente,FolioVenta) IMPFOLIO " & "ON VTA.CodSucursal = IMPFOLIO.CodSucursal AND VTA.CodCliente = IMPFOLIO.CodCliente AND VTA.FolioVenta = IMPFOLIO.FolioVenta " & "INNER JOIN (SELECT CodSucursal,CodCliente," & IIf(chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "SUM(ROUND(PrecioReal * (Cantidad - CantidadDev) + CASE WHEN NumPartida = 1 THEN Redondeo ELSE 0 END,2)) AS ImpVta ", "SUM(ROUND((PrecioListaSinIva - Descuento) * (Cantidad - CantidadDev) + CASE WHEN NumPartida = 1 THEN Redondeo ELSE 0 END,2)) AS ImpVta ") & "FROM DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') " & "GROUP BY CodSucursal,CodCliente) IMPVTA ON VTA.CodSucursal = IMPVTA.CodSucursal AND VTA.CodCliente = IMPVTA.CodCliente " & "WHERE (Cantidad - CantidadDev) > 0 " & IIf(mintCodSucursal <> 0, "AND VTA.CodSucursal = " & mintCodSucursal & " ", "") & IIf(mintCodCliente <> 0, "AND VTA.CodCliente = " & mintCodCliente & " ", "") & "ORDER BY VTA.CodSucursal ASC,IMPVTA.ImpVta DESC,VTA.FolioVenta DESC"
        DevuelveQuery = Sql

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub Imprime()
        Dim rptVentasSalidaDeMercanciaPorCliente As New rptVentasSalidaDeMercanciaPorCliente

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr
        Dim lStrSql As String
        'Declarar vectores para almacenar los parámetros que se le enviarán al reporte
        Dim aParam(5) As Object
        Dim aValues(5) As Object

        If Not ValidaDatos() Then
            Exit Sub
        End If

        lStrSql = DevuelveQuery()
        gStrSql = lStrSql
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
            rptVentasSalidaDeMercanciaPorCliente.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "Mensaje"
        'aValues(1) = Trim(Me.txtMensaje.Text)
        'aParam(2) = "dDesde"
        'aValues(2) = Me.dtpDesde.Value
        'aParam(3) = "dHasta"
        'aValues(3) = Me.dtpHasta.Value
        'aParam(4) = "Empresa"
        'aValues(4) = Trim(gstrNombCortoEmpresa)
        'aParam(5) = "IncluyeImpuestos"
        'aValues(5) = IIf(Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "** Las cantidades expresadas incluyen IVA.", "** Las cantidades expresadas NO incluyen IVA.")


        If (txtMensaje.Text <> Nothing) Then
            pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaPorCliente.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaPorCliente.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        End If

        If (dtpDesde.Value <> Nothing) Then
            pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaPorCliente.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
        End If

        If (dtpHasta.Value <> Nothing) Then
            pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaPorCliente.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
        End If

        If (gstrNombCortoEmpresa <> Nothing) Then
            pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaPorCliente.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
        End If

        'If ("" <> Nothing) Then
        '    pdvNum.Value = "" : pvNum.Add(pdvNum)
        '    rptVentasSalidaDeMercanciaPorCliente.DataDefinition.ParameterFields("MonedaDeCantidades").ApplyCurrentValues(pvNum)
        'End If

        If (chkImpuesto.CheckState <> Nothing) Then
            pdvNum.Value = IIf(Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "** Las cantidades expresadas incluyen IVA.", "** Las cantidades expresadas NO incluyen IVA.") : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaPorCliente.DataDefinition.ParameterFields("IncluyeImpuestos").ApplyCurrentValues(pvNum)
        End If


        frmReportes.reporteActual = rptVentasSalidaDeMercanciaPorCliente 'Es el nombre del archivo que se incluyó en el proyecto
        frmReportes.Show()
        'frmReportes.Imprime(Trim(Me.Text), aParam, aValues)
        Cmd.CommandTimeout = 90

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function ValidaDatos() As Boolean
        If mblnTecleoFechaI Then
            Do While (msglTiempoCambioI) <= 2.1
            Loop
            mblnTecleoFechaI= False
        End If
        If mblnTecleoFechaF Then
            Do While (msglTiempoCambioF) <= 2.1
            Loop
            mblnTecleoFechaF = False
        End If
        System.Windows.Forms.Application.DoEvents()
        Select Case True
            Case Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Unchecked And mintCodSucursal = 0
                MsgBox("Si no quiere imprimir los resultados de todas las sucursales, seleccione una de ellas", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dbcSucursal.Focus()
            Case Me.chkClientes.CheckState = System.Windows.Forms.CheckState.Unchecked And mintCodCliente = 0
                MsgBox("Si no quiere imprimir los resultados de todos los clientes, seleccione uno de ellos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.txtCodCliente.Focus()
            Case Me.dtpDesde.Value > Me.dtpHasta.Value
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Límite", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dtpDesde.Focus()
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Private Sub chkClientes_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkClientes.CheckStateChanged
        Select Case chkClientes.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mintCodCliente = 0
                mblnFueraChange = True
                txtCodCliente.Text = ""
                txtNombre.Text = ""
                txtCodCliente.BackColor = System.Drawing.ColorTranslator.FromOle(ColorGris)
                txtNombre.BackColor = System.Drawing.ColorTranslator.FromOle(ColorGris)
                txtCodCliente.Enabled = False
                txtNombre.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                txtCodCliente.Text = ""
                txtNombre.Text = ""
                txtCodCliente.BackColor = System.Drawing.ColorTranslator.FromOle(ColorBlanco)
                txtNombre.BackColor = System.Drawing.ColorTranslator.FromOle(ColorAmarillo)
                txtCodCliente.Enabled = True
                txtNombre.Enabled = True
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub chkTodas_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodas.CheckStateChanged
        Select Case Me.chkTodas.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                Me.dbcSucursal.Text = "[ Todas ... ]"
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
            Me.chkTodas.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursal_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyUp
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

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcSucursal.text)
        '''    If Me.dbcSucursal.SelectedItem <> 0 Then
        '''        dbcSucursal_LostFocus
        '''    End If
        '''    Me.dbcSucursal.text = Aux
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
        ' msglTiempoCambioF = VB.Timer()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaPorCliente_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaPorCliente_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaPorCliente_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmVtasRPTVentasSalidadeMercanciaPorCliente_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaPorCliente_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Me.dtpDesde.MinDate = C_FECHAINICIAL
        Me.dtpDesde.MaxDate = C_FECHAFINAL
        Me.dtpHasta.MinDate = C_FECHAINICIAL
        Me.dtpHasta.MaxDate = C_FECHAFINAL
        Call Me.Nuevo()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaPorCliente_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If mblnSalir Then
            mblnSalir = False
            Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Sale del Formulario
                    Cancel = 0
                Case MsgBoxResult.No 'No sale del formulario
                    Me.chkTodas.Focus()
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaPorCliente_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
    Private Sub txtCodCliente_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodCliente.TextChanged
        If mblnFueraChange Then Exit Sub
        If gBlnFueraChange Then Exit Sub
        If Trim(txtCodCliente.Text) = "" Then
            mblnFueraChange = True
            txtNombre.Text = ""
            mblnFueraChange = False
        End If
    End Sub

    '''03MAR2008 - MAVF
    Private Sub txtCodCliente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodCliente.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtCodCliente)
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

    Private Sub txtNombre_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.TextChanged
        If mblnFueraChange Then Exit Sub
        If gBlnFueraChange Then Exit Sub

        If Trim(txtNombre.Text) = "" Then
            mblnFueraChange = True
            txtCodCliente.Text = ""
            mblnFueraChange = False
        End If
    End Sub

    Private Sub txtNombre_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.Enter
        strControlActual = UCase("txtNombre")
        ModEstandar.SelTextoTxt(txtNombre)
    End Sub

    Private Sub TxtNombre_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNombre.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, ".,:;{}[]+#$%&()/*\-_<>")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub LimpiarCTE()
        mblnFueraChange = True
        txtCodCliente.Text = ""
        txtNombre.Text = ""
        mblnFueraChange = False
    End Sub

    Sub LlenaDatos()
        On Error GoTo Merr

        'txtCodCliente.Text = Format(txtCodCliente.Text, "00000")
        For i = 1 To 5 - (txtCodCliente.TextLength)
            txtCodCliente.Text = String.Concat("0" + txtCodCliente.Text)
        Next i

        gStrSql = "Select Right('00000' + ltrim(rtrim(CodCliente)),5) as Codigo, DescCliente as Nombre From CatClientes (Nolock) WHERE CodCliente = " & CInt(Numerico(txtCodCliente.Text)) & " "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            mintCodCliente = RsGral.Fields("Codigo").Value
            txtCodCliente.Text = RsGral.Fields("Codigo").Value
            txtNombre.Text = Trim(RsGral.Fields("Nombre").Value)
        Else
            MsjNoExiste("El cliente no existe." & vbNewLine & "Favor de verificar...", gstrNombCortoEmpresa)
            LimpiarCTE()
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub BuscaCliente()
        On Error GoTo Merr

        mintCodCliente = 0
        gStrSql = "SELECT CodCliente,DescCliente,AlmacenVExt FROM CatClientes WHERE CodCliente = " & txtCodCliente.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("CodCliente").Value = 1 Then
                MsgBox("El Cliente Publico en General No Tiene Registradas Cuentas X Cobrar, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodCliente.Text = ""
                txtCodCliente.Focus()
                Exit Sub
            End If
            If Not IsDBNull(RsGral.Fields("AlmacenVExt").Value) Then
                MsgBox("Los Clientes Registrados como Vendedores Externo No Tienen Registradas Cuentas X Cobrar, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodCliente.Text = ""
                txtCodCliente.Focus()
                Exit Sub
            End If
            mintCodCliente = RsGral.Fields("CodCliente").Value

            'txtCodCliente.Text = Format(txtCodCliente.Text, "00000")
            txtCodCliente.Text = Format(txtCodCliente.Text)

            For i = 1 To 5 - (txtCodCliente.TextLength)
                txtCodCliente.Text = String.Concat("0" + txtCodCliente.Text)
            Next i

            txtNombre.Text = Trim(RsGral.Fields("DescCliente").Value)
        Else
            MsgBox("Código de Cliente No Existe, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodCliente.Text = ""
            txtCodCliente.Focus()
            Exit Sub
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
    ''' ************************************
End Class