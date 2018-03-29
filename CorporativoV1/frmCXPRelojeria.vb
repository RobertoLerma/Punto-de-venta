'**********************************************************************************************************************'
'*PROGRAMA: RELOJERIA JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmCXPRelojeria
    Inherits System.Windows.Forms.Form

    Public components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtCodigodelProveedor As System.Windows.Forms.TextBox
    Public WithEvents txtPrecioenDolares As System.Windows.Forms.TextBox
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents _fraMoneda_5 As System.Windows.Forms.GroupBox
    Public WithEvents txtCodArtAnterior As System.Windows.Forms.TextBox
    Public WithEvents chkCodAnt As System.Windows.Forms.CheckBox
    Public WithEvents dbcOrigen As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_32 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_31 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents txtCostoActual As System.Windows.Forms.TextBox
    Public WithEvents txtCantidadCompra As System.Windows.Forms.TextBox
    Public WithEvents cmdBuscarImagen As System.Windows.Forms.Button
    Public WithEvents txtImagen As System.Windows.Forms.TextBox
    Public WithEvents _FrameImagen_0 As System.Windows.Forms.GroupBox
    Public WithEvents _lblArticulo_6 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_7 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_8 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_9 As System.Windows.Forms.Label
    Public WithEvents _fraJoyeria_0 As System.Windows.Forms.GroupBox
    Public WithEvents btnAceptar As System.Windows.Forms.Button
    Public WithEvents btnCancelar As System.Windows.Forms.Button
    Public WithEvents txtAdicional As System.Windows.Forms.TextBox
    Public WithEvents _optMovimiento_2 As System.Windows.Forms.RadioButton
    Public WithEvents _optMovimiento_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optMovimiento_0 As System.Windows.Forms.RadioButton
    Public WithEvents _fraArticulo_2 As System.Windows.Forms.GroupBox
    Public WithEvents btnTipoMaterial As System.Windows.Forms.Button
    Public WithEvents btnModelo As System.Windows.Forms.Button
    Public WithEvents btnMarca As System.Windows.Forms.Button
    Public WithEvents chkCrono As System.Windows.Forms.CheckBox
    Public WithEvents _optGenero_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optGenero_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optGenero_2 As System.Windows.Forms.RadioButton
    Public WithEvents _fraArticulo_1 As System.Windows.Forms.GroupBox
    Public WithEvents dbcMarca As System.Windows.Forms.ComboBox
    Public WithEvents dbcModelo As System.Windows.Forms.ComboBox
    Public WithEvents dbcMaterial As System.Windows.Forms.ComboBox
    Public WithEvents dbcUnidad As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_33 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_16 As System.Windows.Forms.Label
    Public WithEvents txtDescripcion As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_4 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_5 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_13 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_12 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_15 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_17 As System.Windows.Forms.Label
    Public WithEvents FrameImagen As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents fraArticulo As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents fraJoyeria As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents fraMoneda As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblArticulo As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optGenero As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optMovimiento As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray


    Const cINDEFINIDA As String = "[ Vacío ... ]"
    Const cINDEFINIDO As String = "[ Vacío ... ]"

    Const C_COLCODIGO As Integer = 0
    Const C_ColDESCRIPCION As Integer = 1
    Const C_ColUNIDAD As Integer = 2
    Const C_ColCANTIDAD As Integer = 3
    Const C_COLPRECIOUNITARIO As Integer = 4
    Const C_COLCODAUX As Integer = 8
    Const C_COLSTATUS As Integer = 9
    Const C_ColCODGRUPO As Integer = 19
    Const C_COLCODFAMILIA As Integer = 20
    Const C_COLCODLINEA As Integer = 21
    Const C_COLCODSUBLINEA As Integer = 22
    Const C_COLCODKILATES As Integer = 52
    Const C_COLCODMARCA As Integer = 23
    Const C_COLCODMODELO As Integer = 24
    Const C_COLCODTIPOMATERIAL As Integer = 25
    Const C_COLGENERO As Integer = 26
    Const C_COLMOVIMIENTO As Integer = 27
    Const C_COLCRONO As Integer = 54
    Const C_COLCODIGOARTICULOPROV As Integer = 28
    Const C_COLDESCTO As Integer = 6
    Const C_COLDESCTOPORC As Integer = 45
    Const C_COLDESCTOPORCTAG As Integer = 46

    Const C_COLADICIONAL As Integer = 58
    Const C_COLPRECIOPUBDOLAR As Integer = 59
    Const C_COLMONEDAPP As Integer = 60
    Const C_COLORIGENANT As Integer = 61
    Const C_ColCODIGOANT As Integer = 62
    Const C_ColIMAGEN As Integer = 63
    Const C_COLSTATUSX As Integer = 76

    Dim cDescripcion As String
    Dim mblnSalir As Boolean
    Dim mblnCancelar As Boolean
    Dim mblnNuevo As Boolean
    Dim nCol As Integer
    Dim nRow As Integer
    Dim nRowAct As Integer

    Public mblnFueraChange As Boolean
    Public mintCodGrupo As Integer
    Dim mintCodUnidad As Integer
    Public mintCodMarca As Integer
    Public mintCodModelo As Integer
    Public mintCodMaterial As Integer
    Dim intCodAlmacenOrigen As Integer

    Dim cMovimientoTag As String
    Dim cGeneroTag As String
    Dim cCronoTag As String
    Dim cTipoMaterial As String
    Dim lCrono As Boolean
    Dim tecla As Integer
    Dim rsLocal As ADODB.Recordset
    Dim cCodProveedor As String
    Dim cGenero, cModelo, cMarca, cMaterial, cMovimiento As Object
    Dim cCrono, cAdicional As String


    Public Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCXPRelojeria))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me._fraJoyeria_0 = New System.Windows.Forms.GroupBox
        Me.txtCodigodelProveedor = New System.Windows.Forms.TextBox
        Me.txtPrecioenDolares = New System.Windows.Forms.TextBox
        Me._fraMoneda_5 = New System.Windows.Forms.GroupBox
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton
        Me.Frame3 = New System.Windows.Forms.GroupBox
        Me.txtCodArtAnterior = New System.Windows.Forms.TextBox
        Me.chkCodAnt = New System.Windows.Forms.CheckBox
        Me.dbcOrigen = New System.Windows.Forms.ComboBox
        Me._lblArticulo_32 = New System.Windows.Forms.Label
        Me._lblArticulo_31 = New System.Windows.Forms.Label
        Me.txtCostoActual = New System.Windows.Forms.TextBox
        Me.txtCantidadCompra = New System.Windows.Forms.TextBox
        Me._FrameImagen_0 = New System.Windows.Forms.GroupBox
        Me.cmdBuscarImagen = New System.Windows.Forms.Button
        Me.txtImagen = New System.Windows.Forms.TextBox
        Me._lblArticulo_6 = New System.Windows.Forms.Label
        Me._lblArticulo_7 = New System.Windows.Forms.Label
        Me._lblArticulo_8 = New System.Windows.Forms.Label
        Me._lblArticulo_9 = New System.Windows.Forms.Label
        Me.btnAceptar = New System.Windows.Forms.Button
        Me.btnCancelar = New System.Windows.Forms.Button
        Me.txtAdicional = New System.Windows.Forms.TextBox
        Me._fraArticulo_2 = New System.Windows.Forms.GroupBox
        Me._optMovimiento_2 = New System.Windows.Forms.RadioButton
        Me._optMovimiento_1 = New System.Windows.Forms.RadioButton
        Me._optMovimiento_0 = New System.Windows.Forms.RadioButton
        Me.btnTipoMaterial = New System.Windows.Forms.Button
        Me.btnModelo = New System.Windows.Forms.Button
        Me.btnMarca = New System.Windows.Forms.Button
        Me.chkCrono = New System.Windows.Forms.CheckBox
        Me._fraArticulo_1 = New System.Windows.Forms.GroupBox
        Me._optGenero_0 = New System.Windows.Forms.RadioButton
        Me._optGenero_1 = New System.Windows.Forms.RadioButton
        Me._optGenero_2 = New System.Windows.Forms.RadioButton
        Me.dbcMarca = New System.Windows.Forms.ComboBox
        Me.dbcModelo = New System.Windows.Forms.ComboBox
        Me.dbcMaterial = New System.Windows.Forms.ComboBox
        Me.dbcUnidad = New System.Windows.Forms.ComboBox
        Me._lblArticulo_33 = New System.Windows.Forms.Label
        Me._lblArticulo_16 = New System.Windows.Forms.Label
        Me.txtDescripcion = New System.Windows.Forms.Label
        Me._lblArticulo_4 = New System.Windows.Forms.Label
        Me._lblArticulo_5 = New System.Windows.Forms.Label
        Me._lblArticulo_13 = New System.Windows.Forms.Label
        Me._lblArticulo_12 = New System.Windows.Forms.Label
        Me._lblArticulo_15 = New System.Windows.Forms.Label
        Me._lblArticulo_17 = New System.Windows.Forms.Label
        Me.FrameImagen = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
        Me.fraArticulo = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
        Me.fraJoyeria = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
        Me.fraMoneda = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
        Me.lblArticulo = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
        Me.optGenero = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
        Me.optMovimiento = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
        Me._fraJoyeria_0.SuspendLayout()
        Me._fraMoneda_5.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me._FrameImagen_0.SuspendLayout()
        Me._fraArticulo_2.SuspendLayout()
        Me._fraArticulo_1.SuspendLayout()
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        CType(Me.dbcOrigen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dbcMarca, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dbcModelo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dbcMaterial, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dbcUnidad, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FrameImagen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraArticulo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraJoyeria, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblArticulo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optGenero, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMovimiento, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Text = "Definir Relojería"
        Me.ClientSize = New System.Drawing.Size(803, 322)
        Me.Location = New System.Drawing.Point(149, 87)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.MinimizeBox = False
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ControlBox = True
        Me.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = True
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmCXPRelojeria"
        Me._fraJoyeria_0.Text = " Datos Adicionales "
        Me._fraJoyeria_0.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me._fraJoyeria_0.Size = New System.Drawing.Size(382, 276)
        Me._fraJoyeria_0.Location = New System.Drawing.Point(412, 5)
        Me._fraJoyeria_0.TabIndex = 26
        Me._fraJoyeria_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraJoyeria_0.Enabled = True
        Me._fraJoyeria_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraJoyeria_0.Visible = True
        Me._fraJoyeria_0.Name = "_fraJoyeria_0"
        Me.txtCodigodelProveedor.AutoSize = False
        Me.txtCodigodelProveedor.BackColor = System.Drawing.Color.FromArgb(210, 230, 244)
        Me.txtCodigodelProveedor.Size = New System.Drawing.Size(129, 21)
        Me.txtCodigodelProveedor.Location = New System.Drawing.Point(78, 30)
        Me.txtCodigodelProveedor.MaxLength = 20
        Me.txtCodigodelProveedor.TabIndex = 28
        Me.ToolTip1.SetToolTip(Me.txtCodigodelProveedor, "Código del Proveedor para el Artículo")
        Me.txtCodigodelProveedor.AcceptsReturn = True
        Me.txtCodigodelProveedor.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtCodigodelProveedor.CausesValidation = True
        Me.txtCodigodelProveedor.Enabled = True
        Me.txtCodigodelProveedor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodigodelProveedor.HideSelection = True
        Me.txtCodigodelProveedor.ReadOnly = False
        Me.txtCodigodelProveedor.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodigodelProveedor.Multiline = False
        Me.txtCodigodelProveedor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodigodelProveedor.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtCodigodelProveedor.TabStop = True
        Me.txtCodigodelProveedor.Visible = True
        Me.txtCodigodelProveedor.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtCodigodelProveedor.Name = "txtCodigodelProveedor"
        Me.txtPrecioenDolares.AutoSize = False
        Me.txtPrecioenDolares.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrecioenDolares.Size = New System.Drawing.Size(96, 21)
        Me.txtPrecioenDolares.Location = New System.Drawing.Point(115, 134)
        Me.txtPrecioenDolares.TabIndex = 34
        Me.txtPrecioenDolares.Text = "0.00"
        Me.ToolTip1.SetToolTip(Me.txtPrecioenDolares, "Precio al Público en Dólares")
        Me.txtPrecioenDolares.AcceptsReturn = True
        Me.txtPrecioenDolares.BackColor = System.Drawing.SystemColors.Window
        Me.txtPrecioenDolares.CausesValidation = True
        Me.txtPrecioenDolares.Enabled = True
        Me.txtPrecioenDolares.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrecioenDolares.HideSelection = True
        Me.txtPrecioenDolares.ReadOnly = False
        Me.txtPrecioenDolares.MaxLength = 0
        Me.txtPrecioenDolares.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrecioenDolares.Multiline = False
        Me.txtPrecioenDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrecioenDolares.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtPrecioenDolares.TabStop = True
        Me.txtPrecioenDolares.Visible = True
        Me.txtPrecioenDolares.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtPrecioenDolares.Name = "txtPrecioenDolares"
        Me._fraMoneda_5.Text = "  Moneda Precio Público "
        Me._fraMoneda_5.Size = New System.Drawing.Size(196, 40)
        Me._fraMoneda_5.Location = New System.Drawing.Point(17, 162)
        Me._fraMoneda_5.TabIndex = 35
        Me._fraMoneda_5.BackColor = System.Drawing.SystemColors.Control
        Me._fraMoneda_5.Enabled = True
        Me._fraMoneda_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraMoneda_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraMoneda_5.Visible = True
        Me._fraMoneda_5.Name = "_fraMoneda_5"
        Me._optMoneda_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMoneda_1.Text = "Pesos"
        Me._optMoneda_1.Size = New System.Drawing.Size(52, 17)
        Me._optMoneda_1.Location = New System.Drawing.Point(115, 17)
        Me._optMoneda_1.TabIndex = 37
        Me._optMoneda_1.Tag = "0"
        Me.ToolTip1.SetToolTip(Me._optMoneda_1, "Moneda del Precio Público - Pes")
        Me._optMoneda_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.CausesValidation = True
        Me._optMoneda_1.Enabled = True
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Checked = False
        Me._optMoneda_1.Visible = True
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMoneda_0.Text = "Dólares"
        Me._optMoneda_0.Size = New System.Drawing.Size(60, 17)
        Me._optMoneda_0.Location = New System.Drawing.Point(15, 17)
        Me._optMoneda_0.TabIndex = 36
        Me._optMoneda_0.Tag = "1"
        Me.ToolTip1.SetToolTip(Me._optMoneda_0, "Moneda del Precio Público - Dol")
        Me._optMoneda_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.CausesValidation = True
        Me._optMoneda_0.Enabled = True
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Checked = False
        Me._optMoneda_0.Visible = True
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me.Frame3.Text = "    Codigo Anterior"
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.Frame3.Size = New System.Drawing.Size(124, 69)
        Me.Frame3.Location = New System.Drawing.Point(242, 24)
        Me.Frame3.TabIndex = 38
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Enabled = True
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Visible = True
        Me.Frame3.Name = "Frame3"
        Me.txtCodArtAnterior.AutoSize = False
        Me.txtCodArtAnterior.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCodArtAnterior.Enabled = False
        Me.txtCodArtAnterior.Size = New System.Drawing.Size(43, 21)
        Me.txtCodArtAnterior.Location = New System.Drawing.Point(70, 41)
        Me.txtCodArtAnterior.MaxLength = 5
        Me.txtCodArtAnterior.TabIndex = 43
        Me.ToolTip1.SetToolTip(Me.txtCodArtAnterior, "Codigo Anterior del Artículo")
        Me.txtCodArtAnterior.AcceptsReturn = True
        Me.txtCodArtAnterior.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodArtAnterior.CausesValidation = True
        Me.txtCodArtAnterior.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodArtAnterior.HideSelection = True
        Me.txtCodArtAnterior.ReadOnly = False
        Me.txtCodArtAnterior.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodArtAnterior.Multiline = False
        Me.txtCodArtAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodArtAnterior.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtCodArtAnterior.TabStop = True
        Me.txtCodArtAnterior.Visible = True
        Me.txtCodArtAnterior.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtCodArtAnterior.Name = "txtCodArtAnterior"
        Me.chkCodAnt.Size = New System.Drawing.Size(17, 17)
        Me.chkCodAnt.Location = New System.Drawing.Point(8, -2)
        Me.chkCodAnt.TabIndex = 39
        Me.chkCodAnt.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.chkCodAnt.FlatStyle = System.Windows.Forms.FlatStyle.Standard
        Me.chkCodAnt.BackColor = System.Drawing.SystemColors.Control
        Me.chkCodAnt.Text = ""
        Me.chkCodAnt.CausesValidation = True
        Me.chkCodAnt.Enabled = True
        Me.chkCodAnt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCodAnt.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCodAnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCodAnt.Appearance = System.Windows.Forms.Appearance.Normal
        Me.chkCodAnt.TabStop = True
        Me.chkCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked
        Me.chkCodAnt.Visible = True
        Me.chkCodAnt.Name = "chkCodAnt"
        '.OcxState = CType(resources.GetObject("dbcOrigen.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dbcOrigen.Size = New System.Drawing.Size(43, 21)
        Me.dbcOrigen.Location = New System.Drawing.Point(70, 17)
        Me.dbcOrigen.TabIndex = 41
        Me.dbcOrigen.Name = "dbcOrigen"
        Me._lblArticulo_32.Text = "Código: "
        Me._lblArticulo_32.Size = New System.Drawing.Size(39, 13)
        Me._lblArticulo_32.Location = New System.Drawing.Point(26, 44)
        Me._lblArticulo_32.TabIndex = 42
        Me._lblArticulo_32.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_32.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_32.Enabled = True
        Me._lblArticulo_32.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_32.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_32.UseMnemonic = True
        Me._lblArticulo_32.Visible = True
        Me._lblArticulo_32.AutoSize = True
        Me._lblArticulo_32.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_32.Name = "_lblArticulo_32"
        Me._lblArticulo_31.Text = "Origen : "
        Me._lblArticulo_31.Size = New System.Drawing.Size(40, 13)
        Me._lblArticulo_31.Location = New System.Drawing.Point(25, 21)
        Me._lblArticulo_31.TabIndex = 40
        Me._lblArticulo_31.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_31.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_31.Enabled = True
        Me._lblArticulo_31.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_31.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_31.UseMnemonic = True
        Me._lblArticulo_31.Visible = True
        Me._lblArticulo_31.AutoSize = True
        Me._lblArticulo_31.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_31.Name = "_lblArticulo_31"
        Me.txtCostoActual.AutoSize = False
        Me.txtCostoActual.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCostoActual.Size = New System.Drawing.Size(96, 21)
        Me.txtCostoActual.Location = New System.Drawing.Point(112, 70)
        Me.txtCostoActual.TabIndex = 30
        Me.txtCostoActual.Text = "0.00"
        Me.ToolTip1.SetToolTip(Me.txtCostoActual, "Costo Factura sin Iva")
        Me.txtCostoActual.AcceptsReturn = True
        Me.txtCostoActual.BackColor = System.Drawing.SystemColors.Window
        Me.txtCostoActual.CausesValidation = True
        Me.txtCostoActual.Enabled = True
        Me.txtCostoActual.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCostoActual.HideSelection = True
        Me.txtCostoActual.ReadOnly = False
        Me.txtCostoActual.MaxLength = 0
        Me.txtCostoActual.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCostoActual.Multiline = False
        Me.txtCostoActual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCostoActual.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtCostoActual.TabStop = True
        Me.txtCostoActual.Visible = True
        Me.txtCostoActual.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtCostoActual.Name = "txtCostoActual"
        Me.txtCantidadCompra.AutoSize = False
        Me.txtCantidadCompra.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCantidadCompra.Size = New System.Drawing.Size(66, 21)
        Me.txtCantidadCompra.Location = New System.Drawing.Point(143, 97)
        Me.txtCantidadCompra.TabIndex = 32
        Me.txtCantidadCompra.Text = "0"
        Me.ToolTip1.SetToolTip(Me.txtCantidadCompra, "Cantidad de la Compra")
        Me.txtCantidadCompra.AcceptsReturn = True
        Me.txtCantidadCompra.BackColor = System.Drawing.SystemColors.Window
        Me.txtCantidadCompra.CausesValidation = True
        Me.txtCantidadCompra.Enabled = True
        Me.txtCantidadCompra.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCantidadCompra.HideSelection = True
        Me.txtCantidadCompra.ReadOnly = False
        Me.txtCantidadCompra.MaxLength = 0
        Me.txtCantidadCompra.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCantidadCompra.Multiline = False
        Me.txtCantidadCompra.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCantidadCompra.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtCantidadCompra.TabStop = True
        Me.txtCantidadCompra.Visible = True
        Me.txtCantidadCompra.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtCantidadCompra.Name = "txtCantidadCompra"
        Me._FrameImagen_0.Text = "Imagen"
        Me._FrameImagen_0.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me._FrameImagen_0.Size = New System.Drawing.Size(290, 44)
        Me._FrameImagen_0.Location = New System.Drawing.Point(76, 215)
        Me._FrameImagen_0.TabIndex = 44
        Me._FrameImagen_0.BackColor = System.Drawing.SystemColors.Control
        Me._FrameImagen_0.Enabled = True
        Me._FrameImagen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._FrameImagen_0.Visible = True
        Me._FrameImagen_0.Name = "_FrameImagen_0"
        Me.cmdBuscarImagen.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdBuscarImagen.Text = "..."
        Me.cmdBuscarImagen.Size = New System.Drawing.Size(22, 21)
        Me.cmdBuscarImagen.Location = New System.Drawing.Point(260, 15)
        Me.cmdBuscarImagen.TabIndex = 48
        Me.cmdBuscarImagen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBuscarImagen.CausesValidation = True
        Me.cmdBuscarImagen.Enabled = True
        Me.cmdBuscarImagen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBuscarImagen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBuscarImagen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBuscarImagen.TabStop = True
        Me.cmdBuscarImagen.Name = "cmdBuscarImagen"
        Me.txtImagen.AutoSize = False
        Me.txtImagen.Size = New System.Drawing.Size(245, 21)
        Me.txtImagen.Location = New System.Drawing.Point(9, 15)
        Me.txtImagen.ReadOnly = True
        Me.txtImagen.TabIndex = 45
        Me.ToolTip1.SetToolTip(Me.txtImagen, "Imagen del artículo")
        Me.txtImagen.AcceptsReturn = True
        Me.txtImagen.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtImagen.BackColor = System.Drawing.SystemColors.Window
        Me.txtImagen.CausesValidation = True
        Me.txtImagen.Enabled = True
        Me.txtImagen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImagen.HideSelection = True
        Me.txtImagen.MaxLength = 0
        Me.txtImagen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImagen.Multiline = False
        Me.txtImagen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImagen.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtImagen.TabStop = True
        Me.txtImagen.Visible = True
        Me.txtImagen.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtImagen.Name = "txtImagen"
        Me._lblArticulo_6.Text = "Código del Proveedor :"
        Me._lblArticulo_6.Size = New System.Drawing.Size(64, 28)
        Me._lblArticulo_6.Location = New System.Drawing.Point(14, 24)
        Me._lblArticulo_6.TabIndex = 27
        Me._lblArticulo_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_6.Enabled = True
        Me._lblArticulo_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_6.UseMnemonic = True
        Me._lblArticulo_6.Visible = True
        Me._lblArticulo_6.AutoSize = False
        Me._lblArticulo_6.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_6.Name = "_lblArticulo_6"
        Me._lblArticulo_7.Text = "Precio Público :"
        Me._lblArticulo_7.Size = New System.Drawing.Size(76, 20)
        Me._lblArticulo_7.Location = New System.Drawing.Point(17, 139)
        Me._lblArticulo_7.TabIndex = 33
        Me._lblArticulo_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_7.Enabled = True
        Me._lblArticulo_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_7.UseMnemonic = True
        Me._lblArticulo_7.Visible = True
        Me._lblArticulo_7.AutoSize = False
        Me._lblArticulo_7.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_7.Name = "_lblArticulo_7"
        Me._lblArticulo_8.Text = "Cto Fact S/Iva"
        Me._lblArticulo_8.Size = New System.Drawing.Size(76, 17)
        Me._lblArticulo_8.Location = New System.Drawing.Point(22, 74)
        Me._lblArticulo_8.TabIndex = 29
        Me._lblArticulo_8.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_8.Enabled = True
        Me._lblArticulo_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_8.UseMnemonic = True
        Me._lblArticulo_8.Visible = True
        Me._lblArticulo_8.AutoSize = False
        Me._lblArticulo_8.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_8.Name = "_lblArticulo_8"
        Me._lblArticulo_9.Text = "Cantidad Compra :"
        Me._lblArticulo_9.Size = New System.Drawing.Size(90, 15)
        Me._lblArticulo_9.Location = New System.Drawing.Point(44, 102)
        Me._lblArticulo_9.TabIndex = 31
        Me._lblArticulo_9.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_9.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_9.Enabled = True
        Me._lblArticulo_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_9.UseMnemonic = True
        Me._lblArticulo_9.Visible = True
        Me._lblArticulo_9.AutoSize = False
        Me._lblArticulo_9.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_9.Name = "_lblArticulo_9"
        Me.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnAceptar.Text = "&Aceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(88, 25)
        Me.btnAceptar.Location = New System.Drawing.Point(506, 291)
        Me.btnAceptar.TabIndex = 46
        Me.btnAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.btnAceptar.CausesValidation = True
        Me.btnAceptar.Enabled = True
        Me.btnAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAceptar.TabStop = True
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnCancelar.Text = "&Cancelar"
        Me.btnCancelar.Size = New System.Drawing.Size(88, 25)
        Me.btnCancelar.Location = New System.Drawing.Point(661, 291)
        Me.btnCancelar.TabIndex = 47
        Me.btnCancelar.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancelar.CausesValidation = True
        Me.btnCancelar.Enabled = True
        Me.btnCancelar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancelar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancelar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancelar.TabStop = True
        Me.btnCancelar.Name = "btnCancelar"
        Me.txtAdicional.AutoSize = False
        Me.txtAdicional.BackColor = System.Drawing.Color.FromArgb(210, 230, 244)
        Me.txtAdicional.Size = New System.Drawing.Size(117, 21)
        Me.txtAdicional.Location = New System.Drawing.Point(95, 208)
        Me.txtAdicional.MaxLength = 15
        Me.txtAdicional.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.txtAdicional, "Dato Adicional")
        Me.txtAdicional.AcceptsReturn = True
        Me.txtAdicional.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtAdicional.CausesValidation = True
        Me.txtAdicional.Enabled = True
        Me.txtAdicional.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAdicional.HideSelection = True
        Me.txtAdicional.ReadOnly = False
        Me.txtAdicional.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAdicional.Multiline = False
        Me.txtAdicional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAdicional.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtAdicional.TabStop = True
        Me.txtAdicional.Visible = True
        Me.txtAdicional.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtAdicional.Name = "txtAdicional"
        Me._fraArticulo_2.Size = New System.Drawing.Size(265, 33)
        Me._fraArticulo_2.Location = New System.Drawing.Point(95, 137)
        Me._fraArticulo_2.TabIndex = 14
        Me._fraArticulo_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraArticulo_2.Enabled = True
        Me._fraArticulo_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraArticulo_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraArticulo_2.Visible = True
        Me._fraArticulo_2.Name = "_fraArticulo_2"
        Me._optMovimiento_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMovimiento_2.Text = "Manual"
        Me._optMovimiento_2.Size = New System.Drawing.Size(65, 21)
        Me._optMovimiento_2.Location = New System.Drawing.Point(184, 8)
        Me._optMovimiento_2.TabIndex = 17
        Me.ToolTip1.SetToolTip(Me._optMovimiento_2, "Manual")
        Me._optMovimiento_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMovimiento_2.BackColor = System.Drawing.SystemColors.Control
        Me._optMovimiento_2.CausesValidation = True
        Me._optMovimiento_2.Enabled = True
        Me._optMovimiento_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMovimiento_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMovimiento_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMovimiento_2.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optMovimiento_2.TabStop = True
        Me._optMovimiento_2.Checked = False
        Me._optMovimiento_2.Visible = True
        Me._optMovimiento_2.Name = "_optMovimiento_2"
        Me._optMovimiento_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMovimiento_1.Text = "Automático"
        Me._optMovimiento_1.Size = New System.Drawing.Size(81, 21)
        Me._optMovimiento_1.Location = New System.Drawing.Point(104, 8)
        Me._optMovimiento_1.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me._optMovimiento_1, "Movimiento Automatizado")
        Me._optMovimiento_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMovimiento_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMovimiento_1.CausesValidation = True
        Me._optMovimiento_1.Enabled = True
        Me._optMovimiento_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMovimiento_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMovimiento_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMovimiento_1.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optMovimiento_1.TabStop = True
        Me._optMovimiento_1.Checked = False
        Me._optMovimiento_1.Visible = True
        Me._optMovimiento_1.Name = "_optMovimiento_1"
        Me._optMovimiento_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMovimiento_0.Text = "Cuarzo"
        Me._optMovimiento_0.Size = New System.Drawing.Size(57, 21)
        Me._optMovimiento_0.Location = New System.Drawing.Point(24, 8)
        Me._optMovimiento_0.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me._optMovimiento_0, "Movimiento por Cuarzo")
        Me._optMovimiento_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optMovimiento_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMovimiento_0.CausesValidation = True
        Me._optMovimiento_0.Enabled = True
        Me._optMovimiento_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMovimiento_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMovimiento_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMovimiento_0.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optMovimiento_0.TabStop = True
        Me._optMovimiento_0.Checked = False
        Me._optMovimiento_0.Visible = True
        Me._optMovimiento_0.Name = "_optMovimiento_0"
        Me.btnTipoMaterial.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnTipoMaterial.Text = "..."
        Me.btnTipoMaterial.Size = New System.Drawing.Size(21, 21)
        Me.btnTipoMaterial.Location = New System.Drawing.Point(374, 176)
        Me.btnTipoMaterial.TabIndex = 20
        Me.btnTipoMaterial.TabStop = False
        Me.btnTipoMaterial.BackColor = System.Drawing.SystemColors.Control
        Me.btnTipoMaterial.CausesValidation = True
        Me.btnTipoMaterial.Enabled = True
        Me.btnTipoMaterial.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnTipoMaterial.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnTipoMaterial.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnTipoMaterial.Name = "btnTipoMaterial"
        Me.btnModelo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnModelo.Text = "..."
        Me.btnModelo.Size = New System.Drawing.Size(21, 21)
        Me.btnModelo.Location = New System.Drawing.Point(374, 80)
        Me.btnModelo.TabIndex = 7
        Me.btnModelo.TabStop = False
        Me.btnModelo.BackColor = System.Drawing.SystemColors.Control
        Me.btnModelo.CausesValidation = True
        Me.btnModelo.Enabled = True
        Me.btnModelo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnModelo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnModelo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnModelo.Name = "btnModelo"
        Me.btnMarca.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnMarca.Text = "..."
        Me.btnMarca.Size = New System.Drawing.Size(21, 21)
        Me.btnMarca.Location = New System.Drawing.Point(374, 48)
        Me.btnMarca.TabIndex = 4
        Me.btnMarca.TabStop = False
        Me.btnMarca.BackColor = System.Drawing.SystemColors.Control
        Me.btnMarca.CausesValidation = True
        Me.btnMarca.Enabled = True
        Me.btnMarca.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnMarca.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnMarca.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnMarca.Name = "btnMarca"
        Me.chkCrono.Text = "Cronógrafo"
        Me.chkCrono.Size = New System.Drawing.Size(81, 17)
        Me.chkCrono.Location = New System.Drawing.Point(267, 213)
        Me.chkCrono.TabIndex = 23
        Me.chkCrono.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.chkCrono.FlatStyle = System.Windows.Forms.FlatStyle.Standard
        Me.chkCrono.BackColor = System.Drawing.SystemColors.Control
        Me.chkCrono.CausesValidation = True
        Me.chkCrono.Enabled = True
        Me.chkCrono.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCrono.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCrono.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCrono.Appearance = System.Windows.Forms.Appearance.Normal
        Me.chkCrono.TabStop = True
        Me.chkCrono.CheckState = System.Windows.Forms.CheckState.Unchecked
        Me.chkCrono.Visible = True
        Me.chkCrono.Name = "chkCrono"
        Me._fraArticulo_1.Size = New System.Drawing.Size(265, 33)
        Me._fraArticulo_1.Location = New System.Drawing.Point(95, 103)
        Me._fraArticulo_1.TabIndex = 9
        Me._fraArticulo_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraArticulo_1.Enabled = True
        Me._fraArticulo_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraArticulo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraArticulo_1.Visible = True
        Me._fraArticulo_1.Name = "_fraArticulo_1"
        Me._optGenero_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optGenero_0.Text = "Caballero"
        Me._optGenero_0.Size = New System.Drawing.Size(65, 21)
        Me._optGenero_0.Location = New System.Drawing.Point(24, 8)
        Me._optGenero_0.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me._optGenero_0, "Para Hombre")
        Me._optGenero_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optGenero_0.BackColor = System.Drawing.SystemColors.Control
        Me._optGenero_0.CausesValidation = True
        Me._optGenero_0.Enabled = True
        Me._optGenero_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optGenero_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optGenero_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optGenero_0.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optGenero_0.TabStop = True
        Me._optGenero_0.Checked = False
        Me._optGenero_0.Visible = True
        Me._optGenero_0.Name = "_optGenero_0"
        Me._optGenero_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optGenero_1.Text = "Dama"
        Me._optGenero_1.Size = New System.Drawing.Size(57, 21)
        Me._optGenero_1.Location = New System.Drawing.Point(104, 8)
        Me._optGenero_1.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me._optGenero_1, "Para Mujer")
        Me._optGenero_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optGenero_1.BackColor = System.Drawing.SystemColors.Control
        Me._optGenero_1.CausesValidation = True
        Me._optGenero_1.Enabled = True
        Me._optGenero_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optGenero_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optGenero_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optGenero_1.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optGenero_1.TabStop = True
        Me._optGenero_1.Checked = False
        Me._optGenero_1.Visible = True
        Me._optGenero_1.Name = "_optGenero_1"
        Me._optGenero_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optGenero_2.Text = "Mediano"
        Me._optGenero_2.Size = New System.Drawing.Size(65, 21)
        Me._optGenero_2.Location = New System.Drawing.Point(184, 8)
        Me._optGenero_2.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me._optGenero_2, "Para cualquier tipo de sexo")
        Me._optGenero_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._optGenero_2.BackColor = System.Drawing.SystemColors.Control
        Me._optGenero_2.CausesValidation = True
        Me._optGenero_2.Enabled = True
        Me._optGenero_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optGenero_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optGenero_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optGenero_2.Appearance = System.Windows.Forms.Appearance.Normal
        Me._optGenero_2.TabStop = True
        Me._optGenero_2.Checked = False
        Me._optGenero_2.Visible = True
        Me._optGenero_2.Name = "_optGenero_2"
        'dbcMarca.OcxState = CType(resources.GetObject("dbcMarca.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dbcMarca.Size = New System.Drawing.Size(265, 21)
        Me.dbcMarca.Location = New System.Drawing.Point(95, 48)
        Me.dbcMarca.TabIndex = 3
        Me.dbcMarca.Name = "dbcMarca"
        'dbcModelo.OcxState = CType(resources.GetObject("dbcModelo.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dbcModelo.Size = New System.Drawing.Size(265, 21)
        Me.dbcModelo.Location = New System.Drawing.Point(95, 80)
        Me.dbcModelo.TabIndex = 6
        Me.dbcModelo.Name = "dbcModelo"
        'dbcMaterial.OcxState = CType(resources.GetObject("dbcMaterial.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dbcMaterial.Size = New System.Drawing.Size(265, 21)
        Me.dbcMaterial.Location = New System.Drawing.Point(95, 177)
        Me.dbcMaterial.TabIndex = 19
        Me.dbcMaterial.Name = "dbcMaterial"
        ' dbcUnidad.OcxState = CType(resources.GetObject("dbcUnidad.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dbcUnidad.Size = New System.Drawing.Size(100, 21)
        Me.dbcUnidad.Location = New System.Drawing.Point(95, 16)
        Me.dbcUnidad.TabIndex = 1
        Me.dbcUnidad.Name = "dbcUnidad"
        Me._lblArticulo_33.Text = "Dato Adicional"
        Me._lblArticulo_33.Size = New System.Drawing.Size(69, 13)
        Me._lblArticulo_33.Location = New System.Drawing.Point(11, 213)
        Me._lblArticulo_33.TabIndex = 21
        Me._lblArticulo_33.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_33.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_33.Enabled = True
        Me._lblArticulo_33.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_33.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_33.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_33.UseMnemonic = True
        Me._lblArticulo_33.Visible = True
        Me._lblArticulo_33.AutoSize = True
        Me._lblArticulo_33.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_33.Name = "_lblArticulo_33"
        Me._lblArticulo_16.Text = "Funcionamiento"
        Me._lblArticulo_16.Size = New System.Drawing.Size(75, 13)
        Me._lblArticulo_16.Location = New System.Drawing.Point(11, 152)
        Me._lblArticulo_16.TabIndex = 13
        Me._lblArticulo_16.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_16.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_16.Enabled = True
        Me._lblArticulo_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_16.UseMnemonic = True
        Me._lblArticulo_16.Visible = True
        Me._lblArticulo_16.AutoSize = True
        Me._lblArticulo_16.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_16.Name = "_lblArticulo_16"
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Info
        Me.txtDescripcion.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.txtDescripcion.Size = New System.Drawing.Size(265, 57)
        Me.txtDescripcion.Location = New System.Drawing.Point(95, 256)
        Me.txtDescripcion.TabIndex = 25
        Me.txtDescripcion.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.txtDescripcion.Enabled = True
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.UseMnemonic = True
        Me.txtDescripcion.Visible = True
        Me.txtDescripcion.AutoSize = False
        Me.txtDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDescripcion.Name = "txtDescripcion"
        Me._lblArticulo_4.Text = "Descripción"
        Me._lblArticulo_4.Size = New System.Drawing.Size(56, 13)
        Me._lblArticulo_4.Location = New System.Drawing.Point(11, 259)
        Me._lblArticulo_4.TabIndex = 24
        Me._lblArticulo_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_4.Enabled = True
        Me._lblArticulo_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_4.UseMnemonic = True
        Me._lblArticulo_4.Visible = True
        Me._lblArticulo_4.AutoSize = True
        Me._lblArticulo_4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_4.Name = "_lblArticulo_4"
        Me._lblArticulo_5.Text = "Unidad"
        Me._lblArticulo_5.Size = New System.Drawing.Size(34, 13)
        Me._lblArticulo_5.Location = New System.Drawing.Point(11, 20)
        Me._lblArticulo_5.TabIndex = 0
        Me._lblArticulo_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_5.Enabled = True
        Me._lblArticulo_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_5.UseMnemonic = True
        Me._lblArticulo_5.Visible = True
        Me._lblArticulo_5.AutoSize = True
        Me._lblArticulo_5.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_5.Name = "_lblArticulo_5"
        Me._lblArticulo_13.Text = "Modelo"
        Me._lblArticulo_13.Size = New System.Drawing.Size(35, 13)
        Me._lblArticulo_13.Location = New System.Drawing.Point(11, 84)
        Me._lblArticulo_13.TabIndex = 5
        Me._lblArticulo_13.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_13.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_13.Enabled = True
        Me._lblArticulo_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_13.UseMnemonic = True
        Me._lblArticulo_13.Visible = True
        Me._lblArticulo_13.AutoSize = True
        Me._lblArticulo_13.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_13.Name = "_lblArticulo_13"
        Me._lblArticulo_12.Text = "Marca"
        Me._lblArticulo_12.Size = New System.Drawing.Size(30, 13)
        Me._lblArticulo_12.Location = New System.Drawing.Point(11, 52)
        Me._lblArticulo_12.TabIndex = 2
        Me._lblArticulo_12.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_12.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_12.Enabled = True
        Me._lblArticulo_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_12.UseMnemonic = True
        Me._lblArticulo_12.Visible = True
        Me._lblArticulo_12.AutoSize = True
        Me._lblArticulo_12.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_12.Name = "_lblArticulo_12"
        Me._lblArticulo_15.Text = "Género"
        Me._lblArticulo_15.Size = New System.Drawing.Size(35, 13)
        Me._lblArticulo_15.Location = New System.Drawing.Point(11, 116)
        Me._lblArticulo_15.TabIndex = 8
        Me._lblArticulo_15.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_15.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_15.Enabled = True
        Me._lblArticulo_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_15.UseMnemonic = True
        Me._lblArticulo_15.Visible = True
        Me._lblArticulo_15.AutoSize = True
        Me._lblArticulo_15.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_15.Name = "_lblArticulo_15"
        Me._lblArticulo_17.Text = "Tipo de Material"
        Me._lblArticulo_17.Size = New System.Drawing.Size(76, 13)
        Me._lblArticulo_17.Location = New System.Drawing.Point(11, 182)
        Me._lblArticulo_17.TabIndex = 18
        Me._lblArticulo_17.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._lblArticulo_17.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_17.Enabled = True
        Me._lblArticulo_17.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_17.UseMnemonic = True
        Me._lblArticulo_17.Visible = True
        Me._lblArticulo_17.AutoSize = True
        Me._lblArticulo_17.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._lblArticulo_17.Name = "_lblArticulo_17"
        Me.FrameImagen.SetIndex(_FrameImagen_0, CType(0, Integer))
        Me.fraArticulo.SetIndex(_fraArticulo_2, CType(2, Integer))
        Me.fraArticulo.SetIndex(_fraArticulo_1, CType(1, Integer))
        Me.fraJoyeria.SetIndex(_fraJoyeria_0, CType(0, Integer))
        Me.fraMoneda.SetIndex(_fraMoneda_5, CType(5, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_32, CType(32, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_31, CType(31, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_6, CType(6, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_7, CType(7, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_8, CType(8, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_9, CType(9, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_33, CType(33, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_16, CType(16, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_4, CType(4, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_5, CType(5, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_13, CType(13, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_12, CType(12, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_15, CType(15, Integer))
        Me.lblArticulo.SetIndex(_lblArticulo_17, CType(17, Integer))
        Me.optGenero.SetIndex(_optGenero_0, CType(0, Integer))
        Me.optGenero.SetIndex(_optGenero_1, CType(1, Integer))
        Me.optGenero.SetIndex(_optGenero_2, CType(2, Integer))
        Me.optMoneda.SetIndex(_optMoneda_1, CType(1, Integer))
        Me.optMoneda.SetIndex(_optMoneda_0, CType(0, Integer))
        Me.optMovimiento.SetIndex(_optMovimiento_2, CType(2, Integer))
        Me.optMovimiento.SetIndex(_optMovimiento_1, CType(1, Integer))
        Me.optMovimiento.SetIndex(_optMovimiento_0, CType(0, Integer))
        CType(Me.optMovimiento, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optGenero, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblArticulo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraJoyeria, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraArticulo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FrameImagen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dbcUnidad, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dbcMaterial, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dbcModelo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dbcMarca, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dbcOrigen, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Controls.Add(_fraJoyeria_0)
        Me.Controls.Add(btnAceptar)
        Me.Controls.Add(btnCancelar)
        Me.Controls.Add(txtAdicional)
        Me.Controls.Add(_fraArticulo_2)
        Me.Controls.Add(btnTipoMaterial)
        Me.Controls.Add(btnModelo)
        Me.Controls.Add(btnMarca)
        Me.Controls.Add(chkCrono)
        Me.Controls.Add(_fraArticulo_1)
        Me.Controls.Add(dbcMarca)
        Me.Controls.Add(dbcModelo)
        Me.Controls.Add(dbcMaterial)
        Me.Controls.Add(dbcUnidad)
        Me.Controls.Add(_lblArticulo_33)
        Me.Controls.Add(_lblArticulo_16)
        Me.Controls.Add(txtDescripcion)
        Me.Controls.Add(_lblArticulo_4)
        Me.Controls.Add(_lblArticulo_5)
        Me.Controls.Add(_lblArticulo_13)
        Me.Controls.Add(_lblArticulo_12)
        Me.Controls.Add(_lblArticulo_15)
        Me.Controls.Add(_lblArticulo_17)
        Me._fraJoyeria_0.Controls.Add(txtCodigodelProveedor)
        Me._fraJoyeria_0.Controls.Add(txtPrecioenDolares)
        Me._fraJoyeria_0.Controls.Add(_fraMoneda_5)
        Me._fraJoyeria_0.Controls.Add(Frame3)
        Me._fraJoyeria_0.Controls.Add(txtCostoActual)
        Me._fraJoyeria_0.Controls.Add(txtCantidadCompra)
        Me._fraJoyeria_0.Controls.Add(_FrameImagen_0)
        Me._fraJoyeria_0.Controls.Add(_lblArticulo_6)
        Me._fraJoyeria_0.Controls.Add(_lblArticulo_7)
        Me._fraJoyeria_0.Controls.Add(_lblArticulo_8)
        Me._fraJoyeria_0.Controls.Add(_lblArticulo_9)
        Me._fraMoneda_5.Controls.Add(_optMoneda_1)
        Me._fraMoneda_5.Controls.Add(_optMoneda_0)
        Me.Frame3.Controls.Add(txtCodArtAnterior)
        Me.Frame3.Controls.Add(chkCodAnt)
        Me.Frame3.Controls.Add(dbcOrigen)
        Me.Frame3.Controls.Add(_lblArticulo_32)
        Me.Frame3.Controls.Add(_lblArticulo_31)
        Me._FrameImagen_0.Controls.Add(cmdBuscarImagen)
        Me._FrameImagen_0.Controls.Add(txtImagen)
        Me._fraArticulo_2.Controls.Add(_optMovimiento_2)
        Me._fraArticulo_2.Controls.Add(_optMovimiento_1)
        Me._fraArticulo_2.Controls.Add(_optMovimiento_0)
        Me._fraArticulo_1.Controls.Add(_optGenero_0)
        Me._fraArticulo_1.Controls.Add(_optGenero_1)
        Me._fraArticulo_1.Controls.Add(_optGenero_2)
        Me._fraJoyeria_0.ResumeLayout(False)
        Me._fraMoneda_5.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me._FrameImagen_0.ResumeLayout(False)
        Me._fraArticulo_2.ResumeLayout(False)
        Me._fraArticulo_1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Public Function BuscaUnidad(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codUnidad, DescUnidad FROM CatUnidades WHERE CodUnidad = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaUnidad = Trim(rsLocal.Fields("DescUnidad").Value)
        Else
            BuscaUnidad = cINDEFINIDA
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function BuscaTipoMaterialDescCorta(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codTipoMaterial, DescCorta FROM CatTipoMaterial WHERE CodTipoMaterial = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaTipoMaterialDescCorta = Trim(rsLocal.Fields("DescCorta").Value)
        Else
            BuscaTipoMaterialDescCorta = cINDEFINIDO
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaMarca(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codMarca, DescMarca FROM CatMarcas WHERE CodGrupo = " & gCODRELOJERIA & " AND CodMarca = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaMarca = Trim(rsLocal.Fields("DescMarca").Value)
        Else
            BuscaMarca = cINDEFINIDA
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Function BuscaModelo(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codModelo, DescModelo FROM CatModelos WHERE CodGrupo = " & gCODRELOJERIA & " AND CodMarca = " & mintCodMarca & " AND CodModelo = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaModelo = Trim(rsLocal.Fields("DescModelo").Value)
        Else
            BuscaModelo = cINDEFINIDO
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaTipoMaterial(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codTipoMaterial, DescTipoMaterial FROM CatTipoMaterial WHERE CodTipoMaterial = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaTipoMaterial = Trim(rsLocal.Fields("DescTipoMaterial").Value)
        Else
            BuscaTipoMaterial = cINDEFINIDO
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub LLenaForma(ByRef Columna As Integer, ByRef Renglon As Integer, ByRef RenglonAct As Integer, ByRef nCodUnidad As Integer, ByRef nCodMarca As Integer, ByRef nCodModelo As Integer, ByRef nCodMaterial As Integer, ByRef cDescripcion As String, ByRef pGenero As String, ByRef pMovimiento As String, ByRef pCrono As String, ByRef cCodArticuloProv As String, ByRef cAdicional As String, ByRef nPrecioPublico As Decimal, ByRef bMonedaPP As String, ByRef nOrigenAnt As Integer, ByRef nCodigoAnt As Integer, ByRef nCantidadCompra As Integer, ByRef nCostoActual As Decimal, ByRef cImagen As String)

        On Error GoTo Merr
        nCol = Columna
        nRow = Renglon
        nRowAct = RenglonAct

        If CInt(Numerico(frmCXPOrdenCompra.mshFlex.get_TextMatrix(nRowAct, C_COLCODIGO))) = 0 Then
            '''nuevo
            dbcUnidad.Text = False
            dbcMarca.Text = False
            dbcModelo.Text = False
            dbcMaterial.Text = False
            fraArticulo(1).Enabled = True
            fraArticulo(2).Enabled = True
            chkCrono.Enabled = True
            txtCodigodelProveedor.ReadOnly = False
            txtAdicional.ReadOnly = False
        Else
            '''resurtido
            dbcUnidad.Text = True
            dbcMarca.Text = True
            dbcModelo.Text = True
            dbcMaterial.Text = True
            fraArticulo(1).Enabled = False
            fraArticulo(2).Enabled = False
            chkCrono.Enabled = False
            txtCodigodelProveedor.ReadOnly = True
            txtAdicional.ReadOnly = True
        End If

        gstrNombreForma = "FRMCXPRELOJERIA"
        mblnFueraChange = True
        mintCodUnidad = nCodUnidad
        If mintCodUnidad = 0 Then
            dbcUnidad_Enter(dbcUnidad, New System.EventArgs())
            dbcUnidad.Text = Trim(cINDEFINIDA)
        Else
            dbcUnidad.Text = BuscaUnidad(mintCodUnidad)
        End If
        dbcUnidad.Tag = dbcUnidad.Text

        mintCodMarca = nCodMarca
        dbcMarca.Text = BuscaMarca(mintCodMarca)
        dbcMarca.Tag = dbcMarca.Text
        mintCodModelo = nCodModelo
        dbcModelo.Text = BuscaModelo(mintCodModelo)
        dbcModelo.Tag = dbcModelo.Text
        mintCodMaterial = nCodMaterial
        dbcMaterial.Text = BuscaTipoMaterial(mintCodMaterial)
        dbcMaterial.Tag = dbcMaterial.Text
        cMaterial = BuscaTipoMaterialDescCorta(mintCodMaterial)
        txtDescripcion.Text = Trim(cDescripcion)
        txtDescripcion.Tag = txtDescripcion.Text
        mblnFueraChange = False
        'Género
        Select Case pGenero
            Case ""
                optGenero(0).Checked = False
                optGenero(1).Checked = False
                optGenero(2).Checked = False
            Case "H"
                optGenero(0).Checked = True
                cGenero = Trim(pGenero)
                cGeneroTag = cGenero
            Case "D"
                optGenero(1).Checked = True
                cGenero = Trim(pGenero)
                cGeneroTag = cGenero
            Case "M"
                cGenero = pGenero
                cGeneroTag = cGenero
                optGenero(2).Checked = True
        End Select

        'Movimiento
        Select Case pMovimiento
            Case "Q"
                optMovimiento(0).Checked = True
                cMovimiento = Trim(pMovimiento)
                cMovimientoTag = cMovimiento
            Case "AUT"
                optMovimiento(1).Checked = True
                cMovimiento = Trim(pMovimiento)
                cMovimientoTag = cMovimiento
            Case "MAN"
                cMovimiento = Trim(pMovimiento)
                cMovimientoTag = cMovimiento
                optMovimiento(2).Checked = True
            Case ""
                optMovimiento(0).Checked = False
                optMovimiento(1).Checked = False
                optMovimiento(2).Checked = False
        End Select

        If CBool(pCrono) Then
            chkCrono.CheckState = System.Windows.Forms.CheckState.Checked
            lCrono = True
            cCrono = "CHR"
        Else
            chkCrono.CheckState = System.Windows.Forms.CheckState.Unchecked
            lCrono = False
            cCrono = ""
        End If
        cCronoTag = cCrono

        '''Datos Adicionales
        txtAdicional.Text = cAdicional
        txtAdicional.Tag = cAdicional
        If nRowAct = nRow Then '''cuando es el mismo renglon es consulta de un articulo y se deben mostrar los datos completos
            txtCodigodelProveedor.Text = Trim(cCodArticuloProv)
            txtCodigodelProveedor.Tag = txtCodigodelProveedor.Text
            txtPrecioenDolares.Text = Format(nPrecioPublico, "###,##0.00")
            txtPrecioenDolares.Tag = Format(nPrecioPublico, "###,##0.00")

            If CDbl(Trim(CStr(nCodigoAnt))) > 0 Then
                chkCodAnt.CheckState = System.Windows.Forms.CheckState.Checked
                dbcOrigen.Text = nOrigenAnt
                dbcOrigen.Tag = nOrigenAnt
                txtCodArtAnterior.Text = CStr(nCodigoAnt)
                txtCodArtAnterior.Tag = nCodigoAnt
            Else
                chkCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked
                dbcOrigen.Text = ""
                dbcOrigen.Tag = ""
                txtCodArtAnterior.Text = ""
                txtCodArtAnterior.Tag = ""
            End If
            txtCantidadCompra.Text = CStr(nCantidadCompra)
            txtCantidadCompra.Tag = nCantidadCompra
            txtCostoActual.Text = Format(nCostoActual, "###,##0.00")
            txtCostoActual.Tag = Format(nCostoActual, "###,##0.00")
            If nRow = nRowAct Then
                txtImagen.Text = cImagen
                txtImagen.Tag = cImagen
            Else
                txtImagen.Text = ""
                txtImagen.Tag = ""
            End If
            If bMonedaPP = "" Then
                optMoneda(0).Checked = True
            Else
                If bMonedaPP = "D" Then optMoneda(0).Checked = True Else optMoneda(1).Checked = True
            End If
        Else
            '''Datos Adicionales - siempre deben de ser inicializados cuando es un articulo nuevo de la orden de compra
            txtCodigodelProveedor.Text = ""
            txtCodigodelProveedor.Tag = ""
            txtPrecioenDolares.Text = Format(0, "###,##0.00")
            txtPrecioenDolares.Tag = Format(0, "###,##0.00")
            FormaDescripcion()

            chkCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked
            dbcOrigen.Text = ""
            dbcOrigen.Tag = ""
            txtCodArtAnterior.Text = ""
            txtCodArtAnterior.Tag = ""

            txtCantidadCompra.Text = "0"
            txtCantidadCompra.Tag = "0"
            txtCostoActual.Text = Format(0, "###,##0.00")
            txtCostoActual.Tag = Format(0, "###,##0.00")
            txtImagen.Text = ""
            txtImagen.Tag = ""
            optMoneda(0).Checked = False
            optMoneda(1).Checked = False
        End If
        FormaDescripcion()

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub LlenaDatos(ByRef CodFolio As String, ByRef CodArticulo As String, ByRef Columna As Integer, ByRef Renglon As Integer, ByRef RenglonAct As Integer, ByRef nCantidadCompra As Integer, ByRef nCostoActual As Decimal, ByRef cImagen As String, ByRef Tipo As String)
        On Error GoTo Merr
        Dim lResurtido As Boolean

        nCol = Columna
        nRow = Renglon
        nRowAct = RenglonAct

        If Tipo = "1C" Then
            lResurtido = True
            gStrSql = "SELECT * FROM CatArticulos (Nolock) WHERE CodArticulo = " & Trim(frmCXPOrdenCompra.mshFlex.get_TextMatrix(nRow, C_COLCODIGO))
        ElseIf Tipo = "2C" Then
            '''ojo quitar - aux
            lResurtido = False
            gStrSql = "SELECT * FROM OrdenesCompraPreCat (Nolock) WHERE FolioOrdenCompra = '" & Trim(CodFolio) & "' and NumPartida = " & CodArticulo
        End If

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            mblnFueraChange = True
            mintCodUnidad = IIf(IsDBNull(RsGral.Fields("CodUnidad").Value), 0, RsGral.Fields("CodUnidad").Value)
            dbcUnidad.Text = BuscaUnidad(mintCodUnidad)
            dbcUnidad.Tag = dbcUnidad.Text
            mintCodMarca = IIf(IsDBNull(RsGral.Fields("CodMArca").Value), 0, RsGral.Fields("CodMArca").Value)
            dbcMarca.Text = BuscaMarca(mintCodMarca)
            dbcMarca.Tag = dbcMarca.Text
            mintCodModelo = IIf(IsDBNull(RsGral.Fields("CodModelo").Value), 0, RsGral.Fields("CodModelo").Value)
            dbcModelo.Text = BuscaModelo(mintCodModelo)
            dbcModelo.Tag = dbcModelo.Text
            'Género
            Select Case Trim(RsGral.Fields("Genero").Value)
                Case "H"
                    optGenero(0).Checked = True
                Case "D"
                    optGenero(1).Checked = True
                Case "M"
                    optGenero(2).Checked = True
                Case ""
                    optGenero(0).Checked = False
                    optGenero(1).Checked = False
                    optGenero(2).Checked = False
            End Select
            cGenero = RsGral.Fields("Genero").Value
            cGeneroTag = RsGral.Fields("Genero").Value
            'Movimiento
            Select Case Trim(RsGral.Fields("Movimiento").Value)
                Case "Q"
                    optMovimiento(0).Checked = True
                Case "AUT"
                    optMovimiento(1).Checked = True
                Case "MAN"
                    optMovimiento(2).Checked = True
                Case ""
                    optMovimiento(0).Checked = False
                    optMovimiento(1).Checked = False
                    optMovimiento(2).Checked = False
            End Select
            cMovimiento = RsGral.Fields("Movimiento").Value
            cMovimientoTag = cMovimiento

            mintCodMaterial = IIf(IsDBNull(RsGral.Fields("CodTipoMaterial").Value), 0, RsGral.Fields("CodTipoMaterial").Value)
            dbcMaterial.Text = BuscaTipoMaterial(mintCodMaterial)
            dbcMaterial.Tag = dbcMaterial.Text
            cMaterial = BuscaTipoMaterialDescCorta(mintCodMaterial)

            txtAdicional.Text = RsGral.Fields("Adicional").Value
            txtAdicional.Tag = RsGral.Fields("Adicional").Value

            'Cronómetro
            Select Case True
                Case RsGral.Fields("Crono").Value
                    chkCrono.CheckState = System.Windows.Forms.CheckState.Checked
                    lCrono = True
                    cCrono = "CHR"
                Case Else
                    chkCrono.CheckState = System.Windows.Forms.CheckState.Unchecked
                    lCrono = False
                    cCrono = ""
            End Select
            cCronoTag = cCrono
            FormaDescripcion()
            txtDescripcion.Text = Trim(RsGral.Fields("DescArticulo").Value)
            txtDescripcion.Tag = txtDescripcion.Text
            mblnFueraChange = False

            txtCodigodelProveedor.Text = Trim(RsGral.Fields("CodigoArticuloProv").Value)
            txtCodigodelProveedor.Tag = txtCodigodelProveedor.Text

            txtPrecioenDolares.Text = Format(RsGral.Fields("PrecioPubDolar").Value, "###,##0.00")
            txtPrecioenDolares.Tag = Format(RsGral.Fields("PrecioPubDolar").Value, "###,##0.00")

            If CDbl(Trim(RsGral.Fields("CodigoAnt").Value)) > 0 Then
                chkCodAnt.CheckState = System.Windows.Forms.CheckState.Checked
                dbcOrigen.Text = RsGral.Fields("OrigenAnt").Value
                dbcOrigen.Tag = RsGral.Fields("OrigenAnt").Value
                txtCodArtAnterior.Text = RsGral.Fields("CodigoAnt").Value
                txtCodArtAnterior.Tag = RsGral.Fields("CodigoAnt").Value
            Else
                chkCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked
                dbcOrigen.Text = ""
                dbcOrigen.Tag = ""
                txtCodArtAnterior.Text = ""
                txtCodArtAnterior.Tag = ""
            End If

            txtCantidadCompra.Text = CStr(nCantidadCompra)
            txtCantidadCompra.Tag = nCantidadCompra
            txtCostoActual.Text = Format(nCostoActual, "###,##0.00")
            txtCostoActual.Tag = Format(nCostoActual, "###,##0.00")
            If nRow = nRowAct Then
                txtImagen.Text = cImagen
                txtImagen.Tag = cImagen
            Else
                txtImagen.Text = ""
                txtImagen.Tag = ""
            End If

            '''Si es resurtido sale del AbcArticulos si no del PreCat
            If lResurtido Then
                If RsGral.Fields("PesosFijos").Value Then optMoneda(1).Checked = True Else optMoneda(0).Checked = True
            Else
                If RsGral.Fields("MonedaPP").Value = "P" Then optMoneda(1).Checked = True Else optMoneda(0).Checked = True
            End If

            If CInt(Numerico(frmCXPOrdenCompra.mshFlex.get_TextMatrix(nRow, C_COLCODIGO))) = 0 Then
                dbcUnidad.Text = False
                dbcMarca.Text = False
                dbcModelo.Text = False
                dbcMaterial.Text = False
                fraArticulo(1).Enabled = True
                fraArticulo(2).Enabled = True
                chkCrono.Enabled = True
                txtCodigodelProveedor.ReadOnly = False
                txtAdicional.ReadOnly = False
            Else
                dbcUnidad.Text = True
                dbcMarca.Text = True
                dbcModelo.Text = True
                dbcMaterial.Text = True
                fraArticulo(1).Enabled = False
                fraArticulo(2).Enabled = False
                chkCrono.Enabled = False
                txtCodigodelProveedor.ReadOnly = True
                txtAdicional.ReadOnly = True
            End If

        Else
            MsgBox("No se ha localizado el artículo al que hace referencia", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function Cambios() As Boolean
        On Error Resume Next
        Select Case True
            Case Trim(dbcUnidad.Text) <> Trim(dbcUnidad.Tag)
                Cambios = True
            Case Trim(dbcMarca.Text) <> Trim(dbcMarca.Tag)
                Cambios = True
            Case Trim(dbcModelo.Text) <> Trim(dbcModelo.Tag)
                Cambios = True
            Case Trim(dbcMaterial.Text) <> Trim(dbcMaterial.Tag)
                Cambios = True
            Case Trim(cMovimiento) <> Trim(cMovimientoTag)
                Cambios = True
            Case Trim(cCrono) <> Trim(cCronoTag)
                Cambios = True
            Case Trim(cGenero) <> Trim(cGeneroTag)
                Cambios = True
            Case Trim(txtAdicional.Text) <> Trim(txtAdicional.Tag)
                Cambios = True
            Case Trim(txtCodigodelProveedor.Text) <> Trim(txtCodigodelProveedor.Tag)
                Cambios = True

            Case Trim(txtCostoActual.Text) <> Trim(txtCostoActual.Tag)
                Cambios = True
            Case Trim(txtCantidadCompra.Text) <> Trim(txtCantidadCompra.Tag)
                Cambios = True
            Case Trim(txtPrecioenDolares.Text) <> Trim(txtPrecioenDolares.Tag)
                Cambios = True
            Case optMoneda(0).Checked <> CBool(IIf(optMoneda(0).Tag = "", "0", optMoneda(0).Tag))
                Cambios = True
            Case optMoneda(1).Checked <> CBool(IIf(optMoneda(1).Tag = "", "0", optMoneda(1).Tag))
                Cambios = True

            Case chkCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And Trim(dbcOrigen.Text) <> Trim(dbcOrigen.Tag) And Trim(txtCodArtAnterior.Text) <> Trim(txtCodArtAnterior.Tag)
                Cambios = True
            Case chkCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And Trim(dbcOrigen.Text) <> Trim(dbcOrigen.Tag) And Trim(txtCodArtAnterior.Text) = Trim(txtCodArtAnterior.Tag)
                Cambios = True
            Case chkCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And Trim(dbcOrigen.Text) = Trim(dbcOrigen.Tag) And Trim(txtCodArtAnterior.Text) <> Trim(txtCodArtAnterior.Tag)
                Cambios = True

            Case Trim(txtImagen.Text) <> Trim(txtImagen.Tag) And Trim(txtImagen.Text) <> ""
                Cambios = True

            Case Else
                Cambios = False
        End Select
    End Function

    Public Function ValidaDatos() As Boolean
        Select Case True
            Case mintCodUnidad = 0
                MsgBox("Falta indicar la(s) Unidad(es)", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                If dbcUnidad.Enabled Then dbcUnidad.Focus() Else btnAceptar.Focus()
                Exit Function
            Case mintCodMarca = 0
                MsgBox("Falta indicar la Marca del Artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                If dbcMarca.Enabled Then dbcMarca.Focus() Else btnAceptar.Focus()
                Exit Function
            Case Not optGenero(0).Checked And Not optGenero(1).Checked And Not optGenero(2).Checked
                MsgBox("Debe indicar el género del artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                If fraArticulo(1).Enabled Then optGenero(0).Focus() Else btnAceptar.Focus()
                Exit Function
            Case Not optMovimiento(0).Checked And Not optMovimiento(1).Checked And Not optMovimiento(2).Checked
                MsgBox("Debe indicar el funcionamiento del artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                If fraArticulo(2).Enabled Then optMovimiento(0).Focus() Else btnAceptar.Focus()
                Exit Function
            Case mintCodMaterial = 0
                MsgBox("Falta indicar el Tipo de Material del Artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                If dbcMaterial.Enabled Then dbcMaterial.Focus() Else btnAceptar.Focus()
                Exit Function
            Case (Trim(txtPrecioenDolares.Text) <> "") And (Not optMoneda(0).Checked And Not optMoneda(1).Checked)
                MsgBox("Debe indicar la moneda del precio publico del artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                If txtPrecioenDolares.Enabled Then txtPrecioenDolares.Focus() Else btnAceptar.Focus()
                Exit Function
            Case chkCodAnt.CheckState = System.Windows.Forms.CheckState.Checked
                If Trim(dbcOrigen.Text) = "" Then
                    MsgBox("Debe indicar el origen anterior del artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    ValidaDatos = False
                    If dbcOrigen.Enabled Then dbcOrigen.Focus() Else btnAceptar.Focus()
                    Exit Function
                End If
                If Trim(txtCodArtAnterior.Text) = "" Then
                    MsgBox("Debe indicar el código anterior del artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    ValidaDatos = False
                    If txtCodArtAnterior.Enabled Then txtCodArtAnterior.Focus() Else btnAceptar.Focus()
                    Exit Function
                End If
                ValidaDatos = True
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Public Function Guardar() As Boolean
        Dim I As Integer
        Dim lCrono, lGen, lMar, lMod, lMov, lTipoM As String

        lMar = "" : lMod = "" : lGen = "" : lMov = "" : lCrono = "" : lTipoM = ""
        If Not ValidaDatos() Then Exit Function

        'Realiza una búsqueda POR DESCRIPCIÓN para ver si el artículo ya existe en el Grid
        With frmCXPOrdenCompra.mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_ColDESCRIPCION)) = "" Then Exit For

                '''valida la clasificación del artículo
                If frmCXPOrdenCompra.mintRenglonAct <> I Then
                    If mintCodMarca = CDec(Numerico(.get_TextMatrix(I, C_COLCODMARCA))) And mintCodModelo = CDec(Numerico(.get_TextMatrix(I, C_COLCODMODELO))) And Trim(cGenero) = Trim(.get_TextMatrix(I, C_COLGENERO)) And Trim(cMovimiento) = Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)) And IIf(Trim(cCrono) = "CHR", True, False) = CBool(Trim(.get_TextMatrix(I, C_COLCRONO))) And mintCodMaterial = CDec(Numerico(.get_TextMatrix(I, C_COLCODTIPOMATERIAL))) And Trim(txtAdicional.Text) = Trim(.get_TextMatrix(I, C_COLADICIONAL)) And Trim(txtCodigodelProveedor.Text) = Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)) Then

                        Select Case MsgBox("Ya existe un elemento de éstos en la Orden de Compra." & vbNewLine & "¿Desea volver a definir el artículo?", MsgBoxStyle.Information + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                            Case MsgBoxResult.Yes
                                dbcMarca.Focus()
                                ModEstandar.SelTxt()
                                Exit Function
                            Case MsgBoxResult.No
                                mblnCancelar = True
                                Me.Close()
                                Exit Function
                            Case MsgBoxResult.Cancel
                                Exit Function
                        End Select
                    End If
                End If

                '''valida que el codigo anterior no se repita, pero es decisión del usuario si lo duplica o no
                If frmCXPOrdenCompra.mintRenglonAct <> I Then
                    If CDec(Numerico(dbcOrigen.Text)) = CDec(Numerico(.get_TextMatrix(I, C_COLORIGENANT))) And CDec(Numerico((txtCodArtAnterior.Text))) = CDec(Numerico(.get_TextMatrix(I, C_ColCODIGOANT))) And (CDec(Numerico((txtCodArtAnterior.Text))) > 0) Then
                        Select Case MsgBox("Ya se definió este codigo anterior para otro artículo en la Orden de Compra" & vbNewLine & "Desea modificarlo???", MsgBoxStyle.Information + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                            Case MsgBoxResult.Yes
                                dbcOrigen.Focus()
                                ModEstandar.SelTxt()
                                Exit Function
                        End Select
                    End If
                End If
                '''valida que el codigo del articulo del prov no se repita, pero es decisión del usuario si lo duplica o no
                If frmCXPOrdenCompra.mintRenglonAct <> I Then
                    If Trim(txtCodigodelProveedor.Text) = Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)) And Trim(txtCodigodelProveedor.Text) <> "" Then
                        Select Case MsgBox("Ya se definió este codigo del proveedor para otro artículo en la Orden de Compra" & vbNewLine & "Desea modificarlo???", MsgBoxStyle.Information + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                            Case MsgBoxResult.Yes
                                txtCodigodelProveedor.Focus()
                                ModEstandar.SelTxt()
                                Exit Function
                        End Select
                    End If
                End If
                '''valida que el archivo de la imagen no sea el mismo
                If frmCXPOrdenCompra.mintRenglonAct <> I Then
                    If Trim(txtImagen.Text) = Trim(.get_TextMatrix(I, C_ColIMAGEN)) And Trim(txtImagen.Text) <> "" Then
                        Select Case MsgBox("Ya se definió esta imagen para otro artículo en la Orden de Compra" & vbNewLine & "Desea modificarlo???", MsgBoxStyle.Information + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                            Case MsgBoxResult.Yes
                                txtImagen.Focus()
                                ModEstandar.SelTxt()
                                Exit Function
                        End Select
                    End If
                End If

            Next I
        End With

        'Busca el artículo por nombre en la tabla CatArticulos
        '''gStrSql = "SELECT CodArticulo, DescArticulo FROM CatArticulos WHERE CodProveedor = " & frmCXPOrdenCompra.mintCodProveedor & " and LTrim(RTrim(DescArticulo)) = '" & Trim(txtDescripcion.Caption) & "'"
        '''gStrSql = " SELECT * FROM CatArticulos WHERE CodProveedor = " & frmCXPOrdenCompra.mintCodProveedor & _
        '" And CodMarca = " & mintCodMarca & " And CodModelo = " & mintCodModelo & _
        '" And Genero = '" & IIf(optGenero(0).Value, "H", IIf(optGenero(1).Value, "D", IIf(optGenero(0).Value, "M", ""))) & _
        '"' And Movimiento = '" & IIf(optMovimiento(0).Value, "Q", IIf(optMovimiento(1).Value, "AUT", IIf(optMovimiento(2).Value, "MAN", ""))) & "' And Crono = " & IIf(chkCrono.Value = vbChecked, 1, 0) & "" & _
        '" And CodTipoMaterial = " & mintCodMaterial & " And Adicional = '" & Trim(txtAdicional.text) & _
        '"' And CodigoArticuloProv = '" & Trim(txtCodigodelProveedor.text) & "' "

        DefineCondicionesRel(lMar, lMod, lGen, lMov, lCrono, lTipoM)
        gStrSql = " SELECT * FROM CatArticulos (Nolock) WHERE CodProveedor = " & frmCXPOrdenCompra.mintCodProveedor & " And " & lMar & " And " & lMod & " And " & lGen & " And " & lMov & " And " & lCrono & " And " & lTipoM & " And ltrim(rtrim(Adicional)) = '" & Trim(txtAdicional.Text) & "' And ltrim(rtrim(CodigoArticuloProv)) = '" & Trim(txtCodigodelProveedor.Text) & "' "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            'De existir, llena los demás datos y cambia el estatus del registro en el Grid a Resurtido
            With frmCXPOrdenCompra.mshFlex

                .set_TextMatrix(nRowAct, C_ColDESCRIPCION, Trim(txtDescripcion.Text))
                .set_TextMatrix(nRowAct, C_ColUNIDAD, Trim(dbcUnidad.Text))
                .set_TextMatrix(nRowAct, C_ColCODGRUPO, gCODRELOJERIA)
                .set_TextMatrix(nRowAct, C_COLCODFAMILIA, 0)
                .set_TextMatrix(nRowAct, C_COLCODLINEA, 0)
                .set_TextMatrix(nRowAct, C_COLCODSUBLINEA, 0)
                .set_TextMatrix(nRowAct, C_COLCODKILATES, 0)
                .set_TextMatrix(nRowAct, C_COLCODMARCA, mintCodMarca)
                .set_TextMatrix(nRowAct, C_COLCODMODELO, mintCodModelo)
                .set_TextMatrix(nRowAct, C_COLCODTIPOMATERIAL, mintCodMaterial)
                .set_TextMatrix(nRowAct, C_COLGENERO, cGenero)
                .set_TextMatrix(nRowAct, C_COLMOVIMIENTO, cMovimiento)
                .set_TextMatrix(nRowAct, C_COLCRONO, IIf(cCrono = "", "False", "True"))
                .set_TextMatrix(nRowAct, C_COLCODIGOARTICULOPROV, Trim(txtCodigodelProveedor.Text))
                .set_TextMatrix(nRowAct, C_COLSTATUS, "")
                .set_TextMatrix(nRowAct, C_COLDESCTO, 0)
                .set_TextMatrix(nRowAct, C_COLDESCTOPORC, Numerico((frmCXPOrdenCompra.txtPorcDescto.Text)))
                .set_TextMatrix(nRowAct, C_COLDESCTOPORCTAG, Numerico((frmCXPOrdenCompra.txtPorcDescto.Text)))

                .set_TextMatrix(nRowAct, C_COLADICIONAL, Trim(txtAdicional.Text))
                .set_TextMatrix(nRowAct, C_COLPRECIOPUBDOLAR, CDec(Numerico(txtPrecioenDolares.Text)))
                .set_TextMatrix(nRowAct, C_COLMONEDAPP, IIf(optMoneda(0).Checked = True, "D", "P"))
                .set_TextMatrix(nRowAct, C_COLPRECIOUNITARIO, CDec(Numerico((txtCostoActual.Text))))
                .set_TextMatrix(nRowAct, C_ColCANTIDAD, CDec(Numerico((txtCantidadCompra.Text))))
                .set_TextMatrix(nRowAct, C_COLORIGENANT, IIf(Trim(dbcOrigen.Text) = "", "", CShort(Numerico((dbcOrigen.Text)))))
                .set_TextMatrix(nRowAct, C_ColCODIGOANT, IIf(Trim(txtCodArtAnterior.Text) = "", "", CInt(Numerico(txtCodArtAnterior.Text))))
                .set_TextMatrix(nRowAct, C_ColIMAGEN, Trim(txtImagen.Text))
                If Cambios() Then .set_TextMatrix(nRowAct, C_COLSTATUSX, "M")

                'Busca el artículo en el inventario, de existir, entonces es un resurtido
                gStrSql = "SELECT CodArticulo FROM Inventario WHERE CodArticulo = " & rsLocal.Fields("CodArticulo").Value
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.UP_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                rsLocal = Cmd.Execute
                If rsLocal.RecordCount > 0 Then
                    .set_TextMatrix(nRowAct, C_COLSTATUS, C_RESURTIDO)
                    MsgBox("Ya existe un artículo con esta descripción en el sistema. Si considera que este artículo" & vbNewLine & "no debe ser un resurtido de mercancía existente, revise sus datos", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Advertencia ...")
                Else
                    .set_TextMatrix(nRowAct, C_COLSTATUS, "")
                End If

                .Row = frmCXPOrdenCompra.mintRenglonAct + 1
                .Col = C_ColDESCRIPCION
            End With
        Else
            'Si no existe deja el estatus en Vigente y llena los demás datos
            With frmCXPOrdenCompra.mshFlex
                .set_TextMatrix(nRowAct, C_COLCODIGO, "")
                .set_TextMatrix(nRowAct, C_ColDESCRIPCION, Trim(txtDescripcion.Text))
                .set_TextMatrix(nRowAct, C_ColUNIDAD, Trim(dbcUnidad.Text))
                .set_TextMatrix(nRowAct, C_ColCODGRUPO, gCODRELOJERIA)
                .set_TextMatrix(nRowAct, C_COLCODFAMILIA, 0)
                .set_TextMatrix(nRowAct, C_COLCODLINEA, 0)
                .set_TextMatrix(nRowAct, C_COLCODSUBLINEA, 0)
                .set_TextMatrix(nRowAct, C_COLCODKILATES, 0)
                .set_TextMatrix(nRowAct, C_COLCODMARCA, mintCodMarca)
                .set_TextMatrix(nRowAct, C_COLCODMODELO, mintCodModelo)
                .set_TextMatrix(nRowAct, C_COLCODTIPOMATERIAL, mintCodMaterial)
                .set_TextMatrix(nRowAct, C_COLGENERO, cGenero)
                .set_TextMatrix(nRowAct, C_COLMOVIMIENTO, cMovimiento)
                .set_TextMatrix(nRowAct, C_COLCRONO, IIf(cCrono = "", "False", "True"))
                .set_TextMatrix(nRowAct, C_COLCODIGOARTICULOPROV, Trim(txtCodigodelProveedor.Text))
                .set_TextMatrix(nRowAct, C_COLSTATUS, "")
                .set_TextMatrix(nRowAct, C_COLDESCTO, 0)
                .set_TextMatrix(nRowAct, C_COLDESCTOPORC, Numerico((frmCXPOrdenCompra.txtPorcDescto.Text)))
                .set_TextMatrix(nRowAct, C_COLDESCTOPORCTAG, Numerico((frmCXPOrdenCompra.txtPorcDescto.Text)))

                .set_TextMatrix(nRowAct, C_COLADICIONAL, Trim(txtAdicional.Text))
                .set_TextMatrix(nRowAct, C_COLPRECIOPUBDOLAR, CDec(Numerico(txtPrecioenDolares.Text)))
                .set_TextMatrix(nRowAct, C_COLMONEDAPP, IIf(optMoneda(0).Checked = True, "D", "P"))
                .set_TextMatrix(nRowAct, C_COLPRECIOUNITARIO, CDec(Numerico((txtCostoActual.Text))))
                .set_TextMatrix(nRowAct, C_ColCANTIDAD, CDec(Numerico((txtCantidadCompra.Text))))
                .set_TextMatrix(nRowAct, C_COLORIGENANT, IIf(Trim(dbcOrigen.Text) = "", "", CShort(Numerico((dbcOrigen.Text)))))
                .set_TextMatrix(nRowAct, C_ColCODIGOANT, IIf(Trim(txtCodArtAnterior.Text) = "", "", CInt(Numerico(txtCodArtAnterior.Text))))
                .set_TextMatrix(nRowAct, C_ColIMAGEN, Trim(txtImagen.Text))
                If Cambios() Then .set_TextMatrix(nRowAct, C_COLSTATUSX, "M")

                '''actualiza datos para que ya no se consideren cambios recientes
                txtCostoActual.Tag = Trim(txtCostoActual.Text)
                txtCantidadCompra.Tag = Trim(txtCantidadCompra.Text)
                txtPrecioenDolares.Tag = Trim(txtPrecioenDolares.Text)
                If optMoneda(0).Checked Then
                    optMoneda(0).Tag = "1"
                    optMoneda(1).Tag = ""
                ElseIf optMoneda(1).Checked Then
                    optMoneda(1).Tag = "1"
                    optMoneda(0).Tag = ""
                End If

                .Row = frmCXPOrdenCompra.mintRenglonAct + 1
                .Col = C_ColDESCRIPCION
            End With
        End If
        frmCXPOrdenCompra.ActualizaCantidades()
        Guardar = True
    End Function

    Public Sub FormaDescripcion()
        cMarca = IIf(Trim(cMarca) = Trim(cINDEFINIDA), "", Trim(cMarca) & " ")
        cModelo = IIf(Trim(cModelo) = Trim(cINDEFINIDO), "", Trim(cModelo) & " ")
        cMaterial = IIf(Trim(cMaterial) = Trim(cINDEFINIDO), "", Trim(cMaterial) & " ")
        cAdicional = Trim(txtAdicional.Text)
        cDescripcion = cMarca & cModelo & Trim(cGenero) & " " & Trim(cMovimiento) & " " & Trim(cCrono) & " " & cMaterial & " " & cAdicional
        txtDescripcion.Text = cDescripcion
    End Sub

    Private Sub btnAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAceptar.Click
        If Guardar() Then
            mblnNuevo = True
            mblnCancelar = True
            Me.Close()
            'frmCXPOrdenCompra.MuestraClasificacion()
        End If
    End Sub

    Private Sub btnCancelar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancelar.Click
        mblnCancelar = True
        Me.Close()
    End Sub

    Private Sub btnMarca_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnMarca.Click
        Me.Enabled = False
        frmCorpoABCMarca.Tag = UCase(Me.Name)
        frmCorpoABCMarca.Show()
    End Sub

    Private Sub btnModelo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnModelo.Click
        If mintCodMarca = 0 Then
            MsgBox("Debe indicar la Marca de Relojería primero", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Me.dbcMarca.Focus()
        End If
        Me.Enabled = False
        'frmCorpoABCModelos.Tag = UCase(Me.Name)
        'frmCorpoABCModelos.Show()
        'frmCorpoABCModelos.mintCodMarca = mintCodMarca
        'frmCorpoABCModelos.LlenaDatos()
    End Sub

    Private Sub btnTipoMaterial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnTipoMaterial.Click
        Me.Enabled = False
        frmCorpoAbcTiposMaterial.Tag = UCase(Me.Name)
        frmCorpoAbcTiposMaterial.Show()
    End Sub

    Private Sub chkCodAnt_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCodAnt.CheckStateChanged
        If chkCodAnt.CheckState = System.Windows.Forms.CheckState.Checked Then
            dbcOrigen.Enabled = True
            txtCodArtAnterior.Enabled = True
        ElseIf chkCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            dbcOrigen.Enabled = False
            txtCodArtAnterior.Enabled = False
        End If
    End Sub

    Private Sub chkCrono_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCrono.CheckStateChanged
        If chkCrono.CheckState = System.Windows.Forms.CheckState.Checked Then
            cCrono = "CHR"
            lCrono = True
        Else
            cCrono = ""
            lCrono = False
        End If
        FormaDescripcion()
    End Sub

    Private Sub cmdBuscarImagen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBuscarImagen.Click
        'frmCorpoBuscarImagen.ShowDialog()
    End Sub

    Private Sub dbcMarca_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMarca.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        cMarca = Trim(Me.dbcMarca.Text)
        Call FormaDescripcion()

        If mblnFueraChange Then
            Exit Sub
        Else
            mblnFueraChange = True
            mintCodModelo = 0
            Me.dbcModelo.Text = cINDEFINIDO
            mblnFueraChange = False
        End If

        lStrSql = "SELECT codMarca, LTrim(RTrim(descMarca)) as descMarca FROM catMarcas Where codGrupo = " & gCODRELOJERIA & " and descMarca LIKE '" & Trim(Me.dbcMarca.Text) & "%' Order by descMarca "
        ModDCombo.DCChange(lStrSql, tecla, dbcMarca)

        If Trim(Me.dbcMarca.Text) = "" Then
            mintCodMarca = 0
        End If

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcMarca_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMarca.Enter
        Pon_Tool()
        gStrSql = "SELECT codMarca, LTrim(RTrim(descMarca)) as descMarca FROM catMarcas Where codGrupo = " & gCODRELOJERIA & " ORDER BY descMarca "
        ModDCombo.DCGotFocus(gStrSql, dbcMarca)
    End Sub

    Private Sub dbcMarca_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMarca.KeyDown
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Me.dbcUnidad.Focus()
            Case Else
                tecla = eventArgs.KeyCode
        End Select
    End Sub

    Private Sub dbcMarca_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMarca.Leave
        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        cDescripcion = Trim(Me.dbcMarca.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDA) Then
            cDescripcion = ""
        End If
        gStrSql = "SELECT codMarca, LTrim(RTrim(descMarca)) as descMarca FROM catMarcas Where codGrupo = " & gCODRELOJERIA & " and descMarca = '" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "'") & " Order by descMarca "
        Aux = mintCodMarca
        mintCodMarca = 0
        ModDCombo.DCLostFocus(dbcMarca, gStrSql, mintCodMarca)
        cMarca = Trim(Me.dbcMarca.Text)
        If Aux <> mintCodMarca Then
            mblnFueraChange = True
            If mintCodMarca = 0 Then
                Me.dbcMarca.Text = cINDEFINIDA
            End If
            Me.dbcModelo.Text = cINDEFINIDO
            mblnFueraChange = False
        Else
            If mintCodMarca = 0 Then
                mblnFueraChange = True
                Me.dbcMarca.Text = cINDEFINIDA
                mblnFueraChange = False
            End If
        End If
        FormaDescripcion()
    End Sub

    Private Sub dbcMaterial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMaterial.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        lStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial LIKE '" & Trim(Me.dbcMaterial.Text) & "%'" & " Order by descTipoMaterial "
        ModDCombo.DCChange(lStrSql, tecla, dbcMaterial)
        cMaterial = BuscaTipoMaterialDescCorta(mintCodMaterial)
        Call FormaDescripcion()
        If Trim(Me.dbcMaterial.Text) = "" Then
            mintCodMaterial = 0
        End If

MError:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcMaterial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMaterial.Enter
        Pon_Tool()
        gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial ORDER BY descTipoMaterial "
        ModDCombo.DCGotFocus(gStrSql, dbcMaterial)
    End Sub

    Private Sub dbcMaterial_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMaterial.KeyDown
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Me.dbcModelo.Focus()
            Case Else
                tecla = eventArgs.KeyCode
        End Select
    End Sub

    Private Sub dbcMaterial_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMaterial.KeyUp
        Dim Aux As String
        Aux = Trim(Me.dbcMaterial.Text)
        If Me.dbcMaterial.SelectedItem <> 0 Then
            dbcMaterial_Leave(dbcMaterial, New System.EventArgs())
        End If
        Me.dbcMaterial.Text = Aux
    End Sub

    Private Sub dbcMaterial_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMaterial.Leave
        Dim I As Integer
        Dim cDescripcion As String
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        cDescripcion = Trim(Me.dbcMaterial.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDO) Then
            cDescripcion = ""
        End If
        gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "'") & " Order by descTipoMaterial "
        mintCodMaterial = 0
        ModDCombo.DCLostFocus(dbcMaterial, gStrSql, mintCodMaterial)
        cMaterial = BuscaTipoMaterialDescCorta(mintCodMaterial)
        If mintCodMaterial = 0 Then
            mblnFueraChange = True
            Me.dbcMaterial.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
        Call FormaDescripcion()
    End Sub

    Private Sub dbcMaterial_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcMaterial.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcMaterial.Text)
        If Me.dbcMaterial.SelectedItem <> 0 Then
            dbcMaterial_Leave(dbcMaterial, New System.EventArgs())
        End If
        Me.dbcMaterial.Text = Aux
    End Sub

    Private Sub dbcModelo_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcModelo.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        cModelo = Trim(Me.dbcModelo.Text)
        FormaDescripcion()
        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codModelo, LTrim(RTrim(descModelo)) as descModelo FROM catModelos Where codGrupo = " & gCODRELOJERIA & " and codMarca = " & mintCodMarca & " and descModelo LIKE '" & Trim(Me.dbcModelo.Text) & "%' Order by descModelo "
        ModDCombo.DCChange(lStrSql, tecla, dbcModelo)

        If Trim(Me.dbcModelo.Text) = "" Then mintCodModelo = 0

MError:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcModelo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcModelo.Enter
        Pon_Tool()
        gStrSql = "SELECT codModelo, LTrim(RTrim(descModelo)) as descModelo FROM catModelos Where codGrupo = " & gCODRELOJERIA & " and codMarca = " & mintCodMarca & " ORDER BY descModelo "
        ModDCombo.DCGotFocus(gStrSql, dbcModelo)
    End Sub

    Private Sub dbcModelo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcModelo.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcMarca.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcModelo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcModelo.KeyPress
        If eventSender.keyAscii = System.Windows.Forms.Keys.Return Then
            If fraArticulo(1).Enabled Then
                If optGenero(0).Checked Then
                    optGenero(0).Focus()
                ElseIf optGenero(1).Checked Then
                    optGenero(1).Focus()
                ElseIf optGenero(2).Checked Then
                    optGenero(2).Focus()
                Else
                    optGenero(0).Focus()
                End If
            Else
                dbcMaterial.Focus()
            End If
        End If
    End Sub

    Private Sub dbcModelo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcModelo.Leave
        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String

        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        cDescripcion = Trim(Me.dbcModelo.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDO) Then cDescripcion = ""

        gStrSql = "SELECT codModelo, LTrim(RTrim(descModelo)) as descModelo FROM catModelos Where codGrupo = " & gCODRELOJERIA & " and codMarca = " & mintCodMarca & " and descModelo = '" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "'") & " Order by descModelo "
        Aux = mintCodModelo
        mintCodModelo = 0
        ModDCombo.DCLostFocus(dbcModelo, gStrSql, mintCodModelo)
        cModelo = Trim(dbcModelo.Text)
        If mintCodModelo = 0 Then
            mblnFueraChange = True
            dbcModelo.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
        FormaDescripcion()
    End Sub

    Private Sub dbcOrigen_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcOrigen.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcOrigen" Then
            Exit Sub
        End If
        gStrSql = "SELECT CodAlmacenOrigen, CodAlmacenOrigen AS DescAlmacen From CatOrigen WHERE DescAlmacenOrigen LIKE '" & Trim(dbcOrigen.Text) & "%' ORDER BY DescAlmacenOrigen "
        DCChange(gStrSql, tecla)
        intCodAlmacenOrigen = 0
    End Sub

    Private Sub dbcOrigen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen.Enter
        gStrSql = "SELECT CodAlmacenOrigen, CodAlmacenOrigen  AS DescAlmacen From CatOrigen ORDER BY DescAlmacenOrigen "
        DCGotFocus(gStrSql)
        Pon_Tool()
        mblnFueraChange = False
    End Sub

    Private Sub dbcOrigen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcOrigen.KeyPress
        ModEstandar.gp_CampoNumerico(eventSender.keyAscii)
        eventSender.keyAscii = ModEstandar.MskCantidad((dbcOrigen.Text), eventSender.keyAscii, 1, 0, (dbcOrigen.SelectionStart))
    End Sub

    Private Sub dbcOrigen_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcOrigen.KeyUp
        Dim Aux As String
        Aux = dbcOrigen.Text
        If dbcOrigen.SelectedItem <> 0 Then
            dbcOrigen_Leave(dbcOrigen, New System.EventArgs())
        End If
        dbcOrigen.Text = Aux
    End Sub

    Private Sub dbcOrigen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen.Leave
        gStrSql = "SELECT CodAlmacenOrigen, CodAlmacenOrigen AS DescAlmacen From CatOrigen  WHERE CodAlmacenOrigen LIKE '" & Trim(dbcOrigen.Text) & "'  ORDER BY DescAlmacenOrigen "
        DCLostFocus(dbcOrigen, gStrSql, intCodAlmacenOrigen)
    End Sub

    Private Sub dbcOrigen_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcOrigen.MouseUp
        Dim Aux As String
        Aux = dbcOrigen.Text
        If dbcOrigen.SelectedItem <> 0 Then
            dbcOrigen_Leave(dbcOrigen, New System.EventArgs())
        End If
        dbcOrigen.Text = Aux
    End Sub

    Private Sub dbcUnidad_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcUnidad.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad LIKE '" & Trim(Me.dbcUnidad.Text) & "%' Order by DescUnidad "
        ModDCombo.DCChange(lStrSql, tecla, dbcUnidad)

        If Trim(Me.dbcUnidad.Text) = "" Then mintCodUnidad = 0

MError:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcUnidad_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcUnidad.Enter
        Pon_Tool()
        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades ORDER BY descUnidad "
        ModDCombo.DCGotFocus(gStrSql, dbcUnidad)
    End Sub

    Private Sub dbcUnidad_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcUnidad.KeyDown
        Dim Aux As String
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                mblnSalir = True
                Me.Close()
                eventSender.KeyCode = 0
            Case System.Windows.Forms.Keys.Return
                Aux = Trim(Me.dbcUnidad.Text)
                If Me.dbcUnidad.SelectedItem <> 0 Then
                    dbcUnidad_Leave(dbcUnidad, New System.EventArgs())
                End If
                Me.dbcUnidad.Text = Aux
                Exit Sub
            Case System.Windows.Forms.Keys.Tab
                Aux = Trim(Me.dbcUnidad.Text)
                If Me.dbcUnidad.SelectedItem <> 0 Then
                    dbcUnidad_Leave(dbcUnidad, New System.EventArgs())
                End If
                Me.dbcUnidad.Text = Aux
                Exit Sub
        End Select
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcUnidad_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcUnidad.Leave
        Dim cDescripcion As String
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        cDescripcion = Trim(Me.dbcUnidad.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDA) Then
            cDescripcion = ""
        End If
        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad = '" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "'") & " Order by DescUnidad "
        mintCodUnidad = 0
        ModDCombo.DCLostFocus((Me.dbcUnidad), gStrSql, mintCodUnidad)
        If mintCodUnidad = 0 Then
            mblnFueraChange = True
            Me.dbcUnidad.Text = cINDEFINIDA
            mblnFueraChange = False
        End If

    End Sub

    Private Sub dbcUnidad_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcUnidad.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcUnidad.Text)
        If Me.dbcUnidad.SelectedItem <> 0 Then
            dbcUnidad_Leave(dbcUnidad, New System.EventArgs())
        End If
        Me.dbcUnidad.Text = Aux
    End Sub

    Private Sub frmCXPRelojeria_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Me.BringToFront()
    End Sub

    Private Sub frmCXPRelojeria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                Select Case UCase(Trim(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name))
                    Case Is = "DBCMODELO", Is = "OPTGENERO", Is = "TXTPRECIOENDOLARES"
                    Case Else
                        ModEstandar.AvanzarTab(Me)
                End Select
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmCXPRelojeria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPRelojeria_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        mintCodGrupo = gCODRELOJERIA
        intCodAlmacenOrigen = 0
        Select Case True
            Case optGenero(0).Checked
                cGenero = "H"
            Case optGenero(1).Checked
                cGenero = "D"
            Case optGenero(2).Checked
                cGenero = "M"
            Case Else
                cGenero = ""
        End Select
        cGeneroTag = cGenero
        Select Case True
            Case optMovimiento(0).Checked
                cMovimiento = "Q"
            Case optMovimiento(1).Checked
                cMovimiento = "AUT"
            Case optMovimiento(2).Checked
                cMovimiento = "MAN"
            Case Else
                cMovimiento = ""
        End Select
        cMovimientoTag = cMovimiento
        Select Case True
            Case chkCrono.CheckState = System.Windows.Forms.CheckState.Checked
                cCrono = "CHR"
                lCrono = True
            Case Else
                cCrono = ""
                lCrono = False
        End Select
        cCronoTag = cCrono
        gstrNombreForma = "FRMCXPRELOJERIA"
        FormaDescripcion()
    End Sub

    Private Sub frmCXPRelojeria_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If mblnCancelar Then
            Cancel = 0
            mblnCancelar = False
            Exit Sub
        End If
        mblnSalir = False
        Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
            Case MsgBoxResult.Yes 'Sale del Formulario
                Cancel = 0
            Case MsgBoxResult.No 'No sale del formulario
                Me.dbcUnidad.Focus()
                ModEstandar.SelTxt()
                Cancel = 1
        End Select
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCXPRelojeria_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.LimpiaDescBarraEstado()
        'frmCXPOrdenCompra.Enabled = True
        'frmCXPOrdenCompra.mshFlex.Focus()
        gstrNombreForma = ""
        'Me = Nothing
        IsNothing(Me)
    End Sub

    'UPGRADE_WARNING: Event optGenero.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optGenero_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optGenero.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = optGenero.GetIndex(eventSender)
            Select Case Index
                Case 0 'Hombre
                    cGenero = "H"
                Case 1 'Mujer
                    cGenero = "D"
                Case 2 'Unisex
                    cGenero = "M"
            End Select
            FormaDescripcion()
        End If
    End Sub

    Private Sub optGenero_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optGenero.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = optGenero.GetIndex(eventSender)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            If fraArticulo(2).Enabled Then
                If optMovimiento(0).Checked Then
                    optMovimiento(0).Focus()
                ElseIf optMovimiento(1).Checked Then
                    optMovimiento(1).Focus()
                ElseIf optMovimiento(2).Checked Then
                    optMovimiento(2).Focus()
                Else
                    optMovimiento(0).Focus()
                End If
            Else
                dbcMaterial.Focus()
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub optMovimiento_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMovimiento.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = optMovimiento.GetIndex(eventSender)
            Select Case Index
                Case 0 'Cuarzo
                    cMovimiento = "Q"
                Case 1 'Automático
                    cMovimiento = "AUT"
                Case 2 'Manual
                    cMovimiento = "MAN"
            End Select
            FormaDescripcion()
        End If
    End Sub

    Private Sub txtAdicional_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAdicional.TextChanged
        FormaDescripcion()
    End Sub

    Private Sub txtAdicional_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAdicional.Enter
        ModEstandar.SelTextoTxt(txtAdicional)
    End Sub

    Private Sub txtAdicional_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAdicional.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, ":.\/_-(){}[];@#$%&")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtAdicional_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAdicional.Leave
        cAdicional = Trim(txtAdicional.Text)
    End Sub

    Private Sub txtCantidadCompra_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCantidadCompra.Enter
        ModEstandar.SelTextoTxt(txtCantidadCompra)
    End Sub

    Private Sub txtCantidadCompra_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCantidadCompra.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        KeyAscii = ModEstandar.MskCantidad((txtCantidadCompra.Text), KeyAscii, 5, 0, (txtCantidadCompra.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodArtAnterior_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArtAnterior.Enter
        ModEstandar.SelTextoTxt(txtCodArtAnterior)
    End Sub

    Private Sub txtCodArtAnterior_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodArtAnterior.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        KeyAscii = ModEstandar.MskCantidad((Me.txtCodArtAnterior.Text), KeyAscii, 5, 0, (Me.txtCodArtAnterior.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodArtAnterior_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArtAnterior.Leave
        txtCodArtAnterior.Text = CStr(CInt(Numerico((txtCodArtAnterior.Text))))
        If Trim(txtCodArtAnterior.Text) = "0" Then txtCodArtAnterior.Text = ""
    End Sub

    Private Sub txtCodigodelProveedor_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigodelProveedor.TextChanged
        FormaDescripcion()
    End Sub

    Private Sub txtCodigodelProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigodelProveedor.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtCodigodelProveedor)
    End Sub

    Private Sub txtCodigodelProveedor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigodelProveedor.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "-_()[]#$%&/\.")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCostoActual_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoActual.Enter
        ModEstandar.SelTextoTxt(txtCostoActual)
    End Sub

    Private Sub txtCostoActual_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCostoActual.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, ".")
        KeyAscii = ModEstandar.MskCantidad((txtCostoActual.Text), KeyAscii, 9, 2, (txtCostoActual.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCostoActual_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoActual.Leave
        txtCostoActual.Text = Format(Numerico((txtCostoActual.Text)), "###,###,##0.00")
    End Sub

    Private Sub txtImagen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImagen.Enter
        ModEstandar.SelTextoTxt(txtImagen)
    End Sub

    Private Sub txtImagen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImagen.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, ":.\/_-")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPrecioenDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrecioenDolares.Enter
        Pon_Tool()
        SelTextoTxt((Me.txtPrecioenDolares))
    End Sub

    Private Sub txtPrecioenDolares_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPrecioenDolares.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            txtPrecioenDolares.Text = Format(Numerico((Me.txtPrecioenDolares.Text)), "###,###,##0.00")
            If optMoneda(0).Checked Then
                optMoneda(0).Focus()
            ElseIf optMoneda(1).Checked Then
                optMoneda(1).Focus()
            Else
                optMoneda(0).Focus()
            End If
        End If
        ModEstandar.gp_CampoNumerico(KeyAscii, ".")
        KeyAscii = ModEstandar.MskCantidad((txtPrecioenDolares.Text), KeyAscii, 9, 2, (Me.txtPrecioenDolares.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPrecioenDolares_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrecioenDolares.Leave
        txtPrecioenDolares.Text = Format(Numerico((txtPrecioenDolares.Text)), "###,###,##0.00")
    End Sub

    Private Sub DefineCondicionesRel(ByRef nMar As String, ByRef nMod As String, ByRef nGen As String, ByRef nMov As String, ByRef nCrono As String, ByRef nTipoM As String)
        If mintCodMarca <> 0 Then nMar = " CodMarca = " & mintCodMarca Else nMar = " CodMarca Is Null "
        If mintCodModelo <> 0 Then nMod = " CodModelo = " & mintCodModelo Else nMod = " CodModelo Is Null "
        If mintCodMaterial <> 0 Then nTipoM = " CodTipoMaterial = " & mintCodMaterial Else nTipoM = " CodTipoMaterial Is Null "
        nGen = " Genero = '" & cGenero & "'"
        nMov = " Movimiento = '" & cMovimiento & "'"
        If cCrono = "" Then nCrono = " Crono = 0 " Else nCrono = " Crono = 1 "
    End Sub
End Class