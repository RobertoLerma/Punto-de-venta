Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCXPJoyeria
    Inherits System.Windows.Forms.Form
    ''' ****************************************************************************************************************************************************
    ''' SE AGREGARON 4 CAMPOS NUEVOS PARA EL MANEJO DE DIAMANTE SUELTO ( MDS )
    ''' SE MODIFICO FUNCIONES:  GUARDAR-VALIDARDATOSMANEJODIAMANTESUELTO-VALIDADATOS-LLENAFORMA-LLENADATOS-CAMBIOS
    ''' 27OCT2010 - MAVF Ver
    '''
    ''' MDS CORRECCION - SE ELIMINO VALIDACIÓN DE PESO 0.00 YA QUE ES UN DATO NO REQUERIDO PARA TODOS LOS ARTICULOS DEL TIPO JOYERIA ( SOLO APLICA EN JOYERIA DIAMANTE SUELTO )
    ''' NO SE VALIDARA PARA RESURTIDOS
    ''' SE CONSIDERO ADICIONALMENTE  -SIN KILATAJE-  PARA EVITAR QUE PONGA 0K EN LA DESCRIPCIÓN - PETICION DE MRB DE ULTIMA HORA ( SIN $$$)
    ''' 08NOV2010 - MAVF Ver
    '''
    ''' Ver 1.1       Estatus:  Aprobado
    ''' ****************************************************************************************************************************************************

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtMDSPeso As System.Windows.Forms.TextBox
    Public WithEvents txtMDSColor As System.Windows.Forms.TextBox
    Public WithEvents txtMDSPureza As System.Windows.Forms.TextBox
    Public WithEvents txtMDSCertificado As System.Windows.Forms.TextBox
    Public WithEvents lblMDSPeso As System.Windows.Forms.Label
    Public WithEvents lblMDSColor As System.Windows.Forms.Label
    Public WithEvents lblMDSPureza As System.Windows.Forms.Label
    Public WithEvents lblMDSCertificado As System.Windows.Forms.Label
    Public WithEvents lblEstatus As System.Windows.Forms.Label
    Public WithEvents fraDiamanteSuelto As System.Windows.Forms.GroupBox
    Public WithEvents txtAdicional As System.Windows.Forms.TextBox
    Public WithEvents btnTipoMaterial As System.Windows.Forms.Button
    Public WithEvents btnSubLinea As System.Windows.Forms.Button
    Public WithEvents btnLinea As System.Windows.Forms.Button
    Public WithEvents btnFamilia As System.Windows.Forms.Button
    Public WithEvents btnCancelar As System.Windows.Forms.Button
    Public WithEvents btnAceptar As System.Windows.Forms.Button
    Public WithEvents txtImagen As System.Windows.Forms.TextBox
    Public WithEvents cmdBuscarImagen As System.Windows.Forms.Button
    Public WithEvents _FrameImagen_0 As System.Windows.Forms.GroupBox
    Public WithEvents txtCantidadCompra As System.Windows.Forms.TextBox
    Public WithEvents txtCostoActual As System.Windows.Forms.TextBox
    Public WithEvents chkCodAnt As System.Windows.Forms.CheckBox
    Public WithEvents txtCodArtAnterior As System.Windows.Forms.TextBox
    Public WithEvents dbcOrigen As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_31 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_32 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents _fraMoneda_5 As System.Windows.Forms.GroupBox
    Public WithEvents txtPrecioenDolares As System.Windows.Forms.TextBox
    Public WithEvents txtCodigodelProveedor As System.Windows.Forms.TextBox
    Public WithEvents _lblArticulo_9 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_8 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_7 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_6 As System.Windows.Forms.Label
    Public WithEvents _fraJoyeria_0 As System.Windows.Forms.GroupBox
    Public WithEvents dbcFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcSubLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcUnidad As System.Windows.Forms.ComboBox
    Public WithEvents dbcKilates As System.Windows.Forms.ComboBox
    Public WithEvents dbcMaterial As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_33 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_0 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_26 As System.Windows.Forms.Label
    Public WithEvents txtDescripcion As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_5 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_1 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_2 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_3 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_4 As System.Windows.Forms.Label
    Public WithEvents FrameImagen As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents fraJoyeria As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents fraMoneda As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblArticulo As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtAdicional = New System.Windows.Forms.TextBox()
        Me.txtImagen = New System.Windows.Forms.TextBox()
        Me.txtCantidadCompra = New System.Windows.Forms.TextBox()
        Me.txtCostoActual = New System.Windows.Forms.TextBox()
        Me.txtCodArtAnterior = New System.Windows.Forms.TextBox()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me.txtPrecioenDolares = New System.Windows.Forms.TextBox()
        Me.txtCodigodelProveedor = New System.Windows.Forms.TextBox()
        Me.fraDiamanteSuelto = New System.Windows.Forms.GroupBox()
        Me.txtMDSPeso = New System.Windows.Forms.TextBox()
        Me.txtMDSColor = New System.Windows.Forms.TextBox()
        Me.txtMDSPureza = New System.Windows.Forms.TextBox()
        Me.txtMDSCertificado = New System.Windows.Forms.TextBox()
        Me.lblMDSPeso = New System.Windows.Forms.Label()
        Me.lblMDSColor = New System.Windows.Forms.Label()
        Me.lblMDSPureza = New System.Windows.Forms.Label()
        Me.lblMDSCertificado = New System.Windows.Forms.Label()
        Me.lblEstatus = New System.Windows.Forms.Label()
        Me.btnTipoMaterial = New System.Windows.Forms.Button()
        Me.btnSubLinea = New System.Windows.Forms.Button()
        Me.btnLinea = New System.Windows.Forms.Button()
        Me.btnFamilia = New System.Windows.Forms.Button()
        Me.btnCancelar = New System.Windows.Forms.Button()
        Me.btnAceptar = New System.Windows.Forms.Button()
        Me._fraJoyeria_0 = New System.Windows.Forms.GroupBox()
        Me._FrameImagen_0 = New System.Windows.Forms.GroupBox()
        Me.cmdBuscarImagen = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.chkCodAnt = New System.Windows.Forms.CheckBox()
        Me.dbcOrigen = New System.Windows.Forms.ComboBox()
        Me._lblArticulo_31 = New System.Windows.Forms.Label()
        Me._lblArticulo_32 = New System.Windows.Forms.Label()
        Me._fraMoneda_5 = New System.Windows.Forms.GroupBox()
        Me._lblArticulo_9 = New System.Windows.Forms.Label()
        Me._lblArticulo_8 = New System.Windows.Forms.Label()
        Me._lblArticulo_7 = New System.Windows.Forms.Label()
        Me._lblArticulo_6 = New System.Windows.Forms.Label()
        Me.dbcFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcLinea = New System.Windows.Forms.ComboBox()
        Me.dbcSubLinea = New System.Windows.Forms.ComboBox()
        Me.dbcUnidad = New System.Windows.Forms.ComboBox()
        Me.dbcKilates = New System.Windows.Forms.ComboBox()
        Me.dbcMaterial = New System.Windows.Forms.ComboBox()
        Me._lblArticulo_33 = New System.Windows.Forms.Label()
        Me._lblArticulo_0 = New System.Windows.Forms.Label()
        Me._lblArticulo_26 = New System.Windows.Forms.Label()
        Me.txtDescripcion = New System.Windows.Forms.Label()
        Me._lblArticulo_5 = New System.Windows.Forms.Label()
        Me._lblArticulo_1 = New System.Windows.Forms.Label()
        Me._lblArticulo_2 = New System.Windows.Forms.Label()
        Me._lblArticulo_3 = New System.Windows.Forms.Label()
        Me._lblArticulo_4 = New System.Windows.Forms.Label()
        Me.FrameImagen = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.fraJoyeria = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.fraMoneda = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblArticulo = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.fraDiamanteSuelto.SuspendLayout()
        Me._fraJoyeria_0.SuspendLayout()
        Me._FrameImagen_0.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me._fraMoneda_5.SuspendLayout()
        CType(Me.FrameImagen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraJoyeria, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblArticulo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtAdicional
        '
        Me.txtAdicional.AcceptsReturn = True
        Me.txtAdicional.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
        Me.txtAdicional.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAdicional.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAdicional.Location = New System.Drawing.Point(87, 208)
        Me.txtAdicional.MaxLength = 15
        Me.txtAdicional.Name = "txtAdicional"
        Me.txtAdicional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAdicional.Size = New System.Drawing.Size(120, 21)
        Me.txtAdicional.TabIndex = 17
        Me.ToolTip1.SetToolTip(Me.txtAdicional, "Dato Adicional")
        '
        'txtImagen
        '
        Me.txtImagen.AcceptsReturn = True
        Me.txtImagen.BackColor = System.Drawing.SystemColors.Window
        Me.txtImagen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImagen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImagen.Location = New System.Drawing.Point(9, 15)
        Me.txtImagen.MaxLength = 0
        Me.txtImagen.Name = "txtImagen"
        Me.txtImagen.ReadOnly = True
        Me.txtImagen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImagen.Size = New System.Drawing.Size(245, 21)
        Me.txtImagen.TabIndex = 49
        Me.ToolTip1.SetToolTip(Me.txtImagen, "Imagen del artículo...")
        '
        'txtCantidadCompra
        '
        Me.txtCantidadCompra.AcceptsReturn = True
        Me.txtCantidadCompra.BackColor = System.Drawing.SystemColors.Window
        Me.txtCantidadCompra.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCantidadCompra.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCantidadCompra.Location = New System.Drawing.Point(143, 97)
        Me.txtCantidadCompra.MaxLength = 0
        Me.txtCantidadCompra.Name = "txtCantidadCompra"
        Me.txtCantidadCompra.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCantidadCompra.Size = New System.Drawing.Size(66, 21)
        Me.txtCantidadCompra.TabIndex = 36
        Me.txtCantidadCompra.Text = "0"
        Me.txtCantidadCompra.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCantidadCompra, "Cantidad de la compra")
        '
        'txtCostoActual
        '
        Me.txtCostoActual.AcceptsReturn = True
        Me.txtCostoActual.BackColor = System.Drawing.SystemColors.Window
        Me.txtCostoActual.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCostoActual.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCostoActual.Location = New System.Drawing.Point(112, 70)
        Me.txtCostoActual.MaxLength = 0
        Me.txtCostoActual.Name = "txtCostoActual"
        Me.txtCostoActual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCostoActual.Size = New System.Drawing.Size(96, 21)
        Me.txtCostoActual.TabIndex = 34
        Me.txtCostoActual.Text = "0.00"
        Me.txtCostoActual.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCostoActual, "Costo Factura sin Iva")
        '
        'txtCodArtAnterior
        '
        Me.txtCodArtAnterior.AcceptsReturn = True
        Me.txtCodArtAnterior.BackColor = System.Drawing.Color.White
        Me.txtCodArtAnterior.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodArtAnterior.Enabled = False
        Me.txtCodArtAnterior.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodArtAnterior.Location = New System.Drawing.Point(70, 41)
        Me.txtCodArtAnterior.MaxLength = 5
        Me.txtCodArtAnterior.Name = "txtCodArtAnterior"
        Me.txtCodArtAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodArtAnterior.Size = New System.Drawing.Size(43, 21)
        Me.txtCodArtAnterior.TabIndex = 47
        Me.txtCodArtAnterior.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCodArtAnterior, "Codigo anterior del artículo")
        '
        '_optMoneda_0
        '
        Me._optMoneda_0.BackColor = System.Drawing.Color.Silver
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_0, CType(0, Integer))
        Me._optMoneda_0.Location = New System.Drawing.Point(15, 17)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(60, 17)
        Me._optMoneda_0.TabIndex = 40
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Tag = "1"
        Me._optMoneda_0.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me._optMoneda_0, "Moneda del Precio Público - Dol")
        Me._optMoneda_0.UseVisualStyleBackColor = False
        '
        '_optMoneda_1
        '
        Me._optMoneda_1.BackColor = System.Drawing.Color.Silver
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_1, CType(1, Integer))
        Me._optMoneda_1.Location = New System.Drawing.Point(115, 17)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(52, 17)
        Me._optMoneda_1.TabIndex = 41
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Tag = "0"
        Me._optMoneda_1.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me._optMoneda_1, "Moneda del Precio Público - Pes")
        Me._optMoneda_1.UseVisualStyleBackColor = False
        '
        'txtPrecioenDolares
        '
        Me.txtPrecioenDolares.AcceptsReturn = True
        Me.txtPrecioenDolares.BackColor = System.Drawing.SystemColors.Window
        Me.txtPrecioenDolares.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrecioenDolares.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrecioenDolares.Location = New System.Drawing.Point(115, 134)
        Me.txtPrecioenDolares.MaxLength = 0
        Me.txtPrecioenDolares.Name = "txtPrecioenDolares"
        Me.txtPrecioenDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrecioenDolares.Size = New System.Drawing.Size(96, 21)
        Me.txtPrecioenDolares.TabIndex = 38
        Me.txtPrecioenDolares.Text = "0.00"
        Me.txtPrecioenDolares.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtPrecioenDolares, "Precio al Público en Dólares")
        '
        'txtCodigodelProveedor
        '
        Me.txtCodigodelProveedor.AcceptsReturn = True
        Me.txtCodigodelProveedor.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
        Me.txtCodigodelProveedor.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodigodelProveedor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodigodelProveedor.Location = New System.Drawing.Point(78, 30)
        Me.txtCodigodelProveedor.MaxLength = 20
        Me.txtCodigodelProveedor.Name = "txtCodigodelProveedor"
        Me.txtCodigodelProveedor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodigodelProveedor.Size = New System.Drawing.Size(129, 21)
        Me.txtCodigodelProveedor.TabIndex = 32
        Me.ToolTip1.SetToolTip(Me.txtCodigodelProveedor, "Código del Proveedor para el Artículo")
        '
        'fraDiamanteSuelto
        '
        Me.fraDiamanteSuelto.BackColor = System.Drawing.Color.Silver
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSPeso)
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSColor)
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSPureza)
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSCertificado)
        Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSPeso)
        Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSColor)
        Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSPureza)
        Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSCertificado)
        Me.fraDiamanteSuelto.Controls.Add(Me.lblEstatus)
        Me.fraDiamanteSuelto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraDiamanteSuelto.Location = New System.Drawing.Point(235, 137)
        Me.fraDiamanteSuelto.Name = "fraDiamanteSuelto"
        Me.fraDiamanteSuelto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraDiamanteSuelto.Size = New System.Drawing.Size(216, 144)
        Me.fraDiamanteSuelto.TabIndex = 18
        Me.fraDiamanteSuelto.TabStop = False
        Me.fraDiamanteSuelto.Text = " "
        '
        'txtMDSPeso
        '
        Me.txtMDSPeso.AcceptsReturn = True
        Me.txtMDSPeso.BackColor = System.Drawing.SystemColors.Window
        Me.txtMDSPeso.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSPeso.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSPeso.Location = New System.Drawing.Point(71, 14)
        Me.txtMDSPeso.MaxLength = 6
        Me.txtMDSPeso.Name = "txtMDSPeso"
        Me.txtMDSPeso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSPeso.Size = New System.Drawing.Size(50, 22)
        Me.txtMDSPeso.TabIndex = 23
        Me.txtMDSPeso.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMDSColor
        '
        Me.txtMDSColor.AcceptsReturn = True
        Me.txtMDSColor.BackColor = System.Drawing.SystemColors.Window
        Me.txtMDSColor.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSColor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSColor.Location = New System.Drawing.Point(71, 40)
        Me.txtMDSColor.MaxLength = 1
        Me.txtMDSColor.Name = "txtMDSColor"
        Me.txtMDSColor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSColor.Size = New System.Drawing.Size(50, 22)
        Me.txtMDSColor.TabIndex = 24
        '
        'txtMDSPureza
        '
        Me.txtMDSPureza.AcceptsReturn = True
        Me.txtMDSPureza.BackColor = System.Drawing.SystemColors.Window
        Me.txtMDSPureza.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSPureza.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSPureza.Location = New System.Drawing.Point(71, 66)
        Me.txtMDSPureza.MaxLength = 4
        Me.txtMDSPureza.Name = "txtMDSPureza"
        Me.txtMDSPureza.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSPureza.Size = New System.Drawing.Size(50, 22)
        Me.txtMDSPureza.TabIndex = 25
        '
        'txtMDSCertificado
        '
        Me.txtMDSCertificado.AcceptsReturn = True
        Me.txtMDSCertificado.BackColor = System.Drawing.SystemColors.Window
        Me.txtMDSCertificado.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSCertificado.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSCertificado.Location = New System.Drawing.Point(71, 92)
        Me.txtMDSCertificado.MaxLength = 20
        Me.txtMDSCertificado.Name = "txtMDSCertificado"
        Me.txtMDSCertificado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSCertificado.Size = New System.Drawing.Size(136, 22)
        Me.txtMDSCertificado.TabIndex = 26
        '
        'lblMDSPeso
        '
        Me.lblMDSPeso.BackColor = System.Drawing.Color.Silver
        Me.lblMDSPeso.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMDSPeso.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMDSPeso.Location = New System.Drawing.Point(9, 18)
        Me.lblMDSPeso.Name = "lblMDSPeso"
        Me.lblMDSPeso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMDSPeso.Size = New System.Drawing.Size(63, 18)
        Me.lblMDSPeso.TabIndex = 19
        Me.lblMDSPeso.Text = "Peso - CT"
        '
        'lblMDSColor
        '
        Me.lblMDSColor.BackColor = System.Drawing.Color.Silver
        Me.lblMDSColor.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMDSColor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMDSColor.Location = New System.Drawing.Point(9, 44)
        Me.lblMDSColor.Name = "lblMDSColor"
        Me.lblMDSColor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMDSColor.Size = New System.Drawing.Size(63, 18)
        Me.lblMDSColor.TabIndex = 20
        Me.lblMDSColor.Text = "Color"
        '
        'lblMDSPureza
        '
        Me.lblMDSPureza.BackColor = System.Drawing.Color.Silver
        Me.lblMDSPureza.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMDSPureza.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMDSPureza.Location = New System.Drawing.Point(9, 71)
        Me.lblMDSPureza.Name = "lblMDSPureza"
        Me.lblMDSPureza.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMDSPureza.Size = New System.Drawing.Size(63, 18)
        Me.lblMDSPureza.TabIndex = 21
        Me.lblMDSPureza.Text = "Pureza - Q"
        '
        'lblMDSCertificado
        '
        Me.lblMDSCertificado.BackColor = System.Drawing.Color.Silver
        Me.lblMDSCertificado.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMDSCertificado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMDSCertificado.Location = New System.Drawing.Point(7, 98)
        Me.lblMDSCertificado.Name = "lblMDSCertificado"
        Me.lblMDSCertificado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMDSCertificado.Size = New System.Drawing.Size(63, 18)
        Me.lblMDSCertificado.TabIndex = 22
        Me.lblMDSCertificado.Text = "Certificado"
        '
        'lblEstatus
        '
        Me.lblEstatus.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblEstatus.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEstatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEstatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblEstatus.Location = New System.Drawing.Point(71, 117)
        Me.lblEstatus.Name = "lblEstatus"
        Me.lblEstatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEstatus.Size = New System.Drawing.Size(136, 19)
        Me.lblEstatus.TabIndex = 27
        Me.lblEstatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnTipoMaterial
        '
        Me.btnTipoMaterial.BackColor = System.Drawing.SystemColors.Control
        Me.btnTipoMaterial.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnTipoMaterial.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnTipoMaterial.Location = New System.Drawing.Point(210, 177)
        Me.btnTipoMaterial.Name = "btnTipoMaterial"
        Me.btnTipoMaterial.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnTipoMaterial.Size = New System.Drawing.Size(21, 21)
        Me.btnTipoMaterial.TabIndex = 15
        Me.btnTipoMaterial.TabStop = False
        Me.btnTipoMaterial.Text = "..."
        Me.btnTipoMaterial.UseVisualStyleBackColor = False
        '
        'btnSubLinea
        '
        Me.btnSubLinea.BackColor = System.Drawing.SystemColors.Control
        Me.btnSubLinea.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSubLinea.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSubLinea.Location = New System.Drawing.Point(374, 112)
        Me.btnSubLinea.Name = "btnSubLinea"
        Me.btnSubLinea.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSubLinea.Size = New System.Drawing.Size(21, 21)
        Me.btnSubLinea.TabIndex = 10
        Me.btnSubLinea.TabStop = False
        Me.btnSubLinea.Text = "..."
        Me.btnSubLinea.UseVisualStyleBackColor = False
        '
        'btnLinea
        '
        Me.btnLinea.BackColor = System.Drawing.SystemColors.Control
        Me.btnLinea.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLinea.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLinea.Location = New System.Drawing.Point(374, 80)
        Me.btnLinea.Name = "btnLinea"
        Me.btnLinea.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLinea.Size = New System.Drawing.Size(21, 21)
        Me.btnLinea.TabIndex = 7
        Me.btnLinea.TabStop = False
        Me.btnLinea.Text = "..."
        Me.btnLinea.UseVisualStyleBackColor = False
        '
        'btnFamilia
        '
        Me.btnFamilia.BackColor = System.Drawing.SystemColors.Control
        Me.btnFamilia.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnFamilia.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnFamilia.Location = New System.Drawing.Point(374, 48)
        Me.btnFamilia.Name = "btnFamilia"
        Me.btnFamilia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnFamilia.Size = New System.Drawing.Size(21, 21)
        Me.btnFamilia.TabIndex = 4
        Me.btnFamilia.TabStop = False
        Me.btnFamilia.Text = "..."
        Me.btnFamilia.UseVisualStyleBackColor = False
        '
        'btnCancelar
        '
        Me.btnCancelar.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancelar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancelar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancelar.Location = New System.Drawing.Point(693, 312)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancelar.Size = New System.Drawing.Size(88, 25)
        Me.btnCancelar.TabIndex = 51
        Me.btnCancelar.Text = "&Cancelar"
        Me.btnCancelar.UseVisualStyleBackColor = False
        '
        'btnAceptar
        '
        Me.btnAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.btnAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAceptar.Location = New System.Drawing.Point(538, 312)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAceptar.Size = New System.Drawing.Size(88, 25)
        Me.btnAceptar.TabIndex = 50
        Me.btnAceptar.Text = "&Aceptar"
        Me.btnAceptar.UseVisualStyleBackColor = False
        '
        '_fraJoyeria_0
        '
        Me._fraJoyeria_0.BackColor = System.Drawing.Color.Silver
        Me._fraJoyeria_0.Controls.Add(Me._FrameImagen_0)
        Me._fraJoyeria_0.Controls.Add(Me.txtCantidadCompra)
        Me._fraJoyeria_0.Controls.Add(Me.txtCostoActual)
        Me._fraJoyeria_0.Controls.Add(Me.Frame3)
        Me._fraJoyeria_0.Controls.Add(Me._fraMoneda_5)
        Me._fraJoyeria_0.Controls.Add(Me.txtPrecioenDolares)
        Me._fraJoyeria_0.Controls.Add(Me.txtCodigodelProveedor)
        Me._fraJoyeria_0.Controls.Add(Me._lblArticulo_9)
        Me._fraJoyeria_0.Controls.Add(Me._lblArticulo_8)
        Me._fraJoyeria_0.Controls.Add(Me._lblArticulo_7)
        Me._fraJoyeria_0.Controls.Add(Me._lblArticulo_6)
        Me._fraJoyeria_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraJoyeria.SetIndex(Me._fraJoyeria_0, CType(0, Integer))
        Me._fraJoyeria_0.Location = New System.Drawing.Point(459, 5)
        Me._fraJoyeria_0.Name = "_fraJoyeria_0"
        Me._fraJoyeria_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraJoyeria_0.Size = New System.Drawing.Size(382, 276)
        Me._fraJoyeria_0.TabIndex = 30
        Me._fraJoyeria_0.TabStop = False
        Me._fraJoyeria_0.Text = " Datos Adicionales "
        '
        '_FrameImagen_0
        '
        Me._FrameImagen_0.BackColor = System.Drawing.Color.Silver
        Me._FrameImagen_0.Controls.Add(Me.txtImagen)
        Me._FrameImagen_0.Controls.Add(Me.cmdBuscarImagen)
        Me._FrameImagen_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.FrameImagen.SetIndex(Me._FrameImagen_0, CType(0, Integer))
        Me._FrameImagen_0.Location = New System.Drawing.Point(76, 215)
        Me._FrameImagen_0.Name = "_FrameImagen_0"
        Me._FrameImagen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._FrameImagen_0.Size = New System.Drawing.Size(290, 44)
        Me._FrameImagen_0.TabIndex = 48
        Me._FrameImagen_0.TabStop = False
        Me._FrameImagen_0.Text = "Imagen"
        '
        'cmdBuscarImagen
        '
        Me.cmdBuscarImagen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBuscarImagen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBuscarImagen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBuscarImagen.Location = New System.Drawing.Point(260, 15)
        Me.cmdBuscarImagen.Name = "cmdBuscarImagen"
        Me.cmdBuscarImagen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBuscarImagen.Size = New System.Drawing.Size(22, 21)
        Me.cmdBuscarImagen.TabIndex = 52
        Me.cmdBuscarImagen.Text = "..."
        Me.cmdBuscarImagen.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.Color.Silver
        Me.Frame3.Controls.Add(Me.chkCodAnt)
        Me.Frame3.Controls.Add(Me.txtCodArtAnterior)
        Me.Frame3.Controls.Add(Me.dbcOrigen)
        Me.Frame3.Controls.Add(Me._lblArticulo_31)
        Me.Frame3.Controls.Add(Me._lblArticulo_32)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(242, 24)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(124, 69)
        Me.Frame3.TabIndex = 42
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "    Codigo Anterior"
        '
        'chkCodAnt
        '
        Me.chkCodAnt.BackColor = System.Drawing.SystemColors.Control
        Me.chkCodAnt.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCodAnt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCodAnt.Location = New System.Drawing.Point(8, -2)
        Me.chkCodAnt.Name = "chkCodAnt"
        Me.chkCodAnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCodAnt.Size = New System.Drawing.Size(17, 17)
        Me.chkCodAnt.TabIndex = 43
        Me.chkCodAnt.UseVisualStyleBackColor = False
        '
        'dbcOrigen
        '
        Me.dbcOrigen.Location = New System.Drawing.Point(70, 17)
        Me.dbcOrigen.Name = "dbcOrigen"
        Me.dbcOrigen.Size = New System.Drawing.Size(43, 21)
        Me.dbcOrigen.TabIndex = 45
        '
        '_lblArticulo_31
        '
        Me._lblArticulo_31.AutoSize = True
        Me._lblArticulo_31.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_31.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_31.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_31, CType(31, Integer))
        Me._lblArticulo_31.Location = New System.Drawing.Point(25, 21)
        Me._lblArticulo_31.Name = "_lblArticulo_31"
        Me._lblArticulo_31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_31.Size = New System.Drawing.Size(47, 13)
        Me._lblArticulo_31.TabIndex = 44
        Me._lblArticulo_31.Text = "Origen : "
        '
        '_lblArticulo_32
        '
        Me._lblArticulo_32.AutoSize = True
        Me._lblArticulo_32.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_32.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_32.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_32, CType(32, Integer))
        Me._lblArticulo_32.Location = New System.Drawing.Point(26, 44)
        Me._lblArticulo_32.Name = "_lblArticulo_32"
        Me._lblArticulo_32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_32.Size = New System.Drawing.Size(46, 13)
        Me._lblArticulo_32.TabIndex = 46
        Me._lblArticulo_32.Text = "Código: "
        '
        '_fraMoneda_5
        '
        Me._fraMoneda_5.BackColor = System.Drawing.Color.Silver
        Me._fraMoneda_5.Controls.Add(Me._optMoneda_0)
        Me._fraMoneda_5.Controls.Add(Me._optMoneda_1)
        Me._fraMoneda_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMoneda.SetIndex(Me._fraMoneda_5, CType(5, Integer))
        Me._fraMoneda_5.Location = New System.Drawing.Point(17, 162)
        Me._fraMoneda_5.Name = "_fraMoneda_5"
        Me._fraMoneda_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraMoneda_5.Size = New System.Drawing.Size(196, 40)
        Me._fraMoneda_5.TabIndex = 39
        Me._fraMoneda_5.TabStop = False
        Me._fraMoneda_5.Text = "  Moneda Precio Público "
        '
        '_lblArticulo_9
        '
        Me._lblArticulo_9.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_9, CType(9, Integer))
        Me._lblArticulo_9.Location = New System.Drawing.Point(44, 102)
        Me._lblArticulo_9.Name = "_lblArticulo_9"
        Me._lblArticulo_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_9.Size = New System.Drawing.Size(90, 15)
        Me._lblArticulo_9.TabIndex = 35
        Me._lblArticulo_9.Text = "Cantidad Compra :"
        '
        '_lblArticulo_8
        '
        Me._lblArticulo_8.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_8, CType(8, Integer))
        Me._lblArticulo_8.Location = New System.Drawing.Point(22, 74)
        Me._lblArticulo_8.Name = "_lblArticulo_8"
        Me._lblArticulo_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_8.Size = New System.Drawing.Size(76, 17)
        Me._lblArticulo_8.TabIndex = 33
        Me._lblArticulo_8.Text = "Cto Fact S/Iva :"
        '
        '_lblArticulo_7
        '
        Me._lblArticulo_7.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_7, CType(7, Integer))
        Me._lblArticulo_7.Location = New System.Drawing.Point(17, 139)
        Me._lblArticulo_7.Name = "_lblArticulo_7"
        Me._lblArticulo_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_7.Size = New System.Drawing.Size(76, 20)
        Me._lblArticulo_7.TabIndex = 37
        Me._lblArticulo_7.Text = "Precio Público :"
        '
        '_lblArticulo_6
        '
        Me._lblArticulo_6.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_6, CType(6, Integer))
        Me._lblArticulo_6.Location = New System.Drawing.Point(14, 24)
        Me._lblArticulo_6.Name = "_lblArticulo_6"
        Me._lblArticulo_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_6.Size = New System.Drawing.Size(64, 28)
        Me._lblArticulo_6.TabIndex = 31
        Me._lblArticulo_6.Text = "Código del Proveedor :"
        '
        'dbcFamilia
        '
        Me.dbcFamilia.Location = New System.Drawing.Point(87, 48)
        Me.dbcFamilia.Name = "dbcFamilia"
        Me.dbcFamilia.Size = New System.Drawing.Size(265, 21)
        Me.dbcFamilia.TabIndex = 3
        '
        'dbcLinea
        '
        Me.dbcLinea.Location = New System.Drawing.Point(87, 80)
        Me.dbcLinea.Name = "dbcLinea"
        Me.dbcLinea.Size = New System.Drawing.Size(265, 21)
        Me.dbcLinea.TabIndex = 6
        '
        'dbcSubLinea
        '
        Me.dbcSubLinea.Location = New System.Drawing.Point(87, 112)
        Me.dbcSubLinea.Name = "dbcSubLinea"
        Me.dbcSubLinea.Size = New System.Drawing.Size(265, 21)
        Me.dbcSubLinea.TabIndex = 9
        '
        'dbcUnidad
        '
        Me.dbcUnidad.Location = New System.Drawing.Point(87, 16)
        Me.dbcUnidad.Name = "dbcUnidad"
        Me.dbcUnidad.Size = New System.Drawing.Size(100, 21)
        Me.dbcUnidad.TabIndex = 1
        '
        'dbcKilates
        '
        Me.dbcKilates.Location = New System.Drawing.Point(87, 144)
        Me.dbcKilates.Name = "dbcKilates"
        Me.dbcKilates.Size = New System.Drawing.Size(120, 21)
        Me.dbcKilates.TabIndex = 12
        '
        'dbcMaterial
        '
        Me.dbcMaterial.Location = New System.Drawing.Point(87, 177)
        Me.dbcMaterial.Name = "dbcMaterial"
        Me.dbcMaterial.Size = New System.Drawing.Size(120, 21)
        Me.dbcMaterial.TabIndex = 14
        '
        '_lblArticulo_33
        '
        Me._lblArticulo_33.AutoSize = True
        Me._lblArticulo_33.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_33.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_33.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_33, CType(33, Integer))
        Me._lblArticulo_33.Location = New System.Drawing.Point(4, 213)
        Me._lblArticulo_33.Name = "_lblArticulo_33"
        Me._lblArticulo_33.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_33.Size = New System.Drawing.Size(76, 13)
        Me._lblArticulo_33.TabIndex = 16
        Me._lblArticulo_33.Text = "Dato Adicional"
        '
        '_lblArticulo_0
        '
        Me._lblArticulo_0.AutoSize = True
        Me._lblArticulo_0.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_0, CType(0, Integer))
        Me._lblArticulo_0.Location = New System.Drawing.Point(4, 182)
        Me._lblArticulo_0.Name = "_lblArticulo_0"
        Me._lblArticulo_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_0.Size = New System.Drawing.Size(83, 13)
        Me._lblArticulo_0.TabIndex = 13
        Me._lblArticulo_0.Text = "Tipo de Material"
        '
        '_lblArticulo_26
        '
        Me._lblArticulo_26.AutoSize = True
        Me._lblArticulo_26.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_26.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_26, CType(26, Integer))
        Me._lblArticulo_26.Location = New System.Drawing.Point(4, 148)
        Me._lblArticulo_26.Name = "_lblArticulo_26"
        Me._lblArticulo_26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_26.Size = New System.Drawing.Size(38, 13)
        Me._lblArticulo_26.TabIndex = 11
        Me._lblArticulo_26.Text = "Kilates"
        '
        'txtDescripcion
        '
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Info
        Me.txtDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDescripcion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtDescripcion.Location = New System.Drawing.Point(87, 289)
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(265, 57)
        Me.txtDescripcion.TabIndex = 29
        '
        '_lblArticulo_5
        '
        Me._lblArticulo_5.AutoSize = True
        Me._lblArticulo_5.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_5, CType(5, Integer))
        Me._lblArticulo_5.Location = New System.Drawing.Point(4, 20)
        Me._lblArticulo_5.Name = "_lblArticulo_5"
        Me._lblArticulo_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_5.Size = New System.Drawing.Size(41, 13)
        Me._lblArticulo_5.TabIndex = 0
        Me._lblArticulo_5.Text = "Unidad"
        '
        '_lblArticulo_1
        '
        Me._lblArticulo_1.AutoSize = True
        Me._lblArticulo_1.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_1, CType(1, Integer))
        Me._lblArticulo_1.Location = New System.Drawing.Point(4, 52)
        Me._lblArticulo_1.Name = "_lblArticulo_1"
        Me._lblArticulo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_1.Size = New System.Drawing.Size(39, 13)
        Me._lblArticulo_1.TabIndex = 2
        Me._lblArticulo_1.Text = "Familia"
        '
        '_lblArticulo_2
        '
        Me._lblArticulo_2.AutoSize = True
        Me._lblArticulo_2.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_2, CType(2, Integer))
        Me._lblArticulo_2.Location = New System.Drawing.Point(4, 84)
        Me._lblArticulo_2.Name = "_lblArticulo_2"
        Me._lblArticulo_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_2.Size = New System.Drawing.Size(35, 13)
        Me._lblArticulo_2.TabIndex = 5
        Me._lblArticulo_2.Text = "Línea"
        '
        '_lblArticulo_3
        '
        Me._lblArticulo_3.AutoSize = True
        Me._lblArticulo_3.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_3, CType(3, Integer))
        Me._lblArticulo_3.Location = New System.Drawing.Point(4, 116)
        Me._lblArticulo_3.Name = "_lblArticulo_3"
        Me._lblArticulo_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_3.Size = New System.Drawing.Size(54, 13)
        Me._lblArticulo_3.TabIndex = 8
        Me._lblArticulo_3.Text = "SubLínea"
        '
        '_lblArticulo_4
        '
        Me._lblArticulo_4.AutoSize = True
        Me._lblArticulo_4.BackColor = System.Drawing.Color.Silver
        Me._lblArticulo_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArticulo.SetIndex(Me._lblArticulo_4, CType(4, Integer))
        Me._lblArticulo_4.Location = New System.Drawing.Point(3, 292)
        Me._lblArticulo_4.Name = "_lblArticulo_4"
        Me._lblArticulo_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_4.Size = New System.Drawing.Size(63, 13)
        Me._lblArticulo_4.TabIndex = 28
        Me._lblArticulo_4.Text = "Descripción"
        '
        'frmCXPJoyeria
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(851, 354)
        Me.Controls.Add(Me.fraDiamanteSuelto)
        Me.Controls.Add(Me.txtAdicional)
        Me.Controls.Add(Me.btnTipoMaterial)
        Me.Controls.Add(Me.btnSubLinea)
        Me.Controls.Add(Me.btnLinea)
        Me.Controls.Add(Me.btnFamilia)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me._fraJoyeria_0)
        Me.Controls.Add(Me.dbcFamilia)
        Me.Controls.Add(Me.dbcLinea)
        Me.Controls.Add(Me.dbcSubLinea)
        Me.Controls.Add(Me.dbcUnidad)
        Me.Controls.Add(Me.dbcKilates)
        Me.Controls.Add(Me.dbcMaterial)
        Me.Controls.Add(Me._lblArticulo_33)
        Me.Controls.Add(Me._lblArticulo_0)
        Me.Controls.Add(Me._lblArticulo_26)
        Me.Controls.Add(Me.txtDescripcion)
        Me.Controls.Add(Me._lblArticulo_5)
        Me.Controls.Add(Me._lblArticulo_1)
        Me.Controls.Add(Me._lblArticulo_2)
        Me.Controls.Add(Me._lblArticulo_3)
        Me.Controls.Add(Me._lblArticulo_4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(197, 171)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCXPJoyeria"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Definir Joyería"
        Me.fraDiamanteSuelto.ResumeLayout(False)
        Me._fraJoyeria_0.ResumeLayout(False)
        Me._FrameImagen_0.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me._fraMoneda_5.ResumeLayout(False)
        CType(Me.FrameImagen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraJoyeria, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblArticulo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Const cINDEFINIDO As String = "[ Vacío ... ]"
    Const cINDEFINIDA As String = "[ Vacío ... ]"
    Const cSINKILATES As String = "(SIN KILATES)" 'Para evitar poner Kiltaje en 0k    '''08NOV2010 - MAVF

    Const C_COLCODIGO As Integer = 0
    Const C_COLDESCRIPCION As Integer = 1
    Const C_COLUNIDAD As Integer = 2
    Const C_COLCANTIDAD As Integer = 3
    Const C_COLPRECIOUNITARIO As Integer = 4
    Const C_COLCODAUX As Integer = 8
    Const C_ColSTATUS As Integer = 9
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

    Const C_ColMDSPESO As Integer = 83 '''27OCT2010 - MAVF
    Const C_ColMDSCOLOR As Integer = 84 '''27OCT2010 - MAVF
    Const C_ColMDSPUREZA As Integer = 85 '''27OCT2010 - MAVF
    Const C_ColMDSCERTIFICADO As Integer = 86 '''27OCT2010 - MAVF

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
    Public mintCodFamilia As Integer
    Public mintCodLinea As Integer
    Public mintCodSubLinea As Integer
    Dim mintCodKilates As Integer
    Public mintCodMaterial As Integer
    Dim tecla As Integer
    Dim cMarca As String
    Dim cModelo As String
    Dim cTipoMaterial As String
    Dim rsLocal As ADODB.Recordset
    Dim strAdicional As String
    Dim cLinea, cFamilia, cSubLinea As Object
    Dim cKilates As String
    Dim intCodAlmacenOrigen As Integer

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

    Function BuscaSubLineaDescCorta(ByRef CodGrupo As Integer, ByRef CodFamilia As Integer, ByRef COdLinea As Integer, ByRef CodSubLinea As Integer) As String
        On Error GoTo Err_Renamed
        gStrSql = "SELECT CodSubLinea,DescCorta FROM CatSubLineas WHERE CodGrupo = " & CodGrupo & " AND CodFamilia = " & CodFamilia & " AND CodLinea = " & COdLinea & " AND CodSubLinea = " & CodSubLinea
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaSubLineaDescCorta = Trim(rsLocal.Fields("DescCorta").Value)
        Else
            BuscaSubLineaDescCorta = cINDEFINIDO
        End If
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

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
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaFamilia(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codFamilia, DescFamilia FROM CatFamilias WHERE CodGrupo = " & gCODJOYERIA & " AND CodFamilia = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaFamilia = Trim(rsLocal.Fields("DescFamilia").Value)
        Else
            BuscaFamilia = cINDEFINIDA
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaLinea(ByRef Codigo As Integer, ByRef nCodFamilia As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codFamilia, codLinea, DescLinea FROM CatLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & nCodFamilia & " AND CodLinea = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaLinea = Trim(rsLocal.Fields("DescLinea").Value)
        Else
            BuscaLinea = cINDEFINIDA
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaSubLinea(ByRef Codigo As Integer, ByRef nCodFamilia As Integer, ByRef nCodLinea As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codFamilia, codLinea, codSubLinea, DescSubLinea FROM CatSubLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & nCodFamilia & " AND CodLinea = " & nCodLinea & " AND CodSubLinea = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaSubLinea = Trim(rsLocal.Fields("DescSubLinea").Value)
        Else
            BuscaSubLinea = cINDEFINIDA
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaKilataje(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codKilates, DescKilates FROM CatKilates WHERE CodKilates = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaKilataje = Trim(rsLocal.Fields("DescKilates").Value)
        Else
            BuscaKilataje = cINDEFINIDO
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

    Public Sub LLenaForma(ByRef Columna As Integer, ByRef Renglon As Integer, ByRef RenglonAct As Integer, ByRef nCodUnidad As Integer, ByRef nCodFamilia As Integer, ByRef nCodLinea As Integer, ByRef nCodSubLinea As Integer, ByRef nCodKilates As Integer, ByRef cDescripcion As String, ByRef nCodMaterial As Integer, ByRef cCodArticuloProv As String, ByRef cAdicional As String, ByRef nPrecioPublico As Decimal, ByRef bMonedaPP As String, ByRef nOrigenAnt As Integer, ByRef nCodigoAnt As Integer, ByRef nCantidadCompra As Integer, ByRef nCostoActual As Decimal, ByRef cImagen As String, ByRef nmdsPeso As Decimal, ByRef cmdsColor As String, ByRef cmdsPureza As String, ByRef cmdsCertificado As String)
        On Error GoTo Merr

        nCol = Columna
        nRow = Renglon
        nRowAct = RenglonAct

        If CInt(Numerico(frmCXPOrdenCompra.mshFlex.get_TextMatrix(nRowAct, C_COLCODIGO))) = 0 Then
            '''elemento nuevo
            dbcUnidad.Text = False
            dbcFamilia.Text = False
            dbcLinea.Text = False
            dbcSubLinea.Text = False
            dbcKilates.Text = False
            dbcMaterial.Text = False
            txtCodigodelProveedor.ReadOnly = False
            txtAdicional.ReadOnly = False
            txtMDSPeso.ReadOnly = False '''27OCT2010 - MAVF
            txtMDSColor.ReadOnly = False '''27OCT2010 - MAVF
            txtMDSPureza.ReadOnly = False '''27OCT2010 - MAVF
            txtMDSCertificado.ReadOnly = False '''27OCT2010 - MAVF
        Else
            '''resurtido
            dbcUnidad.Text = True
            dbcFamilia.Text = True
            dbcLinea.Text = True
            dbcSubLinea.Text = True
            dbcKilates.Text = True
            dbcMaterial.Text = True
            txtCodigodelProveedor.ReadOnly = True
            txtAdicional.ReadOnly = True
            txtMDSPeso.ReadOnly = True '''27OCT2010 - MAVF
            txtMDSColor.ReadOnly = True '''27OCT2010 - MAVF
            txtMDSPureza.ReadOnly = True '''27OCT2010 - MAVF
            txtMDSCertificado.ReadOnly = True '''27OCT2010 - MAVF
        End If

        gstrNombreForma = "FRMCXPJOYERIA"
        mblnFueraChange = True

        mintCodUnidad = nCodUnidad
        If mintCodUnidad = 0 Then
            dbcUnidad_Enter(dbcUnidad, New System.EventArgs())
            dbcUnidad.Text = Trim(cINDEFINIDA)
        Else
            dbcUnidad.Text = BuscaUnidad(mintCodUnidad)
        End If
        dbcUnidad.Tag = dbcUnidad.Text

        mintCodFamilia = nCodFamilia
        dbcFamilia.Text = BuscaFamilia(mintCodFamilia)
        dbcFamilia.Tag = dbcFamilia.Text
        mintCodLinea = nCodLinea
        dbcLinea.Text = BuscaLinea(mintCodLinea, mintCodFamilia)
        dbcLinea.Tag = dbcLinea.Text
        mintCodSubLinea = nCodSubLinea
        dbcSubLinea.Text = BuscaSubLinea(mintCodSubLinea, mintCodFamilia, mintCodLinea)
        dbcSubLinea.Tag = dbcSubLinea.Text
        cSubLinea = Trim(BuscaSubLineaDescCorta(mintCodGrupo, mintCodFamilia, mintCodLinea, mintCodSubLinea))
        mintCodKilates = nCodKilates
        dbcKilates.Text = BuscaKilataje(mintCodKilates)
        dbcKilates.Tag = dbcKilates.Text
        txtDescripcion.Text = Trim(cDescripcion)
        txtDescripcion.Tag = txtDescripcion.Text
        mintCodMaterial = nCodMaterial
        dbcMaterial.Text = BuscaTipoMaterial(mintCodMaterial)
        dbcMaterial.Tag = dbcMaterial.Text
        cTipoMaterial = BuscaTipoMaterialDescCorta(mintCodMaterial)
        mblnFueraChange = False
        txtAdicional.Text = cAdicional
        txtAdicional.Tag = cAdicional
        '''27OCT2010 - MAVF
        txtMDSPeso.Text = VB6.Format(nmdsPeso, "##0.00")
        txtMDSPeso.Tag = VB6.Format(nmdsPeso, "##0.00")
        txtMDSColor.Text = cmdsColor
        txtMDSColor.Tag = cmdsColor
        txtMDSPureza.Text = cmdsPureza
        txtMDSPureza.Tag = cmdsPureza
        txtMDSCertificado.Text = cmdsCertificado
        txtMDSCertificado.Tag = cmdsCertificado
        ''' ********************************************************
        txtCodigodelProveedor.Text = cCodArticuloProv
        txtCodigodelProveedor.Tag = cCodArticuloProv

        If nRowAct = nRow Then
            txtPrecioenDolares.Text = VB6.Format(nPrecioPublico, "###,##0.00")
            txtPrecioenDolares.Tag = VB6.Format(nPrecioPublico, "###,##0.00")

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
            txtCostoActual.Text = VB6.Format(nCostoActual, "###,##0.00")
            txtCostoActual.Tag = VB6.Format(nCostoActual, "###,##0.00")
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
            txtPrecioenDolares.Text = VB6.Format(0, "###,##0.00")
            txtPrecioenDolares.Tag = VB6.Format(0, "###,##0.00")
            FormaDescripcion()

            chkCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked
            dbcOrigen.Text = ""
            dbcOrigen.Tag = ""
            txtCodArtAnterior.Text = ""
            txtCodArtAnterior.Tag = ""

            txtCantidadCompra.Text = "0"
            txtCantidadCompra.Tag = "0"
            txtCostoActual.Text = VB6.Format(0, "###,##0.00")
            txtCostoActual.Tag = VB6.Format(0, "###,##0.00")
            txtImagen.Text = ""
            txtImagen.Tag = ""
            optMoneda(0).Checked = False
            optMoneda(1).Checked = False
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub LlenaDatos(ByRef CodFolio As String, ByRef CodArticulo As String, ByRef Columna As Integer, ByRef Renglon As Integer, ByRef RenglonAct As Integer, ByRef nCantidadCompra As Integer, ByRef nCostoActual As Decimal, ByRef cImagen As String, ByRef Tipo As String, ByRef nmdsPeso As Decimal, ByRef cmdsColor As String, ByRef cmdsPureza As String, ByRef cmdsCertificado As String)
        On Error GoTo Merr
        Dim lResurtido As Boolean

        nCol = Columna
        nRow = Renglon
        nRowAct = RenglonAct

        If Tipo = "1C" Then
            lResurtido = True
            gStrSql = "SELECT * FROM CatArticulos (Nolock) WHERE CodArticulo = " & Trim(frmCXPOrdenCompra.mshFlex.get_TextMatrix(nRow, C_COLCODIGO))
        ElseIf Tipo = "2C" Then
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
            mintCodFamilia = IIf(IsDBNull(RsGral.Fields("CodFamilia").Value), 0, RsGral.Fields("CodFamilia").Value)
            dbcFamilia.Text = BuscaFamilia(mintCodFamilia)
            dbcFamilia.Tag = dbcFamilia.Text
            mintCodLinea = IIf(IsDBNull(RsGral.Fields("COdLinea").Value), 0, RsGral.Fields("COdLinea").Value)
            dbcLinea.Text = BuscaLinea(mintCodLinea, mintCodFamilia)
            dbcLinea.Tag = dbcLinea.Text
            mintCodSubLinea = IIf(IsDBNull(RsGral.Fields("CodSubLinea").Value), 0, RsGral.Fields("CodSubLinea").Value)
            dbcSubLinea.Text = BuscaSubLinea(mintCodSubLinea, mintCodFamilia, mintCodLinea)
            dbcSubLinea.Tag = dbcSubLinea.Text
            mintCodKilates = IIf(IsDBNull(RsGral.Fields("CodKilates").Value), 0, RsGral.Fields("CodKilates").Value)
            dbcKilates.Text = BuscaKilataje(mintCodKilates)
            dbcKilates.Tag = dbcKilates.Text
            txtDescripcion.Text = Trim(RsGral.Fields("DescArticulo").Value)
            txtDescripcion.Tag = txtDescripcion.Text
            mintCodMaterial = IIf(IsDBNull(RsGral.Fields("CodTipoMaterial").Value), 0, RsGral.Fields("CodTipoMaterial").Value)
            dbcMaterial.Text = BuscaTipoMaterial(mintCodMaterial)
            dbcMaterial.Tag = dbcMaterial.Text
            mblnFueraChange = False
            txtCodigodelProveedor.Text = Trim(RsGral.Fields("CodigoArticuloProv").Value)
            txtCodigodelProveedor.Tag = txtCodigodelProveedor.Text
            txtAdicional.Text = Trim(RsGral.Fields("Adicional").Value)
            txtAdicional.Tag = Trim(RsGral.Fields("Adicional").Value)
            ''' 27OCT2010 - MAVF
            txtMDSPeso.Text = VB6.Format(RsGral.Fields("mdsPeso").Value, "##0.00")
            txtMDSPeso.Tag = VB6.Format(RsGral.Fields("mdsPeso").Value, "##0.00")
            txtMDSColor.Text = Trim(RsGral.Fields("mdsColor").Value)
            txtMDSColor.Tag = Trim(RsGral.Fields("mdsColor").Value)
            txtMDSPureza.Text = Trim(RsGral.Fields("mdsPureza").Value)
            txtMDSPureza.Tag = Trim(RsGral.Fields("mdsPureza").Value)
            txtMDSCertificado.Text = Trim(RsGral.Fields("mdsCertificado").Value)
            txtMDSCertificado.Tag = Trim(RsGral.Fields("mdsCertificado").Value)
            ''' ***************************************************

            txtPrecioenDolares.Text = VB6.Format(RsGral.Fields("PrecioPubDolar").Value, "###,##0.00")
            txtPrecioenDolares.Tag = VB6.Format(RsGral.Fields("PrecioPubDolar").Value, "###,##0.00")

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
            txtCostoActual.Text = VB6.Format(nCostoActual, "###,##0.00")
            txtCostoActual.Tag = VB6.Format(nCostoActual, "###,##0.00")
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
                dbcFamilia.Text = False
                dbcLinea.Text = False
                dbcSubLinea.Text = False
                dbcKilates.Text = False
                dbcMaterial.Text = False
                txtCodigodelProveedor.ReadOnly = False
                txtAdicional.ReadOnly = False
                txtMDSPeso.ReadOnly = False '''27OCT2010 - MAVF
                txtMDSColor.ReadOnly = False '''27OCT2010 - MAVF
                txtMDSPureza.ReadOnly = False '''27OCT2010 - MAVF
                txtMDSCertificado.ReadOnly = False '''27OCT2010 - MAVF
            Else
                dbcUnidad.Text = True
                dbcFamilia.Text = True
                dbcLinea.Text = True
                dbcSubLinea.Text = True
                dbcKilates.Text = True
                dbcMaterial.Text = True
                txtCodigodelProveedor.ReadOnly = True
                txtAdicional.ReadOnly = True
                txtMDSPeso.ReadOnly = True '''27OCT2010 - MAVF
                txtMDSColor.ReadOnly = True '''27OCT2010 - MAVF
                txtMDSPureza.ReadOnly = True '''27OCT2010 - MAVF
                txtMDSCertificado.ReadOnly = True '''27OCT2010 - MAVF
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
            Case Trim(dbcFamilia.Text) <> Trim(dbcFamilia.Tag)
                Cambios = True
            Case Trim(dbcLinea.Text) <> Trim(dbcLinea.Tag)
                Cambios = True
            Case Trim(dbcSubLinea.Text) <> Trim(dbcSubLinea.Tag)
                Cambios = True
            Case Trim(dbcKilates.Text) <> Trim(dbcKilates.Tag)
                Cambios = True
            Case Trim(dbcMaterial.Text) <> Trim(dbcMaterial.Tag)
                Cambios = True
            Case Trim(txtAdicional.Text) <> Trim(txtAdicional.Tag)
                Cambios = True
            Case Trim(txtMDSPeso.Text) <> Trim(txtMDSPeso.Tag) '''27OCT2010 - MAVF
                Cambios = True
            Case Trim(txtMDSColor.Text) <> Trim(txtMDSColor.Tag) '''27OCT2010 - MAVF
                Cambios = True
            Case Trim(txtMDSPureza.Text) <> Trim(txtMDSPureza.Tag) '''27OCT2010 - MAVF
                Cambios = True
            Case Trim(txtMDSCertificado.Text) <> Trim(txtMDSCertificado.Tag) '''27OCT2010 - MAVF
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

    ''' 27OCT2010 - MAVF Ver
    Private Function ValidaDatosManejoDiamanteSuelto() As Boolean
        Dim vlResult As Byte

        lblEstatus.Text = ""
        If Trim(txtMDSPeso.Text) = "" Or CDec(ModEstandar.Numerico((txtMDSPeso.Text))) = 0 Then lblEstatus.Text = "CT-" '''08NOV2010 - MAVF
        If Trim(txtMDSColor.Text) = "" Then lblEstatus.Text = lblEstatus.Text & "COLOR-"
        If Trim(txtMDSPureza.Text) = "" Then lblEstatus.Text = lblEstatus.Text & "Q-"

        ''' Si no estan capturados todos los datos o si falta alguno, entonces notifica y pregunta si kiere capturarlossss
        If Trim(lblEstatus.Text) <> "" Then
            lblEstatus.Text = Mid(lblEstatus.Text, 1, Len(lblEstatus.Text) - 1)
            vlResult = MsgBox("Datos de Diamante Suelto no capturados..." & vbNewLine & "Desea ingresarlos ???", MsgBoxStyle.Information + MsgBoxStyle.YesNo, gstrNombCortoEmpresa)
            If vlResult = MsgBoxResult.Yes Then
                txtMDSPeso.Focus()
                ValidaDatosManejoDiamanteSuelto = False '''Si kiere capturarlos
            Else
                ValidaDatosManejoDiamanteSuelto = True '''No kiere capturarlos
            End If
        Else
            ValidaDatosManejoDiamanteSuelto = True '''Todos los datos estan capturados
        End If

    End Function

    Public Function ValidaDatos() As Boolean
        On Error Resume Next
        Select Case True
            Case mintCodUnidad = 0
                MsgBox("Debe indicar la unidad del artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                If dbcUnidad.Enabled Then dbcUnidad.Focus() Else btnAceptar.Focus()
                Exit Function
            Case mintCodFamilia = 0
                MsgBox("Falta indicar la Familia del Artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                If dbcFamilia.Enabled Then dbcFamilia.Focus() Else btnAceptar.Focus()
                Exit Function
                ''' 27OCT2010 - MAVF
                '''Case mintCodKilates = 0
                '''    MsgBox "Falta indicar el kilataje del artículo", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                '''    ValidaDatos = False
                '''    If dbcKilates.Enabled Then dbcKilates.SetFocus Else btnAceptar.SetFocus
                '''    Exit Function
            Case mintCodMaterial = 0
                MsgBox("Falta indicar el Tipo de Material del Artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                If dbcMaterial.Enabled Then dbcMaterial.Focus() Else btnAceptar.Focus()
                Exit Function
            Case (Trim(txtPrecioenDolares.Text) <> "") And (Not optMoneda(0).Checked And Not optMoneda(1).Checked)
                MsgBox("Debe indicar la moneda del precio publico del artículo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                If txtPrecioenDolares.Enabled Then txtPrecioenDolares.Focus()
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
        Dim lFam As String
        Dim lLin As String
        Dim lSubL As String
        Dim lKil As String
        Dim lTipoM As String

        lFam = "" : lLin = "" : lSubL = "" : lKil = "" : lTipoM = ""

        If Trim(frmCXPOrdenCompra.mshFlex.get_TextMatrix(nRowAct, C_ColSTATUS)) <> C_RESURTIDO Then '''08NOV2010 - MAVF
            If Not ValidaDatosManejoDiamanteSuelto() Then Exit Function '''27OCT2010 - MAVF
        End If

        If Not ValidaDatos() Then Exit Function

        'Realiza una búsqueda POR DESCRIPCIÓN para ver si el artículo ya existe en el Grid
        With frmCXPOrdenCompra.mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then Exit For

                ''' 27OCT2010 - MAVF - SE AGREGARON 4 CAMPOS PARA MANEJO DE DIAMANTE SUELTO
                '''valida la clasificación del artículo
                If frmCXPOrdenCompra.mintRenglonAct <> I Then
                    If mintCodFamilia = CDec(Numerico(.get_TextMatrix(I, C_COLCODFAMILIA))) And mintCodLinea = CDec(Numerico(.get_TextMatrix(I, C_COLCODLINEA))) And mintCodSubLinea = CDec(Numerico(.get_TextMatrix(I, C_COLCODSUBLINEA))) And mintCodKilates = CDec(Numerico(.get_TextMatrix(I, C_COLCODKILATES))) And mintCodMaterial = CDec(Numerico(.get_TextMatrix(I, C_COLCODTIPOMATERIAL))) And Trim(txtAdicional.Text) = Trim(.get_TextMatrix(I, C_COLADICIONAL)) And Trim(txtCodigodelProveedor.Text) = Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)) And Trim(txtMDSPeso.Text) = Trim(.get_TextMatrix(I, C_ColMDSPESO)) And Trim(txtMDSColor.Text) = Trim(.get_TextMatrix(I, C_ColMDSCOLOR)) And Trim(txtMDSPureza.Text) = Trim(.get_TextMatrix(I, C_ColMDSPUREZA)) And Trim(txtMDSCertificado.Text) = Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)) Then
                        Select Case MsgBox("Ya existe un elemento de éstos en la Orden de Compra." & vbNewLine & "¿Desea volver a definir el artículo?", MsgBoxStyle.Information + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                            Case MsgBoxResult.Yes
                                dbcFamilia.Focus()
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
                    If Trim(txtImagen.Text) <> "" And Trim(txtImagen.Text) = Trim(.get_TextMatrix(I, C_ColIMAGEN)) Then
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

        DefineCondicionesJoy(lFam, lLin, lSubL, lKil, lTipoM)
        gStrSql = "SELECT * FROM CatArticulos (Nolock) WHERE CodProveedor = " & frmCXPOrdenCompra.mintCodProveedor & " And " & lFam & " And " & lLin & " And " & lSubL & " And " & lKil & " And " & lTipoM & " And ltrim(rtrim(Adicional)) = '" & Trim(txtAdicional.Text) & "' And ltrim(rtrim(CodigoArticuloProv)) = '" & Trim(txtCodigodelProveedor.Text) & "' "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            'De existir, llena los demás datos y cambia el estatus del registro en el Grid a Resurtido
            With frmCXPOrdenCompra.mshFlex

                .set_TextMatrix(nRowAct, C_COLDESCRIPCION, Trim(txtDescripcion.Text))
                .set_TextMatrix(nRowAct, C_COLUNIDAD, Trim(dbcUnidad.Text))
                .set_TextMatrix(nRowAct, C_ColCODGRUPO, gCODJOYERIA)
                .set_TextMatrix(nRowAct, C_COLCODFAMILIA, mintCodFamilia)
                .set_TextMatrix(nRowAct, C_COLCODLINEA, mintCodLinea)
                .set_TextMatrix(nRowAct, C_COLCODSUBLINEA, mintCodSubLinea)
                .set_TextMatrix(nRowAct, C_COLCODKILATES, mintCodKilates)
                .set_TextMatrix(nRowAct, C_COLCODMARCA, 0)
                .set_TextMatrix(nRowAct, C_COLCODMODELO, 0)
                .set_TextMatrix(nRowAct, C_COLCODTIPOMATERIAL, mintCodMaterial)
                .set_TextMatrix(nRowAct, C_COLGENERO, 0)
                .set_TextMatrix(nRowAct, C_COLMOVIMIENTO, 0)
                .set_TextMatrix(nRowAct, C_COLCRONO, False)
                .set_TextMatrix(nRowAct, C_COLCODIGOARTICULOPROV, Trim(txtCodigodelProveedor.Text))
                .set_TextMatrix(nRowAct, C_ColSTATUS, "")
                .set_TextMatrix(nRowAct, C_COLDESCTO, 0)
                .set_TextMatrix(nRowAct, C_COLDESCTOPORC, CDec(Numerico((frmCXPOrdenCompra.txtPorcDescto.Text))))
                .set_TextMatrix(nRowAct, C_COLDESCTOPORCTAG, CDec(Numerico((frmCXPOrdenCompra.txtPorcDescto.Text))))

                .set_TextMatrix(nRowAct, C_COLADICIONAL, Trim(txtAdicional.Text))
                .set_TextMatrix(nRowAct, C_COLPRECIOPUBDOLAR, CDec(Numerico(txtPrecioenDolares.Text)))
                .set_TextMatrix(nRowAct, C_COLMONEDAPP, IIf(optMoneda(0).Checked = True, "D", "P"))
                .set_TextMatrix(nRowAct, C_COLPRECIOUNITARIO, CDec(Numerico((txtCostoActual.Text))))
                .set_TextMatrix(nRowAct, C_COLCANTIDAD, CDec(Numerico((txtCantidadCompra.Text))))
                .set_TextMatrix(nRowAct, C_COLORIGENANT, IIf(Trim(dbcOrigen.Text) = "", "", CInt(Numerico((dbcOrigen.Text)))))
                .set_TextMatrix(nRowAct, C_ColCODIGOANT, IIf(Trim(txtCodArtAnterior.Text) = "", "", CInt(Numerico(txtCodArtAnterior.Text))))
                .set_TextMatrix(nRowAct, C_ColIMAGEN, Trim(txtImagen.Text))

                ''' 27OCT2010 - MAVF
                .set_TextMatrix(nRowAct, C_ColMDSPESO, ModEstandar.Numerico((txtMDSPeso.Text)))
                .set_TextMatrix(nRowAct, C_ColMDSCOLOR, Trim(txtMDSColor.Text))
                .set_TextMatrix(nRowAct, C_ColMDSPUREZA, Trim(txtMDSPureza.Text))
                .set_TextMatrix(nRowAct, C_ColMDSCERTIFICADO, Trim(txtMDSCertificado.Text))
                ''' *********************************************************************

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
                    .set_TextMatrix(nRowAct, C_ColSTATUS, C_RESURTIDO)
                    MsgBox("Ya existe un artículo con esta descripción en el sistema. Si considera que este artículo" & vbNewLine & "no debe ser un resurtido de mercancía existente, revise sus datos", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Advertencia ...")
                Else
                    .set_TextMatrix(nRowAct, C_ColSTATUS, "")
                End If

                .Row = frmCXPOrdenCompra.mintRenglonAct + 1
                .Col = C_COLDESCRIPCION
            End With
        Else
            'Si no existe deja el estatus en Vigente y llena los demás datos
            With frmCXPOrdenCompra.mshFlex
                .set_TextMatrix(nRowAct, C_COLCODIGO, "")
                .set_TextMatrix(nRowAct, C_COLDESCRIPCION, Trim(txtDescripcion.Text))
                .set_TextMatrix(nRowAct, C_COLUNIDAD, Trim(dbcUnidad.Text))
                .set_TextMatrix(nRowAct, C_ColCODGRUPO, gCODJOYERIA)
                .set_TextMatrix(nRowAct, C_COLCODFAMILIA, mintCodFamilia)
                .set_TextMatrix(nRowAct, C_COLCODLINEA, mintCodLinea)
                .set_TextMatrix(nRowAct, C_COLCODSUBLINEA, mintCodSubLinea)
                .set_TextMatrix(nRowAct, C_COLCODKILATES, mintCodKilates)
                .set_TextMatrix(nRowAct, C_COLCODMARCA, 0)
                .set_TextMatrix(nRowAct, C_COLCODMODELO, 0)
                .set_TextMatrix(nRowAct, C_COLCODTIPOMATERIAL, mintCodMaterial)
                .set_TextMatrix(nRowAct, C_COLGENERO, 0)
                .set_TextMatrix(nRowAct, C_COLMOVIMIENTO, 0)
                .set_TextMatrix(nRowAct, C_COLCRONO, False)
                .set_TextMatrix(nRowAct, C_COLCODIGOARTICULOPROV, Trim(txtCodigodelProveedor.Text))
                .set_TextMatrix(nRowAct, C_ColSTATUS, "")
                .set_TextMatrix(nRowAct, C_COLDESCTO, 0)
                .set_TextMatrix(nRowAct, C_COLDESCTOPORC, CDec(Numerico((frmCXPOrdenCompra.txtPorcDescto.Text))))
                .set_TextMatrix(nRowAct, C_COLDESCTOPORCTAG, CDec(Numerico((frmCXPOrdenCompra.txtPorcDescto.Text))))

                .set_TextMatrix(nRowAct, C_COLADICIONAL, Trim(txtAdicional.Text))
                .set_TextMatrix(nRowAct, C_COLPRECIOPUBDOLAR, CDec(Numerico(txtPrecioenDolares.Text)))
                .set_TextMatrix(nRowAct, C_COLMONEDAPP, IIf(optMoneda(0).Checked = True, "D", "P"))
                .set_TextMatrix(nRowAct, C_COLPRECIOUNITARIO, CDec(Numerico((txtCostoActual.Text))))
                .set_TextMatrix(nRowAct, C_COLCANTIDAD, CDec(Numerico((txtCantidadCompra.Text))))
                .set_TextMatrix(nRowAct, C_COLORIGENANT, IIf(Trim(dbcOrigen.Text) = "", "", CInt(Numerico((dbcOrigen.Text)))))
                .set_TextMatrix(nRowAct, C_ColCODIGOANT, IIf(Trim(txtCodArtAnterior.Text) = "", "", CInt(Numerico(txtCodArtAnterior.Text))))
                .set_TextMatrix(nRowAct, C_ColIMAGEN, Trim(txtImagen.Text))

                ''' 27OCT2010 - MAVF
                .set_TextMatrix(nRowAct, C_ColMDSPESO, Trim(txtMDSPeso.Text))
                .set_TextMatrix(nRowAct, C_ColMDSCOLOR, Trim(txtMDSColor.Text))
                .set_TextMatrix(nRowAct, C_ColMDSPUREZA, Trim(txtMDSPureza.Text))
                .set_TextMatrix(nRowAct, C_ColMDSCERTIFICADO, Trim(txtMDSCertificado.Text))
                ''' *********************************************************************

                If Cambios() Then .set_TextMatrix(nRowAct, C_COLSTATUSX, "M")

                .Row = frmCXPOrdenCompra.mintRenglonAct + 1
                .Col = C_COLDESCRIPCION
            End With
        End If

        frmCXPOrdenCompra.ActualizaCantidades()
        Guardar = True
    End Function

    Public Sub FormaDescripcion()
        If Trim(cKilates) = Trim(cSINKILATES) Then cKilates = " "
        cFamilia = ""
        cLinea = IIf(Trim(cLinea) = cINDEFINIDA, "", cLinea)
        cSubLinea = IIf(Trim(cSubLinea) = cINDEFINIDO, "", cSubLinea)
        cKilates = IIf(Trim(cKilates) = cINDEFINIDO, "", cKilates)
        cTipoMaterial = IIf(Trim(cTipoMaterial) = Trim(cINDEFINIDO), "", Trim(cTipoMaterial)) & " "
        strAdicional = Trim(txtAdicional.Text)
        cDescripcion = Trim(cLinea) & " " & Trim(cSubLinea) & " " & Trim(cKilates) & " " & Trim(cTipoMaterial) & " " & strAdicional
        txtDescripcion.Text = Trim(cDescripcion)
    End Sub

    Private Sub btnAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAceptar.Click
        If Guardar() Then
            mblnNuevo = True
            mblnCancelar = True
            Me.Close()
            frmCXPOrdenCompra.MuestraClasificacion()
        End If
    End Sub

    Private Sub btnCancelar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancelar.Click
        mblnCancelar = True
        Me.Close()
    End Sub

    Private Sub btnFamilia_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnFamilia.Click
        Me.Enabled = False
        frmCorpoABCFamilias.Tag = UCase(Me.Name)
        frmCorpoABCFamilias.Show()
        frmCorpoABCFamilias.mintCodGrupo = gCODJOYERIA
        frmCorpoABCFamilias.LlenaDatos()
    End Sub

    Private Sub btnLinea_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnLinea.Click
        If mintCodFamilia = 0 Then
            MsgBox("Antes de definir una nueva Línea de Joyería, debe seleccionar la familia", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Me.dbcFamilia.Focus()
            Exit Sub
        End If
        Me.Enabled = False
        frmCorpoABCLineas.Tag = UCase(Me.Name)
        frmCorpoABCLineas.Show()
        frmCorpoABCLineas.mintCodGrupo = gCODJOYERIA
        frmCorpoABCLineas.mintCodFamilia = mintCodFamilia
        frmCorpoABCLineas.LlenaDatos()
    End Sub

    Private Sub btnSubLinea_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSubLinea.Click
        If mintCodFamilia = 0 Then
            MsgBox("Antes de definir una nueva SubLínea, debe seleccionar la familia", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Me.dbcFamilia.Focus()
            Exit Sub
        ElseIf mintCodLinea = 0 Then
            MsgBox("Antes de definir una nueva SubLínea, debe seleccionar la Línea a la que pertenecerá", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Me.dbcLinea.Focus()
            Exit Sub
        End If
        Me.Enabled = False
        frmCorpoABCSubLineas.Tag = UCase(Me.Name)
        frmCorpoABCSubLineas.Show()
        'frmCorpoABCSubLineas.mintCodGrupo = gCODJOYERIA
        frmCorpoABCSubLineas.mintCodFamilia = mintCodFamilia
        frmCorpoABCSubLineas.mintCodLinea = mintCodLinea
        frmCorpoABCSubLineas.LlenaDatos()
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

    Private Sub cmdBuscarImagen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBuscarImagen.Click
        'frmCorpoBuscarImagen.ShowDialog()
    End Sub

    Private Sub dbcKilates_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcKilates.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If Trim(cKilates) = Trim(cSINKILATES) Then
            cKilates = " "
        Else
            cKilates = Trim(Me.dbcKilates.Text)
        End If
        Call FormaDescripcion()

        If mblnFueraChange Then
            Exit Sub
        End If

        lStrSql = "SELECT codKilates, LTrim(RTrim(descKilates)) as descKilates FROM CatKilates Where LTrim(RTrim(descKilates)) LIKE '" & Trim(Me.dbcKilates.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcKilates))

        If Trim(Me.dbcKilates.Text) = "" Then
            mintCodKilates = 0
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcKilates_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcKilates.Enter
        Pon_Tool()
        gStrSql = "SELECT codKilates, LTrim(RTrim(descKilates)) as descKilates FROM CatKilates"
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcKilates))
    End Sub

    Private Sub dbcKilates_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcKilates.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcSubLinea.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcKilates_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcKilates.KeyUp
        Dim Aux As String
        Aux = Trim(Me.dbcKilates.Text)
        If Me.dbcKilates.SelectedItem <> 0 Then
            dbcKilates_Leave(dbcKilates, New System.EventArgs())
        End If
        Me.dbcKilates.Text = Aux
    End Sub

    Private Sub dbcKilates_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcKilates.Leave
        Dim Aux As Integer
        Dim cDescripcion As String
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        cDescripcion = Trim(Me.dbcKilates.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDO) Then
            cDescripcion = ""
        End If
        ''" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "%'")
        gStrSql = "SELECT codKilates, LTrim(RTrim(descKilates)) as descKilates FROM CatKilates Where LTrim(RTrim(descKilates)) LIKE '" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "%'")
        Aux = mintCodKilates
        mintCodKilates = 0
        ModDCombo.DCLostFocus((Me.dbcKilates), gStrSql, mintCodKilates)

        If Trim(cKilates) = Trim(cSINKILATES) Then
            cKilates = " "
        Else
            cKilates = Trim(Me.dbcKilates.Text)
        End If

        If mintCodKilates = 0 Then
            mblnFueraChange = True
            Me.dbcKilates.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
        Call FormaDescripcion()
    End Sub

    Private Sub dbcKilates_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcKilates.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcKilates.Text)
        'If Me.dbcKilates.SelectedItem <> 0 Then
        '    dbcKilates_Leave(dbcKilates, New System.EventArgs())
        'End If
        Me.dbcKilates.Text = Aux
    End Sub

    Private Sub dbcMaterial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMaterial.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String
        If mblnFueraChange Then Exit Sub
        cTipoMaterial = Trim(Me.dbcMaterial.Text)
        lStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial LIKE '" & Trim(Me.dbcMaterial.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcMaterial)
        cTipoMaterial = BuscaTipoMaterialDescCorta(mintCodMaterial)
        FormaDescripcion()
        If Trim(Me.dbcMaterial.Text) = "" Then
            mintCodMaterial = 0
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcMaterial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMaterial.Enter
        Pon_Tool()
        gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial ORDER BY descTipoMaterial"
        ModDCombo.DCGotFocus(gStrSql, dbcMaterial)
    End Sub

    Private Sub dbcMaterial_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcMaterial.KeyDown
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Me.dbcKilates.Focus()
            Case Else
                tecla = eventArgs.KeyCode
        End Select
    End Sub

    Private Sub dbcMaterial_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcMaterial.KeyUp
        Dim Aux As String
        Aux = Trim(Me.dbcMaterial.Text)
        'If Me.dbcMaterial.SelectedItem <> 0 Then
        '    dbcMaterial_Leave(dbcMaterial, New System.EventArgs())
        'End If
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
        gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial LIKE '" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "'")
        mintCodMaterial = 0
        ModDCombo.DCLostFocus(dbcMaterial, gStrSql, mintCodMaterial)
        cTipoMaterial = BuscaTipoMaterialDescCorta(mintCodMaterial)
        If mintCodMaterial = 0 Then
            mblnFueraChange = True
            Me.dbcMaterial.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
        FormaDescripcion()
    End Sub

    Private Sub dbcMaterial_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcMaterial.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcMaterial.Text)
        If Me.dbcMaterial.SelectedItem <> 0 Then
            dbcMaterial_Leave(dbcMaterial, New System.EventArgs())
        End If
        Me.dbcMaterial.Text = Aux
    End Sub

    Private Sub dbcOrigen_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen.CursorChanged
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

    Private Sub dbcOrigen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dbcOrigen.KeyPress
        ModEstandar.gp_CampoNumerico(eventSender.keyAscii)
        eventSender.keyAscii = ModEstandar.MskCantidad((dbcOrigen.Text), eventSender.keyAscii, 1, 0, (dbcOrigen.SelectionStart))
    End Sub

    Private Sub dbcOrigen_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcOrigen.KeyUp
        Dim Aux As String
        Aux = dbcOrigen.Text
        If dbcOrigen.SelectedItem <> 0 Then
            dbcOrigen_Leave(dbcOrigen, New System.EventArgs())
        End If
        dbcOrigen.Text = Aux
    End Sub

    Private Sub dbcOrigen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen.Leave
        gStrSql = "SELECT CodAlmacenOrigen, CodAlmacenOrigen AS DescAlmacen From CatOrigen  WHERE CodAlmacenOrigen LIKE '" & Trim(dbcOrigen.Text) & "%'  ORDER BY DescAlmacenOrigen "
        DCLostFocus(dbcOrigen, gStrSql, intCodAlmacenOrigen)
    End Sub

    Private Sub dbcOrigen_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcOrigen.MouseUp
        Dim Aux As String
        Aux = dbcOrigen.Text
        If dbcOrigen.SelectedItem <> 0 Then
            dbcOrigen_Leave(dbcOrigen, New System.EventArgs())
        End If
        dbcOrigen.Text = Aux
    End Sub

    Private Sub dbcSubLinea_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSubLinea.KeyUp
        Dim Aux As String
        Aux = Trim(Me.dbcSubLinea.Text)
        If Me.dbcSubLinea.SelectedItem <> 0 Then
            dbcSubLinea_Leave(dbcSubLinea, New System.EventArgs())
        End If
        Me.dbcSubLinea.Text = Aux
    End Sub

    Private Sub dbcSubLinea_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcSubLinea.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcSubLinea.Text)
        If Me.dbcSubLinea.SelectedItem <> 0 Then
            dbcSubLinea_Leave(dbcSubLinea, New System.EventArgs())
        End If
        Me.dbcSubLinea.Text = Aux
    End Sub

    Private Sub dbcUnidad_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcUnidad.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad LIKE '" & Trim(Me.dbcUnidad.Text) & "%' Order by  descUnidad "
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcUnidad))

        If Trim(Me.dbcUnidad.Text) = "" Then
            mintCodUnidad = 0
        End If

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcUnidad_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcUnidad.Enter
        Pon_Tool()
        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades ORDER BY descUnidad "
        ModDCombo.DCGotFocus(gStrSql, dbcUnidad)
    End Sub

    Private Sub dbcUnidad_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcUnidad.KeyDown
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
        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad = '" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "'") & " Order by descUnidad "
        mintCodUnidad = 0
        ModDCombo.DCLostFocus((Me.dbcUnidad), gStrSql, mintCodUnidad)
        If mintCodUnidad = 0 Then
            mblnFueraChange = True
            Me.dbcUnidad.Text = cINDEFINIDA
            mblnFueraChange = False
        End If
    End Sub

    Private Sub dbcUnidad_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcUnidad.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcUnidad.Text)
        If Me.dbcUnidad.SelectedItem <> 0 Then
            dbcUnidad_Leave(dbcUnidad, New System.EventArgs())
        End If
        Me.dbcUnidad.Text = Aux
    End Sub

    Private Sub dbcFamilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcFamilia.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        cFamilia = Trim(Me.dbcFamilia.Text)
        Call FormaDescripcion()

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia LIKE '" & Trim(Me.dbcFamilia.Text) & "%' Order by descFamilia "
        ModDCombo.DCChange(lStrSql, tecla, dbcFamilia)

        If Trim(Me.dbcFamilia.Text) = "" Then
            mintCodFamilia = 0
            mblnFueraChange = True
            Me.dbcLinea.Text = cINDEFINIDA
            mintCodLinea = 0
            Me.dbcSubLinea.Text = cINDEFINIDA
            mintCodSubLinea = 0
            mblnFueraChange = False
        End If

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcFamilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcFamilia.Enter
        Pon_Tool()
        gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " ORDER BY descFamilia "
        ModDCombo.DCGotFocus(gStrSql, dbcFamilia)
    End Sub

    Private Sub dbcFamilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcFamilia.KeyDown
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Me.dbcUnidad.Focus()
            Case Else
                tecla = eventArgs.KeyCode
        End Select
    End Sub

    Private Sub dbcFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcFamilia.Leave
        Dim I As Integer
        Dim Aux As Integer 'Almacena el anterior
        Dim cDescripcion As String
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        cDescripcion = Trim(Me.dbcFamilia.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDA) Then
            cDescripcion = ""
        End If
        gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia = '" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "'") & " Order by descFamilia "
        Aux = mintCodFamilia
        mintCodFamilia = 0
        ModDCombo.DCLostFocus(dbcFamilia, gStrSql, mintCodFamilia)
        cFamilia = Trim(Me.dbcFamilia.Text)
        If Aux <> mintCodFamilia Then
            mblnFueraChange = True
            If mintCodFamilia = 0 Then
                Me.dbcFamilia.Text = cINDEFINIDA
            End If
            mintCodLinea = 0
            Me.dbcLinea.Text = cINDEFINIDA
            mintCodSubLinea = 0
            Me.dbcSubLinea.Text = cINDEFINIDA
            mblnFueraChange = False
        Else
            If mintCodFamilia = 0 Then
                mblnFueraChange = True
                Me.dbcFamilia.Text = cINDEFINIDA
                mblnFueraChange = False
            End If
        End If
        FormaDescripcion()
    End Sub

    Private Sub dbcLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcLinea.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String
        cLinea = Trim(Me.dbcLinea.Text)
        Call FormaDescripcion()

        If mblnFueraChange Then
            Exit Sub
        Else
            mblnFueraChange = True
            mintCodSubLinea = 0
            Me.dbcSubLinea.Text = cINDEFINIDA
            mblnFueraChange = False
        End If

        lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintCodFamilia & " and descLinea LIKE '" & Trim(Me.dbcLinea.Text) & "%' Order by descLinea "
        ModDCombo.DCChange(lStrSql, tecla, dbcLinea)

        If Trim(Me.dbcLinea.Text) = "" Then mintCodLinea = 0

MError:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcLinea.Enter
        Pon_Tool()
        gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintCodFamilia & " ORDER BY descLinea "
        ModDCombo.DCGotFocus(gStrSql, dbcLinea)
    End Sub

    Private Sub dbcLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcLinea.KeyDown
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Me.dbcFamilia.Focus()
            Case Else
                tecla = eventArgs.KeyCode
        End Select
    End Sub

    Private Sub dbcLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcLinea.Leave
        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        cDescripcion = Trim(Me.dbcLinea.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDA) Then
            cDescripcion = ""
        End If

        gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintCodFamilia & " and descLinea = '" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "'") & " Order by descLinea "
        Aux = mintCodLinea
        mintCodLinea = 0
        ModDCombo.DCLostFocus(dbcLinea, gStrSql, mintCodLinea)
        cLinea = Trim(Me.dbcLinea.Text)
        If Aux <> mintCodLinea Then 'Cambia la descripción
            mblnFueraChange = True
            If mintCodLinea = 0 Then
                Me.dbcLinea.Text = cINDEFINIDA
            End If
            mintCodSubLinea = 0
            Me.dbcSubLinea.Text = cINDEFINIDA
            mblnFueraChange = False
        Else
            If mintCodLinea = 0 Then
                mblnFueraChange = True
                Me.dbcLinea.Text = cINDEFINIDA
                mblnFueraChange = False
            End If
        End If
        Call FormaDescripcion()
    End Sub

    Private Sub dbcSubLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSubLinea.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        cSubLinea = Trim(dbcSubLinea.Text)
        FormaDescripcion()
        If mblnFueraChange Then Exit Sub
        lStrSql = "SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as descSubLinea FROM catSubLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintCodFamilia & " and codLinea = " & mintCodLinea & " and descSubLinea LIKE '" & Trim(Me.dbcSubLinea.Text) & "%' Order by descSubLinea "
        ModDCombo.DCChange(lStrSql, tecla, dbcSubLinea)
        If Trim(Me.dbcSubLinea.Text) = "" Then
            mintCodSubLinea = 0
        End If
        cSubLinea = Trim(BuscaSubLineaDescCorta(mintCodGrupo, mintCodFamilia, mintCodLinea, mintCodSubLinea))
        FormaDescripcion()

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcSubLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSubLinea.Enter
        Pon_Tool()
        gStrSql = "SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as descSubLinea FROM catSubLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintCodFamilia & " and codLinea = " & mintCodLinea & " Order by descSubLinea "
        ModDCombo.DCGotFocus(gStrSql, dbcSubLinea)
    End Sub

    Private Sub dbcSubLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSubLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcLinea.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSubLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSubLinea.Leave
        Dim I As Integer
        Dim cDescripcion As String
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        cDescripcion = Trim(Me.dbcSubLinea.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDA) Then
            cDescripcion = ""
        End If
        gStrSql = "SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as descSubLinea FROM catSubLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintCodFamilia & " and codLinea = " & mintCodLinea & " and descSubLinea = '" & IIf(Trim(cDescripcion) = "", "'", cDescripcion & "'") & " Order by descSubLinea "
        mintCodSubLinea = 0
        ModDCombo.DCLostFocus(dbcSubLinea, gStrSql, mintCodSubLinea)
        cSubLinea = Trim(BuscaSubLineaDescCorta(mintCodGrupo, mintCodFamilia, mintCodLinea, mintCodSubLinea))
        If mintCodSubLinea = 0 Then
            mblnFueraChange = True
            Me.dbcSubLinea.Text = cINDEFINIDA
            mblnFueraChange = False
        End If
        FormaDescripcion()
    End Sub

    Private Sub frmCXPJoyeria_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Me.BringToFront()
    End Sub

    Private Sub frmCXPJoyeria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                Select Case UCase(Trim(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name))
                    Case Is = "TXTPRECIOENDOLARES"
                    Case Else
                        ModEstandar.AvanzarTab(Me)
                End Select
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmCXPJoyeria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPJoyeria_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        mintCodGrupo = gCODJOYERIA
        gstrNombreForma = "FRMCXPJOYERIA"
        intCodAlmacenOrigen = 0
    End Sub

    Private Sub frmCXPJoyeria_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmCXPJoyeria_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.LimpiaDescBarraEstado()
        frmCXPOrdenCompra.Enabled = True
        frmCXPOrdenCompra.mshFlex.Focus()
        gstrNombreForma = ""
        'Me = Nothing
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
        txtCostoActual.Text = VB6.Format(Numerico((txtCostoActual.Text)), "###,###,##0.00")
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

    Private Sub txtMDSCertificado_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSCertificado.Enter
        SelTextoTxt(txtMDSCertificado)
    End Sub

    Private Sub txtMDSCertificado_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSCertificado.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "-")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMDSColor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSColor.Enter
        SelTextoTxt(txtMDSColor)
    End Sub

    Private Sub txtMDSColor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSColor.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
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
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = 13 Then txtMDSPeso.Text = VB6.Format(ModEstandar.Numerico((txtMDSPeso.Text)), "##0.00")
    End Sub

    Private Sub txtMDSPeso_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSPeso.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, ".")
        KeyAscii = ModEstandar.MskCantidad((txtMDSPeso.Text), KeyAscii, 3, 2, (txtMDSPeso.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMDSPureza_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPureza.Enter
        SelTextoTxt(txtMDSPureza)
    End Sub

    Private Sub txtMDSPureza_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSPureza.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii)
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
            txtPrecioenDolares.Text = VB6.Format(Numerico((Me.txtPrecioenDolares.Text)), "###,###,##0.00")
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
        txtPrecioenDolares.Text = VB6.Format(Numerico((txtPrecioenDolares.Text)), "###,###,##0.00")
    End Sub

    Private Sub DefineCondicionesJoy(ByRef nFam As String, ByRef nLin As String, ByRef nSubL As String, ByRef nKil As String, ByRef nTipoM As String)
        If mintCodFamilia <> 0 Then nFam = " CodFamilia = " & mintCodFamilia Else nFam = " CodFamilia Is Null "
        If mintCodLinea <> 0 Then nLin = " CodLinea = " & mintCodLinea Else nLin = " CodLinea Is Null "
        If mintCodSubLinea <> 0 Then nSubL = " CodSubLinea = " & mintCodSubLinea Else nSubL = " CodSubLinea Is Null "
        If mintCodKilates <> 0 Then nKil = " CodKilates = " & mintCodKilates Else nKil = " CodKilates Is Null "
        If mintCodMaterial <> 0 Then nTipoM = " CodTipoMaterial = " & mintCodMaterial Else nTipoM = " CodTipoMaterial Is Null "
    End Sub
End Class