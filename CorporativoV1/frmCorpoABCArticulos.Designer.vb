Imports Microsoft.VisualBasic.Compatibility

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCorpoABCArticulos
#Region "Windows Form Designer generated code "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        'InitializeComponent()
    End Sub
    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer


    ''Required by the Windows Form Designer
    'Private components As System.ComponentModel.IContainer
    'Public ToolTip1 As System.Windows.Forms.ToolTip
    'Public WithEvents chkCodigoAnterior As System.Windows.Forms.CheckBox
    'Public WithEvents txtCodArtAnterior As System.Windows.Forms.TextBox
    'Public WithEvents dbcOrigen As System.Windows.Forms.ComboBox
    'Public WithEvents _lblArticulo_32 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_31 As System.Windows.Forms.Label
    'Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    'Public WithEvents txtDescArticulo As System.Windows.Forms.TextBox
    'Public WithEvents txtCodArticulo As System.Windows.Forms.TextBox
    'Public WithEvents txtMDSCertificado As System.Windows.Forms.TextBox
    'Public WithEvents txtMDSPureza As System.Windows.Forms.TextBox
    'Public WithEvents txtMDSColor As System.Windows.Forms.TextBox
    'Public WithEvents txtMDSPeso As System.Windows.Forms.TextBox
    'Public WithEvents lblEstatus As System.Windows.Forms.Label
    'Public WithEvents lblMDSCertificado As System.Windows.Forms.Label
    'Public WithEvents lblMDSPureza As System.Windows.Forms.Label
    'Public WithEvents lblMDSColor As System.Windows.Forms.Label
    'Public WithEvents lblMDSPeso As System.Windows.Forms.Label
    'Public WithEvents fraDiamanteSuelto As System.Windows.Forms.GroupBox
    'Public WithEvents _txtAdicional_0 As System.Windows.Forms.TextBox
    'Public WithEvents _optMoneda_11 As System.Windows.Forms.RadioButton
    'Public WithEvents _optMoneda_10 As System.Windows.Forms.RadioButton
    'Public WithEvents _fraMoneda_5 As System.Windows.Forms.GroupBox
    'Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    'Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    'Public WithEvents _fraMoneda_0 As System.Windows.Forms.GroupBox
    'Public WithEvents Image1 As System.Windows.Forms.PictureBox
    'Public WithEvents _fraImagen_0 As System.Windows.Forms.GroupBox
    'Public WithEvents _cmdBuscarImagen_0 As System.Windows.Forms.Button
    'Public WithEvents _txtImagen_0 As System.Windows.Forms.TextBox
    'Public WithEvents _Frame4_0 As System.Windows.Forms.GroupBox
    'Public WithEvents _txtCodigodelProveedor_0 As System.Windows.Forms.TextBox
    'Public WithEvents _dbcProveedor_0 As System.Windows.Forms.ComboBox
    'Public WithEvents _cboUnidad_0 As System.Windows.Forms.ComboBox
    'Public WithEvents _cboAlmacen_0 As System.Windows.Forms.ComboBox
    'Public WithEvents _lblArticulo_36 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_35 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_11 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_10 As System.Windows.Forms.Label
    'Public WithEvents _Frame2_0 As System.Windows.Forms.GroupBox
    'Public WithEvents _txtDescripcion_0 As System.Windows.Forms.TextBox
    'Public WithEvents _txtCostoReal_0 As System.Windows.Forms.TextBox
    'Public WithEvents _txtPrecioenDolares_0 As System.Windows.Forms.TextBox
    'Public WithEvents _txtCostoIndirecto_0 As System.Windows.Forms.TextBox
    'Public WithEvents _txtCostoAdicional_0 As System.Windows.Forms.TextBox
    'Public WithEvents _txtCostoFactura_0 As System.Windows.Forms.TextBox
    'Public WithEvents _lblMargen_0 As System.Windows.Forms.Label
    'Public WithEvents Label1 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_34 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_5 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_8 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_7 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_6 As System.Windows.Forms.Label
    'Public WithEvents _Frame1_0 As System.Windows.Forms.GroupBox
    'Public WithEvents _dbcFamilia_0 As System.Windows.Forms.ComboBox
    'Public WithEvents _dbcLinea_0 As System.Windows.Forms.ComboBox
    'Public WithEvents dbcSubLinea As System.Windows.Forms.ComboBox
    'Public WithEvents dbcKilates As System.Windows.Forms.ComboBox
    'Public WithEvents _dbcMaterial_0 As System.Windows.Forms.ComboBox
    'Public WithEvents _lblArticulo_33 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_9 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_29 As System.Windows.Forms.Label
    'Public WithEvents _lblDescripcion_0 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_26 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_37 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_4 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_3 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_2 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_1 As System.Windows.Forms.Label
    'Public WithEvents _fraContenedor_0 As System.Windows.Forms.Panel
    'Public WithEvents _sstArticulo_TabPage0 As System.Windows.Forms.TabPage
    'Public WithEvents _txtAdicional_1 As System.Windows.Forms.TextBox
    'Public WithEvents _optMoneda_7 As System.Windows.Forms.RadioButton
    'Public WithEvents _optMoneda_6 As System.Windows.Forms.RadioButton
    'Public WithEvents _fraMoneda_3 As System.Windows.Forms.GroupBox
    'Public WithEvents _txtDescripcion_1 As System.Windows.Forms.TextBox
    'Public WithEvents _optGenero_0 As System.Windows.Forms.RadioButton
    'Public WithEvents _optGenero_1 As System.Windows.Forms.RadioButton
    'Public WithEvents _optGenero_2 As System.Windows.Forms.RadioButton
    'Public WithEvents _fraArticulo_1 As System.Windows.Forms.GroupBox
    'Public WithEvents _optMovimiento_0 As System.Windows.Forms.RadioButton
    'Public WithEvents _optMovimiento_1 As System.Windows.Forms.RadioButton
    'Public WithEvents _optMovimiento_2 As System.Windows.Forms.RadioButton
    'Public WithEvents _fraArticulo_2 As System.Windows.Forms.GroupBox
    'Public WithEvents Image2 As System.Windows.Forms.PictureBox
    'Public WithEvents _fraImagen_1 As System.Windows.Forms.GroupBox
    'Public WithEvents _optMoneda_3 As System.Windows.Forms.RadioButton
    'Public WithEvents _optMoneda_2 As System.Windows.Forms.RadioButton
    'Public WithEvents _fraMoneda_1 As System.Windows.Forms.GroupBox
    'Public WithEvents _txtCostoReal_1 As System.Windows.Forms.TextBox
    'Public WithEvents _txtPrecioenDolares_1 As System.Windows.Forms.TextBox
    'Public WithEvents _txtCostoIndirecto_1 As System.Windows.Forms.TextBox
    'Public WithEvents _txtCostoAdicional_1 As System.Windows.Forms.TextBox
    'Public WithEvents _txtCostoFactura_1 As System.Windows.Forms.TextBox
    'Public WithEvents _lblMargen_1 As System.Windows.Forms.Label
    'Public WithEvents Label2 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_40 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_41 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_42 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_43 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_44 As System.Windows.Forms.Label
    'Public WithEvents _Frame1_2 As System.Windows.Forms.GroupBox
    'Public WithEvents _txtImagen_1 As System.Windows.Forms.TextBox
    'Public WithEvents _cmdBuscarImagen_1 As System.Windows.Forms.Button
    'Public WithEvents _Frame4_1 As System.Windows.Forms.GroupBox
    'Public WithEvents _txtCodigodelProveedor_1 As System.Windows.Forms.TextBox
    'Public WithEvents _dbcProveedor_1 As System.Windows.Forms.ComboBox
    'Public WithEvents _cboUnidad_1 As System.Windows.Forms.ComboBox
    'Public WithEvents _cboAlmacen_1 As System.Windows.Forms.ComboBox
    'Public WithEvents _lblArticulo_18 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_19 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_20 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_21 As System.Windows.Forms.Label
    'Public WithEvents _Frame2_2 As System.Windows.Forms.GroupBox
    'Public WithEvents chkCrono As System.Windows.Forms.CheckBox
    'Public WithEvents dbcMarca As System.Windows.Forms.ComboBox
    'Public WithEvents dbcModelo As System.Windows.Forms.ComboBox
    'Public WithEvents _dbcMaterial_1 As System.Windows.Forms.ComboBox
    'Public WithEvents _lblArticulo_45 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_27 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_13 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_12 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_14 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_15 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_16 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_17 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_38 As System.Windows.Forms.Label
    'Public WithEvents _lblDescripcion_1 As System.Windows.Forms.Label
    'Public WithEvents _fraContenedor_1 As System.Windows.Forms.Panel
    'Public WithEvents _sstArticulo_TabPage1 As System.Windows.Forms.TabPage
    'Public WithEvents _txtAdicional_2 As System.Windows.Forms.TextBox
    'Public WithEvents _optMoneda_9 As System.Windows.Forms.RadioButton
    'Public WithEvents _optMoneda_8 As System.Windows.Forms.RadioButton
    'Public WithEvents _fraMoneda_4 As System.Windows.Forms.GroupBox
    'Public WithEvents _optMoneda_4 As System.Windows.Forms.RadioButton
    'Public WithEvents _optMoneda_5 As System.Windows.Forms.RadioButton
    'Public WithEvents _fraMoneda_2 As System.Windows.Forms.GroupBox
    'Public WithEvents _txtImagen_2 As System.Windows.Forms.TextBox
    'Public WithEvents _cmdBuscarImagen_2 As System.Windows.Forms.Button
    'Public WithEvents _Frame4_2 As System.Windows.Forms.GroupBox
    'Public WithEvents _txtCodigodelProveedor_2 As System.Windows.Forms.TextBox
    'Public WithEvents _dbcProveedor_2 As System.Windows.Forms.ComboBox
    'Public WithEvents _cboUnidad_2 As System.Windows.Forms.ComboBox
    'Public WithEvents _cboAlmacen_2 As System.Windows.Forms.ComboBox
    'Public WithEvents _lblArticulo_50 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_51 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_52 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_53 As System.Windows.Forms.Label
    'Public WithEvents _Frame2_3 As System.Windows.Forms.GroupBox
    'Public WithEvents _txtCostoFactura_2 As System.Windows.Forms.TextBox
    'Public WithEvents _txtCostoAdicional_2 As System.Windows.Forms.TextBox
    'Public WithEvents _txtCostoIndirecto_2 As System.Windows.Forms.TextBox
    'Public WithEvents _txtPrecioenDolares_2 As System.Windows.Forms.TextBox
    'Public WithEvents _txtCostoReal_2 As System.Windows.Forms.TextBox
    'Public WithEvents _lblMargen_2 As System.Windows.Forms.Label
    'Public WithEvents Label3 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_22 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_23 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_46 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_47 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_48 As System.Windows.Forms.Label
    'Public WithEvents _Frame1_3 As System.Windows.Forms.GroupBox
    'Public WithEvents _txtDescripcion_2 As System.Windows.Forms.TextBox
    'Public WithEvents Image3 As System.Windows.Forms.PictureBox
    'Public WithEvents _fraImagen_2 As System.Windows.Forms.GroupBox
    'Public WithEvents _dbcFamilia_1 As System.Windows.Forms.ComboBox
    'Public WithEvents _dbcLinea_1 As System.Windows.Forms.ComboBox
    'Public WithEvents _dbcMaterial_2 As System.Windows.Forms.ComboBox
    'Public WithEvents _lblArticulo_54 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_49 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_28 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_24 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_25 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_30 As System.Windows.Forms.Label
    'Public WithEvents _lblArticulo_39 As System.Windows.Forms.Label
    'Public WithEvents _lblDescripcion_2 As System.Windows.Forms.Label
    'Public WithEvents _fraContenedor_2 As System.Windows.Forms.Panel
    'Public WithEvents _sstArticulo_TabPage2 As System.Windows.Forms.TabPage
    'Public WithEvents sstArticulo As System.Windows.Forms.TabControl
    'Public WithEvents _lblArticulo_0 As System.Windows.Forms.Label
    'Public WithEvents Frame1 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    'Public WithEvents Frame2 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    'Public WithEvents Frame4 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    'Public WithEvents cboAlmacen As System.Windows.Forms.ComboBox
    'Public WithEvents cboUnidad As System.Windows.Forms.ComboBox
    'Public WithEvents cmdBuscarImagen As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    'Public WithEvents dbcFamilia As System.Windows.Forms.ComboBox
    'Public WithEvents dbcLinea As System.Windows.Forms.ComboBox
    'Public WithEvents dbcMaterial As System.Windows.Forms.ComboBox
    'Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    'Public WithEvents fraArticulo As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    'Public WithEvents fraContenedor As Microsoft.VisualBasic.Compatibility.VB6.PanelArray
    'Public WithEvents fraImagen As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    'Public WithEvents fraMoneda As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    'Public WithEvents lblArticulo As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'Public WithEvents lblDescripcion As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'Public WithEvents lblMargen As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'Public WithEvents optGenero As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    'Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    'Public WithEvents optMovimiento As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    'Public WithEvents txtAdicional As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtCodigodelProveedor As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtCostoAdicional As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtCostoFactura As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtCostoIndirecto As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtCostoReal As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtDescripcion As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtImagen As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtPrecioenDolares As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Friend WithEvents btnLimpiar As Button
    'Friend WithEvents btnEliminar As Button
    'Friend WithEvents btnGuardar As Button

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    '<System.Diagnostics.DebuggerStepThrough()>
    'Private Sub InitializeComponent()
    '    Me.components = New System.ComponentModel.Container()
    '    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    '    Me.txtDescArticulo = New System.Windows.Forms.TextBox()
    '    Me.sstArticulo = New System.Windows.Forms.TabControl()
    '    Me._sstArticulo_TabPage0 = New System.Windows.Forms.TabPage()
    '    Me._fraContenedor_0 = New System.Windows.Forms.Panel()
    '    Me.fraDiamanteSuelto = New System.Windows.Forms.GroupBox()
    '    Me.txtMDSCertificado = New System.Windows.Forms.TextBox()
    '    Me.txtMDSPureza = New System.Windows.Forms.TextBox()
    '    Me.txtMDSColor = New System.Windows.Forms.TextBox()
    '    Me.txtMDSPeso = New System.Windows.Forms.TextBox()
    '    Me.lblEstatus = New System.Windows.Forms.Label()
    '    Me.lblMDSCertificado = New System.Windows.Forms.Label()
    '    Me.lblMDSPureza = New System.Windows.Forms.Label()
    '    Me.lblMDSColor = New System.Windows.Forms.Label()
    '    Me.lblMDSPeso = New System.Windows.Forms.Label()
    '    Me._txtAdicional_0 = New System.Windows.Forms.TextBox()
    '    Me._fraMoneda_5 = New System.Windows.Forms.GroupBox()
    '    Me._optMoneda_11 = New System.Windows.Forms.RadioButton()
    '    Me._optMoneda_10 = New System.Windows.Forms.RadioButton()
    '    Me._fraMoneda_0 = New System.Windows.Forms.GroupBox()
    '    Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
    '    Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
    '    Me._fraImagen_0 = New System.Windows.Forms.GroupBox()
    '    Me.Image1 = New System.Windows.Forms.PictureBox()
    '    Me._Frame2_0 = New System.Windows.Forms.GroupBox()
    '    Me._Frame4_0 = New System.Windows.Forms.GroupBox()
    '    Me._cmdBuscarImagen_0 = New System.Windows.Forms.Button()
    '    Me._txtImagen_0 = New System.Windows.Forms.TextBox()
    '    Me._txtCodigodelProveedor_0 = New System.Windows.Forms.TextBox()
    '    Me._dbcProveedor_0 = New System.Windows.Forms.ComboBox()
    '    Me._cboUnidad_0 = New System.Windows.Forms.ComboBox()
    '    Me._cboAlmacen_0 = New System.Windows.Forms.ComboBox()
    '    Me._lblArticulo_36 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_35 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_11 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_10 = New System.Windows.Forms.Label()
    '    Me._txtDescripcion_0 = New System.Windows.Forms.TextBox()
    '    Me._Frame1_0 = New System.Windows.Forms.GroupBox()
    '    Me._txtCostoReal_0 = New System.Windows.Forms.TextBox()
    '    Me._txtPrecioenDolares_0 = New System.Windows.Forms.TextBox()
    '    Me._txtCostoIndirecto_0 = New System.Windows.Forms.TextBox()
    '    Me._txtCostoAdicional_0 = New System.Windows.Forms.TextBox()
    '    Me._txtCostoFactura_0 = New System.Windows.Forms.TextBox()
    '    Me._lblMargen_0 = New System.Windows.Forms.Label()
    '    Me.Label1 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_34 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_5 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_8 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_7 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_6 = New System.Windows.Forms.Label()
    '    Me._dbcFamilia_0 = New System.Windows.Forms.ComboBox()
    '    Me._dbcLinea_0 = New System.Windows.Forms.ComboBox()
    '    Me.dbcSubLinea = New System.Windows.Forms.ComboBox()
    '    Me.dbcKilates = New System.Windows.Forms.ComboBox()
    '    Me._dbcMaterial_0 = New System.Windows.Forms.ComboBox()
    '    Me._lblArticulo_33 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_9 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_29 = New System.Windows.Forms.Label()
    '    Me._lblDescripcion_0 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_26 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_37 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_4 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_3 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_2 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_1 = New System.Windows.Forms.Label()
    '    Me._sstArticulo_TabPage1 = New System.Windows.Forms.TabPage()
    '    Me._fraContenedor_1 = New System.Windows.Forms.Panel()
    '    Me._txtAdicional_1 = New System.Windows.Forms.TextBox()
    '    Me._fraMoneda_3 = New System.Windows.Forms.GroupBox()
    '    Me._optMoneda_7 = New System.Windows.Forms.RadioButton()
    '    Me._optMoneda_6 = New System.Windows.Forms.RadioButton()
    '    Me._txtDescripcion_1 = New System.Windows.Forms.TextBox()
    '    Me._fraArticulo_1 = New System.Windows.Forms.GroupBox()
    '    Me._optGenero_0 = New System.Windows.Forms.RadioButton()
    '    Me._optGenero_1 = New System.Windows.Forms.RadioButton()
    '    Me._optGenero_2 = New System.Windows.Forms.RadioButton()
    '    Me._fraArticulo_2 = New System.Windows.Forms.GroupBox()
    '    Me._optMovimiento_0 = New System.Windows.Forms.RadioButton()
    '    Me._optMovimiento_1 = New System.Windows.Forms.RadioButton()
    '    Me._optMovimiento_2 = New System.Windows.Forms.RadioButton()
    '    Me._fraImagen_1 = New System.Windows.Forms.GroupBox()
    '    Me.Image2 = New System.Windows.Forms.PictureBox()
    '    Me._fraMoneda_1 = New System.Windows.Forms.GroupBox()
    '    Me._optMoneda_3 = New System.Windows.Forms.RadioButton()
    '    Me._optMoneda_2 = New System.Windows.Forms.RadioButton()
    '    Me._Frame1_2 = New System.Windows.Forms.GroupBox()
    '    Me._txtCostoReal_1 = New System.Windows.Forms.TextBox()
    '    Me._txtPrecioenDolares_1 = New System.Windows.Forms.TextBox()
    '    Me._txtCostoIndirecto_1 = New System.Windows.Forms.TextBox()
    '    Me._txtCostoAdicional_1 = New System.Windows.Forms.TextBox()
    '    Me._txtCostoFactura_1 = New System.Windows.Forms.TextBox()
    '    Me._lblMargen_1 = New System.Windows.Forms.Label()
    '    Me.Label2 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_40 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_41 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_42 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_43 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_44 = New System.Windows.Forms.Label()
    '    Me._Frame2_2 = New System.Windows.Forms.GroupBox()
    '    Me._Frame4_1 = New System.Windows.Forms.GroupBox()
    '    Me._txtImagen_1 = New System.Windows.Forms.TextBox()
    '    Me._cmdBuscarImagen_1 = New System.Windows.Forms.Button()
    '    Me._txtCodigodelProveedor_1 = New System.Windows.Forms.TextBox()
    '    Me._dbcProveedor_1 = New System.Windows.Forms.ComboBox()
    '    Me._cboUnidad_1 = New System.Windows.Forms.ComboBox()
    '    Me._cboAlmacen_1 = New System.Windows.Forms.ComboBox()
    '    Me._lblArticulo_18 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_19 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_20 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_21 = New System.Windows.Forms.Label()
    '    Me.chkCrono = New System.Windows.Forms.CheckBox()
    '    Me.dbcMarca = New System.Windows.Forms.ComboBox()
    '    Me.dbcModelo = New System.Windows.Forms.ComboBox()
    '    Me._dbcMaterial_1 = New System.Windows.Forms.ComboBox()
    '    Me._lblArticulo_45 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_27 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_13 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_12 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_14 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_15 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_16 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_17 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_38 = New System.Windows.Forms.Label()
    '    Me._lblDescripcion_1 = New System.Windows.Forms.Label()
    '    Me._sstArticulo_TabPage2 = New System.Windows.Forms.TabPage()
    '    Me._fraContenedor_2 = New System.Windows.Forms.Panel()
    '    Me._txtAdicional_2 = New System.Windows.Forms.TextBox()
    '    Me._fraMoneda_4 = New System.Windows.Forms.GroupBox()
    '    Me._optMoneda_9 = New System.Windows.Forms.RadioButton()
    '    Me._optMoneda_8 = New System.Windows.Forms.RadioButton()
    '    Me._fraMoneda_2 = New System.Windows.Forms.GroupBox()
    '    Me._optMoneda_4 = New System.Windows.Forms.RadioButton()
    '    Me._optMoneda_5 = New System.Windows.Forms.RadioButton()
    '    Me._Frame2_3 = New System.Windows.Forms.GroupBox()
    '    Me._Frame4_2 = New System.Windows.Forms.GroupBox()
    '    Me._txtImagen_2 = New System.Windows.Forms.TextBox()
    '    Me._cmdBuscarImagen_2 = New System.Windows.Forms.Button()
    '    Me._txtCodigodelProveedor_2 = New System.Windows.Forms.TextBox()
    '    Me._dbcProveedor_2 = New System.Windows.Forms.ComboBox()
    '    Me._cboUnidad_2 = New System.Windows.Forms.ComboBox()
    '    Me._cboAlmacen_2 = New System.Windows.Forms.ComboBox()
    '    Me._lblArticulo_50 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_51 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_52 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_53 = New System.Windows.Forms.Label()
    '    Me._Frame1_3 = New System.Windows.Forms.GroupBox()
    '    Me._txtCostoFactura_2 = New System.Windows.Forms.TextBox()
    '    Me._txtCostoAdicional_2 = New System.Windows.Forms.TextBox()
    '    Me._txtCostoIndirecto_2 = New System.Windows.Forms.TextBox()
    '    Me._txtPrecioenDolares_2 = New System.Windows.Forms.TextBox()
    '    Me._txtCostoReal_2 = New System.Windows.Forms.TextBox()
    '    Me._lblMargen_2 = New System.Windows.Forms.Label()
    '    Me.Label3 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_22 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_23 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_46 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_47 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_48 = New System.Windows.Forms.Label()
    '    Me._txtDescripcion_2 = New System.Windows.Forms.TextBox()
    '    Me._fraImagen_2 = New System.Windows.Forms.GroupBox()
    '    Me.Image3 = New System.Windows.Forms.PictureBox()
    '    Me._dbcFamilia_1 = New System.Windows.Forms.ComboBox()
    '    Me._dbcLinea_1 = New System.Windows.Forms.ComboBox()
    '    Me._dbcMaterial_2 = New System.Windows.Forms.ComboBox()
    '    Me._lblArticulo_54 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_49 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_28 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_24 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_25 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_30 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_39 = New System.Windows.Forms.Label()
    '    Me._lblDescripcion_2 = New System.Windows.Forms.Label()
    '    Me.chkCodigoAnterior = New System.Windows.Forms.CheckBox()
    '    Me.Frame3 = New System.Windows.Forms.GroupBox()
    '    Me.txtCodArtAnterior = New System.Windows.Forms.TextBox()
    '    Me.dbcOrigen = New System.Windows.Forms.ComboBox()
    '    Me._lblArticulo_32 = New System.Windows.Forms.Label()
    '    Me._lblArticulo_31 = New System.Windows.Forms.Label()
    '    Me.txtCodArticulo = New System.Windows.Forms.TextBox()
    '    Me._lblArticulo_0 = New System.Windows.Forms.Label()
    '    Me.Frame1 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
    '    Me.Frame2 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
    '    Me.Frame4 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
    '    Me.cboAlmacen = New System.Windows.Forms.ComboBox()
    '    Me.cboUnidad = New System.Windows.Forms.ComboBox()
    '    Me.cmdBuscarImagen = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
    '    Me.dbcFamilia = New System.Windows.Forms.ComboBox()
    '    Me.dbcLinea = New System.Windows.Forms.ComboBox()
    '    Me.dbcMaterial = New System.Windows.Forms.ComboBox()
    '    Me.dbcProveedor = New System.Windows.Forms.ComboBox()
    '    Me.fraArticulo = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
    '    Me.fraContenedor = New Microsoft.VisualBasic.Compatibility.VB6.PanelArray(Me.components)
    '    Me.fraImagen = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
    '    Me.fraMoneda = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
    '    Me.lblArticulo = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
    '    Me.lblDescripcion = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
    '    Me.lblMargen = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
    '    Me.optGenero = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
    '    Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
    '    Me.optMovimiento = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
    '    Me.txtAdicional = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
    '    Me.txtCodigodelProveedor = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
    '    Me.txtCostoAdicional = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
    '    Me.txtCostoFactura = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
    '    Me.txtCostoIndirecto = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
    '    Me.txtCostoReal = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
    '    Me.txtDescripcion = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
    '    Me.txtImagen = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
    '    Me.txtPrecioenDolares = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
    '    Me.btnLimpiar = New System.Windows.Forms.Button()
    '    Me.btnEliminar = New System.Windows.Forms.Button()
    '    Me.btnGuardar = New System.Windows.Forms.Button()
    '    Me.sstArticulo.SuspendLayout()
    '    Me._sstArticulo_TabPage0.SuspendLayout()
    '    Me._fraContenedor_0.SuspendLayout()
    '    Me.fraDiamanteSuelto.SuspendLayout()
    '    Me._fraMoneda_5.SuspendLayout()
    '    Me._fraMoneda_0.SuspendLayout()
    '    Me._fraImagen_0.SuspendLayout()
    '    CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
    '    Me._Frame2_0.SuspendLayout()
    '    Me._Frame4_0.SuspendLayout()
    '    Me._Frame1_0.SuspendLayout()
    '    Me._sstArticulo_TabPage1.SuspendLayout()
    '    Me._fraContenedor_1.SuspendLayout()
    '    Me._fraMoneda_3.SuspendLayout()
    '    Me._fraArticulo_1.SuspendLayout()
    '    Me._fraArticulo_2.SuspendLayout()
    '    Me._fraImagen_1.SuspendLayout()
    '    CType(Me.Image2, System.ComponentModel.ISupportInitialize).BeginInit()
    '    Me._fraMoneda_1.SuspendLayout()
    '    Me._Frame1_2.SuspendLayout()
    '    Me._Frame2_2.SuspendLayout()
    '    Me._Frame4_1.SuspendLayout()
    '    Me._sstArticulo_TabPage2.SuspendLayout()
    '    Me._fraContenedor_2.SuspendLayout()
    '    Me._fraMoneda_4.SuspendLayout()
    '    Me._fraMoneda_2.SuspendLayout()
    '    Me._Frame2_3.SuspendLayout()
    '    Me._Frame4_2.SuspendLayout()
    '    Me._Frame1_3.SuspendLayout()
    '    Me._fraImagen_2.SuspendLayout()
    '    CType(Me.Image3, System.ComponentModel.ISupportInitialize).BeginInit()
    '    Me.Frame3.SuspendLayout()
    '    CType(Me.Frame1, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.Frame2, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.Frame4, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.cmdBuscarImagen, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.fraArticulo, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.fraContenedor, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.fraImagen, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.fraMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.lblArticulo, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.lblDescripcion, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.lblMargen, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.optGenero, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.optMovimiento, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.txtAdicional, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.txtCodigodelProveedor, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.txtCostoAdicional, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.txtCostoFactura, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.txtCostoIndirecto, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.txtCostoReal, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.txtDescripcion, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.txtImagen, System.ComponentModel.ISupportInitialize).BeginInit()
    '    CType(Me.txtPrecioenDolares, System.ComponentModel.ISupportInitialize).BeginInit()
    '    Me.SuspendLayout()
    '    '
    '    'txtDescArticulo
    '    '
    '    Me.txtDescArticulo.AcceptsReturn = True
    '    Me.txtDescArticulo.BackColor = System.Drawing.SystemColors.Window
    '    Me.txtDescArticulo.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me.txtDescArticulo.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtDescArticulo.Location = New System.Drawing.Point(168, 8)
    '    Me.txtDescArticulo.MaxLength = 0
    '    Me.txtDescArticulo.Name = "txtDescArticulo"
    '    Me.txtDescArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.txtDescArticulo.Size = New System.Drawing.Size(328, 20)
    '    Me.txtDescArticulo.TabIndex = 2
    '    Me.ToolTip1.SetToolTip(Me.txtDescArticulo, "Descripción del artículo")
    '    '
    '    'sstArticulo
    '    '
    '    Me.sstArticulo.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
    '    Me.sstArticulo.Controls.Add(Me._sstArticulo_TabPage0)
    '    Me.sstArticulo.Controls.Add(Me._sstArticulo_TabPage1)
    '    Me.sstArticulo.Controls.Add(Me._sstArticulo_TabPage2)
    '    Me.sstArticulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
    '    Me.sstArticulo.ItemSize = New System.Drawing.Size(42, 18)
    '    Me.sstArticulo.Location = New System.Drawing.Point(10, 64)
    '    Me.sstArticulo.Name = "sstArticulo"
    '    Me.sstArticulo.SelectedIndex = 0
    '    Me.sstArticulo.Size = New System.Drawing.Size(708, 567)
    '    Me.sstArticulo.TabIndex = 9
    '    Me.ToolTip1.SetToolTip(Me.sstArticulo, "Grupo de Artículo al que pertenece")
    '    '
    '    '_sstArticulo_TabPage0
    '    '
    '    Me._sstArticulo_TabPage0.Controls.Add(Me._fraContenedor_0)
    '    Me._sstArticulo_TabPage0.Location = New System.Drawing.Point(4, 22)
    '    Me._sstArticulo_TabPage0.Name = "_sstArticulo_TabPage0"
    '    Me._sstArticulo_TabPage0.Size = New System.Drawing.Size(700, 541)
    '    Me._sstArticulo_TabPage0.TabIndex = 0
    '    Me._sstArticulo_TabPage0.Text = "Joyería"
    '    '
    '    '_fraContenedor_0
    '    '
    '    Me._fraContenedor_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraContenedor_0.Controls.Add(Me.fraDiamanteSuelto)
    '    Me._fraContenedor_0.Controls.Add(Me._txtAdicional_0)
    '    Me._fraContenedor_0.Controls.Add(Me._fraMoneda_5)
    '    Me._fraContenedor_0.Controls.Add(Me._fraMoneda_0)
    '    Me._fraContenedor_0.Controls.Add(Me._fraImagen_0)
    '    Me._fraContenedor_0.Controls.Add(Me._Frame2_0)
    '    Me._fraContenedor_0.Controls.Add(Me._txtDescripcion_0)
    '    Me._fraContenedor_0.Controls.Add(Me._Frame1_0)
    '    Me._fraContenedor_0.Controls.Add(Me._dbcFamilia_0)
    '    Me._fraContenedor_0.Controls.Add(Me._dbcLinea_0)
    '    Me._fraContenedor_0.Controls.Add(Me.dbcSubLinea)
    '    Me._fraContenedor_0.Controls.Add(Me.dbcKilates)
    '    Me._fraContenedor_0.Controls.Add(Me._dbcMaterial_0)
    '    Me._fraContenedor_0.Controls.Add(Me._lblArticulo_33)
    '    Me._fraContenedor_0.Controls.Add(Me._lblArticulo_9)
    '    Me._fraContenedor_0.Controls.Add(Me._lblArticulo_29)
    '    Me._fraContenedor_0.Controls.Add(Me._lblDescripcion_0)
    '    Me._fraContenedor_0.Controls.Add(Me._lblArticulo_26)
    '    Me._fraContenedor_0.Controls.Add(Me._lblArticulo_37)
    '    Me._fraContenedor_0.Controls.Add(Me._lblArticulo_4)
    '    Me._fraContenedor_0.Controls.Add(Me._lblArticulo_3)
    '    Me._fraContenedor_0.Controls.Add(Me._lblArticulo_2)
    '    Me._fraContenedor_0.Controls.Add(Me._lblArticulo_1)
    '    Me._fraContenedor_0.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._fraContenedor_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraContenedor.SetIndex(Me._fraContenedor_0, CType(0, Short))
    '    Me._fraContenedor_0.Location = New System.Drawing.Point(8, 24)
    '    Me._fraContenedor_0.Name = "_fraContenedor_0"
    '    Me._fraContenedor_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraContenedor_0.Size = New System.Drawing.Size(680, 504)
    '    Me._fraContenedor_0.TabIndex = 10
    '    '
    '    'fraDiamanteSuelto
    '    '
    '    Me.fraDiamanteSuelto.BackColor = System.Drawing.SystemColors.Control
    '    Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSCertificado)
    '    Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSPureza)
    '    Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSColor)
    '    Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSPeso)
    '    Me.fraDiamanteSuelto.Controls.Add(Me.lblEstatus)
    '    Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSCertificado)
    '    Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSPureza)
    '    Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSColor)
    '    Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSPeso)
    '    Me.fraDiamanteSuelto.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraDiamanteSuelto.Location = New System.Drawing.Point(236, 104)
    '    Me.fraDiamanteSuelto.Name = "fraDiamanteSuelto"
    '    Me.fraDiamanteSuelto.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.fraDiamanteSuelto.Size = New System.Drawing.Size(268, 84)
    '    Me.fraDiamanteSuelto.TabIndex = 23
    '    Me.fraDiamanteSuelto.TabStop = False
    '    Me.fraDiamanteSuelto.Text = " "
    '    '
    '    'txtMDSCertificado
    '    '
    '    Me.txtMDSCertificado.AcceptsReturn = True
    '    Me.txtMDSCertificado.BackColor = System.Drawing.SystemColors.Window
    '    Me.txtMDSCertificado.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me.txtMDSCertificado.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtMDSCertificado.Location = New System.Drawing.Point(125, 40)
    '    Me.txtMDSCertificado.MaxLength = 20
    '    Me.txtMDSCertificado.Name = "txtMDSCertificado"
    '    Me.txtMDSCertificado.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.txtMDSCertificado.Size = New System.Drawing.Size(136, 20)
    '    Me.txtMDSCertificado.TabIndex = 31
    '    '
    '    'txtMDSPureza
    '    '
    '    Me.txtMDSPureza.AcceptsReturn = True
    '    Me.txtMDSPureza.BackColor = System.Drawing.SystemColors.Window
    '    Me.txtMDSPureza.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me.txtMDSPureza.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtMDSPureza.Location = New System.Drawing.Point(71, 66)
    '    Me.txtMDSPureza.MaxLength = 4
    '    Me.txtMDSPureza.Name = "txtMDSPureza"
    '    Me.txtMDSPureza.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.txtMDSPureza.Size = New System.Drawing.Size(50, 20)
    '    Me.txtMDSPureza.TabIndex = 29
    '    '
    '    'txtMDSColor
    '    '
    '    Me.txtMDSColor.AcceptsReturn = True
    '    Me.txtMDSColor.BackColor = System.Drawing.SystemColors.Window
    '    Me.txtMDSColor.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me.txtMDSColor.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtMDSColor.Location = New System.Drawing.Point(71, 40)
    '    Me.txtMDSColor.MaxLength = 1
    '    Me.txtMDSColor.Name = "txtMDSColor"
    '    Me.txtMDSColor.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.txtMDSColor.Size = New System.Drawing.Size(50, 20)
    '    Me.txtMDSColor.TabIndex = 27
    '    '
    '    'txtMDSPeso
    '    '
    '    Me.txtMDSPeso.AcceptsReturn = True
    '    Me.txtMDSPeso.BackColor = System.Drawing.SystemColors.Window
    '    Me.txtMDSPeso.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me.txtMDSPeso.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtMDSPeso.Location = New System.Drawing.Point(71, 14)
    '    Me.txtMDSPeso.MaxLength = 6
    '    Me.txtMDSPeso.Name = "txtMDSPeso"
    '    Me.txtMDSPeso.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.txtMDSPeso.Size = New System.Drawing.Size(50, 20)
    '    Me.txtMDSPeso.TabIndex = 25
    '    Me.txtMDSPeso.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    '
    '    'lblEstatus
    '    '
    '    Me.lblEstatus.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
    '    Me.lblEstatus.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    '    Me.lblEstatus.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.lblEstatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
    '    Me.lblEstatus.Location = New System.Drawing.Point(125, 68)
    '    Me.lblEstatus.Name = "lblEstatus"
    '    Me.lblEstatus.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.lblEstatus.Size = New System.Drawing.Size(136, 19)
    '    Me.lblEstatus.TabIndex = 32
    '    Me.lblEstatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
    '    '
    '    'lblMDSCertificado
    '    '
    '    Me.lblMDSCertificado.BackColor = System.Drawing.SystemColors.Control
    '    Me.lblMDSCertificado.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.lblMDSCertificado.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblMDSCertificado.Location = New System.Drawing.Point(127, 17)
    '    Me.lblMDSCertificado.Name = "lblMDSCertificado"
    '    Me.lblMDSCertificado.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.lblMDSCertificado.Size = New System.Drawing.Size(63, 18)
    '    Me.lblMDSCertificado.TabIndex = 30
    '    Me.lblMDSCertificado.Text = "Certificado"
    '    '
    '    'lblMDSPureza
    '    '
    '    Me.lblMDSPureza.BackColor = System.Drawing.SystemColors.Control
    '    Me.lblMDSPureza.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.lblMDSPureza.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblMDSPureza.Location = New System.Drawing.Point(9, 71)
    '    Me.lblMDSPureza.Name = "lblMDSPureza"
    '    Me.lblMDSPureza.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.lblMDSPureza.Size = New System.Drawing.Size(63, 18)
    '    Me.lblMDSPureza.TabIndex = 28
    '    Me.lblMDSPureza.Text = "Pureza - Q"
    '    '
    '    'lblMDSColor
    '    '
    '    Me.lblMDSColor.BackColor = System.Drawing.SystemColors.Control
    '    Me.lblMDSColor.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.lblMDSColor.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblMDSColor.Location = New System.Drawing.Point(9, 44)
    '    Me.lblMDSColor.Name = "lblMDSColor"
    '    Me.lblMDSColor.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.lblMDSColor.Size = New System.Drawing.Size(63, 18)
    '    Me.lblMDSColor.TabIndex = 26
    '    Me.lblMDSColor.Text = "Color"
    '    '
    '    'lblMDSPeso
    '    '
    '    Me.lblMDSPeso.BackColor = System.Drawing.SystemColors.Control
    '    Me.lblMDSPeso.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.lblMDSPeso.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblMDSPeso.Location = New System.Drawing.Point(9, 18)
    '    Me.lblMDSPeso.Name = "lblMDSPeso"
    '    Me.lblMDSPeso.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.lblMDSPeso.Size = New System.Drawing.Size(63, 18)
    '    Me.lblMDSPeso.TabIndex = 24
    '    Me.lblMDSPeso.Text = "Peso - CT"
    '    '
    '    '_txtAdicional_0
    '    '
    '    Me._txtAdicional_0.AcceptsReturn = True
    '    Me._txtAdicional_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
    '    Me._txtAdicional_0.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtAdicional_0.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtAdicional.SetIndex(Me._txtAdicional_0, CType(0, Short))
    '    Me._txtAdicional_0.Location = New System.Drawing.Point(89, 168)
    '    Me._txtAdicional_0.MaxLength = 15
    '    Me._txtAdicional_0.Name = "_txtAdicional_0"
    '    Me._txtAdicional_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtAdicional_0.Size = New System.Drawing.Size(120, 20)
    '    Me._txtAdicional_0.TabIndex = 22
    '    '
    '    '_fraMoneda_5
    '    '
    '    Me._fraMoneda_5.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraMoneda_5.Controls.Add(Me._optMoneda_11)
    '    Me._fraMoneda_5.Controls.Add(Me._optMoneda_10)
    '    Me._fraMoneda_5.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraMoneda.SetIndex(Me._fraMoneda_5, CType(5, Short))
    '    Me._fraMoneda_5.Location = New System.Drawing.Point(89, 256)
    '    Me._fraMoneda_5.Name = "_fraMoneda_5"
    '    Me._fraMoneda_5.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraMoneda_5.Size = New System.Drawing.Size(218, 33)
    '    Me._fraMoneda_5.TabIndex = 37
    '    Me._fraMoneda_5.TabStop = False
    '    '
    '    '_optMoneda_11
    '    '
    '    Me._optMoneda_11.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_11.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_11.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_11, CType(11, Short))
    '    Me._optMoneda_11.Location = New System.Drawing.Point(129, 11)
    '    Me._optMoneda_11.Name = "_optMoneda_11"
    '    Me._optMoneda_11.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_11.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_11.TabIndex = 39
    '    Me._optMoneda_11.TabStop = True
    '    Me._optMoneda_11.Tag = "0"
    '    Me._optMoneda_11.Text = "Pesos"
    '    Me._optMoneda_11.UseVisualStyleBackColor = False
    '    '
    '    '_optMoneda_10
    '    '
    '    Me._optMoneda_10.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_10.Checked = True
    '    Me._optMoneda_10.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_10.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_10, CType(10, Short))
    '    Me._optMoneda_10.Location = New System.Drawing.Point(36, 11)
    '    Me._optMoneda_10.Name = "_optMoneda_10"
    '    Me._optMoneda_10.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_10.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_10.TabIndex = 38
    '    Me._optMoneda_10.TabStop = True
    '    Me._optMoneda_10.Tag = "1"
    '    Me._optMoneda_10.Text = "Dólares"
    '    Me._optMoneda_10.UseVisualStyleBackColor = False
    '    '
    '    '_fraMoneda_0
    '    '
    '    Me._fraMoneda_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraMoneda_0.Controls.Add(Me._optMoneda_1)
    '    Me._fraMoneda_0.Controls.Add(Me._optMoneda_0)
    '    Me._fraMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraMoneda.SetIndex(Me._fraMoneda_0, CType(0, Short))
    '    Me._fraMoneda_0.Location = New System.Drawing.Point(416, 256)
    '    Me._fraMoneda_0.Name = "_fraMoneda_0"
    '    Me._fraMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraMoneda_0.Size = New System.Drawing.Size(209, 33)
    '    Me._fraMoneda_0.TabIndex = 41
    '    Me._fraMoneda_0.TabStop = False
    '    '
    '    '_optMoneda_1
    '    '
    '    Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_1, CType(1, Short))
    '    Me._optMoneda_1.Location = New System.Drawing.Point(134, 11)
    '    Me._optMoneda_1.Name = "_optMoneda_1"
    '    Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_1.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_1.TabIndex = 43
    '    Me._optMoneda_1.TabStop = True
    '    Me._optMoneda_1.Text = "Pesos"
    '    Me._optMoneda_1.UseVisualStyleBackColor = False
    '    '
    '    '_optMoneda_0
    '    '
    '    Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_0.Checked = True
    '    Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_0, CType(0, Short))
    '    Me._optMoneda_0.Location = New System.Drawing.Point(34, 11)
    '    Me._optMoneda_0.Name = "_optMoneda_0"
    '    Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_0.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_0.TabIndex = 42
    '    Me._optMoneda_0.TabStop = True
    '    Me._optMoneda_0.Text = "Dólares"
    '    Me._optMoneda_0.UseVisualStyleBackColor = False
    '    '
    '    '_fraImagen_0
    '    '
    '    Me._fraImagen_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraImagen_0.Controls.Add(Me.Image1)
    '    Me._fraImagen_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
    '    Me.fraImagen.SetIndex(Me._fraImagen_0, CType(0, Short))
    '    Me._fraImagen_0.Location = New System.Drawing.Point(510, 8)
    '    Me._fraImagen_0.Name = "_fraImagen_0"
    '    Me._fraImagen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraImagen_0.Size = New System.Drawing.Size(178, 186)
    '    Me._fraImagen_0.TabIndex = 69
    '    Me._fraImagen_0.TabStop = False
    '    Me._fraImagen_0.Text = "Imagen del Artículo"
    '    '
    '    'Image1
    '    '
    '    Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    '    Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.Image1.Location = New System.Drawing.Point(7, 21)
    '    Me.Image1.Name = "Image1"
    '    Me.Image1.Size = New System.Drawing.Size(163, 157)
    '    Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
    '    Me.Image1.TabIndex = 0
    '    Me.Image1.TabStop = False
    '    '
    '    '_Frame2_0
    '    '
    '    Me._Frame2_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._Frame2_0.Controls.Add(Me._Frame4_0)
    '    Me._Frame2_0.Controls.Add(Me._txtCodigodelProveedor_0)
    '    Me._Frame2_0.Controls.Add(Me._dbcProveedor_0)
    '    Me._Frame2_0.Controls.Add(Me._cboUnidad_0)
    '    Me._Frame2_0.Controls.Add(Me._cboAlmacen_0)
    '    Me._Frame2_0.Controls.Add(Me._lblArticulo_36)
    '    Me._Frame2_0.Controls.Add(Me._lblArticulo_35)
    '    Me._Frame2_0.Controls.Add(Me._lblArticulo_11)
    '    Me._Frame2_0.Controls.Add(Me._lblArticulo_10)
    '    Me._Frame2_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.Frame2.SetIndex(Me._Frame2_0, CType(0, Short))
    '    Me._Frame2_0.Location = New System.Drawing.Point(312, 294)
    '    Me._Frame2_0.Name = "_Frame2_0"
    '    Me._Frame2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._Frame2_0.Size = New System.Drawing.Size(313, 185)
    '    Me._Frame2_0.TabIndex = 57
    '    Me._Frame2_0.TabStop = False
    '    '
    '    '_Frame4_0
    '    '
    '    Me._Frame4_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._Frame4_0.Controls.Add(Me._cmdBuscarImagen_0)
    '    Me._Frame4_0.Controls.Add(Me._txtImagen_0)
    '    Me._Frame4_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
    '    Me.Frame4.SetIndex(Me._Frame4_0, CType(0, Short))
    '    Me._Frame4_0.Location = New System.Drawing.Point(12, 132)
    '    Me._Frame4_0.Name = "_Frame4_0"
    '    Me._Frame4_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._Frame4_0.Size = New System.Drawing.Size(290, 44)
    '    Me._Frame4_0.TabIndex = 66
    '    Me._Frame4_0.TabStop = False
    '    Me._Frame4_0.Text = "Imagen"
    '    '
    '    '_cmdBuscarImagen_0
    '    '
    '    Me._cmdBuscarImagen_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._cmdBuscarImagen_0.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._cmdBuscarImagen_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.cmdBuscarImagen.SetIndex(Me._cmdBuscarImagen_0, CType(0, Short))
    '    Me._cmdBuscarImagen_0.Location = New System.Drawing.Point(260, 15)
    '    Me._cmdBuscarImagen_0.Name = "_cmdBuscarImagen_0"
    '    Me._cmdBuscarImagen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._cmdBuscarImagen_0.Size = New System.Drawing.Size(22, 21)
    '    Me._cmdBuscarImagen_0.TabIndex = 68
    '    Me._cmdBuscarImagen_0.Text = "..."
    '    Me._cmdBuscarImagen_0.UseVisualStyleBackColor = False
    '    '
    '    '_txtImagen_0
    '    '
    '    Me._txtImagen_0.AcceptsReturn = True
    '    Me._txtImagen_0.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtImagen_0.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtImagen_0.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtImagen.SetIndex(Me._txtImagen_0, CType(0, Short))
    '    Me._txtImagen_0.Location = New System.Drawing.Point(9, 15)
    '    Me._txtImagen_0.MaxLength = 0
    '    Me._txtImagen_0.Name = "_txtImagen_0"
    '    Me._txtImagen_0.ReadOnly = True
    '    Me._txtImagen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtImagen_0.Size = New System.Drawing.Size(245, 20)
    '    Me._txtImagen_0.TabIndex = 67
    '    '
    '    '_txtCodigodelProveedor_0
    '    '
    '    Me._txtCodigodelProveedor_0.AcceptsReturn = True
    '    Me._txtCodigodelProveedor_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
    '    Me._txtCodigodelProveedor_0.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCodigodelProveedor_0.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtCodigodelProveedor.SetIndex(Me._txtCodigodelProveedor_0, CType(0, Short))
    '    Me._txtCodigodelProveedor_0.Location = New System.Drawing.Point(171, 102)
    '    Me._txtCodigodelProveedor_0.MaxLength = 20
    '    Me._txtCodigodelProveedor_0.Name = "_txtCodigodelProveedor_0"
    '    Me._txtCodigodelProveedor_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCodigodelProveedor_0.Size = New System.Drawing.Size(129, 20)
    '    Me._txtCodigodelProveedor_0.TabIndex = 65
    '    Me.ToolTip1.SetToolTip(Me._txtCodigodelProveedor_0, "Código que usa el Proveedor para el Artículo")
    '    '
    '    '_dbcProveedor_0
    '    '
    '    Me._dbcProveedor_0.Location = New System.Drawing.Point(100, 74)
    '    Me._dbcProveedor_0.Name = "_dbcProveedor_0"
    '    Me._dbcProveedor_0.Size = New System.Drawing.Size(201, 21)
    '    Me._dbcProveedor_0.TabIndex = 63
    '    '
    '    '_cboUnidad_0
    '    '
    '    Me._cboUnidad_0.Location = New System.Drawing.Point(100, 17)
    '    Me._cboUnidad_0.Name = "_cboUnidad_0"
    '    Me._cboUnidad_0.Size = New System.Drawing.Size(78, 21)
    '    Me._cboUnidad_0.TabIndex = 59
    '    '
    '    '_cboAlmacen_0
    '    '
    '    Me._cboAlmacen_0.Location = New System.Drawing.Point(100, 46)
    '    Me._cboAlmacen_0.Name = "_cboAlmacen_0"
    '    Me._cboAlmacen_0.Size = New System.Drawing.Size(201, 21)
    '    Me._cboAlmacen_0.TabIndex = 61
    '    '
    '    '_lblArticulo_36
    '    '
    '    Me._lblArticulo_36.AutoSize = True
    '    Me._lblArticulo_36.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_36.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_36.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_36, CType(36, Short))
    '    Me._lblArticulo_36.Location = New System.Drawing.Point(12, 50)
    '    Me._lblArticulo_36.Name = "_lblArticulo_36"
    '    Me._lblArticulo_36.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_36.Size = New System.Drawing.Size(84, 13)
    '    Me._lblArticulo_36.TabIndex = 60
    '    Me._lblArticulo_36.Text = "Almacén/Origen"
    '    '
    '    '_lblArticulo_35
    '    '
    '    Me._lblArticulo_35.AutoSize = True
    '    Me._lblArticulo_35.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_35.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_35.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_35, CType(35, Short))
    '    Me._lblArticulo_35.Location = New System.Drawing.Point(12, 21)
    '    Me._lblArticulo_35.Name = "_lblArticulo_35"
    '    Me._lblArticulo_35.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_35.Size = New System.Drawing.Size(41, 13)
    '    Me._lblArticulo_35.TabIndex = 58
    '    Me._lblArticulo_35.Text = "Unidad"
    '    '
    '    '_lblArticulo_11
    '    '
    '    Me._lblArticulo_11.AutoSize = True
    '    Me._lblArticulo_11.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_11.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_11.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_11, CType(11, Short))
    '    Me._lblArticulo_11.Location = New System.Drawing.Point(19, 106)
    '    Me._lblArticulo_11.Name = "_lblArticulo_11"
    '    Me._lblArticulo_11.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_11.Size = New System.Drawing.Size(156, 13)
    '    Me._lblArticulo_11.TabIndex = 64
    '    Me._lblArticulo_11.Text = "Código artículo del proveedor : "
    '    '
    '    '_lblArticulo_10
    '    '
    '    Me._lblArticulo_10.AutoSize = True
    '    Me._lblArticulo_10.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_10.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_10.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_10, CType(10, Short))
    '    Me._lblArticulo_10.Location = New System.Drawing.Point(12, 79)
    '    Me._lblArticulo_10.Name = "_lblArticulo_10"
    '    Me._lblArticulo_10.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_10.Size = New System.Drawing.Size(56, 13)
    '    Me._lblArticulo_10.TabIndex = 62
    '    Me._lblArticulo_10.Text = "Proveedor"
    '    '
    '    '_txtDescripcion_0
    '    '
    '    Me._txtDescripcion_0.AcceptsReturn = True
    '    Me._txtDescripcion_0.BackColor = System.Drawing.SystemColors.Info
    '    Me._txtDescripcion_0.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtDescripcion_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
    '    Me.txtDescripcion.SetIndex(Me._txtDescripcion_0, CType(0, Short))
    '    Me._txtDescripcion_0.Location = New System.Drawing.Point(89, 198)
    '    Me._txtDescripcion_0.MaxLength = 150
    '    Me._txtDescripcion_0.Name = "_txtDescripcion_0"
    '    Me._txtDescripcion_0.ReadOnly = True
    '    Me._txtDescripcion_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtDescripcion_0.Size = New System.Drawing.Size(537, 20)
    '    Me._txtDescripcion_0.TabIndex = 34
    '    Me.ToolTip1.SetToolTip(Me._txtDescripcion_0, "Descripción del Artículo")
    '    '
    '    '_Frame1_0
    '    '
    '    Me._Frame1_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._Frame1_0.Controls.Add(Me._txtCostoReal_0)
    '    Me._Frame1_0.Controls.Add(Me._txtPrecioenDolares_0)
    '    Me._Frame1_0.Controls.Add(Me._txtCostoIndirecto_0)
    '    Me._Frame1_0.Controls.Add(Me._txtCostoAdicional_0)
    '    Me._Frame1_0.Controls.Add(Me._txtCostoFactura_0)
    '    Me._Frame1_0.Controls.Add(Me._lblMargen_0)
    '    Me._Frame1_0.Controls.Add(Me.Label1)
    '    Me._Frame1_0.Controls.Add(Me._lblArticulo_34)
    '    Me._Frame1_0.Controls.Add(Me._lblArticulo_5)
    '    Me._Frame1_0.Controls.Add(Me._lblArticulo_8)
    '    Me._Frame1_0.Controls.Add(Me._lblArticulo_7)
    '    Me._Frame1_0.Controls.Add(Me._lblArticulo_6)
    '    Me._Frame1_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.Frame1.SetIndex(Me._Frame1_0, CType(0, Short))
    '    Me._Frame1_0.Location = New System.Drawing.Point(5, 294)
    '    Me._Frame1_0.Name = "_Frame1_0"
    '    Me._Frame1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._Frame1_0.Size = New System.Drawing.Size(300, 185)
    '    Me._Frame1_0.TabIndex = 44
    '    Me._Frame1_0.TabStop = False
    '    '
    '    '_txtCostoReal_0
    '    '
    '    Me._txtCostoReal_0.AcceptsReturn = True
    '    Me._txtCostoReal_0.BackColor = System.Drawing.SystemColors.Info
    '    Me._txtCostoReal_0.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoReal_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
    '    Me.txtCostoReal.SetIndex(Me._txtCostoReal_0, CType(0, Short))
    '    Me._txtCostoReal_0.Location = New System.Drawing.Point(92, 152)
    '    Me._txtCostoReal_0.MaxLength = 0
    '    Me._txtCostoReal_0.Name = "_txtCostoReal_0"
    '    Me._txtCostoReal_0.ReadOnly = True
    '    Me._txtCostoReal_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoReal_0.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoReal_0.TabIndex = 54
    '    Me._txtCostoReal_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoReal_0, "Costo Real del artículo")
    '    '
    '    '_txtPrecioenDolares_0
    '    '
    '    Me._txtPrecioenDolares_0.AcceptsReturn = True
    '    Me._txtPrecioenDolares_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
    '    Me._txtPrecioenDolares_0.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtPrecioenDolares_0.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtPrecioenDolares.SetIndex(Me._txtPrecioenDolares_0, CType(0, Short))
    '    Me._txtPrecioenDolares_0.Location = New System.Drawing.Point(92, 24)
    '    Me._txtPrecioenDolares_0.MaxLength = 0
    '    Me._txtPrecioenDolares_0.Name = "_txtPrecioenDolares_0"
    '    Me._txtPrecioenDolares_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtPrecioenDolares_0.Size = New System.Drawing.Size(113, 20)
    '    Me._txtPrecioenDolares_0.TabIndex = 46
    '    Me._txtPrecioenDolares_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtPrecioenDolares_0, "Precio al Público en Dólares")
    '    '
    '    '_txtCostoIndirecto_0
    '    '
    '    Me._txtCostoIndirecto_0.AcceptsReturn = True
    '    Me._txtCostoIndirecto_0.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtCostoIndirecto_0.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoIndirecto_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.txtCostoIndirecto.SetIndex(Me._txtCostoIndirecto_0, CType(0, Short))
    '    Me._txtCostoIndirecto_0.Location = New System.Drawing.Point(92, 120)
    '    Me._txtCostoIndirecto_0.MaxLength = 0
    '    Me._txtCostoIndirecto_0.Name = "_txtCostoIndirecto_0"
    '    Me._txtCostoIndirecto_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoIndirecto_0.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoIndirecto_0.TabIndex = 52
    '    Me._txtCostoIndirecto_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoIndirecto_0, "Gastos Indirectos en Dólares")
    '    '
    '    '_txtCostoAdicional_0
    '    '
    '    Me._txtCostoAdicional_0.AcceptsReturn = True
    '    Me._txtCostoAdicional_0.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtCostoAdicional_0.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoAdicional_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.txtCostoAdicional.SetIndex(Me._txtCostoAdicional_0, CType(0, Short))
    '    Me._txtCostoAdicional_0.Location = New System.Drawing.Point(92, 88)
    '    Me._txtCostoAdicional_0.MaxLength = 0
    '    Me._txtCostoAdicional_0.Name = "_txtCostoAdicional_0"
    '    Me._txtCostoAdicional_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoAdicional_0.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoAdicional_0.TabIndex = 50
    '    Me._txtCostoAdicional_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoAdicional_0, "Costo en Dólares")
    '    '
    '    '_txtCostoFactura_0
    '    '
    '    Me._txtCostoFactura_0.AcceptsReturn = True
    '    Me._txtCostoFactura_0.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtCostoFactura_0.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoFactura_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.txtCostoFactura.SetIndex(Me._txtCostoFactura_0, CType(0, Short))
    '    Me._txtCostoFactura_0.Location = New System.Drawing.Point(92, 56)
    '    Me._txtCostoFactura_0.MaxLength = 0
    '    Me._txtCostoFactura_0.Name = "_txtCostoFactura_0"
    '    Me._txtCostoFactura_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoFactura_0.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoFactura_0.TabIndex = 48
    '    Me._txtCostoFactura_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoFactura_0, "Costo en Pesos")
    '    '
    '    '_lblMargen_0
    '    '
    '    Me._lblMargen_0.BackColor = System.Drawing.SystemColors.Window
    '    Me._lblMargen_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    '    Me._lblMargen_0.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblMargen_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblMargen.SetIndex(Me._lblMargen_0, CType(0, Short))
    '    Me._lblMargen_0.Location = New System.Drawing.Point(244, 152)
    '    Me._lblMargen_0.Name = "_lblMargen_0"
    '    Me._lblMargen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblMargen_0.Size = New System.Drawing.Size(49, 21)
    '    Me._lblMargen_0.TabIndex = 56
    '    Me._lblMargen_0.TextAlign = System.Drawing.ContentAlignment.TopRight
    '    '
    '    'Label1
    '    '
    '    Me.Label1.BackColor = System.Drawing.SystemColors.Control
    '    Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.Label1.Location = New System.Drawing.Point(232, 120)
    '    Me.Label1.Name = "Label1"
    '    Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.Label1.Size = New System.Drawing.Size(61, 29)
    '    Me.Label1.TabIndex = 55
    '    Me.Label1.Text = "% Margen de Venta "
    '    Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
    '    '
    '    '_lblArticulo_34
    '    '
    '    Me._lblArticulo_34.AutoSize = True
    '    Me._lblArticulo_34.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_34.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_34.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_34, CType(34, Short))
    '    Me._lblArticulo_34.Location = New System.Drawing.Point(12, 156)
    '    Me._lblArticulo_34.Name = "_lblArticulo_34"
    '    Me._lblArticulo_34.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_34.Size = New System.Drawing.Size(59, 13)
    '    Me._lblArticulo_34.TabIndex = 53
    '    Me._lblArticulo_34.Text = "Costo Real"
    '    '
    '    '_lblArticulo_5
    '    '
    '    Me._lblArticulo_5.AutoSize = True
    '    Me._lblArticulo_5.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_5.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_5.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_5, CType(5, Short))
    '    Me._lblArticulo_5.Location = New System.Drawing.Point(12, 28)
    '    Me._lblArticulo_5.Name = "_lblArticulo_5"
    '    Me._lblArticulo_5.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_5.Size = New System.Drawing.Size(75, 13)
    '    Me._lblArticulo_5.TabIndex = 45
    '    Me._lblArticulo_5.Text = "Precio Público"
    '    '
    '    '_lblArticulo_8
    '    '
    '    Me._lblArticulo_8.AutoSize = True
    '    Me._lblArticulo_8.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_8.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_8.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_8, CType(8, Short))
    '    Me._lblArticulo_8.Location = New System.Drawing.Point(12, 124)
    '    Me._lblArticulo_8.Name = "_lblArticulo_8"
    '    Me._lblArticulo_8.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_8.Size = New System.Drawing.Size(78, 13)
    '    Me._lblArticulo_8.TabIndex = 51
    '    Me._lblArticulo_8.Text = "Costo Indirecto"
    '    '
    '    '_lblArticulo_7
    '    '
    '    Me._lblArticulo_7.AutoSize = True
    '    Me._lblArticulo_7.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_7.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_7.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_7, CType(7, Short))
    '    Me._lblArticulo_7.Location = New System.Drawing.Point(12, 92)
    '    Me._lblArticulo_7.Name = "_lblArticulo_7"
    '    Me._lblArticulo_7.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_7.Size = New System.Drawing.Size(80, 13)
    '    Me._lblArticulo_7.TabIndex = 49
    '    Me._lblArticulo_7.Text = "Costo Adicional"
    '    '
    '    '_lblArticulo_6
    '    '
    '    Me._lblArticulo_6.AutoSize = True
    '    Me._lblArticulo_6.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_6.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_6.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_6, CType(6, Short))
    '    Me._lblArticulo_6.Location = New System.Drawing.Point(12, 60)
    '    Me._lblArticulo_6.Name = "_lblArticulo_6"
    '    Me._lblArticulo_6.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_6.Size = New System.Drawing.Size(73, 13)
    '    Me._lblArticulo_6.TabIndex = 47
    '    Me._lblArticulo_6.Text = "Costo Factura"
    '    '
    '    '_dbcFamilia_0
    '    '
    '    Me._dbcFamilia_0.Location = New System.Drawing.Point(89, 14)
    '    Me._dbcFamilia_0.Name = "_dbcFamilia_0"
    '    Me._dbcFamilia_0.Size = New System.Drawing.Size(265, 21)
    '    Me._dbcFamilia_0.TabIndex = 12
    '    '
    '    '_dbcLinea_0
    '    '
    '    Me._dbcLinea_0.Location = New System.Drawing.Point(89, 46)
    '    Me._dbcLinea_0.Name = "_dbcLinea_0"
    '    Me._dbcLinea_0.Size = New System.Drawing.Size(265, 21)
    '    Me._dbcLinea_0.TabIndex = 14
    '    '
    '    'dbcSubLinea
    '    '
    '    Me.dbcSubLinea.Location = New System.Drawing.Point(89, 78)
    '    Me.dbcSubLinea.Name = "dbcSubLinea"
    '    Me.dbcSubLinea.Size = New System.Drawing.Size(265, 21)
    '    Me.dbcSubLinea.TabIndex = 16
    '    '
    '    'dbcKilates
    '    '
    '    Me.dbcKilates.Location = New System.Drawing.Point(89, 110)
    '    Me.dbcKilates.Name = "dbcKilates"
    '    Me.dbcKilates.Size = New System.Drawing.Size(134, 21)
    '    Me.dbcKilates.TabIndex = 18
    '    '
    '    '_dbcMaterial_0
    '    '
    '    Me._dbcMaterial_0.Location = New System.Drawing.Point(89, 140)
    '    Me._dbcMaterial_0.Name = "_dbcMaterial_0"
    '    Me._dbcMaterial_0.Size = New System.Drawing.Size(134, 21)
    '    Me._dbcMaterial_0.TabIndex = 20
    '    '
    '    '_lblArticulo_33
    '    '
    '    Me._lblArticulo_33.AutoSize = True
    '    Me._lblArticulo_33.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_33.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_33.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_33, CType(33, Short))
    '    Me._lblArticulo_33.Location = New System.Drawing.Point(2, 172)
    '    Me._lblArticulo_33.Name = "_lblArticulo_33"
    '    Me._lblArticulo_33.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_33.Size = New System.Drawing.Size(76, 13)
    '    Me._lblArticulo_33.TabIndex = 21
    '    Me._lblArticulo_33.Text = "Dato Adicional"
    '    '
    '    '_lblArticulo_9
    '    '
    '    Me._lblArticulo_9.AutoSize = True
    '    Me._lblArticulo_9.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_9.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_9.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_9, CType(9, Short))
    '    Me._lblArticulo_9.Location = New System.Drawing.Point(2, 144)
    '    Me._lblArticulo_9.Name = "_lblArticulo_9"
    '    Me._lblArticulo_9.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_9.Size = New System.Drawing.Size(83, 13)
    '    Me._lblArticulo_9.TabIndex = 19
    '    Me._lblArticulo_9.Text = "Tipo de Material"
    '    '
    '    '_lblArticulo_29
    '    '
    '    Me._lblArticulo_29.AutoSize = True
    '    Me._lblArticulo_29.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_29.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_29.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_29, CType(29, Short))
    '    Me._lblArticulo_29.Location = New System.Drawing.Point(0, 270)
    '    Me._lblArticulo_29.Name = "_lblArticulo_29"
    '    Me._lblArticulo_29.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_29.Size = New System.Drawing.Size(92, 13)
    '    Me._lblArticulo_29.TabIndex = 36
    '    Me._lblArticulo_29.Text = "Precio público en "
    '    '
    '    '_lblDescripcion_0
    '    '
    '    Me._lblDescripcion_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
    '    Me._lblDescripcion_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    '    Me._lblDescripcion_0.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblDescripcion_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
    '    Me.lblDescripcion.SetIndex(Me._lblDescripcion_0, CType(0, Short))
    '    Me._lblDescripcion_0.Location = New System.Drawing.Point(88, 230)
    '    Me._lblDescripcion_0.Name = "_lblDescripcion_0"
    '    Me._lblDescripcion_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblDescripcion_0.Size = New System.Drawing.Size(537, 21)
    '    Me._lblDescripcion_0.TabIndex = 35
    '    '
    '    '_lblArticulo_26
    '    '
    '    Me._lblArticulo_26.AutoSize = True
    '    Me._lblArticulo_26.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_26.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_26.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_26, CType(26, Short))
    '    Me._lblArticulo_26.Location = New System.Drawing.Point(2, 114)
    '    Me._lblArticulo_26.Name = "_lblArticulo_26"
    '    Me._lblArticulo_26.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_26.Size = New System.Drawing.Size(38, 13)
    '    Me._lblArticulo_26.TabIndex = 17
    '    Me._lblArticulo_26.Text = "Kilates"
    '    '
    '    '_lblArticulo_37
    '    '
    '    Me._lblArticulo_37.AutoSize = True
    '    Me._lblArticulo_37.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_37.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_37.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_37, CType(37, Short))
    '    Me._lblArticulo_37.Location = New System.Drawing.Point(313, 270)
    '    Me._lblArticulo_37.Name = "_lblArticulo_37"
    '    Me._lblArticulo_37.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_37.Size = New System.Drawing.Size(85, 13)
    '    Me._lblArticulo_37.TabIndex = 40
    '    Me._lblArticulo_37.Text = "Moneda Compra"
    '    '
    '    '_lblArticulo_4
    '    '
    '    Me._lblArticulo_4.AutoSize = True
    '    Me._lblArticulo_4.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_4.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_4.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_4, CType(4, Short))
    '    Me._lblArticulo_4.Location = New System.Drawing.Point(2, 202)
    '    Me._lblArticulo_4.Name = "_lblArticulo_4"
    '    Me._lblArticulo_4.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_4.Size = New System.Drawing.Size(63, 13)
    '    Me._lblArticulo_4.TabIndex = 33
    '    Me._lblArticulo_4.Text = "Descripción"
    '    '
    '    '_lblArticulo_3
    '    '
    '    Me._lblArticulo_3.AutoSize = True
    '    Me._lblArticulo_3.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_3.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_3.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_3, CType(3, Short))
    '    Me._lblArticulo_3.Location = New System.Drawing.Point(2, 82)
    '    Me._lblArticulo_3.Name = "_lblArticulo_3"
    '    Me._lblArticulo_3.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_3.Size = New System.Drawing.Size(54, 13)
    '    Me._lblArticulo_3.TabIndex = 15
    '    Me._lblArticulo_3.Text = "SubLínea"
    '    '
    '    '_lblArticulo_2
    '    '
    '    Me._lblArticulo_2.AutoSize = True
    '    Me._lblArticulo_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_2.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_2, CType(2, Short))
    '    Me._lblArticulo_2.Location = New System.Drawing.Point(2, 50)
    '    Me._lblArticulo_2.Name = "_lblArticulo_2"
    '    Me._lblArticulo_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_2.Size = New System.Drawing.Size(35, 13)
    '    Me._lblArticulo_2.TabIndex = 13
    '    Me._lblArticulo_2.Text = "Línea"
    '    '
    '    '_lblArticulo_1
    '    '
    '    Me._lblArticulo_1.AutoSize = True
    '    Me._lblArticulo_1.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_1.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_1, CType(1, Short))
    '    Me._lblArticulo_1.Location = New System.Drawing.Point(2, 18)
    '    Me._lblArticulo_1.Name = "_lblArticulo_1"
    '    Me._lblArticulo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_1.Size = New System.Drawing.Size(39, 13)
    '    Me._lblArticulo_1.TabIndex = 11
    '    Me._lblArticulo_1.Text = "Familia"
    '    '
    '    '_sstArticulo_TabPage1
    '    '
    '    Me._sstArticulo_TabPage1.Controls.Add(Me._fraContenedor_1)
    '    Me._sstArticulo_TabPage1.Location = New System.Drawing.Point(4, 22)
    '    Me._sstArticulo_TabPage1.Name = "_sstArticulo_TabPage1"
    '    Me._sstArticulo_TabPage1.Size = New System.Drawing.Size(700, 527)
    '    Me._sstArticulo_TabPage1.TabIndex = 1
    '    Me._sstArticulo_TabPage1.Text = "Relojería"
    '    '
    '    '_fraContenedor_1
    '    '
    '    Me._fraContenedor_1.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraContenedor_1.Controls.Add(Me._txtAdicional_1)
    '    Me._fraContenedor_1.Controls.Add(Me._fraMoneda_3)
    '    Me._fraContenedor_1.Controls.Add(Me._txtDescripcion_1)
    '    Me._fraContenedor_1.Controls.Add(Me._fraArticulo_1)
    '    Me._fraContenedor_1.Controls.Add(Me._fraArticulo_2)
    '    Me._fraContenedor_1.Controls.Add(Me._fraImagen_1)
    '    Me._fraContenedor_1.Controls.Add(Me._fraMoneda_1)
    '    Me._fraContenedor_1.Controls.Add(Me._Frame1_2)
    '    Me._fraContenedor_1.Controls.Add(Me._Frame2_2)
    '    Me._fraContenedor_1.Controls.Add(Me.chkCrono)
    '    Me._fraContenedor_1.Controls.Add(Me.dbcMarca)
    '    Me._fraContenedor_1.Controls.Add(Me.dbcModelo)
    '    Me._fraContenedor_1.Controls.Add(Me._dbcMaterial_1)
    '    Me._fraContenedor_1.Controls.Add(Me._lblArticulo_45)
    '    Me._fraContenedor_1.Controls.Add(Me._lblArticulo_27)
    '    Me._fraContenedor_1.Controls.Add(Me._lblArticulo_13)
    '    Me._fraContenedor_1.Controls.Add(Me._lblArticulo_12)
    '    Me._fraContenedor_1.Controls.Add(Me._lblArticulo_14)
    '    Me._fraContenedor_1.Controls.Add(Me._lblArticulo_15)
    '    Me._fraContenedor_1.Controls.Add(Me._lblArticulo_16)
    '    Me._fraContenedor_1.Controls.Add(Me._lblArticulo_17)
    '    Me._fraContenedor_1.Controls.Add(Me._lblArticulo_38)
    '    Me._fraContenedor_1.Controls.Add(Me._lblDescripcion_1)
    '    Me._fraContenedor_1.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._fraContenedor_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraContenedor.SetIndex(Me._fraContenedor_1, CType(1, Short))
    '    Me._fraContenedor_1.Location = New System.Drawing.Point(8, 24)
    '    Me._fraContenedor_1.Name = "_fraContenedor_1"
    '    Me._fraContenedor_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraContenedor_1.Size = New System.Drawing.Size(689, 481)
    '    Me._fraContenedor_1.TabIndex = 70
    '    '
    '    '_txtAdicional_1
    '    '
    '    Me._txtAdicional_1.AcceptsReturn = True
    '    Me._txtAdicional_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
    '    Me._txtAdicional_1.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtAdicional_1.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtAdicional.SetIndex(Me._txtAdicional_1, CType(1, Short))
    '    Me._txtAdicional_1.Location = New System.Drawing.Point(89, 168)
    '    Me._txtAdicional_1.MaxLength = 15
    '    Me._txtAdicional_1.Name = "_txtAdicional_1"
    '    Me._txtAdicional_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtAdicional_1.Size = New System.Drawing.Size(120, 20)
    '    Me._txtAdicional_1.TabIndex = 88
    '    '
    '    '_fraMoneda_3
    '    '
    '    Me._fraMoneda_3.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraMoneda_3.Controls.Add(Me._optMoneda_7)
    '    Me._fraMoneda_3.Controls.Add(Me._optMoneda_6)
    '    Me._fraMoneda_3.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraMoneda.SetIndex(Me._fraMoneda_3, CType(3, Short))
    '    Me._fraMoneda_3.Location = New System.Drawing.Point(89, 256)
    '    Me._fraMoneda_3.Name = "_fraMoneda_3"
    '    Me._fraMoneda_3.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraMoneda_3.Size = New System.Drawing.Size(218, 33)
    '    Me._fraMoneda_3.TabIndex = 94
    '    Me._fraMoneda_3.TabStop = False
    '    '
    '    '_optMoneda_7
    '    '
    '    Me._optMoneda_7.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_7.Checked = True
    '    Me._optMoneda_7.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_7.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_7, CType(7, Short))
    '    Me._optMoneda_7.Location = New System.Drawing.Point(36, 11)
    '    Me._optMoneda_7.Name = "_optMoneda_7"
    '    Me._optMoneda_7.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_7.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_7.TabIndex = 95
    '    Me._optMoneda_7.TabStop = True
    '    Me._optMoneda_7.Tag = "1"
    '    Me._optMoneda_7.Text = "Dólares"
    '    Me._optMoneda_7.UseVisualStyleBackColor = False
    '    '
    '    '_optMoneda_6
    '    '
    '    Me._optMoneda_6.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_6.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_6.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_6, CType(6, Short))
    '    Me._optMoneda_6.Location = New System.Drawing.Point(129, 11)
    '    Me._optMoneda_6.Name = "_optMoneda_6"
    '    Me._optMoneda_6.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_6.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_6.TabIndex = 96
    '    Me._optMoneda_6.TabStop = True
    '    Me._optMoneda_6.Text = "Pesos"
    '    Me._optMoneda_6.UseVisualStyleBackColor = False
    '    '
    '    '_txtDescripcion_1
    '    '
    '    Me._txtDescripcion_1.AcceptsReturn = True
    '    Me._txtDescripcion_1.BackColor = System.Drawing.SystemColors.Info
    '    Me._txtDescripcion_1.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtDescripcion_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
    '    Me.txtDescripcion.SetIndex(Me._txtDescripcion_1, CType(1, Short))
    '    Me._txtDescripcion_1.Location = New System.Drawing.Point(89, 198)
    '    Me._txtDescripcion_1.MaxLength = 0
    '    Me._txtDescripcion_1.Name = "_txtDescripcion_1"
    '    Me._txtDescripcion_1.ReadOnly = True
    '    Me._txtDescripcion_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtDescripcion_1.Size = New System.Drawing.Size(537, 20)
    '    Me._txtDescripcion_1.TabIndex = 91
    '    Me.ToolTip1.SetToolTip(Me._txtDescripcion_1, "Descripción del Reloj")
    '    '
    '    '_fraArticulo_1
    '    '
    '    Me._fraArticulo_1.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraArticulo_1.Controls.Add(Me._optGenero_0)
    '    Me._fraArticulo_1.Controls.Add(Me._optGenero_1)
    '    Me._fraArticulo_1.Controls.Add(Me._optGenero_2)
    '    Me._fraArticulo_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraArticulo.SetIndex(Me._fraArticulo_1, CType(1, Short))
    '    Me._fraArticulo_1.Location = New System.Drawing.Point(89, 68)
    '    Me._fraArticulo_1.Name = "_fraArticulo_1"
    '    Me._fraArticulo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraArticulo_1.Size = New System.Drawing.Size(265, 33)
    '    Me._fraArticulo_1.TabIndex = 76
    '    Me._fraArticulo_1.TabStop = False
    '    '
    '    '_optGenero_0
    '    '
    '    Me._optGenero_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._optGenero_0.Checked = True
    '    Me._optGenero_0.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optGenero_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optGenero.SetIndex(Me._optGenero_0, CType(0, Short))
    '    Me._optGenero_0.Location = New System.Drawing.Point(8, 8)
    '    Me._optGenero_0.Name = "_optGenero_0"
    '    Me._optGenero_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optGenero_0.Size = New System.Drawing.Size(86, 21)
    '    Me._optGenero_0.TabIndex = 77
    '    Me._optGenero_0.TabStop = True
    '    Me._optGenero_0.Text = "Caballero"
    '    Me.ToolTip1.SetToolTip(Me._optGenero_0, "Para Hombre")
    '    Me._optGenero_0.UseVisualStyleBackColor = False
    '    '
    '    '_optGenero_1
    '    '
    '    Me._optGenero_1.BackColor = System.Drawing.SystemColors.Control
    '    Me._optGenero_1.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optGenero_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optGenero.SetIndex(Me._optGenero_1, CType(1, Short))
    '    Me._optGenero_1.Location = New System.Drawing.Point(100, 8)
    '    Me._optGenero_1.Name = "_optGenero_1"
    '    Me._optGenero_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optGenero_1.Size = New System.Drawing.Size(57, 21)
    '    Me._optGenero_1.TabIndex = 78
    '    Me._optGenero_1.TabStop = True
    '    Me._optGenero_1.Text = "Dama"
    '    Me.ToolTip1.SetToolTip(Me._optGenero_1, "Para Mujer")
    '    Me._optGenero_1.UseVisualStyleBackColor = False
    '    '
    '    '_optGenero_2
    '    '
    '    Me._optGenero_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._optGenero_2.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optGenero_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optGenero.SetIndex(Me._optGenero_2, CType(2, Short))
    '    Me._optGenero_2.Location = New System.Drawing.Point(190, 8)
    '    Me._optGenero_2.Name = "_optGenero_2"
    '    Me._optGenero_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optGenero_2.Size = New System.Drawing.Size(73, 21)
    '    Me._optGenero_2.TabIndex = 79
    '    Me._optGenero_2.TabStop = True
    '    Me._optGenero_2.Text = "Mediano"
    '    Me.ToolTip1.SetToolTip(Me._optGenero_2, "Para cualquier tipo de sexo")
    '    Me._optGenero_2.UseVisualStyleBackColor = False
    '    '
    '    '_fraArticulo_2
    '    '
    '    Me._fraArticulo_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraArticulo_2.Controls.Add(Me._optMovimiento_0)
    '    Me._fraArticulo_2.Controls.Add(Me._optMovimiento_1)
    '    Me._fraArticulo_2.Controls.Add(Me._optMovimiento_2)
    '    Me._fraArticulo_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraArticulo.SetIndex(Me._fraArticulo_2, CType(2, Short))
    '    Me._fraArticulo_2.Location = New System.Drawing.Point(89, 101)
    '    Me._fraArticulo_2.Name = "_fraArticulo_2"
    '    Me._fraArticulo_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraArticulo_2.Size = New System.Drawing.Size(265, 33)
    '    Me._fraArticulo_2.TabIndex = 81
    '    Me._fraArticulo_2.TabStop = False
    '    '
    '    '_optMovimiento_0
    '    '
    '    Me._optMovimiento_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMovimiento_0.Checked = True
    '    Me._optMovimiento_0.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMovimiento_0.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMovimiento.SetIndex(Me._optMovimiento_0, CType(0, Short))
    '    Me._optMovimiento_0.Location = New System.Drawing.Point(8, 8)
    '    Me._optMovimiento_0.Name = "_optMovimiento_0"
    '    Me._optMovimiento_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMovimiento_0.Size = New System.Drawing.Size(57, 21)
    '    Me._optMovimiento_0.TabIndex = 82
    '    Me._optMovimiento_0.TabStop = True
    '    Me._optMovimiento_0.Text = "Cuarzo"
    '    Me.ToolTip1.SetToolTip(Me._optMovimiento_0, "Movimiento por Cuarzo")
    '    Me._optMovimiento_0.UseVisualStyleBackColor = False
    '    '
    '    '_optMovimiento_1
    '    '
    '    Me._optMovimiento_1.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMovimiento_1.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMovimiento_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMovimiento.SetIndex(Me._optMovimiento_1, CType(1, Short))
    '    Me._optMovimiento_1.Location = New System.Drawing.Point(100, 8)
    '    Me._optMovimiento_1.Name = "_optMovimiento_1"
    '    Me._optMovimiento_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMovimiento_1.Size = New System.Drawing.Size(81, 21)
    '    Me._optMovimiento_1.TabIndex = 83
    '    Me._optMovimiento_1.TabStop = True
    '    Me._optMovimiento_1.Text = "Automático"
    '    Me.ToolTip1.SetToolTip(Me._optMovimiento_1, "Movimiento Automatizado")
    '    Me._optMovimiento_1.UseVisualStyleBackColor = False
    '    '
    '    '_optMovimiento_2
    '    '
    '    Me._optMovimiento_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMovimiento_2.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMovimiento_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMovimiento.SetIndex(Me._optMovimiento_2, CType(2, Short))
    '    Me._optMovimiento_2.Location = New System.Drawing.Point(190, 8)
    '    Me._optMovimiento_2.Name = "_optMovimiento_2"
    '    Me._optMovimiento_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMovimiento_2.Size = New System.Drawing.Size(65, 21)
    '    Me._optMovimiento_2.TabIndex = 84
    '    Me._optMovimiento_2.TabStop = True
    '    Me._optMovimiento_2.Text = "Manual"
    '    Me.ToolTip1.SetToolTip(Me._optMovimiento_2, "Manual")
    '    Me._optMovimiento_2.UseVisualStyleBackColor = False
    '    '
    '    '_fraImagen_1
    '    '
    '    Me._fraImagen_1.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraImagen_1.Controls.Add(Me.Image2)
    '    Me._fraImagen_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
    '    Me.fraImagen.SetIndex(Me._fraImagen_1, CType(1, Short))
    '    Me._fraImagen_1.Location = New System.Drawing.Point(510, 8)
    '    Me._fraImagen_1.Name = "_fraImagen_1"
    '    Me._fraImagen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraImagen_1.Size = New System.Drawing.Size(178, 186)
    '    Me._fraImagen_1.TabIndex = 126
    '    Me._fraImagen_1.TabStop = False
    '    Me._fraImagen_1.Text = "Imagen del Artículo"
    '    '
    '    'Image2
    '    '
    '    Me.Image2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    '    Me.Image2.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.Image2.Location = New System.Drawing.Point(7, 21)
    '    Me.Image2.Name = "Image2"
    '    Me.Image2.Size = New System.Drawing.Size(163, 157)
    '    Me.Image2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
    '    Me.Image2.TabIndex = 0
    '    Me.Image2.TabStop = False
    '    '
    '    '_fraMoneda_1
    '    '
    '    Me._fraMoneda_1.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraMoneda_1.Controls.Add(Me._optMoneda_3)
    '    Me._fraMoneda_1.Controls.Add(Me._optMoneda_2)
    '    Me._fraMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraMoneda.SetIndex(Me._fraMoneda_1, CType(1, Short))
    '    Me._fraMoneda_1.Location = New System.Drawing.Point(416, 256)
    '    Me._fraMoneda_1.Name = "_fraMoneda_1"
    '    Me._fraMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraMoneda_1.Size = New System.Drawing.Size(209, 33)
    '    Me._fraMoneda_1.TabIndex = 98
    '    Me._fraMoneda_1.TabStop = False
    '    '
    '    '_optMoneda_3
    '    '
    '    Me._optMoneda_3.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_3.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_3.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_3, CType(3, Short))
    '    Me._optMoneda_3.Location = New System.Drawing.Point(134, 11)
    '    Me._optMoneda_3.Name = "_optMoneda_3"
    '    Me._optMoneda_3.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_3.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_3.TabIndex = 100
    '    Me._optMoneda_3.TabStop = True
    '    Me._optMoneda_3.Text = "Pesos"
    '    Me._optMoneda_3.UseVisualStyleBackColor = False
    '    '
    '    '_optMoneda_2
    '    '
    '    Me._optMoneda_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_2.Checked = True
    '    Me._optMoneda_2.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_2, CType(2, Short))
    '    Me._optMoneda_2.Location = New System.Drawing.Point(34, 11)
    '    Me._optMoneda_2.Name = "_optMoneda_2"
    '    Me._optMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_2.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_2.TabIndex = 99
    '    Me._optMoneda_2.TabStop = True
    '    Me._optMoneda_2.Text = "Dólares"
    '    Me._optMoneda_2.UseVisualStyleBackColor = False
    '    '
    '    '_Frame1_2
    '    '
    '    Me._Frame1_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._Frame1_2.Controls.Add(Me._txtCostoReal_1)
    '    Me._Frame1_2.Controls.Add(Me._txtPrecioenDolares_1)
    '    Me._Frame1_2.Controls.Add(Me._txtCostoIndirecto_1)
    '    Me._Frame1_2.Controls.Add(Me._txtCostoAdicional_1)
    '    Me._Frame1_2.Controls.Add(Me._txtCostoFactura_1)
    '    Me._Frame1_2.Controls.Add(Me._lblMargen_1)
    '    Me._Frame1_2.Controls.Add(Me.Label2)
    '    Me._Frame1_2.Controls.Add(Me._lblArticulo_40)
    '    Me._Frame1_2.Controls.Add(Me._lblArticulo_41)
    '    Me._Frame1_2.Controls.Add(Me._lblArticulo_42)
    '    Me._Frame1_2.Controls.Add(Me._lblArticulo_43)
    '    Me._Frame1_2.Controls.Add(Me._lblArticulo_44)
    '    Me._Frame1_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.Frame1.SetIndex(Me._Frame1_2, CType(2, Short))
    '    Me._Frame1_2.Location = New System.Drawing.Point(0, 294)
    '    Me._Frame1_2.Name = "_Frame1_2"
    '    Me._Frame1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._Frame1_2.Size = New System.Drawing.Size(305, 185)
    '    Me._Frame1_2.TabIndex = 101
    '    Me._Frame1_2.TabStop = False
    '    '
    '    '_txtCostoReal_1
    '    '
    '    Me._txtCostoReal_1.AcceptsReturn = True
    '    Me._txtCostoReal_1.BackColor = System.Drawing.SystemColors.Info
    '    Me._txtCostoReal_1.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoReal_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
    '    Me.txtCostoReal.SetIndex(Me._txtCostoReal_1, CType(1, Short))
    '    Me._txtCostoReal_1.Location = New System.Drawing.Point(92, 152)
    '    Me._txtCostoReal_1.MaxLength = 0
    '    Me._txtCostoReal_1.Name = "_txtCostoReal_1"
    '    Me._txtCostoReal_1.ReadOnly = True
    '    Me._txtCostoReal_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoReal_1.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoReal_1.TabIndex = 111
    '    Me._txtCostoReal_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    '
    '    '_txtPrecioenDolares_1
    '    '
    '    Me._txtPrecioenDolares_1.AcceptsReturn = True
    '    Me._txtPrecioenDolares_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
    '    Me._txtPrecioenDolares_1.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtPrecioenDolares_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.txtPrecioenDolares.SetIndex(Me._txtPrecioenDolares_1, CType(1, Short))
    '    Me._txtPrecioenDolares_1.Location = New System.Drawing.Point(92, 24)
    '    Me._txtPrecioenDolares_1.MaxLength = 0
    '    Me._txtPrecioenDolares_1.Name = "_txtPrecioenDolares_1"
    '    Me._txtPrecioenDolares_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtPrecioenDolares_1.Size = New System.Drawing.Size(113, 20)
    '    Me._txtPrecioenDolares_1.TabIndex = 103
    '    Me._txtPrecioenDolares_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtPrecioenDolares_1, "Precio al Público en Dólares")
    '    '
    '    '_txtCostoIndirecto_1
    '    '
    '    Me._txtCostoIndirecto_1.AcceptsReturn = True
    '    Me._txtCostoIndirecto_1.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtCostoIndirecto_1.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoIndirecto_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.txtCostoIndirecto.SetIndex(Me._txtCostoIndirecto_1, CType(1, Short))
    '    Me._txtCostoIndirecto_1.Location = New System.Drawing.Point(92, 120)
    '    Me._txtCostoIndirecto_1.MaxLength = 0
    '    Me._txtCostoIndirecto_1.Name = "_txtCostoIndirecto_1"
    '    Me._txtCostoIndirecto_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoIndirecto_1.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoIndirecto_1.TabIndex = 109
    '    Me._txtCostoIndirecto_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoIndirecto_1, "Gastos Indirectos en Dólares")
    '    '
    '    '_txtCostoAdicional_1
    '    '
    '    Me._txtCostoAdicional_1.AcceptsReturn = True
    '    Me._txtCostoAdicional_1.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtCostoAdicional_1.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoAdicional_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.txtCostoAdicional.SetIndex(Me._txtCostoAdicional_1, CType(1, Short))
    '    Me._txtCostoAdicional_1.Location = New System.Drawing.Point(92, 88)
    '    Me._txtCostoAdicional_1.MaxLength = 0
    '    Me._txtCostoAdicional_1.Name = "_txtCostoAdicional_1"
    '    Me._txtCostoAdicional_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoAdicional_1.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoAdicional_1.TabIndex = 107
    '    Me._txtCostoAdicional_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoAdicional_1, "Costo en Dólares")
    '    '
    '    '_txtCostoFactura_1
    '    '
    '    Me._txtCostoFactura_1.AcceptsReturn = True
    '    Me._txtCostoFactura_1.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtCostoFactura_1.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoFactura_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.txtCostoFactura.SetIndex(Me._txtCostoFactura_1, CType(1, Short))
    '    Me._txtCostoFactura_1.Location = New System.Drawing.Point(92, 56)
    '    Me._txtCostoFactura_1.MaxLength = 0
    '    Me._txtCostoFactura_1.Name = "_txtCostoFactura_1"
    '    Me._txtCostoFactura_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoFactura_1.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoFactura_1.TabIndex = 105
    '    Me._txtCostoFactura_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoFactura_1, "Costo en Pesos")
    '    '
    '    '_lblMargen_1
    '    '
    '    Me._lblMargen_1.BackColor = System.Drawing.SystemColors.Window
    '    Me._lblMargen_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    '    Me._lblMargen_1.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblMargen_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblMargen.SetIndex(Me._lblMargen_1, CType(1, Short))
    '    Me._lblMargen_1.Location = New System.Drawing.Point(244, 152)
    '    Me._lblMargen_1.Name = "_lblMargen_1"
    '    Me._lblMargen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblMargen_1.Size = New System.Drawing.Size(49, 21)
    '    Me._lblMargen_1.TabIndex = 113
    '    Me._lblMargen_1.TextAlign = System.Drawing.ContentAlignment.TopRight
    '    '
    '    'Label2
    '    '
    '    Me.Label2.BackColor = System.Drawing.SystemColors.Control
    '    Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.Label2.Location = New System.Drawing.Point(232, 120)
    '    Me.Label2.Name = "Label2"
    '    Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.Label2.Size = New System.Drawing.Size(61, 29)
    '    Me.Label2.TabIndex = 112
    '    Me.Label2.Text = "% Margen de Venta "
    '    Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
    '    '
    '    '_lblArticulo_40
    '    '
    '    Me._lblArticulo_40.AutoSize = True
    '    Me._lblArticulo_40.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_40.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_40.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_40, CType(40, Short))
    '    Me._lblArticulo_40.Location = New System.Drawing.Point(12, 156)
    '    Me._lblArticulo_40.Name = "_lblArticulo_40"
    '    Me._lblArticulo_40.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_40.Size = New System.Drawing.Size(59, 13)
    '    Me._lblArticulo_40.TabIndex = 110
    '    Me._lblArticulo_40.Text = "Costo Real"
    '    Me.ToolTip1.SetToolTip(Me._lblArticulo_40, "Costo Real del artículo")
    '    '
    '    '_lblArticulo_41
    '    '
    '    Me._lblArticulo_41.AutoSize = True
    '    Me._lblArticulo_41.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_41.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_41.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_41, CType(41, Short))
    '    Me._lblArticulo_41.Location = New System.Drawing.Point(12, 28)
    '    Me._lblArticulo_41.Name = "_lblArticulo_41"
    '    Me._lblArticulo_41.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_41.Size = New System.Drawing.Size(75, 13)
    '    Me._lblArticulo_41.TabIndex = 102
    '    Me._lblArticulo_41.Text = "Precio Público"
    '    '
    '    '_lblArticulo_42
    '    '
    '    Me._lblArticulo_42.AutoSize = True
    '    Me._lblArticulo_42.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_42.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_42.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_42, CType(42, Short))
    '    Me._lblArticulo_42.Location = New System.Drawing.Point(12, 124)
    '    Me._lblArticulo_42.Name = "_lblArticulo_42"
    '    Me._lblArticulo_42.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_42.Size = New System.Drawing.Size(78, 13)
    '    Me._lblArticulo_42.TabIndex = 108
    '    Me._lblArticulo_42.Text = "Costo Indirecto"
    '    '
    '    '_lblArticulo_43
    '    '
    '    Me._lblArticulo_43.AutoSize = True
    '    Me._lblArticulo_43.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_43.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_43.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_43, CType(43, Short))
    '    Me._lblArticulo_43.Location = New System.Drawing.Point(12, 92)
    '    Me._lblArticulo_43.Name = "_lblArticulo_43"
    '    Me._lblArticulo_43.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_43.Size = New System.Drawing.Size(80, 13)
    '    Me._lblArticulo_43.TabIndex = 106
    '    Me._lblArticulo_43.Text = "Costo Adicional"
    '    '
    '    '_lblArticulo_44
    '    '
    '    Me._lblArticulo_44.AutoSize = True
    '    Me._lblArticulo_44.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_44.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_44.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_44, CType(44, Short))
    '    Me._lblArticulo_44.Location = New System.Drawing.Point(12, 60)
    '    Me._lblArticulo_44.Name = "_lblArticulo_44"
    '    Me._lblArticulo_44.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_44.Size = New System.Drawing.Size(73, 13)
    '    Me._lblArticulo_44.TabIndex = 104
    '    Me._lblArticulo_44.Text = "Costo Factura"
    '    '
    '    '_Frame2_2
    '    '
    '    Me._Frame2_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._Frame2_2.Controls.Add(Me._Frame4_1)
    '    Me._Frame2_2.Controls.Add(Me._txtCodigodelProveedor_1)
    '    Me._Frame2_2.Controls.Add(Me._dbcProveedor_1)
    '    Me._Frame2_2.Controls.Add(Me._cboUnidad_1)
    '    Me._Frame2_2.Controls.Add(Me._cboAlmacen_1)
    '    Me._Frame2_2.Controls.Add(Me._lblArticulo_18)
    '    Me._Frame2_2.Controls.Add(Me._lblArticulo_19)
    '    Me._Frame2_2.Controls.Add(Me._lblArticulo_20)
    '    Me._Frame2_2.Controls.Add(Me._lblArticulo_21)
    '    Me._Frame2_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.Frame2.SetIndex(Me._Frame2_2, CType(2, Short))
    '    Me._Frame2_2.Location = New System.Drawing.Point(312, 294)
    '    Me._Frame2_2.Name = "_Frame2_2"
    '    Me._Frame2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._Frame2_2.Size = New System.Drawing.Size(313, 185)
    '    Me._Frame2_2.TabIndex = 114
    '    Me._Frame2_2.TabStop = False
    '    '
    '    '_Frame4_1
    '    '
    '    Me._Frame4_1.BackColor = System.Drawing.SystemColors.Control
    '    Me._Frame4_1.Controls.Add(Me._txtImagen_1)
    '    Me._Frame4_1.Controls.Add(Me._cmdBuscarImagen_1)
    '    Me._Frame4_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
    '    Me.Frame4.SetIndex(Me._Frame4_1, CType(1, Short))
    '    Me._Frame4_1.Location = New System.Drawing.Point(12, 132)
    '    Me._Frame4_1.Name = "_Frame4_1"
    '    Me._Frame4_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._Frame4_1.Size = New System.Drawing.Size(290, 44)
    '    Me._Frame4_1.TabIndex = 123
    '    Me._Frame4_1.TabStop = False
    '    Me._Frame4_1.Text = "Imagen"
    '    '
    '    '_txtImagen_1
    '    '
    '    Me._txtImagen_1.AcceptsReturn = True
    '    Me._txtImagen_1.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtImagen_1.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtImagen_1.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtImagen.SetIndex(Me._txtImagen_1, CType(1, Short))
    '    Me._txtImagen_1.Location = New System.Drawing.Point(9, 15)
    '    Me._txtImagen_1.MaxLength = 0
    '    Me._txtImagen_1.Name = "_txtImagen_1"
    '    Me._txtImagen_1.ReadOnly = True
    '    Me._txtImagen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtImagen_1.Size = New System.Drawing.Size(245, 20)
    '    Me._txtImagen_1.TabIndex = 124
    '    '
    '    '_cmdBuscarImagen_1
    '    '
    '    Me._cmdBuscarImagen_1.BackColor = System.Drawing.SystemColors.Control
    '    Me._cmdBuscarImagen_1.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._cmdBuscarImagen_1.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.cmdBuscarImagen.SetIndex(Me._cmdBuscarImagen_1, CType(1, Short))
    '    Me._cmdBuscarImagen_1.Location = New System.Drawing.Point(260, 15)
    '    Me._cmdBuscarImagen_1.Name = "_cmdBuscarImagen_1"
    '    Me._cmdBuscarImagen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._cmdBuscarImagen_1.Size = New System.Drawing.Size(22, 21)
    '    Me._cmdBuscarImagen_1.TabIndex = 125
    '    Me._cmdBuscarImagen_1.Text = "..."
    '    Me._cmdBuscarImagen_1.UseVisualStyleBackColor = False
    '    '
    '    '_txtCodigodelProveedor_1
    '    '
    '    Me._txtCodigodelProveedor_1.AcceptsReturn = True
    '    Me._txtCodigodelProveedor_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
    '    Me._txtCodigodelProveedor_1.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCodigodelProveedor_1.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtCodigodelProveedor.SetIndex(Me._txtCodigodelProveedor_1, CType(1, Short))
    '    Me._txtCodigodelProveedor_1.Location = New System.Drawing.Point(171, 102)
    '    Me._txtCodigodelProveedor_1.MaxLength = 20
    '    Me._txtCodigodelProveedor_1.Name = "_txtCodigodelProveedor_1"
    '    Me._txtCodigodelProveedor_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCodigodelProveedor_1.Size = New System.Drawing.Size(129, 20)
    '    Me._txtCodigodelProveedor_1.TabIndex = 122
    '    Me.ToolTip1.SetToolTip(Me._txtCodigodelProveedor_1, "Código que usa el Proveedor para el Artículo")
    '    '
    '    '_dbcProveedor_1
    '    '
    '    Me._dbcProveedor_1.Location = New System.Drawing.Point(100, 74)
    '    Me._dbcProveedor_1.Name = "_dbcProveedor_1"
    '    Me._dbcProveedor_1.Size = New System.Drawing.Size(201, 21)
    '    Me._dbcProveedor_1.TabIndex = 120
    '    '
    '    '_cboUnidad_1
    '    '
    '    Me._cboUnidad_1.Location = New System.Drawing.Point(100, 17)
    '    Me._cboUnidad_1.Name = "_cboUnidad_1"
    '    Me._cboUnidad_1.Size = New System.Drawing.Size(78, 21)
    '    Me._cboUnidad_1.TabIndex = 116
    '    '
    '    '_cboAlmacen_1
    '    '
    '    Me._cboAlmacen_1.Location = New System.Drawing.Point(100, 46)
    '    Me._cboAlmacen_1.Name = "_cboAlmacen_1"
    '    Me._cboAlmacen_1.Size = New System.Drawing.Size(201, 21)
    '    Me._cboAlmacen_1.TabIndex = 118
    '    '
    '    '_lblArticulo_18
    '    '
    '    Me._lblArticulo_18.AutoSize = True
    '    Me._lblArticulo_18.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_18.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_18.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_18, CType(18, Short))
    '    Me._lblArticulo_18.Location = New System.Drawing.Point(12, 50)
    '    Me._lblArticulo_18.Name = "_lblArticulo_18"
    '    Me._lblArticulo_18.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_18.Size = New System.Drawing.Size(84, 13)
    '    Me._lblArticulo_18.TabIndex = 117
    '    Me._lblArticulo_18.Text = "Almacén/Origen"
    '    '
    '    '_lblArticulo_19
    '    '
    '    Me._lblArticulo_19.AutoSize = True
    '    Me._lblArticulo_19.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_19.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_19.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_19, CType(19, Short))
    '    Me._lblArticulo_19.Location = New System.Drawing.Point(12, 21)
    '    Me._lblArticulo_19.Name = "_lblArticulo_19"
    '    Me._lblArticulo_19.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_19.Size = New System.Drawing.Size(41, 13)
    '    Me._lblArticulo_19.TabIndex = 115
    '    Me._lblArticulo_19.Text = "Unidad"
    '    '
    '    '_lblArticulo_20
    '    '
    '    Me._lblArticulo_20.AutoSize = True
    '    Me._lblArticulo_20.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_20.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_20.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_20, CType(20, Short))
    '    Me._lblArticulo_20.Location = New System.Drawing.Point(19, 106)
    '    Me._lblArticulo_20.Name = "_lblArticulo_20"
    '    Me._lblArticulo_20.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_20.Size = New System.Drawing.Size(156, 13)
    '    Me._lblArticulo_20.TabIndex = 121
    '    Me._lblArticulo_20.Text = "Código artículo del proveedor : "
    '    '
    '    '_lblArticulo_21
    '    '
    '    Me._lblArticulo_21.AutoSize = True
    '    Me._lblArticulo_21.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_21.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_21.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_21, CType(21, Short))
    '    Me._lblArticulo_21.Location = New System.Drawing.Point(12, 79)
    '    Me._lblArticulo_21.Name = "_lblArticulo_21"
    '    Me._lblArticulo_21.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_21.Size = New System.Drawing.Size(56, 13)
    '    Me._lblArticulo_21.TabIndex = 119
    '    Me._lblArticulo_21.Text = "Proveedor"
    '    '
    '    'chkCrono
    '    '
    '    Me.chkCrono.BackColor = System.Drawing.SystemColors.Control
    '    Me.chkCrono.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.chkCrono.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.chkCrono.Location = New System.Drawing.Point(279, 174)
    '    Me.chkCrono.Name = "chkCrono"
    '    Me.chkCrono.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.chkCrono.Size = New System.Drawing.Size(81, 17)
    '    Me.chkCrono.TabIndex = 89
    '    Me.chkCrono.Text = "Cronógrafo"
    '    Me.chkCrono.UseVisualStyleBackColor = False
    '    '
    '    'dbcMarca
    '    '
    '    Me.dbcMarca.Location = New System.Drawing.Point(89, 14)
    '    Me.dbcMarca.Name = "dbcMarca"
    '    Me.dbcMarca.Size = New System.Drawing.Size(265, 21)
    '    Me.dbcMarca.TabIndex = 72
    '    '
    '    'dbcModelo
    '    '
    '    Me.dbcModelo.Location = New System.Drawing.Point(89, 46)
    '    Me.dbcModelo.Name = "dbcModelo"
    '    Me.dbcModelo.Size = New System.Drawing.Size(265, 21)
    '    Me.dbcModelo.TabIndex = 74
    '    '
    '    '_dbcMaterial_1
    '    '
    '    Me._dbcMaterial_1.Location = New System.Drawing.Point(89, 140)
    '    Me._dbcMaterial_1.Name = "_dbcMaterial_1"
    '    Me._dbcMaterial_1.Size = New System.Drawing.Size(265, 21)
    '    Me._dbcMaterial_1.TabIndex = 86
    '    '
    '    '_lblArticulo_45
    '    '
    '    Me._lblArticulo_45.AutoSize = True
    '    Me._lblArticulo_45.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_45.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_45.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_45, CType(45, Short))
    '    Me._lblArticulo_45.Location = New System.Drawing.Point(2, 172)
    '    Me._lblArticulo_45.Name = "_lblArticulo_45"
    '    Me._lblArticulo_45.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_45.Size = New System.Drawing.Size(76, 13)
    '    Me._lblArticulo_45.TabIndex = 87
    '    Me._lblArticulo_45.Text = "Dato Adicional"
    '    '
    '    '_lblArticulo_27
    '    '
    '    Me._lblArticulo_27.AutoSize = True
    '    Me._lblArticulo_27.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_27.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_27.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_27, CType(27, Short))
    '    Me._lblArticulo_27.Location = New System.Drawing.Point(0, 270)
    '    Me._lblArticulo_27.Name = "_lblArticulo_27"
    '    Me._lblArticulo_27.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_27.Size = New System.Drawing.Size(92, 13)
    '    Me._lblArticulo_27.TabIndex = 93
    '    Me._lblArticulo_27.Text = "Precio público en "
    '    '
    '    '_lblArticulo_13
    '    '
    '    Me._lblArticulo_13.AutoSize = True
    '    Me._lblArticulo_13.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_13.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_13.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_13, CType(13, Short))
    '    Me._lblArticulo_13.Location = New System.Drawing.Point(2, 50)
    '    Me._lblArticulo_13.Name = "_lblArticulo_13"
    '    Me._lblArticulo_13.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_13.Size = New System.Drawing.Size(42, 13)
    '    Me._lblArticulo_13.TabIndex = 73
    '    Me._lblArticulo_13.Text = "Modelo"
    '    '
    '    '_lblArticulo_12
    '    '
    '    Me._lblArticulo_12.AutoSize = True
    '    Me._lblArticulo_12.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_12.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_12.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_12, CType(12, Short))
    '    Me._lblArticulo_12.Location = New System.Drawing.Point(2, 18)
    '    Me._lblArticulo_12.Name = "_lblArticulo_12"
    '    Me._lblArticulo_12.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_12.Size = New System.Drawing.Size(37, 13)
    '    Me._lblArticulo_12.TabIndex = 71
    '    Me._lblArticulo_12.Text = "Marca"
    '    '
    '    '_lblArticulo_14
    '    '
    '    Me._lblArticulo_14.AutoSize = True
    '    Me._lblArticulo_14.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_14.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_14.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_14, CType(14, Short))
    '    Me._lblArticulo_14.Location = New System.Drawing.Point(2, 202)
    '    Me._lblArticulo_14.Name = "_lblArticulo_14"
    '    Me._lblArticulo_14.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_14.Size = New System.Drawing.Size(63, 13)
    '    Me._lblArticulo_14.TabIndex = 90
    '    Me._lblArticulo_14.Text = "Descripción"
    '    '
    '    '_lblArticulo_15
    '    '
    '    Me._lblArticulo_15.AutoSize = True
    '    Me._lblArticulo_15.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_15.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_15.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_15, CType(15, Short))
    '    Me._lblArticulo_15.Location = New System.Drawing.Point(2, 82)
    '    Me._lblArticulo_15.Name = "_lblArticulo_15"
    '    Me._lblArticulo_15.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_15.Size = New System.Drawing.Size(42, 13)
    '    Me._lblArticulo_15.TabIndex = 75
    '    Me._lblArticulo_15.Text = "Género"
    '    '
    '    '_lblArticulo_16
    '    '
    '    Me._lblArticulo_16.AutoSize = True
    '    Me._lblArticulo_16.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_16.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_16.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_16, CType(16, Short))
    '    Me._lblArticulo_16.Location = New System.Drawing.Point(2, 114)
    '    Me._lblArticulo_16.Name = "_lblArticulo_16"
    '    Me._lblArticulo_16.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_16.Size = New System.Drawing.Size(82, 13)
    '    Me._lblArticulo_16.TabIndex = 80
    '    Me._lblArticulo_16.Text = "Funcionamiento"
    '    '
    '    '_lblArticulo_17
    '    '
    '    Me._lblArticulo_17.AutoSize = True
    '    Me._lblArticulo_17.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_17.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_17.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_17, CType(17, Short))
    '    Me._lblArticulo_17.Location = New System.Drawing.Point(2, 144)
    '    Me._lblArticulo_17.Name = "_lblArticulo_17"
    '    Me._lblArticulo_17.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_17.Size = New System.Drawing.Size(83, 13)
    '    Me._lblArticulo_17.TabIndex = 85
    '    Me._lblArticulo_17.Text = "Tipo de Material"
    '    '
    '    '_lblArticulo_38
    '    '
    '    Me._lblArticulo_38.AutoSize = True
    '    Me._lblArticulo_38.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_38.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_38.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_38, CType(38, Short))
    '    Me._lblArticulo_38.Location = New System.Drawing.Point(313, 270)
    '    Me._lblArticulo_38.Name = "_lblArticulo_38"
    '    Me._lblArticulo_38.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_38.Size = New System.Drawing.Size(85, 13)
    '    Me._lblArticulo_38.TabIndex = 97
    '    Me._lblArticulo_38.Text = "Moneda Compra"
    '    '
    '    '_lblDescripcion_1
    '    '
    '    Me._lblDescripcion_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
    '    Me._lblDescripcion_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    '    Me._lblDescripcion_1.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblDescripcion_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
    '    Me.lblDescripcion.SetIndex(Me._lblDescripcion_1, CType(1, Short))
    '    Me._lblDescripcion_1.Location = New System.Drawing.Point(89, 230)
    '    Me._lblDescripcion_1.Name = "_lblDescripcion_1"
    '    Me._lblDescripcion_1.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblDescripcion_1.Size = New System.Drawing.Size(537, 21)
    '    Me._lblDescripcion_1.TabIndex = 92
    '    '
    '    '_sstArticulo_TabPage2
    '    '
    '    Me._sstArticulo_TabPage2.Controls.Add(Me._fraContenedor_2)
    '    Me._sstArticulo_TabPage2.Location = New System.Drawing.Point(4, 22)
    '    Me._sstArticulo_TabPage2.Name = "_sstArticulo_TabPage2"
    '    Me._sstArticulo_TabPage2.Size = New System.Drawing.Size(700, 527)
    '    Me._sstArticulo_TabPage2.TabIndex = 2
    '    Me._sstArticulo_TabPage2.Text = "Varios"
    '    '
    '    '_fraContenedor_2
    '    '
    '    Me._fraContenedor_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraContenedor_2.Controls.Add(Me._txtAdicional_2)
    '    Me._fraContenedor_2.Controls.Add(Me._fraMoneda_4)
    '    Me._fraContenedor_2.Controls.Add(Me._fraMoneda_2)
    '    Me._fraContenedor_2.Controls.Add(Me._Frame2_3)
    '    Me._fraContenedor_2.Controls.Add(Me._Frame1_3)
    '    Me._fraContenedor_2.Controls.Add(Me._txtDescripcion_2)
    '    Me._fraContenedor_2.Controls.Add(Me._fraImagen_2)
    '    Me._fraContenedor_2.Controls.Add(Me._dbcFamilia_1)
    '    Me._fraContenedor_2.Controls.Add(Me._dbcLinea_1)
    '    Me._fraContenedor_2.Controls.Add(Me._dbcMaterial_2)
    '    Me._fraContenedor_2.Controls.Add(Me._lblArticulo_54)
    '    Me._fraContenedor_2.Controls.Add(Me._lblArticulo_49)
    '    Me._fraContenedor_2.Controls.Add(Me._lblArticulo_28)
    '    Me._fraContenedor_2.Controls.Add(Me._lblArticulo_24)
    '    Me._fraContenedor_2.Controls.Add(Me._lblArticulo_25)
    '    Me._fraContenedor_2.Controls.Add(Me._lblArticulo_30)
    '    Me._fraContenedor_2.Controls.Add(Me._lblArticulo_39)
    '    Me._fraContenedor_2.Controls.Add(Me._lblDescripcion_2)
    '    Me._fraContenedor_2.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._fraContenedor_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraContenedor.SetIndex(Me._fraContenedor_2, CType(2, Short))
    '    Me._fraContenedor_2.Location = New System.Drawing.Point(8, 24)
    '    Me._fraContenedor_2.Name = "_fraContenedor_2"
    '    Me._fraContenedor_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraContenedor_2.Size = New System.Drawing.Size(689, 481)
    '    Me._fraContenedor_2.TabIndex = 127
    '    '
    '    '_txtAdicional_2
    '    '
    '    Me._txtAdicional_2.AcceptsReturn = True
    '    Me._txtAdicional_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
    '    Me._txtAdicional_2.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtAdicional_2.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtAdicional.SetIndex(Me._txtAdicional_2, CType(2, Short))
    '    Me._txtAdicional_2.Location = New System.Drawing.Point(89, 110)
    '    Me._txtAdicional_2.MaxLength = 15
    '    Me._txtAdicional_2.Name = "_txtAdicional_2"
    '    Me._txtAdicional_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtAdicional_2.Size = New System.Drawing.Size(120, 20)
    '    Me._txtAdicional_2.TabIndex = 135
    '    '
    '    '_fraMoneda_4
    '    '
    '    Me._fraMoneda_4.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraMoneda_4.Controls.Add(Me._optMoneda_9)
    '    Me._fraMoneda_4.Controls.Add(Me._optMoneda_8)
    '    Me._fraMoneda_4.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraMoneda.SetIndex(Me._fraMoneda_4, CType(4, Short))
    '    Me._fraMoneda_4.Location = New System.Drawing.Point(89, 256)
    '    Me._fraMoneda_4.Name = "_fraMoneda_4"
    '    Me._fraMoneda_4.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraMoneda_4.Size = New System.Drawing.Size(218, 33)
    '    Me._fraMoneda_4.TabIndex = 140
    '    Me._fraMoneda_4.TabStop = False
    '    '
    '    '_optMoneda_9
    '    '
    '    Me._optMoneda_9.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_9.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_9.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_9, CType(9, Short))
    '    Me._optMoneda_9.Location = New System.Drawing.Point(129, 11)
    '    Me._optMoneda_9.Name = "_optMoneda_9"
    '    Me._optMoneda_9.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_9.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_9.TabIndex = 142
    '    Me._optMoneda_9.TabStop = True
    '    Me._optMoneda_9.Text = "Pesos"
    '    Me._optMoneda_9.UseVisualStyleBackColor = False
    '    '
    '    '_optMoneda_8
    '    '
    '    Me._optMoneda_8.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_8.Checked = True
    '    Me._optMoneda_8.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_8.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_8, CType(8, Short))
    '    Me._optMoneda_8.Location = New System.Drawing.Point(36, 11)
    '    Me._optMoneda_8.Name = "_optMoneda_8"
    '    Me._optMoneda_8.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_8.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_8.TabIndex = 141
    '    Me._optMoneda_8.TabStop = True
    '    Me._optMoneda_8.Tag = "1"
    '    Me._optMoneda_8.Text = "Dólares"
    '    Me._optMoneda_8.UseVisualStyleBackColor = False
    '    '
    '    '_fraMoneda_2
    '    '
    '    Me._fraMoneda_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraMoneda_2.Controls.Add(Me._optMoneda_4)
    '    Me._fraMoneda_2.Controls.Add(Me._optMoneda_5)
    '    Me._fraMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.fraMoneda.SetIndex(Me._fraMoneda_2, CType(2, Short))
    '    Me._fraMoneda_2.Location = New System.Drawing.Point(416, 256)
    '    Me._fraMoneda_2.Name = "_fraMoneda_2"
    '    Me._fraMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraMoneda_2.Size = New System.Drawing.Size(209, 33)
    '    Me._fraMoneda_2.TabIndex = 144
    '    Me._fraMoneda_2.TabStop = False
    '    '
    '    '_optMoneda_4
    '    '
    '    Me._optMoneda_4.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_4.Checked = True
    '    Me._optMoneda_4.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_4.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_4, CType(4, Short))
    '    Me._optMoneda_4.Location = New System.Drawing.Point(34, 11)
    '    Me._optMoneda_4.Name = "_optMoneda_4"
    '    Me._optMoneda_4.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_4.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_4.TabIndex = 145
    '    Me._optMoneda_4.TabStop = True
    '    Me._optMoneda_4.Text = "Dólares"
    '    Me._optMoneda_4.UseVisualStyleBackColor = False
    '    '
    '    '_optMoneda_5
    '    '
    '    Me._optMoneda_5.BackColor = System.Drawing.SystemColors.Control
    '    Me._optMoneda_5.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._optMoneda_5.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.optMoneda.SetIndex(Me._optMoneda_5, CType(5, Short))
    '    Me._optMoneda_5.Location = New System.Drawing.Point(134, 11)
    '    Me._optMoneda_5.Name = "_optMoneda_5"
    '    Me._optMoneda_5.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._optMoneda_5.Size = New System.Drawing.Size(65, 17)
    '    Me._optMoneda_5.TabIndex = 146
    '    Me._optMoneda_5.TabStop = True
    '    Me._optMoneda_5.Text = "Pesos"
    '    Me._optMoneda_5.UseVisualStyleBackColor = False
    '    '
    '    '_Frame2_3
    '    '
    '    Me._Frame2_3.BackColor = System.Drawing.SystemColors.Control
    '    Me._Frame2_3.Controls.Add(Me._Frame4_2)
    '    Me._Frame2_3.Controls.Add(Me._txtCodigodelProveedor_2)
    '    Me._Frame2_3.Controls.Add(Me._dbcProveedor_2)
    '    Me._Frame2_3.Controls.Add(Me._cboUnidad_2)
    '    Me._Frame2_3.Controls.Add(Me._cboAlmacen_2)
    '    Me._Frame2_3.Controls.Add(Me._lblArticulo_50)
    '    Me._Frame2_3.Controls.Add(Me._lblArticulo_51)
    '    Me._Frame2_3.Controls.Add(Me._lblArticulo_52)
    '    Me._Frame2_3.Controls.Add(Me._lblArticulo_53)
    '    Me._Frame2_3.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.Frame2.SetIndex(Me._Frame2_3, CType(3, Short))
    '    Me._Frame2_3.Location = New System.Drawing.Point(312, 294)
    '    Me._Frame2_3.Name = "_Frame2_3"
    '    Me._Frame2_3.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._Frame2_3.Size = New System.Drawing.Size(313, 185)
    '    Me._Frame2_3.TabIndex = 160
    '    Me._Frame2_3.TabStop = False
    '    '
    '    '_Frame4_2
    '    '
    '    Me._Frame4_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._Frame4_2.Controls.Add(Me._txtImagen_2)
    '    Me._Frame4_2.Controls.Add(Me._cmdBuscarImagen_2)
    '    Me._Frame4_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
    '    Me.Frame4.SetIndex(Me._Frame4_2, CType(2, Short))
    '    Me._Frame4_2.Location = New System.Drawing.Point(12, 132)
    '    Me._Frame4_2.Name = "_Frame4_2"
    '    Me._Frame4_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._Frame4_2.Size = New System.Drawing.Size(290, 44)
    '    Me._Frame4_2.TabIndex = 169
    '    Me._Frame4_2.TabStop = False
    '    Me._Frame4_2.Text = "Imagen"
    '    '
    '    '_txtImagen_2
    '    '
    '    Me._txtImagen_2.AcceptsReturn = True
    '    Me._txtImagen_2.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtImagen_2.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtImagen_2.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtImagen.SetIndex(Me._txtImagen_2, CType(2, Short))
    '    Me._txtImagen_2.Location = New System.Drawing.Point(9, 15)
    '    Me._txtImagen_2.MaxLength = 0
    '    Me._txtImagen_2.Name = "_txtImagen_2"
    '    Me._txtImagen_2.ReadOnly = True
    '    Me._txtImagen_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtImagen_2.Size = New System.Drawing.Size(245, 20)
    '    Me._txtImagen_2.TabIndex = 170
    '    '
    '    '_cmdBuscarImagen_2
    '    '
    '    Me._cmdBuscarImagen_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._cmdBuscarImagen_2.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._cmdBuscarImagen_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.cmdBuscarImagen.SetIndex(Me._cmdBuscarImagen_2, CType(2, Short))
    '    Me._cmdBuscarImagen_2.Location = New System.Drawing.Point(260, 15)
    '    Me._cmdBuscarImagen_2.Name = "_cmdBuscarImagen_2"
    '    Me._cmdBuscarImagen_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._cmdBuscarImagen_2.Size = New System.Drawing.Size(22, 21)
    '    Me._cmdBuscarImagen_2.TabIndex = 171
    '    Me._cmdBuscarImagen_2.Text = "..."
    '    Me._cmdBuscarImagen_2.UseVisualStyleBackColor = False
    '    '
    '    '_txtCodigodelProveedor_2
    '    '
    '    Me._txtCodigodelProveedor_2.AcceptsReturn = True
    '    Me._txtCodigodelProveedor_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
    '    Me._txtCodigodelProveedor_2.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCodigodelProveedor_2.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtCodigodelProveedor.SetIndex(Me._txtCodigodelProveedor_2, CType(2, Short))
    '    Me._txtCodigodelProveedor_2.Location = New System.Drawing.Point(171, 102)
    '    Me._txtCodigodelProveedor_2.MaxLength = 20
    '    Me._txtCodigodelProveedor_2.Name = "_txtCodigodelProveedor_2"
    '    Me._txtCodigodelProveedor_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCodigodelProveedor_2.Size = New System.Drawing.Size(129, 20)
    '    Me._txtCodigodelProveedor_2.TabIndex = 168
    '    Me.ToolTip1.SetToolTip(Me._txtCodigodelProveedor_2, "Código que usa el Proveedor para el Artículo")
    '    '
    '    '_dbcProveedor_2
    '    '
    '    Me._dbcProveedor_2.Location = New System.Drawing.Point(100, 74)
    '    Me._dbcProveedor_2.Name = "_dbcProveedor_2"
    '    Me._dbcProveedor_2.Size = New System.Drawing.Size(201, 21)
    '    Me._dbcProveedor_2.TabIndex = 166
    '    '
    '    '_cboUnidad_2
    '    '
    '    Me._cboUnidad_2.Location = New System.Drawing.Point(100, 17)
    '    Me._cboUnidad_2.Name = "_cboUnidad_2"
    '    Me._cboUnidad_2.Size = New System.Drawing.Size(78, 21)
    '    Me._cboUnidad_2.TabIndex = 162
    '    '
    '    '_cboAlmacen_2
    '    '
    '    Me._cboAlmacen_2.Location = New System.Drawing.Point(100, 46)
    '    Me._cboAlmacen_2.Name = "_cboAlmacen_2"
    '    Me._cboAlmacen_2.Size = New System.Drawing.Size(201, 21)
    '    Me._cboAlmacen_2.TabIndex = 164
    '    '
    '    '_lblArticulo_50
    '    '
    '    Me._lblArticulo_50.AutoSize = True
    '    Me._lblArticulo_50.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_50.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_50.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_50, CType(50, Short))
    '    Me._lblArticulo_50.Location = New System.Drawing.Point(12, 79)
    '    Me._lblArticulo_50.Name = "_lblArticulo_50"
    '    Me._lblArticulo_50.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_50.Size = New System.Drawing.Size(56, 13)
    '    Me._lblArticulo_50.TabIndex = 165
    '    Me._lblArticulo_50.Text = "Proveedor"
    '    '
    '    '_lblArticulo_51
    '    '
    '    Me._lblArticulo_51.AutoSize = True
    '    Me._lblArticulo_51.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_51.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_51.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_51, CType(51, Short))
    '    Me._lblArticulo_51.Location = New System.Drawing.Point(19, 106)
    '    Me._lblArticulo_51.Name = "_lblArticulo_51"
    '    Me._lblArticulo_51.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_51.Size = New System.Drawing.Size(156, 13)
    '    Me._lblArticulo_51.TabIndex = 167
    '    Me._lblArticulo_51.Text = "Código artículo del proveedor : "
    '    '
    '    '_lblArticulo_52
    '    '
    '    Me._lblArticulo_52.AutoSize = True
    '    Me._lblArticulo_52.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_52.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_52.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_52, CType(52, Short))
    '    Me._lblArticulo_52.Location = New System.Drawing.Point(12, 21)
    '    Me._lblArticulo_52.Name = "_lblArticulo_52"
    '    Me._lblArticulo_52.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_52.Size = New System.Drawing.Size(41, 13)
    '    Me._lblArticulo_52.TabIndex = 161
    '    Me._lblArticulo_52.Text = "Unidad"
    '    '
    '    '_lblArticulo_53
    '    '
    '    Me._lblArticulo_53.AutoSize = True
    '    Me._lblArticulo_53.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_53.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_53.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_53, CType(53, Short))
    '    Me._lblArticulo_53.Location = New System.Drawing.Point(12, 50)
    '    Me._lblArticulo_53.Name = "_lblArticulo_53"
    '    Me._lblArticulo_53.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_53.Size = New System.Drawing.Size(84, 13)
    '    Me._lblArticulo_53.TabIndex = 163
    '    Me._lblArticulo_53.Text = "Almacén/Origen"
    '    '
    '    '_Frame1_3
    '    '
    '    Me._Frame1_3.BackColor = System.Drawing.SystemColors.Control
    '    Me._Frame1_3.Controls.Add(Me._txtCostoFactura_2)
    '    Me._Frame1_3.Controls.Add(Me._txtCostoAdicional_2)
    '    Me._Frame1_3.Controls.Add(Me._txtCostoIndirecto_2)
    '    Me._Frame1_3.Controls.Add(Me._txtPrecioenDolares_2)
    '    Me._Frame1_3.Controls.Add(Me._txtCostoReal_2)
    '    Me._Frame1_3.Controls.Add(Me._lblMargen_2)
    '    Me._Frame1_3.Controls.Add(Me.Label3)
    '    Me._Frame1_3.Controls.Add(Me._lblArticulo_22)
    '    Me._Frame1_3.Controls.Add(Me._lblArticulo_23)
    '    Me._Frame1_3.Controls.Add(Me._lblArticulo_46)
    '    Me._Frame1_3.Controls.Add(Me._lblArticulo_47)
    '    Me._Frame1_3.Controls.Add(Me._lblArticulo_48)
    '    Me._Frame1_3.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.Frame1.SetIndex(Me._Frame1_3, CType(3, Short))
    '    Me._Frame1_3.Location = New System.Drawing.Point(0, 294)
    '    Me._Frame1_3.Name = "_Frame1_3"
    '    Me._Frame1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._Frame1_3.Size = New System.Drawing.Size(305, 185)
    '    Me._Frame1_3.TabIndex = 147
    '    Me._Frame1_3.TabStop = False
    '    '
    '    '_txtCostoFactura_2
    '    '
    '    Me._txtCostoFactura_2.AcceptsReturn = True
    '    Me._txtCostoFactura_2.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtCostoFactura_2.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoFactura_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.txtCostoFactura.SetIndex(Me._txtCostoFactura_2, CType(2, Short))
    '    Me._txtCostoFactura_2.Location = New System.Drawing.Point(92, 56)
    '    Me._txtCostoFactura_2.MaxLength = 0
    '    Me._txtCostoFactura_2.Name = "_txtCostoFactura_2"
    '    Me._txtCostoFactura_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoFactura_2.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoFactura_2.TabIndex = 151
    '    Me._txtCostoFactura_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoFactura_2, "Costo en Pesos")
    '    '
    '    '_txtCostoAdicional_2
    '    '
    '    Me._txtCostoAdicional_2.AcceptsReturn = True
    '    Me._txtCostoAdicional_2.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtCostoAdicional_2.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoAdicional_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.txtCostoAdicional.SetIndex(Me._txtCostoAdicional_2, CType(2, Short))
    '    Me._txtCostoAdicional_2.Location = New System.Drawing.Point(92, 88)
    '    Me._txtCostoAdicional_2.MaxLength = 0
    '    Me._txtCostoAdicional_2.Name = "_txtCostoAdicional_2"
    '    Me._txtCostoAdicional_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoAdicional_2.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoAdicional_2.TabIndex = 153
    '    Me._txtCostoAdicional_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoAdicional_2, "Costo en Dólares")
    '    '
    '    '_txtCostoIndirecto_2
    '    '
    '    Me._txtCostoIndirecto_2.AcceptsReturn = True
    '    Me._txtCostoIndirecto_2.BackColor = System.Drawing.SystemColors.Window
    '    Me._txtCostoIndirecto_2.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoIndirecto_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.txtCostoIndirecto.SetIndex(Me._txtCostoIndirecto_2, CType(2, Short))
    '    Me._txtCostoIndirecto_2.Location = New System.Drawing.Point(92, 120)
    '    Me._txtCostoIndirecto_2.MaxLength = 0
    '    Me._txtCostoIndirecto_2.Name = "_txtCostoIndirecto_2"
    '    Me._txtCostoIndirecto_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoIndirecto_2.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoIndirecto_2.TabIndex = 155
    '    Me._txtCostoIndirecto_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoIndirecto_2, "Gastos Indirectos en Dólares")
    '    '
    '    '_txtPrecioenDolares_2
    '    '
    '    Me._txtPrecioenDolares_2.AcceptsReturn = True
    '    Me._txtPrecioenDolares_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
    '    Me._txtPrecioenDolares_2.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtPrecioenDolares_2.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtPrecioenDolares.SetIndex(Me._txtPrecioenDolares_2, CType(2, Short))
    '    Me._txtPrecioenDolares_2.Location = New System.Drawing.Point(92, 24)
    '    Me._txtPrecioenDolares_2.MaxLength = 0
    '    Me._txtPrecioenDolares_2.Name = "_txtPrecioenDolares_2"
    '    Me._txtPrecioenDolares_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtPrecioenDolares_2.Size = New System.Drawing.Size(113, 20)
    '    Me._txtPrecioenDolares_2.TabIndex = 149
    '    Me._txtPrecioenDolares_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtPrecioenDolares_2, "Precio al Público en Dólares")
    '    '
    '    '_txtCostoReal_2
    '    '
    '    Me._txtCostoReal_2.AcceptsReturn = True
    '    Me._txtCostoReal_2.BackColor = System.Drawing.SystemColors.Info
    '    Me._txtCostoReal_2.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtCostoReal_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
    '    Me.txtCostoReal.SetIndex(Me._txtCostoReal_2, CType(2, Short))
    '    Me._txtCostoReal_2.Location = New System.Drawing.Point(92, 152)
    '    Me._txtCostoReal_2.MaxLength = 0
    '    Me._txtCostoReal_2.Name = "_txtCostoReal_2"
    '    Me._txtCostoReal_2.ReadOnly = True
    '    Me._txtCostoReal_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtCostoReal_2.Size = New System.Drawing.Size(113, 20)
    '    Me._txtCostoReal_2.TabIndex = 157
    '    Me._txtCostoReal_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    Me.ToolTip1.SetToolTip(Me._txtCostoReal_2, "Costo Real del artículo")
    '    '
    '    '_lblMargen_2
    '    '
    '    Me._lblMargen_2.BackColor = System.Drawing.SystemColors.Window
    '    Me._lblMargen_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    '    Me._lblMargen_2.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblMargen_2.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblMargen.SetIndex(Me._lblMargen_2, CType(2, Short))
    '    Me._lblMargen_2.Location = New System.Drawing.Point(244, 152)
    '    Me._lblMargen_2.Name = "_lblMargen_2"
    '    Me._lblMargen_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblMargen_2.Size = New System.Drawing.Size(49, 21)
    '    Me._lblMargen_2.TabIndex = 159
    '    Me._lblMargen_2.TextAlign = System.Drawing.ContentAlignment.TopRight
    '    '
    '    'Label3
    '    '
    '    Me.Label3.BackColor = System.Drawing.SystemColors.Control
    '    Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.Label3.Location = New System.Drawing.Point(232, 120)
    '    Me.Label3.Name = "Label3"
    '    Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.Label3.Size = New System.Drawing.Size(61, 29)
    '    Me.Label3.TabIndex = 158
    '    Me.Label3.Text = "% Margen de Venta "
    '    Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
    '    '
    '    '_lblArticulo_22
    '    '
    '    Me._lblArticulo_22.AutoSize = True
    '    Me._lblArticulo_22.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_22.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_22.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_22, CType(22, Short))
    '    Me._lblArticulo_22.Location = New System.Drawing.Point(12, 60)
    '    Me._lblArticulo_22.Name = "_lblArticulo_22"
    '    Me._lblArticulo_22.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_22.Size = New System.Drawing.Size(73, 13)
    '    Me._lblArticulo_22.TabIndex = 150
    '    Me._lblArticulo_22.Text = "Costo Factura"
    '    '
    '    '_lblArticulo_23
    '    '
    '    Me._lblArticulo_23.AutoSize = True
    '    Me._lblArticulo_23.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_23.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_23.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_23, CType(23, Short))
    '    Me._lblArticulo_23.Location = New System.Drawing.Point(12, 92)
    '    Me._lblArticulo_23.Name = "_lblArticulo_23"
    '    Me._lblArticulo_23.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_23.Size = New System.Drawing.Size(80, 13)
    '    Me._lblArticulo_23.TabIndex = 152
    '    Me._lblArticulo_23.Text = "Costo Adicional"
    '    '
    '    '_lblArticulo_46
    '    '
    '    Me._lblArticulo_46.AutoSize = True
    '    Me._lblArticulo_46.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_46.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_46.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_46, CType(46, Short))
    '    Me._lblArticulo_46.Location = New System.Drawing.Point(12, 124)
    '    Me._lblArticulo_46.Name = "_lblArticulo_46"
    '    Me._lblArticulo_46.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_46.Size = New System.Drawing.Size(78, 13)
    '    Me._lblArticulo_46.TabIndex = 154
    '    Me._lblArticulo_46.Text = "Costo Indirecto"
    '    '
    '    '_lblArticulo_47
    '    '
    '    Me._lblArticulo_47.AutoSize = True
    '    Me._lblArticulo_47.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_47.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_47.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_47, CType(47, Short))
    '    Me._lblArticulo_47.Location = New System.Drawing.Point(12, 28)
    '    Me._lblArticulo_47.Name = "_lblArticulo_47"
    '    Me._lblArticulo_47.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_47.Size = New System.Drawing.Size(78, 13)
    '    Me._lblArticulo_47.TabIndex = 148
    '    Me._lblArticulo_47.Text = "Precio Público "
    '    '
    '    '_lblArticulo_48
    '    '
    '    Me._lblArticulo_48.AutoSize = True
    '    Me._lblArticulo_48.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_48.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_48.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_48, CType(48, Short))
    '    Me._lblArticulo_48.Location = New System.Drawing.Point(12, 156)
    '    Me._lblArticulo_48.Name = "_lblArticulo_48"
    '    Me._lblArticulo_48.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_48.Size = New System.Drawing.Size(59, 13)
    '    Me._lblArticulo_48.TabIndex = 156
    '    Me._lblArticulo_48.Text = "Costo Real"
    '    '
    '    '_txtDescripcion_2
    '    '
    '    Me._txtDescripcion_2.AcceptsReturn = True
    '    Me._txtDescripcion_2.BackColor = System.Drawing.SystemColors.Info
    '    Me._txtDescripcion_2.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me._txtDescripcion_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
    '    Me.txtDescripcion.SetIndex(Me._txtDescripcion_2, CType(2, Short))
    '    Me._txtDescripcion_2.Location = New System.Drawing.Point(89, 198)
    '    Me._txtDescripcion_2.MaxLength = 0
    '    Me._txtDescripcion_2.Name = "_txtDescripcion_2"
    '    Me._txtDescripcion_2.ReadOnly = True
    '    Me._txtDescripcion_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._txtDescripcion_2.Size = New System.Drawing.Size(537, 20)
    '    Me._txtDescripcion_2.TabIndex = 137
    '    Me.ToolTip1.SetToolTip(Me._txtDescripcion_2, "Descripción")
    '    '
    '    '_fraImagen_2
    '    '
    '    Me._fraImagen_2.BackColor = System.Drawing.SystemColors.Control
    '    Me._fraImagen_2.Controls.Add(Me.Image3)
    '    Me._fraImagen_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
    '    Me.fraImagen.SetIndex(Me._fraImagen_2, CType(2, Short))
    '    Me._fraImagen_2.Location = New System.Drawing.Point(510, 8)
    '    Me._fraImagen_2.Name = "_fraImagen_2"
    '    Me._fraImagen_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._fraImagen_2.Size = New System.Drawing.Size(178, 186)
    '    Me._fraImagen_2.TabIndex = 172
    '    Me._fraImagen_2.TabStop = False
    '    Me._fraImagen_2.Text = "Imagen del Artículo"
    '    '
    '    'Image3
    '    '
    '    Me.Image3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    '    Me.Image3.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.Image3.Location = New System.Drawing.Point(7, 21)
    '    Me.Image3.Name = "Image3"
    '    Me.Image3.Size = New System.Drawing.Size(163, 157)
    '    Me.Image3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
    '    Me.Image3.TabIndex = 0
    '    Me.Image3.TabStop = False
    '    '
    '    '_dbcFamilia_1
    '    '
    '    Me._dbcFamilia_1.Location = New System.Drawing.Point(89, 14)
    '    Me._dbcFamilia_1.Name = "_dbcFamilia_1"
    '    Me._dbcFamilia_1.Size = New System.Drawing.Size(265, 21)
    '    Me._dbcFamilia_1.TabIndex = 129
    '    '
    '    '_dbcLinea_1
    '    '
    '    Me._dbcLinea_1.Location = New System.Drawing.Point(89, 46)
    '    Me._dbcLinea_1.Name = "_dbcLinea_1"
    '    Me._dbcLinea_1.Size = New System.Drawing.Size(265, 21)
    '    Me._dbcLinea_1.TabIndex = 131
    '    '
    '    '_dbcMaterial_2
    '    '
    '    Me._dbcMaterial_2.Location = New System.Drawing.Point(89, 78)
    '    Me._dbcMaterial_2.Name = "_dbcMaterial_2"
    '    Me._dbcMaterial_2.Size = New System.Drawing.Size(265, 21)
    '    Me._dbcMaterial_2.TabIndex = 133
    '    '
    '    '_lblArticulo_54
    '    '
    '    Me._lblArticulo_54.AutoSize = True
    '    Me._lblArticulo_54.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_54.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_54.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_54, CType(54, Short))
    '    Me._lblArticulo_54.Location = New System.Drawing.Point(2, 114)
    '    Me._lblArticulo_54.Name = "_lblArticulo_54"
    '    Me._lblArticulo_54.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_54.Size = New System.Drawing.Size(76, 13)
    '    Me._lblArticulo_54.TabIndex = 134
    '    Me._lblArticulo_54.Text = "Dato Adicional"
    '    '
    '    '_lblArticulo_49
    '    '
    '    Me._lblArticulo_49.AutoSize = True
    '    Me._lblArticulo_49.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_49.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_49.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_49, CType(49, Short))
    '    Me._lblArticulo_49.Location = New System.Drawing.Point(2, 82)
    '    Me._lblArticulo_49.Name = "_lblArticulo_49"
    '    Me._lblArticulo_49.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_49.Size = New System.Drawing.Size(83, 13)
    '    Me._lblArticulo_49.TabIndex = 132
    '    Me._lblArticulo_49.Text = "Tipo de Material"
    '    '
    '    '_lblArticulo_28
    '    '
    '    Me._lblArticulo_28.AutoSize = True
    '    Me._lblArticulo_28.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_28.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_28.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_28, CType(28, Short))
    '    Me._lblArticulo_28.Location = New System.Drawing.Point(0, 270)
    '    Me._lblArticulo_28.Name = "_lblArticulo_28"
    '    Me._lblArticulo_28.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_28.Size = New System.Drawing.Size(92, 13)
    '    Me._lblArticulo_28.TabIndex = 139
    '    Me._lblArticulo_28.Text = "Precio público en "
    '    '
    '    '_lblArticulo_24
    '    '
    '    Me._lblArticulo_24.AutoSize = True
    '    Me._lblArticulo_24.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_24.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_24.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_24, CType(24, Short))
    '    Me._lblArticulo_24.Location = New System.Drawing.Point(2, 18)
    '    Me._lblArticulo_24.Name = "_lblArticulo_24"
    '    Me._lblArticulo_24.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_24.Size = New System.Drawing.Size(39, 13)
    '    Me._lblArticulo_24.TabIndex = 128
    '    Me._lblArticulo_24.Text = "Familia"
    '    '
    '    '_lblArticulo_25
    '    '
    '    Me._lblArticulo_25.AutoSize = True
    '    Me._lblArticulo_25.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_25.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_25.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_25, CType(25, Short))
    '    Me._lblArticulo_25.Location = New System.Drawing.Point(2, 50)
    '    Me._lblArticulo_25.Name = "_lblArticulo_25"
    '    Me._lblArticulo_25.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_25.Size = New System.Drawing.Size(35, 13)
    '    Me._lblArticulo_25.TabIndex = 130
    '    Me._lblArticulo_25.Text = "Línea"
    '    '
    '    '_lblArticulo_30
    '    '
    '    Me._lblArticulo_30.AutoSize = True
    '    Me._lblArticulo_30.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_30.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_30.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_30, CType(30, Short))
    '    Me._lblArticulo_30.Location = New System.Drawing.Point(2, 202)
    '    Me._lblArticulo_30.Name = "_lblArticulo_30"
    '    Me._lblArticulo_30.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_30.Size = New System.Drawing.Size(63, 13)
    '    Me._lblArticulo_30.TabIndex = 136
    '    Me._lblArticulo_30.Text = "Descripción"
    '    '
    '    '_lblArticulo_39
    '    '
    '    Me._lblArticulo_39.AutoSize = True
    '    Me._lblArticulo_39.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_39.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_39.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_39, CType(39, Short))
    '    Me._lblArticulo_39.Location = New System.Drawing.Point(313, 270)
    '    Me._lblArticulo_39.Name = "_lblArticulo_39"
    '    Me._lblArticulo_39.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_39.Size = New System.Drawing.Size(85, 13)
    '    Me._lblArticulo_39.TabIndex = 143
    '    Me._lblArticulo_39.Text = "Moneda Compra"
    '    '
    '    '_lblDescripcion_2
    '    '
    '    Me._lblDescripcion_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
    '    Me._lblDescripcion_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    '    Me._lblDescripcion_2.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblDescripcion_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
    '    Me.lblDescripcion.SetIndex(Me._lblDescripcion_2, CType(2, Short))
    '    Me._lblDescripcion_2.Location = New System.Drawing.Point(89, 230)
    '    Me._lblDescripcion_2.Name = "_lblDescripcion_2"
    '    Me._lblDescripcion_2.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblDescripcion_2.Size = New System.Drawing.Size(537, 21)
    '    Me._lblDescripcion_2.TabIndex = 138
    '    '
    '    'chkCodigoAnterior
    '    '
    '    Me.chkCodigoAnterior.BackColor = System.Drawing.SystemColors.Control
    '    Me.chkCodigoAnterior.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.chkCodigoAnterior.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.chkCodigoAnterior.Location = New System.Drawing.Point(6, 0)
    '    Me.chkCodigoAnterior.Name = "chkCodigoAnterior"
    '    Me.chkCodigoAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.chkCodigoAnterior.Size = New System.Drawing.Size(17, 17)
    '    Me.chkCodigoAnterior.TabIndex = 4
    '    Me.chkCodigoAnterior.Text = "chkCodAnterior"
    '    Me.chkCodigoAnterior.UseVisualStyleBackColor = False
    '    '
    '    'Frame3
    '    '
    '    Me.Frame3.BackColor = System.Drawing.SystemColors.Control
    '    Me.Frame3.Controls.Add(Me.chkCodigoAnterior)
    '    Me.Frame3.Controls.Add(Me.txtCodArtAnterior)
    '    Me.Frame3.Controls.Add(Me.dbcOrigen)
    '    Me.Frame3.Controls.Add(Me._lblArticulo_32)
    '    Me.Frame3.Controls.Add(Me._lblArticulo_31)
    '    Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
    '    Me.Frame3.Location = New System.Drawing.Point(538, 3)
    '    Me.Frame3.Name = "Frame3"
    '    Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.Frame3.Size = New System.Drawing.Size(182, 69)
    '    Me.Frame3.TabIndex = 3
    '    Me.Frame3.TabStop = False
    '    Me.Frame3.Text = "      Código anterior "
    '    '
    '    'txtCodArtAnterior
    '    '
    '    Me.txtCodArtAnterior.AcceptsReturn = True
    '    Me.txtCodArtAnterior.BackColor = System.Drawing.SystemColors.Window
    '    Me.txtCodArtAnterior.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me.txtCodArtAnterior.Enabled = False
    '    Me.txtCodArtAnterior.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtCodArtAnterior.Location = New System.Drawing.Point(76, 41)
    '    Me.txtCodArtAnterior.MaxLength = 5
    '    Me.txtCodArtAnterior.Name = "txtCodArtAnterior"
    '    Me.txtCodArtAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.txtCodArtAnterior.Size = New System.Drawing.Size(88, 20)
    '    Me.txtCodArtAnterior.TabIndex = 8
    '    Me.txtCodArtAnterior.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    '
    '    'dbcOrigen
    '    '
    '    Me.dbcOrigen.Location = New System.Drawing.Point(76, 17)
    '    Me.dbcOrigen.Name = "dbcOrigen"
    '    Me.dbcOrigen.Size = New System.Drawing.Size(88, 21)
    '    Me.dbcOrigen.TabIndex = 6
    '    '
    '    '_lblArticulo_32
    '    '
    '    Me._lblArticulo_32.AutoSize = True
    '    Me._lblArticulo_32.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_32.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_32.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_32, CType(32, Short))
    '    Me._lblArticulo_32.Location = New System.Drawing.Point(31, 44)
    '    Me._lblArticulo_32.Name = "_lblArticulo_32"
    '    Me._lblArticulo_32.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_32.Size = New System.Drawing.Size(46, 13)
    '    Me._lblArticulo_32.TabIndex = 7
    '    Me._lblArticulo_32.Text = "Código: "
    '    '
    '    '_lblArticulo_31
    '    '
    '    Me._lblArticulo_31.AutoSize = True
    '    Me._lblArticulo_31.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_31.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_31.ForeColor = System.Drawing.SystemColors.ControlText
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_31, CType(31, Short))
    '    Me._lblArticulo_31.Location = New System.Drawing.Point(30, 21)
    '    Me._lblArticulo_31.Name = "_lblArticulo_31"
    '    Me._lblArticulo_31.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_31.Size = New System.Drawing.Size(47, 13)
    '    Me._lblArticulo_31.TabIndex = 5
    '    Me._lblArticulo_31.Text = "Origen : "
    '    '
    '    'txtCodArticulo
    '    '
    '    Me.txtCodArticulo.AcceptsReturn = True
    '    Me.txtCodArticulo.BackColor = System.Drawing.SystemColors.Window
    '    Me.txtCodArticulo.Cursor = System.Windows.Forms.Cursors.IBeam
    '    Me.txtCodArticulo.ForeColor = System.Drawing.SystemColors.WindowText
    '    Me.txtCodArticulo.Location = New System.Drawing.Point(72, 8)
    '    Me.txtCodArticulo.MaxLength = 8
    '    Me.txtCodArticulo.Name = "txtCodArticulo"
    '    Me.txtCodArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.txtCodArticulo.Size = New System.Drawing.Size(89, 20)
    '    Me.txtCodArticulo.TabIndex = 1
    '    Me.txtCodArticulo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '    '
    '    '_lblArticulo_0
    '    '
    '    Me._lblArticulo_0.AutoSize = True
    '    Me._lblArticulo_0.BackColor = System.Drawing.SystemColors.Control
    '    Me._lblArticulo_0.Cursor = System.Windows.Forms.Cursors.Default
    '    Me._lblArticulo_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
    '    Me.lblArticulo.SetIndex(Me._lblArticulo_0, CType(0, Short))
    '    Me._lblArticulo_0.Location = New System.Drawing.Point(16, 12)
    '    Me._lblArticulo_0.Name = "_lblArticulo_0"
    '    Me._lblArticulo_0.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me._lblArticulo_0.Size = New System.Drawing.Size(44, 13)
    '    Me._lblArticulo_0.TabIndex = 0
    '    Me._lblArticulo_0.Text = "Artículo"
    '    '
    '    'cboAlmacen
    '    '
    '    Me.cboAlmacen.Location = New System.Drawing.Point(0, 0)
    '    Me.cboAlmacen.Name = "cboAlmacen"
    '    Me.cboAlmacen.Size = New System.Drawing.Size(121, 21)
    '    Me.cboAlmacen.TabIndex = 0
    '    '
    '    'cboUnidad
    '    '
    '    Me.cboUnidad.Location = New System.Drawing.Point(0, 0)
    '    Me.cboUnidad.Name = "cboUnidad"
    '    Me.cboUnidad.Size = New System.Drawing.Size(121, 21)
    '    Me.cboUnidad.TabIndex = 0
    '    '
    '    'cmdBuscarImagen
    '    '
    '    '
    '    'dbcFamilia
    '    '
    '    Me.dbcFamilia.Location = New System.Drawing.Point(0, 0)
    '    Me.dbcFamilia.Name = "dbcFamilia"
    '    Me.dbcFamilia.Size = New System.Drawing.Size(121, 21)
    '    Me.dbcFamilia.TabIndex = 0
    '    '
    '    'dbcLinea
    '    '
    '    Me.dbcLinea.Location = New System.Drawing.Point(0, 0)
    '    Me.dbcLinea.Name = "dbcLinea"
    '    Me.dbcLinea.Size = New System.Drawing.Size(121, 21)
    '    Me.dbcLinea.TabIndex = 0
    '    '
    '    'dbcMaterial
    '    '
    '    Me.dbcMaterial.Location = New System.Drawing.Point(0, 0)
    '    Me.dbcMaterial.Name = "dbcMaterial"
    '    Me.dbcMaterial.Size = New System.Drawing.Size(121, 21)
    '    Me.dbcMaterial.TabIndex = 0
    '    '
    '    'dbcProveedor
    '    '
    '    Me.dbcProveedor.Location = New System.Drawing.Point(0, 0)
    '    Me.dbcProveedor.Name = "dbcProveedor"
    '    Me.dbcProveedor.Size = New System.Drawing.Size(121, 21)
    '    Me.dbcProveedor.TabIndex = 0
    '    '
    '    'optGenero
    '    '
    '    '
    '    'optMoneda
    '    '
    '    '
    '    'optMovimiento
    '    '
    '    '
    '    'txtAdicional
    '    '
    '    '
    '    'txtCodigodelProveedor
    '    '
    '    '
    '    'txtCostoAdicional
    '    '
    '    '
    '    'txtCostoFactura
    '    '
    '    '
    '    'txtCostoIndirecto
    '    '
    '    '
    '    'txtCostoReal
    '    '
    '    '
    '    'txtDescripcion
    '    '
    '    '
    '    'txtPrecioenDolares
    '    '
    '    '
    '    'btnLimpiar
    '    '
    '    Me.btnLimpiar.Location = New System.Drawing.Point(416, 637)
    '    Me.btnLimpiar.Name = "btnLimpiar"
    '    Me.btnLimpiar.Size = New System.Drawing.Size(93, 35)
    '    Me.btnLimpiar.TabIndex = 12
    '    Me.btnLimpiar.Text = "Limpiar"
    '    Me.btnLimpiar.UseVisualStyleBackColor = True
    '    '
    '    'btnEliminar
    '    '
    '    Me.btnEliminar.Location = New System.Drawing.Point(309, 637)
    '    Me.btnEliminar.Name = "btnEliminar"
    '    Me.btnEliminar.Size = New System.Drawing.Size(93, 35)
    '    Me.btnEliminar.TabIndex = 11
    '    Me.btnEliminar.Text = "Eliminar"
    '    Me.btnEliminar.UseVisualStyleBackColor = True
    '    '
    '    'btnGuardar
    '    '
    '    Me.btnGuardar.Location = New System.Drawing.Point(201, 637)
    '    Me.btnGuardar.Name = "btnGuardar"
    '    Me.btnGuardar.Size = New System.Drawing.Size(93, 35)
    '    Me.btnGuardar.TabIndex = 10
    '    Me.btnGuardar.Text = "Guardar"
    '    Me.btnGuardar.UseVisualStyleBackColor = True
    '    '
    '    'frmCorpoABCArticulos
    '    '
    '    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    '    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    '    Me.BackColor = System.Drawing.SystemColors.Control
    '    Me.ClientSize = New System.Drawing.Size(727, 684)
    '    Me.Controls.Add(Me.btnLimpiar)
    '    Me.Controls.Add(Me.btnEliminar)
    '    Me.Controls.Add(Me.btnGuardar)
    '    Me.Controls.Add(Me.Frame3)
    '    Me.Controls.Add(Me.txtDescArticulo)
    '    Me.Controls.Add(Me.txtCodArticulo)
    '    Me.Controls.Add(Me.sstArticulo)
    '    Me.Controls.Add(Me._lblArticulo_0)
    '    Me.Cursor = System.Windows.Forms.Cursors.Default
    '    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
    '    Me.KeyPreview = True
    '    Me.Location = New System.Drawing.Point(298, 150)
    '    Me.MaximizeBox = False
    '    Me.Name = "frmCorpoABCArticulos"
    '    Me.RightToLeft = System.Windows.Forms.RightToLeft.No
    '    Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
    '    Me.Text = "ABC a Artículos"
    '    Me.sstArticulo.ResumeLayout(False)
    '    Me._sstArticulo_TabPage0.ResumeLayout(False)
    '    Me._fraContenedor_0.ResumeLayout(False)
    '    Me._fraContenedor_0.PerformLayout()
    '    Me.fraDiamanteSuelto.ResumeLayout(False)
    '    Me.fraDiamanteSuelto.PerformLayout()
    '    Me._fraMoneda_5.ResumeLayout(False)
    '    Me._fraMoneda_0.ResumeLayout(False)
    '    Me._fraImagen_0.ResumeLayout(False)
    '    CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
    '    Me._Frame2_0.ResumeLayout(False)
    '    Me._Frame2_0.PerformLayout()
    '    Me._Frame4_0.ResumeLayout(False)
    '    Me._Frame4_0.PerformLayout()
    '    Me._Frame1_0.ResumeLayout(False)
    '    Me._Frame1_0.PerformLayout()
    '    Me._sstArticulo_TabPage1.ResumeLayout(False)
    '    Me._fraContenedor_1.ResumeLayout(False)
    '    Me._fraContenedor_1.PerformLayout()
    '    Me._fraMoneda_3.ResumeLayout(False)
    '    Me._fraArticulo_1.ResumeLayout(False)
    '    Me._fraArticulo_2.ResumeLayout(False)
    '    Me._fraImagen_1.ResumeLayout(False)
    '    CType(Me.Image2, System.ComponentModel.ISupportInitialize).EndInit()
    '    Me._fraMoneda_1.ResumeLayout(False)
    '    Me._Frame1_2.ResumeLayout(False)
    '    Me._Frame1_2.PerformLayout()
    '    Me._Frame2_2.ResumeLayout(False)
    '    Me._Frame2_2.PerformLayout()
    '    Me._Frame4_1.ResumeLayout(False)
    '    Me._Frame4_1.PerformLayout()
    '    Me._sstArticulo_TabPage2.ResumeLayout(False)
    '    Me._fraContenedor_2.ResumeLayout(False)
    '    Me._fraContenedor_2.PerformLayout()
    '    Me._fraMoneda_4.ResumeLayout(False)
    '    Me._fraMoneda_2.ResumeLayout(False)
    '    Me._Frame2_3.ResumeLayout(False)
    '    Me._Frame2_3.PerformLayout()
    '    Me._Frame4_2.ResumeLayout(False)
    '    Me._Frame4_2.PerformLayout()
    '    Me._Frame1_3.ResumeLayout(False)
    '    Me._Frame1_3.PerformLayout()
    '    Me._fraImagen_2.ResumeLayout(False)
    '    CType(Me.Image3, System.ComponentModel.ISupportInitialize).EndInit()
    '    Me.Frame3.ResumeLayout(False)
    '    Me.Frame3.PerformLayout()
    '    CType(Me.Frame1, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.Frame2, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.Frame4, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.cmdBuscarImagen, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.fraArticulo, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.fraContenedor, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.fraImagen, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.fraMoneda, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.lblArticulo, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.lblDescripcion, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.lblMargen, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.optGenero, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.optMovimiento, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.txtAdicional, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.txtCodigodelProveedor, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.txtCostoAdicional, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.txtCostoFactura, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.txtCostoIndirecto, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.txtCostoReal, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.txtDescripcion, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.txtImagen, System.ComponentModel.ISupportInitialize).EndInit()
    '    CType(Me.txtPrecioenDolares, System.ComponentModel.ISupportInitialize).EndInit()
    '    Me.ResumeLayout(False)
    '    Me.PerformLayout()

   ' End Sub

#End Region
End Class
