Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVerificadorPrecios
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents msgExistencia As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents msgPromociones As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents _Shape1_2 As System.Windows.Forms.Label
    Public WithEvents _Shape1_10 As System.Windows.Forms.Label
    Public WithEvents lblNoExisteProm As System.Windows.Forms.Label
    Public WithEvents spNoExisteProm As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_10 As System.Windows.Forms.Label
    Public WithEvents lblPrecioPesos As System.Windows.Forms.Label
    Public WithEvents lblPrecioDolares As System.Windows.Forms.Label
    Public WithEvents lblDescuento As System.Windows.Forms.Label
    Public WithEvents lblFechaIncio As System.Windows.Forms.Label
    Public WithEvents lblFechaFin As System.Windows.Forms.Label
    Public WithEvents lblTipoCambioDolar As System.Windows.Forms.Label
    Public WithEvents _Shape1_1 As System.Windows.Forms.Label
    Public WithEvents _Shape1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Shape1_6 As System.Windows.Forms.Label
    Public WithEvents _Shape1_3 As System.Windows.Forms.Label
    Public WithEvents _Shape1_5 As System.Windows.Forms.Label
    Public WithEvents _Shape1_4 As System.Windows.Forms.Label
    Public WithEvents _Label1_11 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_14 As System.Windows.Forms.Label
    Public WithEvents _Shape1_9 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.Panel
    Public WithEvents txtCodigoAnterior As System.Windows.Forms.TextBox
    Public WithEvents txtDescArticulo As System.Windows.Forms.TextBox
    Public WithEvents txtCodArticulo As System.Windows.Forms.TextBox
    Public WithEvents _Label1_4 As System.Windows.Forms.Label
    Public WithEvents _Label1_18 As System.Windows.Forms.Label
    Public WithEvents _Label1_17 As System.Windows.Forms.Label
    Public WithEvents imgImagenArticulo As System.Windows.Forms.PictureBox
    Public WithEvents Frame1 As System.Windows.Forms.Panel
    Public WithEvents _Shape1_8 As System.Windows.Forms.Label
    Public WithEvents Marco As System.Windows.Forms.Panel
    Public WithEvents _Shape1_7 As System.Windows.Forms.Label
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents Shape1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Dim FueraChange As Boolean
    Dim intCodArticulo As Integer
    Dim tecla As Integer

    Const C_ColSUCURSAL As Integer = 1
    Const C_ColEXISTENCIA As Integer = 2
    Const C_ColCODSUCURSAL As Integer = 3
    Const C_ColAPARTADOS As Integer = 0

    Const C_ColPROMOCION As Integer = 0
    Const C_COLDESDE As Integer = 1
    Const C_COLHASTA As Integer = 2
    Public WithEvents btnNuevo As Button
    Public WithEvents btnBuscar As Button
    Const C_ColDESCUENTO As Integer = 3
    Public strControlActual As String 'Nombre del control actual

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVerificadorPrecios))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCodigoAnterior = New System.Windows.Forms.TextBox()
        Me.txtCodArticulo = New System.Windows.Forms.TextBox()
        Me.Marco = New System.Windows.Forms.Panel()
        Me.Frame2 = New System.Windows.Forms.Panel()
        Me.msgExistencia = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.msgPromociones = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._Shape1_2 = New System.Windows.Forms.Label()
        Me._Shape1_10 = New System.Windows.Forms.Label()
        Me.lblNoExisteProm = New System.Windows.Forms.Label()
        Me.spNoExisteProm = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_10 = New System.Windows.Forms.Label()
        Me.lblPrecioPesos = New System.Windows.Forms.Label()
        Me.lblPrecioDolares = New System.Windows.Forms.Label()
        Me.lblDescuento = New System.Windows.Forms.Label()
        Me.lblFechaIncio = New System.Windows.Forms.Label()
        Me.lblFechaFin = New System.Windows.Forms.Label()
        Me.lblTipoCambioDolar = New System.Windows.Forms.Label()
        Me._Shape1_1 = New System.Windows.Forms.Label()
        Me._Shape1_0 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Shape1_6 = New System.Windows.Forms.Label()
        Me._Shape1_3 = New System.Windows.Forms.Label()
        Me._Shape1_5 = New System.Windows.Forms.Label()
        Me._Shape1_4 = New System.Windows.Forms.Label()
        Me._Label1_11 = New System.Windows.Forms.Label()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_14 = New System.Windows.Forms.Label()
        Me._Shape1_9 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.txtDescArticulo = New System.Windows.Forms.TextBox()
        Me._Label1_4 = New System.Windows.Forms.Label()
        Me._Label1_18 = New System.Windows.Forms.Label()
        Me._Label1_17 = New System.Windows.Forms.Label()
        Me.imgImagenArticulo = New System.Windows.Forms.PictureBox()
        Me._Shape1_8 = New System.Windows.Forms.Label()
        Me._Shape1_7 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Shape1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Marco.SuspendLayout()
        Me.Frame2.SuspendLayout()
        CType(Me.msgExistencia, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.msgPromociones, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame1.SuspendLayout()
        CType(Me.imgImagenArticulo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Shape1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCodigoAnterior
        '
        Me.txtCodigoAnterior.AcceptsReturn = True
        Me.txtCodigoAnterior.BackColor = System.Drawing.Color.White
        Me.txtCodigoAnterior.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodigoAnterior.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodigoAnterior.Location = New System.Drawing.Point(125, 32)
        Me.txtCodigoAnterior.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodigoAnterior.MaxLength = 19
        Me.txtCodigoAnterior.Name = "txtCodigoAnterior"
        Me.txtCodigoAnterior.ReadOnly = True
        Me.txtCodigoAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodigoAnterior.Size = New System.Drawing.Size(84, 20)
        Me.txtCodigoAnterior.TabIndex = 24
        Me.txtCodigoAnterior.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCodigoAnterior, "Folio de Artículo")
        '
        'txtCodArticulo
        '
        Me.txtCodArticulo.AcceptsReturn = True
        Me.txtCodArticulo.BackColor = System.Drawing.Color.White
        Me.txtCodArticulo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodArticulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodArticulo.Location = New System.Drawing.Point(12, 32)
        Me.txtCodArticulo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodArticulo.MaxLength = 19
        Me.txtCodArticulo.Name = "txtCodArticulo"
        Me.txtCodArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodArticulo.Size = New System.Drawing.Size(84, 20)
        Me.txtCodArticulo.TabIndex = 1
        Me.txtCodArticulo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCodArticulo, "Folio de Artículo")
        '
        'Marco
        '
        Me.Marco.BackColor = System.Drawing.Color.White
        Me.Marco.Controls.Add(Me.Frame2)
        Me.Marco.Controls.Add(Me.Frame1)
        Me.Marco.Controls.Add(Me._Shape1_8)
        Me.Marco.Cursor = System.Windows.Forms.Cursors.Default
        Me.Marco.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Marco.Location = New System.Drawing.Point(6, 6)
        Me.Marco.Margin = New System.Windows.Forms.Padding(2)
        Me.Marco.Name = "Marco"
        Me.Marco.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Marco.Size = New System.Drawing.Size(632, 391)
        Me.Marco.TabIndex = 6
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.Frame2.Controls.Add(Me.msgExistencia)
        Me.Frame2.Controls.Add(Me.msgPromociones)
        Me.Frame2.Controls.Add(Me._Shape1_2)
        Me.Frame2.Controls.Add(Me._Shape1_10)
        Me.Frame2.Controls.Add(Me.lblNoExisteProm)
        Me.Frame2.Controls.Add(Me.spNoExisteProm)
        Me.Frame2.Controls.Add(Me._Label1_1)
        Me.Frame2.Controls.Add(Me._Label1_10)
        Me.Frame2.Controls.Add(Me.lblPrecioPesos)
        Me.Frame2.Controls.Add(Me.lblPrecioDolares)
        Me.Frame2.Controls.Add(Me.lblDescuento)
        Me.Frame2.Controls.Add(Me.lblFechaIncio)
        Me.Frame2.Controls.Add(Me.lblFechaFin)
        Me.Frame2.Controls.Add(Me.lblTipoCambioDolar)
        Me.Frame2.Controls.Add(Me._Shape1_1)
        Me.Frame2.Controls.Add(Me._Shape1_0)
        Me.Frame2.Controls.Add(Me._Label1_3)
        Me.Frame2.Controls.Add(Me._Label1_2)
        Me.Frame2.Controls.Add(Me._Shape1_6)
        Me.Frame2.Controls.Add(Me._Shape1_3)
        Me.Frame2.Controls.Add(Me._Shape1_5)
        Me.Frame2.Controls.Add(Me._Shape1_4)
        Me.Frame2.Controls.Add(Me._Label1_11)
        Me.Frame2.Controls.Add(Me._Label1_6)
        Me.Frame2.Controls.Add(Me._Label1_0)
        Me.Frame2.Controls.Add(Me._Label1_14)
        Me.Frame2.Controls.Add(Me._Shape1_9)
        Me.Frame2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(258, 6)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(367, 380)
        Me.Frame2.TabIndex = 7
        '
        'msgExistencia
        '
        Me.msgExistencia.DataSource = Nothing
        Me.msgExistencia.Location = New System.Drawing.Point(5, 83)
        Me.msgExistencia.Margin = New System.Windows.Forms.Padding(2)
        Me.msgExistencia.Name = "msgExistencia"
        Me.msgExistencia.OcxState = CType(resources.GetObject("msgExistencia.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msgExistencia.Size = New System.Drawing.Size(347, 82)
        Me.msgExistencia.TabIndex = 4
        '
        'msgPromociones
        '
        Me.msgPromociones.DataSource = Nothing
        Me.msgPromociones.Location = New System.Drawing.Point(11, 207)
        Me.msgPromociones.Margin = New System.Windows.Forms.Padding(2)
        Me.msgPromociones.Name = "msgPromociones"
        Me.msgPromociones.OcxState = CType(resources.GetObject("msgPromociones.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msgPromociones.Size = New System.Drawing.Size(347, 125)
        Me.msgPromociones.TabIndex = 5
        '
        '_Shape1_2
        '
        Me._Shape1_2.BackColor = System.Drawing.Color.White
        Me._Shape1_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_2.Location = New System.Drawing.Point(11, 207)
        Me._Shape1_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_2.Name = "_Shape1_2"
        Me._Shape1_2.Size = New System.Drawing.Size(348, 125)
        Me._Shape1_2.TabIndex = 6
        '
        '_Shape1_10
        '
        Me._Shape1_10.BackColor = System.Drawing.Color.White
        Me._Shape1_10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_10.Location = New System.Drawing.Point(5, 83)
        Me._Shape1_10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_10.Name = "_Shape1_10"
        Me._Shape1_10.Size = New System.Drawing.Size(347, 82)
        Me._Shape1_10.TabIndex = 7
        '
        'lblNoExisteProm
        '
        Me.lblNoExisteProm.BackColor = System.Drawing.Color.White
        Me.lblNoExisteProm.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNoExisteProm.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNoExisteProm.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNoExisteProm.Location = New System.Drawing.Point(72, 225)
        Me.lblNoExisteProm.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNoExisteProm.Name = "lblNoExisteProm"
        Me.lblNoExisteProm.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNoExisteProm.Size = New System.Drawing.Size(215, 16)
        Me.lblNoExisteProm.TabIndex = 23
        Me.lblNoExisteProm.Text = "No Existe"
        Me.lblNoExisteProm.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblNoExisteProm.Visible = False
        '
        'spNoExisteProm
        '
        Me.spNoExisteProm.BackColor = System.Drawing.Color.White
        Me.spNoExisteProm.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.spNoExisteProm.Location = New System.Drawing.Point(12, 224)
        Me.spNoExisteProm.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.spNoExisteProm.Name = "spNoExisteProm"
        Me.spNoExisteProm.Size = New System.Drawing.Size(346, 21)
        Me.spNoExisteProm.TabIndex = 24
        Me.spNoExisteProm.Visible = False
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.HighlightText
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.Color.Blue
        Me._Label1_1.Location = New System.Drawing.Point(57, 2)
        Me._Label1_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(230, 17)
        Me._Label1_1.TabIndex = 16
        Me._Label1_1.Text = "P R E C I O   P Ú B L I C O"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_10
        '
        Me._Label1_10.BackColor = System.Drawing.SystemColors.HighlightText
        Me._Label1_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_10.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_10.ForeColor = System.Drawing.Color.Blue
        Me._Label1_10.Location = New System.Drawing.Point(88, 181)
        Me._Label1_10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_10.Name = "_Label1_10"
        Me._Label1_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_10.Size = New System.Drawing.Size(178, 17)
        Me._Label1_10.TabIndex = 14
        Me._Label1_10.Text = "P R O M O C I O N"
        Me._Label1_10.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblPrecioPesos
        '
        Me.lblPrecioPesos.BackColor = System.Drawing.Color.White
        Me.lblPrecioPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPrecioPesos.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrecioPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPrecioPesos.Location = New System.Drawing.Point(85, 58)
        Me.lblPrecioPesos.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblPrecioPesos.Name = "lblPrecioPesos"
        Me.lblPrecioPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPrecioPesos.Size = New System.Drawing.Size(65, 14)
        Me.lblPrecioPesos.TabIndex = 13
        Me.lblPrecioPesos.Text = "999,999.00"
        Me.lblPrecioPesos.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblPrecioDolares
        '
        Me.lblPrecioDolares.BackColor = System.Drawing.Color.White
        Me.lblPrecioDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPrecioDolares.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrecioDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPrecioDolares.Location = New System.Drawing.Point(211, 58)
        Me.lblPrecioDolares.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblPrecioDolares.Name = "lblPrecioDolares"
        Me.lblPrecioDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPrecioDolares.Size = New System.Drawing.Size(65, 16)
        Me.lblPrecioDolares.TabIndex = 12
        Me.lblPrecioDolares.Text = "999,999.00"
        Me.lblPrecioDolares.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblDescuento
        '
        Me.lblDescuento.BackColor = System.Drawing.Color.White
        Me.lblDescuento.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDescuento.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescuento.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDescuento.Location = New System.Drawing.Point(113, 256)
        Me.lblDescuento.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDescuento.Name = "lblDescuento"
        Me.lblDescuento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDescuento.Size = New System.Drawing.Size(55, 17)
        Me.lblDescuento.TabIndex = 11
        Me.lblDescuento.Text = "$ 9,999.99"
        Me.lblDescuento.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblDescuento.Visible = False
        '
        'lblFechaIncio
        '
        Me.lblFechaIncio.BackColor = System.Drawing.Color.White
        Me.lblFechaIncio.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFechaIncio.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFechaIncio.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFechaIncio.Location = New System.Drawing.Point(48, 228)
        Me.lblFechaIncio.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFechaIncio.Name = "lblFechaIncio"
        Me.lblFechaIncio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFechaIncio.Size = New System.Drawing.Size(70, 15)
        Me.lblFechaIncio.TabIndex = 10
        Me.lblFechaIncio.Text = "01/Julio/2003"
        Me.lblFechaIncio.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblFechaFin
        '
        Me.lblFechaFin.BackColor = System.Drawing.Color.White
        Me.lblFechaFin.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFechaFin.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFechaFin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFechaFin.Location = New System.Drawing.Point(162, 228)
        Me.lblFechaFin.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFechaFin.Name = "lblFechaFin"
        Me.lblFechaFin.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFechaFin.Size = New System.Drawing.Size(70, 14)
        Me.lblFechaFin.TabIndex = 9
        Me.lblFechaFin.Text = "31/Julio/2003"
        Me.lblFechaFin.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblTipoCambioDolar
        '
        Me.lblTipoCambioDolar.BackColor = System.Drawing.Color.White
        Me.lblTipoCambioDolar.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTipoCambioDolar.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTipoCambioDolar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTipoCambioDolar.Location = New System.Drawing.Point(309, 350)
        Me.lblTipoCambioDolar.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTipoCambioDolar.Name = "lblTipoCambioDolar"
        Me.lblTipoCambioDolar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTipoCambioDolar.Size = New System.Drawing.Size(43, 15)
        Me.lblTipoCambioDolar.TabIndex = 8
        Me.lblTipoCambioDolar.Text = "99.99"
        Me.lblTipoCambioDolar.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Shape1_1
        '
        Me._Shape1_1.BackColor = System.Drawing.Color.White
        Me._Shape1_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_1.Location = New System.Drawing.Point(205, 56)
        Me._Shape1_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_1.Name = "_Shape1_1"
        Me._Shape1_1.Size = New System.Drawing.Size(75, 21)
        Me._Shape1_1.TabIndex = 25
        '
        '_Shape1_0
        '
        Me._Shape1_0.BackColor = System.Drawing.Color.White
        Me._Shape1_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_0.Location = New System.Drawing.Point(79, 56)
        Me._Shape1_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_0.Name = "_Shape1_0"
        Me._Shape1_0.Size = New System.Drawing.Size(75, 21)
        Me._Shape1_0.TabIndex = 26
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.ForeColor = System.Drawing.Color.Blue
        Me._Label1_3.Location = New System.Drawing.Point(97, 43)
        Me._Label1_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(37, 14)
        Me._Label1_3.TabIndex = 15
        Me._Label1_3.Text = "Pesos"
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.ForeColor = System.Drawing.Color.Blue
        Me._Label1_2.Location = New System.Drawing.Point(211, 43)
        Me._Label1_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(49, 14)
        Me._Label1_2.TabIndex = 17
        Me._Label1_2.Text = "Dólares"
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Shape1_6
        '
        Me._Shape1_6.BackColor = System.Drawing.Color.White
        Me._Shape1_6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_6.Location = New System.Drawing.Point(303, 346)
        Me._Shape1_6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_6.Name = "_Shape1_6"
        Me._Shape1_6.Size = New System.Drawing.Size(56, 25)
        Me._Shape1_6.TabIndex = 27
        '
        '_Shape1_3
        '
        Me._Shape1_3.BackColor = System.Drawing.Color.White
        Me._Shape1_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_3.Location = New System.Drawing.Point(108, 254)
        Me._Shape1_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_3.Name = "_Shape1_3"
        Me._Shape1_3.Size = New System.Drawing.Size(67, 21)
        Me._Shape1_3.TabIndex = 28
        Me._Shape1_3.Visible = False
        '
        '_Shape1_5
        '
        Me._Shape1_5.BackColor = System.Drawing.Color.White
        Me._Shape1_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_5.Location = New System.Drawing.Point(156, 226)
        Me._Shape1_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_5.Name = "_Shape1_5"
        Me._Shape1_5.Size = New System.Drawing.Size(78, 17)
        Me._Shape1_5.TabIndex = 29
        '
        '_Shape1_4
        '
        Me._Shape1_4.BackColor = System.Drawing.Color.White
        Me._Shape1_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_4.Location = New System.Drawing.Point(42, 227)
        Me._Shape1_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_4.Name = "_Shape1_4"
        Me._Shape1_4.Size = New System.Drawing.Size(78, 17)
        Me._Shape1_4.TabIndex = 30
        '
        '_Label1_11
        '
        Me._Label1_11.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._Label1_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_11.ForeColor = System.Drawing.Color.Blue
        Me._Label1_11.Location = New System.Drawing.Point(126, 228)
        Me._Label1_11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_11.Name = "_Label1_11"
        Me._Label1_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_11.Size = New System.Drawing.Size(31, 11)
        Me._Label1_11.TabIndex = 19
        Me._Label1_11.Text = "Hasta :"
        '
        '_Label1_6
        '
        Me._Label1_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.ForeColor = System.Drawing.Color.Blue
        Me._Label1_6.Location = New System.Drawing.Point(12, 228)
        Me._Label1_6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(31, 11)
        Me._Label1_6.TabIndex = 18
        Me._Label1_6.Text = "Desde :"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.Color.Blue
        Me._Label1_0.Location = New System.Drawing.Point(48, 258)
        Me._Label1_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(61, 11)
        Me._Label1_0.TabIndex = 20
        Me._Label1_0.Text = "Descuento :"
        Me._Label1_0.Visible = False
        '
        '_Label1_14
        '
        Me._Label1_14.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._Label1_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_14.ForeColor = System.Drawing.Color.Blue
        Me._Label1_14.Location = New System.Drawing.Point(202, 352)
        Me._Label1_14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_14.Name = "_Label1_14"
        Me._Label1_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_14.Size = New System.Drawing.Size(103, 11)
        Me._Label1_14.TabIndex = 21
        Me._Label1_14.Text = "Tipo Cambio Dólar:"
        '
        '_Shape1_9
        '
        Me._Shape1_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._Shape1_9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_9.Location = New System.Drawing.Point(0, 0)
        Me._Shape1_9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_9.Name = "_Shape1_9"
        Me._Shape1_9.Size = New System.Drawing.Size(367, 381)
        Me._Shape1_9.TabIndex = 31
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.Frame1.Controls.Add(Me.txtCodigoAnterior)
        Me.Frame1.Controls.Add(Me.txtDescArticulo)
        Me.Frame1.Controls.Add(Me.txtCodArticulo)
        Me.Frame1.Controls.Add(Me._Label1_4)
        Me.Frame1.Controls.Add(Me._Label1_18)
        Me.Frame1.Controls.Add(Me._Label1_17)
        Me.Frame1.Controls.Add(Me.imgImagenArticulo)
        Me.Frame1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(6, 8)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(242, 378)
        Me.Frame1.TabIndex = 22
        '
        'txtDescArticulo
        '
        Me.txtDescArticulo.AcceptsReturn = True
        Me.txtDescArticulo.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescArticulo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescArticulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescArticulo.Location = New System.Drawing.Point(12, 74)
        Me.txtDescArticulo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescArticulo.MaxLength = 0
        Me.txtDescArticulo.Name = "txtDescArticulo"
        Me.txtDescArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescArticulo.Size = New System.Drawing.Size(221, 20)
        Me.txtDescArticulo.TabIndex = 3
        '
        '_Label1_4
        '
        Me._Label1_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_4.ForeColor = System.Drawing.Color.Blue
        Me._Label1_4.Location = New System.Drawing.Point(125, 20)
        Me._Label1_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_4.Size = New System.Drawing.Size(109, 14)
        Me._Label1_4.TabIndex = 25
        Me._Label1_4.Text = "Código anterior"
        '
        '_Label1_18
        '
        Me._Label1_18.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._Label1_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_18.ForeColor = System.Drawing.Color.Blue
        Me._Label1_18.Location = New System.Drawing.Point(12, 20)
        Me._Label1_18.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_18.Name = "_Label1_18"
        Me._Label1_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_18.Size = New System.Drawing.Size(109, 14)
        Me._Label1_18.TabIndex = 0
        Me._Label1_18.Text = "Código del artículo"
        '
        '_Label1_17
        '
        Me._Label1_17.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._Label1_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_17.ForeColor = System.Drawing.Color.Blue
        Me._Label1_17.Location = New System.Drawing.Point(12, 58)
        Me._Label1_17.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_17.Name = "_Label1_17"
        Me._Label1_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_17.Size = New System.Drawing.Size(85, 14)
        Me._Label1_17.TabIndex = 2
        Me._Label1_17.Text = "Descripción"
        '
        'imgImagenArticulo
        '
        Me.imgImagenArticulo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.imgImagenArticulo.Cursor = System.Windows.Forms.Cursors.Default
        Me.imgImagenArticulo.Image = Global.CorporativoV1.My.Resources.Resources.JMR
        Me.imgImagenArticulo.Location = New System.Drawing.Point(15, 137)
        Me.imgImagenArticulo.Margin = New System.Windows.Forms.Padding(2)
        Me.imgImagenArticulo.Name = "imgImagenArticulo"
        Me.imgImagenArticulo.Size = New System.Drawing.Size(213, 222)
        Me.imgImagenArticulo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgImagenArticulo.TabIndex = 26
        Me.imgImagenArticulo.TabStop = False
        '
        '_Shape1_8
        '
        Me._Shape1_8.BackColor = System.Drawing.Color.White
        Me._Shape1_8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_8.Location = New System.Drawing.Point(4, 6)
        Me._Shape1_8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_8.Name = "_Shape1_8"
        Me._Shape1_8.Size = New System.Drawing.Size(244, 381)
        Me._Shape1_8.TabIndex = 23
        '
        '_Shape1_7
        '
        Me._Shape1_7.BackColor = System.Drawing.Color.White
        Me._Shape1_7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Shape1_7.Location = New System.Drawing.Point(2, 1)
        Me._Shape1_7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Shape1_7.Name = "_Shape1_7"
        Me._Shape1_7.Size = New System.Drawing.Size(647, 405)
        Me._Shape1_7.TabIndex = 7
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Window
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(6, 418)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 81
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(121, 419)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 80
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmVerificadorPrecios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(653, 463)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Marco)
        Me.Controls.Add(Me._Shape1_7)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(158, 178)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVerificadorPrecios"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Verificador de Precios"
        Me.Marco.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        CType(Me.msgExistencia, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.msgPromociones, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.imgImagenArticulo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Shape1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Sub Buscar()
        BuscarArticulos(False, CStr(0))
    End Sub

    Private Sub Command1_Click()
        Dim Picture1 As Object
        Dim Line1 As Object
        Line1.BorderColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
        Picture1.b()
    End Sub

    Sub NoExistePromocion()
        spNoExisteProm.Visible = True
        lblNoExisteProm.Visible = True
    End Sub

    Sub Nuevo()
        Dim lblExistenciaArticulos As Object
        On Error GoTo Merr

        'txtCodArticulo.Text = ""
        FueraChange = True
        txtDescArticulo.Text = ""
        FueraChange = False
        lblPrecioDolares.Text = ""
        lblPrecioPesos.Text = ""
        lblExistenciaArticulos = ""
        lblFechaFin.Text = ""
        lblFechaIncio.Text = ""
        lblDescuento.Text = ""
        lblTipoCambioDolar.Text = ""
        txtCodigoAnterior.Text = ""
        lblNoExisteProm.Visible = False
        spNoExisteProm.Visible = False
        msgPromociones.Visible = True
        _Shape1_2.Visible = True
        msgExistencia.Clear()
        msgPromociones.Clear()
        Encabezado()
        'ModCorporativo.BuscaImagen("", imgImagenArticulo)

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub SiExistePromocion()
        lblNoExisteProm.Visible = False
        spNoExisteProm.Visible = False
    End Sub

    Private Sub frmVerificadorPrecios_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVerificadorPrecios_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmVerificadorPrecios_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
        Encabezado()
    End Sub

    Private Sub frmVerificadorPrecios_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        ' En este evento del formulario se valida la tecla presionada.
        ' Si es Enter se simula un tab(Avanza al siguiente control)
        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ' Si el control en que se presiono enter, es el Grid de Detalle de la venta que no se ejecute el avanzar tab
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmVerificadorPrecios_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVerificadorPrecios_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        Dim mblnSalir As Object
        If Not mblnSalir Then
            'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
            ModEstandar.RestaurarForma(Me, False)
        Else 'Se quiere salir con escape
            mblnSalir = False
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
                Case MsgBoxResult.Yes
                    Cancel = 0 'Sale de la Captura, Con 1: Sigue en la captura
                Case MsgBoxResult.No 'No sale del formulario
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVerificadorPrecios_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub MSGEXISTENCIA_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgExistencia.Enter
        'iNDICA LA CELDA QUE APARECERA SELECCIONADA
        msgExistencia.Col = 0
        msgExistencia.ColSel = 3
        msgExistencia.HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightAlways
        msgExistencia.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    End Sub

    Private Sub MSGEXISTENCIA_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgExistencia.Leave
        msgExistencia.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
        msgExistencia.HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
    End Sub

    Private Sub msgPromociones_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgPromociones.Enter
        'iNDICA LA CELDA QUE APARECERA SELECCIONADA
        msgPromociones.Col = 0
        msgPromociones.ColSel = 3
        msgPromociones.HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightAlways
        msgPromociones.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    End Sub

    Private Sub msgPromociones_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgPromociones.Leave
        msgPromociones.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
        msgPromociones.HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
    End Sub

    '''Modific.-  24ENE2005 -  no mostraba promociones de articulos por fecha incorrecta
    Private Sub txtCodArticulo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArticulo.TextChanged
        If FueraChange = True Then Exit Sub
        Nuevo()
        '''If Trim(txtCodArticulo.text) = "" Then Nuevo
    End Sub

    Private Sub txtCodArticulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArticulo.Enter
        strControlActual = UCase("txtCodArticulo")
        SelTextoTxt(txtCodArticulo)
        Pon_Tool()
    End Sub

    Private Sub txtCodArticulo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodArticulo.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Dim mblnSalir As Object
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            KeyCode = 0
        ElseIf KeyCode = System.Windows.Forms.Keys.Delete Then
            'sI La Tecla presionada fue SUPR, se borrará todo el contenido del form. ya que no es posible hacer modificaciones.
            'Unicamnete podran consultarse los datos.
            Nuevo()
        End If
    End Sub

    Private Sub txtCodArticulo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodArticulo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtCodArticulo.Text, KeyAscii, 8, 0, (txtCodArticulo.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub LlenaDatos(ByRef CodArticulo As Integer)
        On Error GoTo Merr
        Dim TotalDescuento As Decimal

        If Trim(txtCodArticulo.Text) = "" Then Exit Sub

        '''gStrSql = "SELECT TOP 1 A.CodArticulo, A.DescArticulo, A.PrecioPubDolar, A.PesosFijos, ISNULL(PV.Importe, 0) AS Importe, ISNULL(PV.Porcentaje, 0) AS Porcentaje, ISNULL(PV.FechaInicio, " & _
        '"'01/01/1900') AS FechaInicio, ISNULL(PV.FechaFin, '01/01/1900') AS FechaFin , CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as CodAnterior " & _
        '"FROM dbo.CatArticulos A LEFT OUTER JOIN  " & _
        '"dbo.PromocionesVentas PV ON A.CodGrupo = PV.CodGrupo AND A.CodFamilia = PV.CodFamilia AND A.CodLinea = PV.CodLinea AND  " & _
        '"A.CodSubLinea = PV.CodSubLinea Or A.CodMarca = PV.CodMarca And A.CodModelo = PV.CodModelo  " & _
        '"Where (A.CodArticulo = " & CodArticulo & ") ORDER BY PV.FechaInicio DESC"

        gStrSql = "SELECT   A.CodArticulo, A.DescArticulo, A.PrecioPubDolar, A.PesosFijos, ISNULL(PV.Importe, 0) AS Importe, " & "         ISNULL(PV.Porcentaje, 0) AS Porcentaje, ISNULL(PV.FechaInicio, '01/01/1900') AS FechaInicio, " & "         ISNULL(PV.FechaFin, '01/01/1900') AS FechaFin , CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as CodAnterior " & "FROM     dbo.CatArticulos A LEFT OUTER JOIN  dbo.PromocionesVentas PV ON A.CodArticulo = PV.CodArticulo And ((PV.FechaInicio <= '" & Format(Today, C_FORMATFECHAGUARDAR) & "' And PV.FechaFin >= '" & Format(Today, C_FORMATFECHAGUARDAR) & "') or (PV.FechaFin > '" & Format(Today, C_FORMATFECHAGUARDAR) & "')) " & "Where    (A.CodArticulo = " & CodArticulo & ") " & "ORDER    BY PV.FechaInicio DESC "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            FueraChange = True
            txtCodArticulo.Text = RsGral.Fields("CodArticulo").Value
            txtCodigoAnterior.Text = RsGral.Fields("CodAnterior").Value
            txtDescArticulo.Text = Trim(RsGral.Fields("DescArticulo").Value)
            If Not RsGral.Fields("PesosFijos").Value Then
                lblPrecioDolares.Text = Format(System.Math.Round(RsGral.Fields("PrecioPubDolar").Value, 2), gstrFormatoCantidad)
                lblPrecioPesos.Text = Format(System.Math.Round(RsGral.Fields("PrecioPubDolar").Value * gcurCorpoTIPOCAMBIODOLAR, 1), gstrFormatoCantidad)
                lblTipoCambioDolar.Text = Format(gcurCorpoTIPOCAMBIODOLAR, gstrFormatoCantidad)
            Else
                lblPrecioPesos.Text = Format(System.Math.Round(RsGral.Fields("PrecioPubDolar").Value, 2), gstrFormatoCantidad)
                lblPrecioDolares.Text = Format(System.Math.Round(RsGral.Fields("PrecioPubDolar").Value / gcurCorpoTIPOCAMBIODOLAR, 2), gstrFormatoCantidad)
                lblTipoCambioDolar.Text = Format(gcurCorpoTIPOCAMBIODOLAR, gstrFormatoCantidad)
            End If
            LlenaDatosExistencia(CInt(Numerico(txtCodArticulo.Text)))
            'ModCorporativo.BuscaImagen(Trim(txtCodArticulo.Text), imgImagenArticulo)
            FueraChange = False
        Else
            MsjNoExiste("El Artículo", gstrCorpoNOMBREEMPRESA)
            Limpiar()
        End If
        Exit Sub

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatosDescripcion(ByRef DesArticulo As String)
        On Error GoTo Merr
        Dim TotalDescuento As Decimal

        If Trim(txtDescArticulo.Text) = "" Then Exit Sub

        '''gStrSql = "SELECT TOP 1 A.CodArticulo, A.DescArticulo, A.PrecioPubDolar, A.PesosFijos, ISNULL(PV.Importe, 0) AS Importe, ISNULL(PV.Porcentaje, 0) AS Porcentaje, ISNULL(PV.FechaInicio, " & _
        '"'01/01/1900') AS FechaInicio, ISNULL(PV.FechaFin, '01/01/1900') AS FechaFin , CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as CodAnterior " & _
        '"FROM dbo.CatArticulos A LEFT OUTER JOIN  " & _
        '"dbo.PromocionesVentas PV ON A.CodGrupo = PV.CodGrupo AND A.CodFamilia = PV.CodFamilia AND A.CodLinea = PV.CodLinea AND  " & _
        '"A.CodSubLinea = PV.CodSubLinea Or A.CodMarca = PV.CodMarca And A.CodModelo = PV.CodModelo  " & _
        '"Where (A.CodArticulo = " & CodArticulo & ") ORDER BY PV.FechaInicio DESC"

        gStrSql = "SELECT   A.CodArticulo, A.DescArticulo, A.PrecioPubDolar, A.PesosFijos, ISNULL(PV.Importe, 0) AS Importe, " & "         ISNULL(PV.Porcentaje, 0) AS Porcentaje, ISNULL(PV.FechaInicio, '01/01/1900') AS FechaInicio, " & "         ISNULL(PV.FechaFin, '01/01/1900') AS FechaFin , CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as CodAnterior " & "FROM     dbo.CatArticulos A LEFT OUTER JOIN  dbo.PromocionesVentas PV ON A.CodArticulo = PV.CodArticulo And ((PV.FechaInicio <= '" & Format(Today, C_FORMATFECHAGUARDAR) & "' And PV.FechaFin >= '" & Format(Today, C_FORMATFECHAGUARDAR) & "') or (PV.FechaFin > '" & Format(Today, C_FORMATFECHAGUARDAR) & "')) " & "Where    (A.DescArticulo like '" & DesArticulo & "%' ) " & "ORDER    BY PV.FechaInicio DESC "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            FueraChange = True
            txtCodArticulo.Text = RsGral.Fields("CodArticulo").Value
            txtCodigoAnterior.Text = RsGral.Fields("CodAnterior").Value
            txtDescArticulo.Text = Trim(RsGral.Fields("DescArticulo").Value)
            If Not RsGral.Fields("PesosFijos").Value Then
                lblPrecioDolares.Text = Format(System.Math.Round(RsGral.Fields("PrecioPubDolar").Value, 2), gstrFormatoCantidad)
                lblPrecioPesos.Text = Format(System.Math.Round(RsGral.Fields("PrecioPubDolar").Value * gcurCorpoTIPOCAMBIODOLAR, 1), gstrFormatoCantidad)
                lblTipoCambioDolar.Text = Format(gcurCorpoTIPOCAMBIODOLAR, gstrFormatoCantidad)
            Else
                lblPrecioPesos.Text = Format(System.Math.Round(RsGral.Fields("PrecioPubDolar").Value, 2), gstrFormatoCantidad)
                lblPrecioDolares.Text = Format(System.Math.Round(RsGral.Fields("PrecioPubDolar").Value / gcurCorpoTIPOCAMBIODOLAR, 2), gstrFormatoCantidad)
                lblTipoCambioDolar.Text = Format(gcurCorpoTIPOCAMBIODOLAR, gstrFormatoCantidad)
            End If
            LlenaDatosExistencia(CInt(Numerico(txtCodArticulo.Text)))
            LlenaDatosPromocion(CInt(Numerico(txtCodArticulo.Text)))
            'ModCorporativo.BuscaImagen(Trim(txtCodArticulo.Text), imgImagenArticulo)
            FueraChange = False
        Else
            MsjNoExiste("El Artículo", gstrCorpoNOMBREEMPRESA)
            Limpiar()
        End If
        Exit Sub

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub txtCodArticulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArticulo.Leave
        Dim ResBusquedaArt As Object
        'If System.Windows.Forms.Form.ActiveForm.Text <> Me.Text Then
        '    Exit Sub
        'End If
        Dim Prefijo As String
        Dim CodAux As Integer
        If Trim(txtCodArticulo.Text) <> "" Then
            ResBusquedaArt = ModCorporativo.BuscarCodigoArticulo(Trim(txtCodArticulo.Text))
            If ResBusquedaArt > 0 Or ResBusquedaArt = -1 Then
                LlenaDatos(CInt(ResBusquedaArt))
            ElseIf ResBusquedaArt = -2 Then
                CodAux = CInt(txtCodArticulo.Text)
                txtCodArticulo.Text = ""
                FueraChange = True
                BuscarArticulos(True, (New String("0", 6) & Trim(CStr(CodAux))))
                FueraChange = False
            End If
        End If
        LlenaDatosPromocion(CInt(Numerico(txtCodArticulo.Text)))
    End Sub

    Sub BuscarArticulos(ByRef BusquedaEspecial As Boolean, ByRef CodArticulo As String)
        Dim mblnFueraChange As Object
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas 
        Dim Columna As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        strCaptionForm = "Consulta de Articulos"
        If strControlActual = "TXTCODARTICULO" And BusquedaEspecial = False Then
            strSQL = "SELECT     A.CodArticulo AS CODIGO, LTRIM(RTRIM(A.DescArticulo)) AS DESCRIPCION, M.DescTipoMaterial AS MATERIAL, LTrim(Rtrim(A.CodigoArticuloProv)) AS [ARTICULO PROV],  CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt) + '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]   " & "FROM dbo.CatArticulos A INNER JOIN dbo.CatTipoMaterial M ON A.CodTipoMaterial = M.CodTipoMaterial  "
        ElseIf strControlActual = "TXTDESCARTICULO" And BusquedaEspecial = False Then
            strSQL = "SELECT     A.CodArticulo AS CODIGO, LTRIM(RTRIM(A.DescArticulo)) AS DESCRIPCION, M.DescTipoMaterial AS MATERIAL, LTrim(Rtrim(A.CodigoArticuloProv)) AS [ARTICULO PROV],  CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+ '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]   " & "FROM dbo.CatArticulos A INNER JOIN dbo.CatTipoMaterial M ON A.CodTipoMaterial = M.CodTipoMaterial  wHERE DescArticulo Like '" & Trim(txtDescArticulo.Text) & "%'"
        ElseIf strControlActual = "TXTDESCARTICULO" And BusquedaEspecial Then
            strSQL = "SELECT     CodArticulo AS CODIGO, RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION, " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR], " & "dbo.FormatCantidad(A.PrecioPubDolar)  AS [PRECIO PÚBLICO] , " & "case PesosFijos WHEN 0 THEN 'DÓLARES' WHEN 1 THEN 'PESOS' END AS [MONEDA] " & "From CatArticulos A cross Join Configuraciongeneral c WHERE (CodArticulo = " & CInt(CodArticulo) & ") " & "OR   (OrigenAnt = " & CInt((CodArticulo)) & ") AND (CodigoAnt = " & CInt((CodArticulo)) & ")"

        Else
            Exit Sub
        End If

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, strSQL))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            If BusquedaEspecial = True Then
                MsgBox("El Artículo no existe." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                RsGral.Close()
                Exit Sub
            Else
                MsgBox(C_msgSINDATOS & vbNewLine & "Verifique por favor....", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                RsGral.Close()
                Exit Sub
            End If
        End If

        'Carga el formulario de consulta
        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 10950, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            .set_ColWidth(0, 0, 900)
            .set_ColWidth(1, 0, 4800)
            .set_ColWidth(2, 0, 1900)
            .set_ColWidth(3, 0, 1620)
            .set_ColWidth(4, 0, 1700)

            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)

            .WordWrap = False
        End With
        mblnFueraChange = True
        CentrarForma(FrmConsultas)
        FrmConsultas.ShowDialog()
        mblnFueraChange = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Limpiar()
        Nuevo()
        FueraChange = True
        txtCodArticulo.Text = ""
        txtCodArticulo.Focus()
        FueraChange = False
    End Sub

    Sub Encabezado()
        Dim C_COLCODIGO As Object
        'Genera el encabezao del Grid, asigna el tamaño y número de columas y centra el texto dentro de ellas
        Dim LnContador As Integer

        With msgExistencia
            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusHeavy 'flexFocusLight 'flexFocusNone
            .WordWrap = False
            .FixedRows = 1
            .FixedCols = 0
            .set_ColWidth(C_ColSUCURSAL, 0, 2615)
            .set_ColWidth(C_ColCODSUCURSAL, 0, 0)
            .set_ColWidth(C_ColEXISTENCIA, 0, 1000)
            .set_ColWidth(C_ColAPARTADOS, 0, 1)

            .set_TextMatrix(0, C_ColSUCURSAL, "Sucursal")
            .set_TextMatrix(0, C_ColCODSUCURSAL, "Sucursal")
            .set_TextMatrix(0, C_ColEXISTENCIA, "Existencia")
            .set_TextMatrix(0, C_ColAPARTADOS, "Apartados")

            .Row = 0
            For LnContador = 0 To C_ColAPARTADOS
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next LnContador
            .Row = 1
            .Col = C_COLCODIGO
            .WordWrap = False 'Hacer esto , para que no se puedan escribir dos o mal lineas de texto en una  sola fila, solo se usa para el encabezado
        End With

        With msgPromociones
            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusHeavy 'flexFocusLight 'flexFocusNone
            .WordWrap = False
            .FixedRows = 0
            .FixedRows = 1
            .FixedCols = 0
            .set_ColWidth(C_ColPROMOCION, 0, 970)
            .set_ColWidth(C_COLDESDE, 0, 1115)
            .set_ColWidth(C_COLHASTA, 0, 1115)
            .set_ColWidth(C_ColDESCUENTO, 0, 1130)

            .set_TextMatrix(0, C_ColPROMOCION, "Promoción")
            .set_TextMatrix(0, C_COLDESDE, "Desde")
            .set_TextMatrix(0, C_COLHASTA, "Hasta")
            .set_TextMatrix(0, C_ColDESCUENTO, "Descuento")

            .set_ColAlignment(C_ColDESCUENTO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .Row = 0
            For LnContador = 0 To C_ColDESCUENTO
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next LnContador
            .Row = 1
            .Col = C_COLCODIGO
            .WordWrap = False 'Hacer esto , para que no se puedan escribir dos o mal lineas de texto en una  sola fila, solo se usa para el encabezado
        End With
    End Sub

    Sub LlenaDatosPromocion(ByRef CodArticulo As Integer)
        Dim TotalDescuento As Object
        On Error GoTo Merr
        Dim I As Integer

        If CodArticulo = 0 Then Exit Sub

        ''' Referencía: QLG V 1.0
        ''' Referencía: QLG V 1.1
        ''' gStrSql = FlArmaquery(CodArticulo)

        '''27MAY2008 - MAVF
        gStrSql = ""
        gStrSql = gStrSql & "Select   * " & vbNewLine
        gStrSql = gStrSql & "From     dbo.VerificadorPromociones(" & CodArticulo & ", '" & Format(Today, C_FORMATFECHAGUARDAR) & "') " & vbNewLine
        gStrSql = gStrSql & "Order    by FechaInicio " & vbNewLine
        '''*****************************************************************************************************************************************

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            lblTipoCambioDolar.Text = Format(gcurCorpoTIPOCAMBIODOLAR, gstrFormatoCantidad)
            msgPromociones.Visible = True
            _Shape1_2.Visible = True
            With msgPromociones
                I = 1
                Do While Not RsGral.EOF
                    TotalDescuento = RsGral.Fields("importe").Value + RsGral.Fields("Porcentaje").Value
                    If TotalDescuento = 0 Then
                        NoExistePromocion()
                        lblDescuento.Text = "0.00"
                    ElseIf RsGral.Fields("importe").Value <> 0 Then
                        lblDescuento.Text = "$" & " " & Format(RsGral.Fields("importe").Value, gstrFormatoCantidad)
                        SiExistePromocion()
                    ElseIf RsGral.Fields("Porcentaje").Value <> 0 Then
                        lblDescuento.Text = Format(RsGral.Fields("Porcentaje").Value, "0.00") & " " & "%"
                        SiExistePromocion()
                    End If
                    lblFechaIncio.Text = Format(RsGral.Fields("FechaInicio").Value, "dd/mmm/yyyy")
                    lblFechaFin.Text = Format(RsGral.Fields("FechaFin").Value, "dd/mmm/yyyy")
                    If RsGral.Fields("TipoProm").Value = "A" Then
                        .set_TextMatrix(I, C_ColPROMOCION, "Por Articulo")
                    ElseIf RsGral.Fields("TipoProm").Value = "G" Then
                        .set_TextMatrix(I, C_ColPROMOCION, "Por Grupo")
                    End If
                    .set_TextMatrix(I, C_COLDESDE, lblFechaIncio.Text)
                    .set_TextMatrix(I, C_COLHASTA, lblFechaFin.Text)
                    .set_TextMatrix(I, C_ColDESCUENTO, lblDescuento.Text)
                    RsGral.MoveNext()
                    If Not RsGral.EOF Then
                        If I = .Rows - 1 Then
                            .Rows = .Rows + 1
                        End If
                        I = I + 1
                    End If
                Loop
                .Visible = True
                .Col = 0
                .Row = 1
                .TopRow = 1
                .Col = C_ColPROMOCION
                .ColSel = C_ColDESCUENTO
                .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightAlways
                .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
            End With
        Else
            NoExistePromocion()
            msgPromociones.Visible = False
            _Shape1_2.Visible = False
            lblDescuento.Text = "0.00"
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatosExistencia(ByRef CodArticulo As Integer)
        Dim I As Object
        On Error GoTo Merr

        gStrSql = "SELECT  I.CodAlmacen, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) AS Existencia, SUM(I.Apartados) AS Apartados, Al.DescAlmacen " & "FROM    dbo.Inventario I INNER JOIN dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen " & "Where   (I.CodArticulo =" & CodArticulo & " ) AND (Al.TipoAlmacen = 'P') " & "GROUP   BY Al.DescAlmacen, I.CodAlmacen " & "ORDER   BY Existencia "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        msgExistencia.Clear()
        Encabezado()

        With msgExistencia
            If RsGral.RecordCount > 0 Then
                If RsGral.RecordCount > .Rows - 1 Then .Rows = RsGral.RecordCount + 1
                For I = 1 To RsGral.RecordCount
                    .set_TextMatrix(I, C_ColCODSUCURSAL, Trim(RsGral.Fields("CodAlmacen").Value))
                    .set_TextMatrix(I, C_ColSUCURSAL, Trim(RsGral.Fields("DescAlmacen").Value))
                    .set_TextMatrix(I, C_ColEXISTENCIA, RsGral.Fields("Existencia").Value)
                    .set_TextMatrix(I, C_ColAPARTADOS, RsGral.Fields("Apartados").Value)
                    RsGral.MoveNext()
                Next
            Else
            End If
            .Row = 1
            .TopRow = 1
            .Col = 0
            .ColSel = 3
            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightAlways
            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
        End With

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub txtDescArticulo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescArticulo.TextChanged
        If FueraChange = True Then Exit Sub
        If Trim(txtDescArticulo.Text) = "" Then
            FueraChange = True
            txtCodArticulo.Text = ""
            FueraChange = False
            Nuevo()
        End If
    End Sub

    Private Sub txtDescArticulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescArticulo.Enter
        strControlActual = UCase("txtDescArticulo")
        SelTextoTxt(txtDescArticulo)
    End Sub

    Private Sub txtDescArticulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescArticulo.Leave
        '''LlenaDatosDescripcion Trim(txtDescArticulo.text)
        ''' NO ES POSIBLE CONSIDERARLO DEBIDO A QUE EXISTEN DESCRIPCIONES DUPLICADAS Y
        ''' CUANDO EL TXTCODIGO PIERDE EL FOCO TOMA EL PRIMER ARTICULO ENCONTRADO CON DICHA DESCRIPCION
        ''' Y CAMBIA EL ARTICULO - MAVF 02MAY2006
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub
    '''FUNCION ANTERIOR - 27MAY2008 - MAVF
    '''Sub LlenaDatosPromocion(CodArticulo As Long)
    '''    On Local Error GoTo Merr
    '''    Dim I As Integer
    ''''    gStrSql = "SELECT   /*TOP 1*/   A.CodArticulo, A.DescArticulo, A.PrecioPubDolar, ISNULL(PV.Importe, 0) AS Importe, ISNULL(PV.Porcentaje, 0) AS Porcentaje, " & _
    '''''            "ISNULL(PV.FechaInicio,'01/01/1900') AS FechaInicio, ISNULL(PV.FechaFin, '01/01/1900') AS FechaFin " & _
    '''''            "FROM         dbo.CatArticulos A LEFT OUTER JOIN  " & _
    '''''            "dbo.PromocionesVentas PV ON A.CodGrupo = PV.CodGrupo AND A.CodFamilia = PV.CodFamilia AND A.CodLinea = PV.CodLinea AND " & _
    '''''            "A.CodSubLinea = PV.CodSubLinea Or A.CodMArca = PV.CodMArca And A.CodModelo = PV.CodModelo " & _
    '''''            "Where (A.CodArticulo = " & CodArticulo & ") and ( '" & Format(gdtmFechaOperacion, C_FORMATFECHAGUARDAR) & "' BetWeen  Pv.FechaInicio and Pv.FechaFin) " & _
    '''''            "ORDER BY PV.FechaInicio DESC "
    '''
    '''
    '''    '''gStrSql = "select * from " & _
    ''''    "(SELECT TOP 1 A.CodArticulo, A.DescArticulo, A.PrecioPubDolar, ISNULL(PV.Importe, 0) AS Importe," & _
    ''''    "ISNULL(PV.Porcentaje, 0) AS Porcentaje, ISNULL(PV.FechaInicio,'01/01/1900') AS FechaInicio," & _
    ''''    "ISNULL(PV.FechaFin, '01/01/1900') AS FechaFin,Tipoprom " & _
    ''''    "FROM dbo.CatArticulos A LEFT OUTER JOIN  dbo.PromocionesVentas PV ON A.CodGrupo = PV.CodGrupo " & _
    ''''    "and A.CodFamilia = PV.CodFamilia and A.CodLinea = PV.CodLinea OR A.CodSubLinea = PV.CodSubLinea " & _
    ''''    "or A.CodMArca = PV.CodMArca OR A.CodModelo = PV.CodModelo " & _
    ''''    "Where (A.CodArticulo = " & CodArticulo & ") and ('" & Format(Date, C_FORMATFECHAGUARDAR) & "' BetWeen  Pv.FechaInicio and Pv.FechaFin) and tipoprom = 'G' and estatus <> 'C' " & _
    ''''    "Union " & _
    ''''    "SELECT TOP 1 A.CodArticulo, A.DescArticulo, A.PrecioPubDolar, ISNULL(PV.Importe, 0) AS Importe," & _
    ''''    "ISNULL(PV.Porcentaje, 0) AS Porcentaje, ISNULL(PV.FechaInicio,'01/01/1900') AS FechaInicio," & _
    ''''    "ISNULL(PV.FechaFin, '01/01/1900') AS FechaFin,tipoprom " & _
    ''''    "FROM dbo.CatArticulos A LEFT OUTER JOIN  dbo.PromocionesVentas PV ON A.CodArticulo = PV.CodArticulo " & _
    ''''    "Where (A.CodArticulo = " & CodArticulo & ") and ('" & Format(Date, C_FORMATFECHAGUARDAR) & "' BetWeen  Pv.FechaInicio and Pv.FechaFin) and tipoprom = 'A' and estatus <> 'C' ) t " & _
    ''''    "order by fechainicio"
    '''
    '''   ' Referencía: QLG V 1.0
    '''
    '''    'gStrSql = "Select * From  ( " & _
    ''''    '          "       SELECT   A.CodArticulo, A.DescArticulo, A.PrecioPubDolar, ISNULL(PV.Importe, 0) AS Importe,ISNULL(PV.Porcentaje, 0) AS Porcentaje, ISNULL(PV.FechaInicio,'01/01/1900') AS FechaInicio,ISNULL(PV.FechaFin, '01/01/1900') AS FechaFin,Tipoprom " & _
    ''''    '          "       FROM  dbo.CatArticulos A LEFT OUTER JOIN  dbo.PromocionesVentas PV ON A.CodArticulo = PV.CodArticulo " & _
    ''''    '          "       Where (A.CodArticulo = " & CodArticulo & ") " & _
    ''''    '          "       And  ((PV.FechaInicio <= '" & Format(Date, C_FORMATFECHAGUARDAR) & "' And PV.FechaFin >= '" & Format(Date, C_FORMATFECHAGUARDAR) & "') or (PV.FechaFin > '" & Format(Date, C_FORMATFECHAGUARDAR) & "')) " & _
    ''''    '          "       And  PV.TipoProm = 'G' And PV.Estatus <> 'C' " & _
    ''''    '          "       Union " & _
    ''''    '          "       SELECT   A.CodArticulo, A.DescArticulo, A.PrecioPubDolar, ISNULL(PV.Importe, 0) AS Importe,ISNULL(PV.Porcentaje, 0) AS Porcentaje, ISNULL(PV.FechaInicio,'01/01/1900') AS FechaInicio,ISNULL(PV.FechaFin, '01/01/1900') AS FechaFin,Tipoprom " & _
    ''''    '          "       FROM  dbo.CatArticulos A LEFT OUTER JOIN  dbo.PromocionesVentas PV ON A.CodArticulo = PV.CodArticulo " & _
    ''''    '          "       Where (A.CodArticulo = " & CodArticulo & ") " & _
    ''''    '          "       And  ((PV.FechaInicio <= '" & Format(Date, C_FORMATFECHAGUARDAR) & "' And PV.FechaFin >= '" & Format(Date, C_FORMATFECHAGUARDAR) & "') or (PV.FechaFin > '" & Format(Date, C_FORMATFECHAGUARDAR) & "')) " & _
    ''''    '          "       And  PV.TipoProm = 'A' And PV.Estatus <> 'C' " & _
    ''''    '          " ) T   " & _
    ''''    '          "Order    By FechaInicio "
    '''
    '''   If CodArticulo = 0 Then Exit Sub
    '''
    '''   '''27MAY2008 - MAVF
    '''   gStrSql = ""
    '''   gStrSql = gStrSql & "Select   * " & vbNewLine
    '''   gStrSql = gStrSql & "From     dbo.VerificadorPromociones(" & CodArticulo & ", '" & Format(Date, C_FORMATFECHAGUARDAR) & "') " & vbNewLine
    '''   gStrSql = gStrSql & "Order    by FechaInicio " & vbNewLine
    '''   '''*****************************************************************************************************************************************
    '''
    '''   ''' Referencía: QLG V 1.1
    '''   '''gStrSql = FlArmaquery(CodArticulo)
    '''
    '''   ModEstandar.BorraCmd
    '''   Cmd.CommandText = "dbo.UP_SELECT_DATOS"
    '''   Cmd.CommandType = adCmdStoredProc
    '''   Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '''   Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
    '''   Set RsGral = Cmd.Execute
    '''
    '''   If RsGral.RecordCount > 0 Then
    '''      lblTipoCambioDolar = Format(gcurCorpoTIPOCAMBIODOLAR, gstrFormatoCantidad)
    '''      msgPromociones.Visible = True
    '''      _Shape1_2.Visible = True
    '''      With msgPromociones
    '''         I = 1
    '''         Do While Not RsGral.EOF
    '''            TotalDescuento = RsGral!importe + RsGral!Porcentaje
    '''            If TotalDescuento = 0 Then
    '''               NoExistePromocion
    '''               lblDescuento = "0.00"
    '''            ElseIf RsGral!importe <> 0 Then
    '''               lblDescuento = "$" & " " & Format(RsGral!importe, gstrFormatoCantidad)
    '''               SiExistePromocion
    '''            ElseIf RsGral!Porcentaje <> 0 Then
    '''               lblDescuento = Format(RsGral!Porcentaje, "0.00") & " " & "%"
    '''               SiExistePromocion
    '''            End If
    '''            lblFechaIncio = Format(RsGral!FechaInicio, "dd/mmm/yyyy")
    '''            lblFechaFin = Format(RsGral!FechaFin, "dd/mmm/yyyy")
    '''            If RsGral!TipoProm = "A" Then
    '''               .TextMatrix(I, C_ColPROMOCION) = "Por Articulo"
    '''            ElseIf RsGral!TipoProm = "G" Then
    '''               .TextMatrix(I, C_ColPROMOCION) = "Por Grupo"
    '''            End If
    '''            .TextMatrix(I, C_COLDESDE) = lblFechaIncio
    '''            .TextMatrix(I, C_COLHASTA) = lblFechaFin
    '''            .TextMatrix(I, C_ColDESCUENTO) = lblDescuento
    '''            RsGral.MoveNext
    '''            If Not RsGral.EOF Then
    '''               If I = .Rows - 1 Then
    '''                  .Rows = .Rows + 1
    '''               End If
    '''               I = I + 1
    '''            End If
    '''         Loop
    '''         .Visible = True
    '''         .Col = 0
    '''         .Row = 1
    '''         .TopRow = 1
    '''         .Col = C_ColPROMOCION
    '''         .ColSel = C_ColDESCUENTO
    '''         .HighLight = flexHighlightAlways
    '''         .FocusRect = flexFocusNone
    '''      End With
    '''   Else
    '''      NoExistePromocion
    '''      msgPromociones.Visible = False
    '''      _Shape1_2.Visible = False
    '''      lblDescuento = "0.00"
    '''   End If
    '''
    '''Merr:
    '''    If Err.Number <> 0 Then ModEstandar.MostrarError
    '''End Sub

    '''Private Function FlArmaquery(ByVal CodArticulo As String) As String
    ''''*************************************************************************************************************
    ''''** Nombre        : L.I José Gilberto Quintero López                                                        **
    ''''** Fecha         : 08 May 2008                                                                             **
    ''''** Referencía    : QLG V 1.1                                                                               **
    ''''*************************************************************************************************************
    '''On Local Error GoTo Merr
    '''
    '''  Const SELECT_CAMPOS = " ca.CodArticulo " & vbCrLf _
    ''''                        & " ,ca.DescArticulo" & vbCrLf _
    ''''                        & ",ca.PrecioPubDolar" & vbCrLf _
    ''''                        & ",isnull(pv.Importe, 0) AS Importe" & vbCrLf _
    ''''                        & ",isnull(pv.Porcentaje, 0) AS Porcentaje" & vbCrLf _
    ''''                        & ",isnull(pv.FechaInicio,'01/01/1900') AS FechaInicio" & vbCrLf _
    ''''                        & ",isnull(pv.FechaFin, '01/01/1900') AS FechaFin" & vbCrLf _
    ''''                        & ",Tipoprom  From PromocionesVentas pv" & vbCrLf _
    ''''
    '''    gStrSql = "Select CodArticulo,CodGrupo,CodFamilia,CodLinea,CodSubLinea,CodMarca,CodModelo From CatArticulos where CodArticulo =  " & CodArticulo
    '''
    '''    ModEstandar.BorraCmd
    '''    Cmd.CommandText = "dbo.UP_SELECT_DATOS"
    '''    Cmd.CommandType = adCmdStoredProc
    '''    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '''    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, gStrSql)
    '''    Set RsGral = Cmd.Execute
    '''
    '''    FlArmaquery = ""
    '''
    '''    If Not RsGral.EOF Then
    '''      If Val(RsGral!CodGrupo & "") = gCODJOYERIA Then
    '''          '--Grupo 1
    '''          FlArmaquery = "Select " & SELECT_CAMPOS & vbCrLf _
    ''''          & "inner join CatArticulos ca on ( " & vbCrLf _
    ''''                & "(ca.CodArticulo = pv.CodArticulo and TipoProm = 'A') " & vbCrLf _
    ''''                & "Or " & vbCrLf _
    ''''                & "(ca.CodGrupo = pv.CodGrupo " & vbCrLf _
    ''''                & "and isnull(ca.CodFamilia,0)  = isnull(pv.CodFamilia,0) " & vbCrLf _
    ''''                & "and TipoProm = 'G') " & vbCrLf _
    ''''                & "Or " & vbCrLf _
    ''''                & "(ca.CodGrupo = pv.CodGrupo " & vbCrLf _
    ''''                & "and isnull(ca.CodFamilia,0)  = isnull(pv.CodFamilia,0) " & vbCrLf _
    ''''                & "and isnull(ca.CodLinea,0)    = isnull(pv.CodLinea,0) " & vbCrLf _
    ''''                & "and TipoProm = 'G') " & vbCrLf _
    ''''                & "Or " & vbCrLf _
    ''''                & "(ca.CodGrupo = pv.CodGrupo " & vbCrLf _
    ''''                & "and isnull(ca.CodFamilia,0)  = isnull(pv.CodFamilia,0) " & vbCrLf _
    ''''                & "and isnull(ca.CodLinea,0)    = isnull(pv.CodLinea,0) " & vbCrLf _
    ''''                & "and isnull(ca.CodSubLinea,0) = isnull(pv.CodsubLinea,0) " & vbCrLf _
    ''''                & "and TipoProm = 'G'))"
    '''
    '''      ElseIf Val(RsGral!CodGrupo & "") = gCODRELOJERIA Then
    '''      '--Grupo 2
    '''       FlArmaquery = "Select " & SELECT_CAMPOS & vbCrLf _
    ''''       & "inner join CatArticulos ca on (" & vbCrLf _
    ''''                & "(ca.CodArticulo = pv.CodArticulo and TipoProm = 'A')" & vbCrLf _
    ''''                & "Or" & vbCrLf _
    ''''                & "(ca.CodGrupo = pv.CodGrupo" & vbCrLf _
    ''''                & "and ca.CodMarca = pv.CodMarca" & vbCrLf _
    ''''                & "and TipoProm = 'G')" & vbCrLf _
    ''''                & "Or" & vbCrLf _
    ''''                & "(ca.CodGrupo = pv.CodGrupo" & vbCrLf _
    ''''                & "and ca.CodMarca = pv.CodMarca" & vbCrLf _
    ''''                & "and isnull(ca.CodModelo,0) = isnull(pv.CodModelo,0)" & vbCrLf _
    ''''                & "and TipoProm = 'G'))"
    '''      ElseIf Val(RsGral!CodGrupo & "") = gCODVARIOS Then
    '''        '--Grupo 3
    '''        FlArmaquery = "Select " & SELECT_CAMPOS & vbCrLf _
    ''''       & "inner join CatArticulos ca on (" & vbCrLf _
    ''''                & "(ca.CodArticulo = pv.CodArticulo and TipoProm = 'A')" & vbCrLf _
    ''''                & "Or" & vbCrLf _
    ''''                & "(ca.CodGrupo = pv.CodGrupo" & vbCrLf _
    ''''                & "and isnull(ca.CodFamilia,0)  = isnull(pv.CodFamilia,0)" & vbCrLf _
    ''''                & "and TipoProm = 'G')" & vbCrLf _
    ''''                & "Or" & vbCrLf _
    ''''                & "(ca.CodGrupo = pv.CodGrupo" & vbCrLf _
    ''''                & "and isnull(ca.CodFamilia,0)  = isnull(pv.CodFamilia,0)" & vbCrLf _
    ''''                & "and isnull(ca.CodLinea,0)    = isnull(pv.CodLinea,0)" & vbCrLf _
    ''''                & "and TipoProm = 'G'))"
    '''
    '''      End If
    '''
    '''      FlArmaquery = FlArmaquery & vbCrLf & "Where ca.CodArticulo = " & CodArticulo & vbCrLf _
    ''''      & "And ca.CodGrupo = " & Val(RsGral!CodGrupo & "") & vbCrLf _
    ''''      & "And pv.Estatus = 'A' " & vbCrLf
    '''
    '''      FlArmaquery = FlArmaquery & IIf(Val(RsGral!CodFamilia & "") <> 0, " And ca.CodFamilia = " & RsGral!CodFamilia & "" & vbCrLf, "")
    '''      FlArmaquery = FlArmaquery & IIf(Val(RsGral!COdLinea & "") <> 0, "And ca.CodLinea = " & RsGral!COdLinea & "" & vbCrLf, "")
    '''      FlArmaquery = FlArmaquery & IIf(Val(RsGral!CodSubLinea & "") <> 0, "And ca.CodSubLinea = " & RsGral!CodSubLinea & "" & vbCrLf, "")
    '''      FlArmaquery = FlArmaquery & IIf(Val(RsGral!CodMArca & "") <> 0, "And ca.CodMarca = " & RsGral!CodMArca & "" & vbCrLf, "")
    '''      FlArmaquery = FlArmaquery & IIf(Val(RsGral!CodModelo & "") <> 0, "And ca.CodModelo = " & RsGral!CodModelo & "" & vbCrLf, "")
    '''      FlArmaquery = FlArmaquery & "And (convert(datetime,'" & Format(Date, "mm/dd/yyyy") & "',101) between FechaInicio and FechaFin or convert(datetime,'" & Format(Date, "mm/dd/yyyy") & "',101) <= FechaFin)" & vbCrLf _
    ''''      & "Order by  pv.FechaInicio"
    '''
    '''    Else
    '''      FlArmaquery = ""
    '''    End If
    '''    RsGral.Close
    '''
    '''    Exit Function
    '''Merr:
    '''    If Err.Number <> 0 Then ModEstandar.MostrarError
    '''End Function
End Class