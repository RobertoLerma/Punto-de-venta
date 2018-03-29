'**********************************************************************************************************************'
'*PROGRAMA: PROGRAMACIÓN DE PROMOCIONES JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.Compatibility
Imports ADODB
Imports Microsoft.VisualStudio.Data

Public Class frmProgramacionPromociones

    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer


    'Programa: ProGramacion de Promociones
    '    Elaboró: Rosaura Torres López
    '    Fecha:18/Julio/2003

    '    **************************************************************************************************************
    '     MODIFICACION A ABC - ASIGNACION DE DESCUENTOS POR PROVEEDOR
    '     SE AGREGARON 2 CAMPOS A LA ESTRUCTURA PARA PODER DISTINGUIR PROMOCIONES DIFERENTES DE UN MISMO PROV.
    '     20ABR2006 - MAVF
    '    **************************************************************************************************************
    Public WithEvents sstGrupos As System.Windows.Forms.TabControl
    Public WithEvents _sstGrupos_TabPage0 As System.Windows.Forms.TabPage
    Public WithEvents _sstGrupos_TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents _sstGrupos_TabPage2 As System.Windows.Forms.TabPage
    Public WithEvents _sstGrupos_TabPage3 As System.Windows.Forms.TabPage
    Public WithEvents _sstGrupos_TabPage4 As System.Windows.Forms.TabPage
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents dtpFechaInIcioJ As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinJ As System.Windows.Forms.DateTimePicker
    Public WithEvents dbcJFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcJLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcJSubLinea As System.Windows.Forms.ComboBox
    Public WithEvents txtJoyeria As System.Windows.Forms.TextBox
    Public WithEvents txtArticulo As System.Windows.Forms.TextBox
    Public WithEvents dbcJArticulo As System.Windows.Forms.ComboBox
    Public WithEvents Panel1 As Panel
    Public WithEvents txtDesArticulo As Label
    Public WithEvents Label3 As Label
    Public WithEvents Label1 As Label
    Public WithEvents lblVigente As System.Windows.Forms.Label
    Public WithEvents lblCancelada As System.Windows.Forms.Label
    Public WithEvents _lblOrden_15 As System.Windows.Forms.Label
    Public WithEvents _lblOrden_17 As System.Windows.Forms.Label
    Public WithEvents msgJoyeria As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents txtRelojeria As System.Windows.Forms.TextBox
    Public WithEvents dbcRMarca As System.Windows.Forms.ComboBox
    Public WithEvents dbcRArticulo As System.Windows.Forms.ComboBox
    Public WithEvents dbcRModelo As System.Windows.Forms.ComboBox
    Public WithEvents txtArticuloR As System.Windows.Forms.TextBox
    Public WithEvents msgRelojeria As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid



    'Public WithEvents dtpFechaInIcioV As AxMSComCtl2.AxDTPicker
    'Public WithEvents dtpFechaFinV As AxMSComCtl2.AxDTPicker
    'Public WithEvents dtpFechaInIcioR As AxMSComCtl2.AxDTPicker
    'Public WithEvents dtpFechaFinR As AxMSComCtl2.AxDTPicker
    'Public WithEvents _Label1_0 As System.Windows.Forms.Label
    'Public WithEvents Label2 As System.Windows.Forms.Label
    'Public WithEvents fraPeriodo As System.Windows.Forms.GroupBox 
    'Public WithEvents msgVarios As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    'Public WithEvents txtArticuloV As System.Windows.Forms.TextBox
    'Public WithEvents msgXArticulo As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    'Public WithEvents txtFlex As System.Windows.Forms.TextBox
    'Public WithEvents _lblOrden_0 As System.Windows.Forms.Label
    'Public WithEvents lblArtsNoSel As System.Windows.Forms.Label
    'Public WithEvents msgArtxProv As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    'Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    'Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    'Public WithEvents chkCancelarP As System.Windows.Forms.CheckBox
    'Public WithEvents chkBorrar As System.Windows.Forms.CheckBox
    'Public WithEvents chkAplicar As System.Windows.Forms.CheckBox
    'Public WithEvents txtDesctoP As System.Windows.Forms.TextBox
    'Public WithEvents lblDesctoP As System.Windows.Forms.Label
    'Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    'Public WithEvents txtDetArtxProv As System.Windows.Forms.TextBox
    'Public WithEvents lblTotArt As System.Windows.Forms.Label
    'Public WithEvents _lblOrden_1 As System.Windows.Forms.Label
    'Public WithEvents Frame5 As System.Windows.Forms.GroupBox

    'Public WithEvents lblVigente As System.Windows.Forms.Label
    'Public WithEvents _lblOrden_15 As System.Windows.Forms.Label
    'Public WithEvents _lblOrden_17 As System.Windows.Forms.Label
    'Public WithEvents lblCancelada As System.Windows.Forms.Label
    'Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'Public WithEvents lblOrden As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProgramacionPromociones))
        Me.sstGrupos = New System.Windows.Forms.TabControl()
        Me._sstGrupos_TabPage0 = New System.Windows.Forms.TabPage()
        Me.dbcJFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcJLinea = New System.Windows.Forms.ComboBox()
        Me.dbcJSubLinea = New System.Windows.Forms.ComboBox()
        Me.txtJoyeria = New System.Windows.Forms.TextBox()
        Me.dbcJArticulo = New System.Windows.Forms.ComboBox()
        Me.msgJoyeria = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._sstGrupos_TabPage1 = New System.Windows.Forms.TabPage()
        Me.dbcRMarca = New System.Windows.Forms.ComboBox()
        Me.dbcRModelo = New System.Windows.Forms.ComboBox()
        Me.txtRelojeria = New System.Windows.Forms.TextBox()
        Me.dbcRArticulo = New System.Windows.Forms.ComboBox()
        Me.txtArticuloR = New System.Windows.Forms.TextBox()
        Me.msgRelojeria = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._sstGrupos_TabPage2 = New System.Windows.Forms.TabPage()
        Me._sstGrupos_TabPage3 = New System.Windows.Forms.TabPage()
        Me._sstGrupos_TabPage4 = New System.Windows.Forms.TabPage()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtDesArticulo = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtpFechaInIcioJ = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinJ = New System.Windows.Forms.DateTimePicker()
        Me.txtArticulo = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblVigente = New System.Windows.Forms.Label()
        Me.lblCancelada = New System.Windows.Forms.Label()
        Me._lblOrden_15 = New System.Windows.Forms.Label()
        Me._lblOrden_17 = New System.Windows.Forms.Label()
        Me.sstGrupos.SuspendLayout()
        Me._sstGrupos_TabPage0.SuspendLayout()
        CType(Me.msgJoyeria, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._sstGrupos_TabPage1.SuspendLayout()
        CType(Me.msgRelojeria, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'sstGrupos
        '
        Me.sstGrupos.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.sstGrupos.Controls.Add(Me._sstGrupos_TabPage0)
        Me.sstGrupos.Controls.Add(Me._sstGrupos_TabPage1)
        Me.sstGrupos.Controls.Add(Me._sstGrupos_TabPage2)
        Me.sstGrupos.Controls.Add(Me._sstGrupos_TabPage3)
        Me.sstGrupos.Controls.Add(Me._sstGrupos_TabPage4)
        Me.sstGrupos.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.sstGrupos.ItemSize = New System.Drawing.Size(42, 18)
        Me.sstGrupos.Location = New System.Drawing.Point(12, 73)
        Me.sstGrupos.Name = "sstGrupos"
        Me.sstGrupos.SelectedIndex = 0
        Me.sstGrupos.Size = New System.Drawing.Size(556, 306)
        Me.sstGrupos.TabIndex = 9
        '
        '_sstGrupos_TabPage0
        '
        Me._sstGrupos_TabPage0.Controls.Add(Me.dbcJFamilia)
        Me._sstGrupos_TabPage0.Controls.Add(Me.dbcJLinea)
        Me._sstGrupos_TabPage0.Controls.Add(Me.dbcJSubLinea)
        Me._sstGrupos_TabPage0.Controls.Add(Me.txtJoyeria)
        Me._sstGrupos_TabPage0.Controls.Add(Me.dbcJArticulo)
        Me._sstGrupos_TabPage0.Controls.Add(Me.msgJoyeria)
        Me._sstGrupos_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._sstGrupos_TabPage0.Name = "_sstGrupos_TabPage0"
        Me._sstGrupos_TabPage0.Size = New System.Drawing.Size(548, 280)
        Me._sstGrupos_TabPage0.TabIndex = 0
        Me._sstGrupos_TabPage0.Text = "Joyería"
        '
        'dbcJFamilia
        '
        Me.dbcJFamilia.Location = New System.Drawing.Point(9, 51)
        Me.dbcJFamilia.Name = "dbcJFamilia"
        Me.dbcJFamilia.Size = New System.Drawing.Size(66, 21)
        Me.dbcJFamilia.TabIndex = 10
        Me.dbcJFamilia.Visible = False
        '
        'dbcJLinea
        '
        Me.dbcJLinea.Location = New System.Drawing.Point(75, 51)
        Me.dbcJLinea.Name = "dbcJLinea"
        Me.dbcJLinea.Size = New System.Drawing.Size(63, 21)
        Me.dbcJLinea.TabIndex = 11
        Me.dbcJLinea.Visible = False
        '
        'dbcJSubLinea
        '
        Me.dbcJSubLinea.Location = New System.Drawing.Point(138, 51)
        Me.dbcJSubLinea.Name = "dbcJSubLinea"
        Me.dbcJSubLinea.Size = New System.Drawing.Size(65, 21)
        Me.dbcJSubLinea.TabIndex = 12
        Me.dbcJSubLinea.Visible = False
        '
        'txtJoyeria
        '
        Me.txtJoyeria.AcceptsReturn = True
        Me.txtJoyeria.BackColor = System.Drawing.SystemColors.Window
        Me.txtJoyeria.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtJoyeria.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtJoyeria.Location = New System.Drawing.Point(201, 52)
        Me.txtJoyeria.MaxLength = 0
        Me.txtJoyeria.Name = "txtJoyeria"
        Me.txtJoyeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtJoyeria.Size = New System.Drawing.Size(64, 20)
        Me.txtJoyeria.TabIndex = 13
        Me.txtJoyeria.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtJoyeria.Visible = False
        '
        'dbcJArticulo
        '
        Me.dbcJArticulo.Location = New System.Drawing.Point(330, 51)
        Me.dbcJArticulo.Name = "dbcJArticulo"
        Me.dbcJArticulo.Size = New System.Drawing.Size(65, 21)
        Me.dbcJArticulo.TabIndex = 26
        Me.dbcJArticulo.Visible = False
        '
        'msgJoyeria
        '
        Me.msgJoyeria.DataSource = Nothing
        Me.msgJoyeria.Location = New System.Drawing.Point(8, 29)
        Me.msgJoyeria.Name = "msgJoyeria"
        Me.msgJoyeria.OcxState = CType(resources.GetObject("msgJoyeria.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msgJoyeria.Size = New System.Drawing.Size(526, 236)
        Me.msgJoyeria.TabIndex = 6
        '
        '_sstGrupos_TabPage1
        '
        'Me._sstGrupos_TabPage1.Controls.Add(Me.dbcRMarca)
        'Me._sstGrupos_TabPage1.Controls.Add(Me.dbcRModelo)
        'Me._sstGrupos_TabPage1.Controls.Add(Me.txtRelojeria)
        'Me._sstGrupos_TabPage1.Controls.Add(Me.dbcRArticulo)
        'Me._sstGrupos_TabPage1.Controls.Add(Me.txtArticuloR)
        'Me._sstGrupos_TabPage1.Controls.Add(Me.msgRelojeria)
        Me._sstGrupos_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._sstGrupos_TabPage1.Name = "_sstGrupos_TabPage1"
        Me._sstGrupos_TabPage1.Size = New System.Drawing.Size(548, 280)
        Me._sstGrupos_TabPage1.TabIndex = 1
        Me._sstGrupos_TabPage1.Text = "Relojería"
        '
        'dbcRMarca
        '
        Me.dbcRMarca.Location = New System.Drawing.Point(10, 51)
        Me.dbcRMarca.Name = "dbcRMarca"
        Me.dbcRMarca.Size = New System.Drawing.Size(66, 21)
        Me.dbcRMarca.TabIndex = 15
        Me.dbcRMarca.Visible = False
        '
        'dbcRModelo
        '
        Me.dbcRModelo.Location = New System.Drawing.Point(75, 51)
        Me.dbcRModelo.Name = "dbcRModelo"
        Me.dbcRModelo.Size = New System.Drawing.Size(66, 21)
        Me.dbcRModelo.TabIndex = 16
        Me.dbcRModelo.Visible = False
        '
        'txtRelojeria
        '
        Me.txtRelojeria.AcceptsReturn = True
        Me.txtRelojeria.BackColor = System.Drawing.SystemColors.Window
        Me.txtRelojeria.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRelojeria.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRelojeria.Location = New System.Drawing.Point(140, 52)
        Me.txtRelojeria.MaxLength = 0
        Me.txtRelojeria.Name = "txtRelojeria"
        Me.txtRelojeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRelojeria.Size = New System.Drawing.Size(62, 20)
        Me.txtRelojeria.TabIndex = 14
        Me.txtRelojeria.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtRelojeria.Visible = False
        '
        'dbcRArticulo
        '
        Me.dbcRArticulo.Location = New System.Drawing.Point(202, 51)
        Me.dbcRArticulo.Name = "dbcRArticulo"
        Me.dbcRArticulo.Size = New System.Drawing.Size(65, 21)
        Me.dbcRArticulo.TabIndex = 27
        Me.dbcRArticulo.Visible = False
        '
        'txtArticuloR
        '
        Me.txtArticuloR.AcceptsReturn = True
        Me.txtArticuloR.BackColor = System.Drawing.SystemColors.Window
        Me.txtArticuloR.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtArticuloR.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtArticuloR.Location = New System.Drawing.Point(267, 51)
        Me.txtArticuloR.MaxLength = 0
        Me.txtArticuloR.Name = "txtArticuloR"
        Me.txtArticuloR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtArticuloR.Size = New System.Drawing.Size(64, 21)
        Me.txtArticuloR.TabIndex = 29
        Me.txtArticuloR.Visible = False
        '
        'msgRelojeria
        '
        Me.msgRelojeria.DataSource = Nothing
        Me.msgRelojeria.Location = New System.Drawing.Point(9, 29)
        Me.msgRelojeria.Name = "msgRelojeria"
        Me.msgRelojeria.OcxState = CType(resources.GetObject("msgRelojeria.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msgRelojeria.Size = New System.Drawing.Size(523, 235)
        Me.msgRelojeria.TabIndex = 7
        '
        '_sstGrupos_TabPage2
        '
        Me._sstGrupos_TabPage2.Location = New System.Drawing.Point(4, 22)
        Me._sstGrupos_TabPage2.Name = "_sstGrupos_TabPage2"
        Me._sstGrupos_TabPage2.Size = New System.Drawing.Size(548, 280)
        Me._sstGrupos_TabPage2.TabIndex = 2
        Me._sstGrupos_TabPage2.Text = "Varios"
        '
        '_sstGrupos_TabPage3
        '
        Me._sstGrupos_TabPage3.Location = New System.Drawing.Point(4, 22)
        Me._sstGrupos_TabPage3.Name = "_sstGrupos_TabPage3"
        Me._sstGrupos_TabPage3.Size = New System.Drawing.Size(548, 280)
        Me._sstGrupos_TabPage3.TabIndex = 2
        Me._sstGrupos_TabPage3.Text = "X Articulo"
        '
        '_sstGrupos_TabPage4
        '
        Me._sstGrupos_TabPage4.Location = New System.Drawing.Point(4, 22)
        Me._sstGrupos_TabPage4.Name = "_sstGrupos_TabPage4"
        Me._sstGrupos_TabPage4.Size = New System.Drawing.Size(548, 280)
        Me._sstGrupos_TabPage4.TabIndex = 2
        Me._sstGrupos_TabPage4.Text = "Articulos X Proveedor"
        '
        'txtDesArticulo
        '
        Me.txtDesArticulo.BackColor = System.Drawing.SystemColors.Info
        Me.txtDesArticulo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDesArticulo.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDesArticulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtDesArticulo.Location = New System.Drawing.Point(184, 424)
        Me.txtDesArticulo.Name = "txtDesArticulo"
        Me.txtDesArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDesArticulo.Size = New System.Drawing.Size(249, 21)
        Me.txtDesArticulo.TabIndex = 25
        Me.txtDesArticulo.Text = "[SEL-SUPR] = Cancelar Bloque"
        Me.txtDesArticulo.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolTip1.SetToolTip(Me.txtDesArticulo, "Descripción de Artículos")
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Info
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(184, 394)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(249, 21)
        Me.Label3.TabIndex = 47
        Me.Label3.Text = "[SUPR] = Cancelar Promoción"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolTip1.SetToolTip(Me.Label3, "Descripción de Artículos")
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(116, 20)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(56, 16)
        Me._Label1_0.TabIndex = 1
        Me._Label1_0.Text = "Desde el :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(324, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(56, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Hasta el :"
        '
        'dtpFechaInIcioJ
        '
        Me.dtpFechaInIcioJ.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInIcioJ.Location = New System.Drawing.Point(172, 16)
        Me.dtpFechaInIcioJ.Name = "dtpFechaInIcioJ"
        Me.dtpFechaInIcioJ.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaInIcioJ.TabIndex = 2
        '
        'dtpFechaFinJ
        '
        Me.dtpFechaFinJ.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinJ.Location = New System.Drawing.Point(380, 16)
        Me.dtpFechaFinJ.Name = "dtpFechaFinJ"
        Me.dtpFechaFinJ.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaFinJ.TabIndex = 4
        '
        'txtArticulo
        '
        Me.txtArticulo.AcceptsReturn = True
        Me.txtArticulo.BackColor = System.Drawing.SystemColors.Window
        Me.txtArticulo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtArticulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtArticulo.Location = New System.Drawing.Point(265, 51)
        Me.txtArticulo.MaxLength = 0
        Me.txtArticulo.Name = "txtArticulo"
        Me.txtArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtArticulo.Size = New System.Drawing.Size(65, 20)
        Me.txtArticulo.TabIndex = 28
        Me.txtArticulo.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me._Label1_0)
        Me.Panel1.Controls.Add(Me.dtpFechaInIcioJ)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.dtpFechaFinJ)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(556, 55)
        Me.Panel1.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 13)
        Me.Label1.TabIndex = 48
        Me.Label1.Text = "Vigencia de la oferta"
        '
        'lblVigente
        '
        Me.lblVigente.BackColor = System.Drawing.Color.White
        Me.lblVigente.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblVigente.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVigente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVigente.Location = New System.Drawing.Point(24, 398)
        Me.lblVigente.Name = "lblVigente"
        Me.lblVigente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVigente.Size = New System.Drawing.Size(17, 17)
        Me.lblVigente.TabIndex = 20
        '
        'lblCancelada
        '
        Me.lblCancelada.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(200, Byte), Integer), CType(CType(145, Byte), Integer))
        Me.lblCancelada.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCancelada.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCancelada.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCancelada.Location = New System.Drawing.Point(24, 420)
        Me.lblCancelada.Name = "lblCancelada"
        Me.lblCancelada.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCancelada.Size = New System.Drawing.Size(17, 17)
        Me.lblCancelada.TabIndex = 17
        '
        '_lblOrden_15
        '
        Me._lblOrden_15.AutoSize = True
        Me._lblOrden_15.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_15.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblOrden_15.Location = New System.Drawing.Point(47, 398)
        Me._lblOrden_15.Name = "_lblOrden_15"
        Me._lblOrden_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_15.Size = New System.Drawing.Size(102, 13)
        Me._lblOrden_15.TabIndex = 19
        Me._lblOrden_15.Text = "Vigentes sin Aplicar "
        '
        '_lblOrden_17
        '
        Me._lblOrden_17.AutoSize = True
        Me._lblOrden_17.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_17.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblOrden_17.Location = New System.Drawing.Point(47, 420)
        Me._lblOrden_17.Name = "_lblOrden_17"
        Me._lblOrden_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_17.Size = New System.Drawing.Size(63, 13)
        Me._lblOrden_17.TabIndex = 18
        Me._lblOrden_17.Text = "Canceladas"
        '
        'frmProgramacionPromociones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(587, 480)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDesArticulo)
        Me.Controls.Add(Me.lblVigente)
        Me.Controls.Add(Me.lblCancelada)
        Me.Controls.Add(Me._lblOrden_15)
        Me.Controls.Add(Me._lblOrden_17)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.sstGrupos)
        Me._sstGrupos_TabPage1.Controls.Add(Me.dbcRMarca)
        Me._sstGrupos_TabPage1.Controls.Add(Me.dbcRModelo)
        Me._sstGrupos_TabPage1.Controls.Add(Me.txtRelojeria)
        Me._sstGrupos_TabPage1.Controls.Add(Me.dbcRArticulo)
        Me._sstGrupos_TabPage1.Controls.Add(Me.txtArticuloR)
        Me._sstGrupos_TabPage1.Controls.Add(Me.msgRelojeria)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(298, 150)
        Me.MaximizeBox = False
        Me.Name = "frmProgramacionPromociones"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Programación de Promociones"
        Me.sstGrupos.ResumeLayout(False)
        Me._sstGrupos_TabPage0.ResumeLayout(False)
        Me._sstGrupos_TabPage0.PerformLayout()
        CType(Me.msgJoyeria, System.ComponentModel.ISupportInitialize).EndInit()
        Me._sstGrupos_TabPage1.ResumeLayout(False)
        Me._sstGrupos_TabPage1.PerformLayout()
        CType(Me.msgRelojeria, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub






    '    Dim mblnSalir As Boolean
    '    Dim FueraChange As Boolean
    '    Dim intCodFamilia As Integer
    '    Dim intCodLinea As Integer
    '    Dim intCodSubLinea As Integer
    '    Dim intCodMarca As Integer
    '    Dim intCodModelo As Integer
    '    Dim intCodArticulo As Integer
    '    Dim tecla As Integer
    '    Dim I As Integer
    '    Dim mblnNuevo As Boolean
    '    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    '    Dim mintCodProveedor As Integer
    '    Dim mintcodRenglon As Integer
    '    Dim mblnFueraChange As Boolean
    '    Dim mintTotalRen As Integer

    '    Public GridACtivo As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    '    Public mProveedor As Integer
    '    Public mrenProv As Integer

    '    'Para Manejar el gRID DE JOYERIA
    '    Const C_ColJFAMILIA As Integer = 0
    '    Const C_ColJLINEA As Integer = 1
    '    Const C_ColJSUBLINEA As Integer = 2
    '    Const C_ColJARTICULO As Integer = 3
    '    Const C_ColJPORCDESCTO As Integer = 4
    '    Const C_ColJPRECIO As Integer = 5
    '    Const C_ColJCODFAMILIA As Integer = 6
    '    Const C_ColJCODLINEA As Integer = 7
    '    Const C_ColJCODSUBLINEA As Integer = 8
    '    Const C_ColJCODARTICULO As Integer = 9
    '    Const C_ColJPORCDESCTOTAG As Integer = 10
    '    Const C_ColJPRECIOTAG As Integer = 11
    '    Const C_ColJESNUEVO As Integer = 12
    '    Const C_ColJESTATUS As Integer = 13
    '    Const C_ColJESTATUSTAG As Integer = 14
    '    Const C_COLJTIPO As Integer = 15

    '    'Para Manejar el gRID DE RELOJERIA

    '    Const C_ColRMARCA As Integer = 0
    '    Const C_ColRMODELO As Integer = 1
    '    Const C_ColRARTICULO As Integer = 2
    '    Const C_ColRPORCDESCTO As Integer = 3
    '    Const C_ColRPRECIO As Integer = 4
    '    Const C_ColRCODMARCA As Integer = 5
    '    Const C_ColRCODMODELO As Integer = 6
    '    Const C_ColRCODARTICULO As Integer = 7
    '    Const C_ColRPORCDESCTOTAG As Integer = 8
    '    Const C_ColRPRECIOTAG As Integer = 9
    '    Const C_ColRESNUEVO As Integer = 10
    '    Const C_ColRESTATUS As Integer = 11
    '    Const C_ColRESTATUSTAG As Integer = 12
    '    Const C_COLRTIPO As Integer = 13

    '    'Para Manejar el Grid por Articulos
    '    Const C_COLXARTCODARTICULO As Integer = 0
    '    Const C_COLXARTDESCARTICULO As Integer = 1
    '    Const C_COLXARTCODANTERIOR As Integer = 2
    '    Const C_COLXARTPORCDESCTO As Integer = 3
    '    Const C_COLXARTPRECIO As Integer = 4
    '    Const C_COLXARTCODGRUPO As Integer = 5
    '    Const C_COLXARTCODFAMILIA As Integer = 6
    '    Const C_COLXARTCODLINEA As Integer = 7
    '    Const C_COLXARTCODSUBLINEA As Integer = 8
    '    Const C_COLXARTCODMARCA As Integer = 9
    '    Const C_COLXARTCODMODELO As Integer = 10
    '    Const C_COLXARTESNUEVO As Integer = 11
    '    Const C_COLXARTPRECIOTAG As Integer = 12
    '    Const C_COLXARTPORCDESCTOTAG As Integer = 13
    '    Const C_COLXARTESTATUS As Integer = 14
    '    Const C_COLXARTESTATUSTAG As Integer = 15
    '    Const C_COLXARTTIPO As Integer = 16

    '    'Para Manejar el Grid por Articulos x Proveedor
    '    Const C_COLXPRVCODARTICULO As Integer = 0
    '    Const C_COLXPRVDESCARTICULO As Integer = 1
    '    Const C_COLXPRVCODANTERIOR As Integer = 2
    '    Const C_COLXPRVPORCDESCTO As Integer = 3
    '    Const C_COLXPRVPRECIO As Integer = 4
    '    Const C_COLXPRVCODGRUPO As Integer = 5
    '    Const C_COLXPRVCODFAMILIA As Integer = 6
    '    Const C_COLXPRVCODLINEA As Integer = 7
    '    Const C_COLXPRVCODSUBLINEA As Integer = 8
    '    Const C_COLXPRVCODMARCA As Integer = 9
    '    Const C_COLXPRVCODMODELO As Integer = 10
    '    Const C_COLXPRVESNUEVO As Integer = 11
    '    Const C_COLXPRVPRECIOTAG As Integer = 12
    '    Const C_COLXPRVPORCDESCTOTAG As Integer = 13
    '    Const C_COLXPRVESTATUS As Integer = 14
    '    Const C_COLXPRVESTATUSTAG As Integer = 15
    '    Const C_COLXPRVTIPO As Integer = 16

    '    'COnstates para el Estatus
    '    Const C_Cancelado As String = "C"
    '    Const C_Aplicado As String = "A"
    '    Const C_Vigente As String = "V"

    '    Sub DesHabilitarFechas()
    '        dtpFechaFinJ.Enabled = False
    '        dtpFechaFinV.Enabled = False
    '        dtpFechaFinR.Enabled = False
    '        dtpFechaInIcioJ.Enabled = False
    '        dtpFechaInIcioR.Enabled = False
    '        dtpFechaInIcioV.Enabled = False
    '    End Sub

    '    Sub HabilitarFechas()
    '        dtpFechaFinJ.Enabled = True
    '        dtpFechaFinV.Enabled = True
    '        dtpFechaFinR.Enabled = True
    '        dtpFechaInIcioJ.Enabled = True
    '        dtpFechaInIcioR.Enabled = True
    '        dtpFechaInIcioV.Enabled = True
    '    End Sub
    '    Sub Encabezado()
    '        On Error GoTo Merr
    '        'Genera el encabezao del Grid, asigna el tamaño y número de columas y centra el texto dentro de ellas
    '        Dim LnContador As Integer

    '        With msgJoyeria
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusHeavy 'flexFocusLight 'flexFocusNone
    '            .FixedRows = 1
    '            .FixedCols = 0
    '            .set_ColWidth(C_ColJFAMILIA, 0, 2000)
    '            .set_ColWidth(C_ColJLINEA, 0, 2000)
    '            .set_ColWidth(C_ColJSUBLINEA, 0, 2000)
    '            .set_ColWidth(C_ColJARTICULO, 0, 3810)
    '            .set_ColWidth(C_ColJPRECIO, 0, 1760)
    '            .set_ColWidth(C_ColJPORCDESCTO, 0, 1200)
    '            .set_ColWidth(C_ColJCODFAMILIA, 0, 1)
    '            .set_ColWidth(C_ColJCODLINEA, 0, 1)
    '            .set_ColWidth(C_ColJCODSUBLINEA, 0, 1)
    '            .set_ColWidth(C_ColJCODARTICULO, 0, 1)
    '            .set_ColWidth(C_ColJPORCDESCTOTAG, 0, 1)
    '            .set_ColWidth(C_ColJPRECIOTAG, 0, 1)
    '            .set_ColWidth(C_ColJESNUEVO, 0, 1)
    '            .set_ColWidth(C_ColJESTATUS, 0, 1)
    '            .set_ColWidth(C_ColJESTATUSTAG, 0, 1)
    '            .set_ColWidth(C_COLJTIPO, 0, 1)

    '            .set_TextMatrix(0, C_ColJFAMILIA, "Familia")
    '            .set_TextMatrix(0, C_ColJLINEA, "Línea")
    '            .set_TextMatrix(0, C_ColJSUBLINEA, "SubLínea")
    '            .set_TextMatrix(0, C_ColJARTICULO, "Artículo")
    '            .set_TextMatrix(0, C_ColJPRECIO, "Precio")
    '            .set_TextMatrix(0, C_ColJPORCDESCTO, "% Descto")
    '            .set_TextMatrix(0, C_ColJCODFAMILIA, "CodFamilia")
    '            .set_TextMatrix(0, C_ColJCODLINEA, "CodLinea")
    '            .set_TextMatrix(0, C_ColJCODSUBLINEA, "CodSubLinea")
    '            .set_TextMatrix(0, C_ColJCODARTICULO, "CodArticulo")
    '            .set_TextMatrix(0, C_ColJPORCDESCTOTAG, "Porc")
    '            .set_TextMatrix(0, C_ColJPRECIOTAG, "PrecioTag")
    '            .set_TextMatrix(0, C_ColJESNUEVO, "Es Nuevo")
    '            .set_TextMatrix(0, C_ColJESTATUS, "Estatus")
    '            .set_TextMatrix(0, C_ColJESTATUSTAG, "Estatustag")

    '            .set_ColAlignment(C_ColJFAMILIA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_ColJLINEA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_ColJSUBLINEA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_ColJARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_ColJPORCDESCTO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '            .set_ColAlignment(C_ColJPRECIO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)

    '            .Row = 0
    '            For LnContador = 0 To C_ColJESTATUS
    '                .Col = LnContador
    '                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
    '                .CellFontBold = True
    '            Next LnContador
    '            For LnContador = 1 To .Rows - 1
    '                .set_TextMatrix(LnContador, C_COLJTIPO, "G")
    '            Next
    '            .Row = 1
    '            .Col = C_ColJFAMILIA
    '            .WordWrap = False 'Hacer esto , para que no se puedan escribir dos o mal lineas de texto en una  sola fila, solo se usa para el encabezado
    '        End With

    '        'ENCABEZADO DE rELOJERIA
    '        With msgRelojeria
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusHeavy 'flexFocusLight 'flexFocusNone
    '            .FixedRows = 1
    '            .FixedCols = 0
    '            .set_ColWidth(C_ColRMARCA, 0, 2100)
    '            .set_ColWidth(C_ColRMODELO, 0, 2100)
    '            .set_ColWidth(C_ColRARTICULO, 0, 4950)
    '            .set_ColWidth(C_ColRPRECIO, 0, 2000)
    '            .set_ColWidth(C_ColRPORCDESCTO, 0, 1600)
    '            .set_ColWidth(C_ColRCODMARCA, 0, 1)
    '            .set_ColWidth(C_ColRCODMODELO, 0, 1)
    '            .set_ColWidth(C_ColRCODARTICULO, 0, 1)
    '            .set_ColWidth(C_ColRPORCDESCTOTAG, 0, 1)
    '            .set_ColWidth(C_ColRPRECIOTAG, 0, 1)
    '            .set_ColWidth(C_ColRESNUEVO, 0, 1)
    '            .set_ColWidth(C_ColRESTATUS, 0, 1)
    '            .set_ColWidth(C_ColRESTATUSTAG, 0, 1)
    '            .set_ColWidth(C_COLRTIPO, 0, 1)

    '            .set_TextMatrix(0, C_ColRMARCA, "Marca")
    '            .set_TextMatrix(0, C_ColRMODELO, "Modelo")
    '            .set_TextMatrix(0, C_ColRARTICULO, "Articulo")
    '            .set_TextMatrix(0, C_ColRPRECIO, "Precio")
    '            .set_TextMatrix(0, C_ColRPORCDESCTO, "% Descto")
    '            .set_TextMatrix(0, C_ColRCODMARCA, "CodMarca")
    '            .set_TextMatrix(0, C_ColRCODMODELO, "CodModelo")
    '            .set_TextMatrix(0, C_ColRCODARTICULO, "CodArticulo")
    '            .set_TextMatrix(0, C_ColRPORCDESCTOTAG, "Porc")
    '            .set_TextMatrix(0, C_ColRPRECIOTAG, "PrecioTag")
    '            .set_TextMatrix(0, C_ColRESNUEVO, "Es Nuevo")
    '            .set_TextMatrix(0, C_ColRESTATUS, "Estatus")
    '            .set_TextMatrix(0, C_ColRESTATUSTAG, "EstatusTag")

    '            .set_ColAlignment(C_ColRMARCA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_ColRMODELO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_ColRARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_ColRPORCDESCTO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '            .set_ColAlignment(C_ColRPRECIO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)

    '            .Row = 0
    '            For LnContador = 0 To C_ColRESTATUS
    '                .Col = LnContador
    '                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
    '                .CellFontBold = True
    '            Next LnContador
    '            For LnContador = 1 To .Rows - 1
    '                .set_TextMatrix(LnContador, C_COLRTIPO, "G")
    '            Next
    '            .TopRow = 1
    '            .Row = 1
    '            .Col = C_ColJFAMILIA
    '            .WordWrap = False 'Hacer esto , para que no se puedan escribir dos o mal lineas de texto en una  sola fila, solo se usa para el encabezado
    '        End With

    '        'Encabezado de Varios
    '        With msgVarios
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusHeavy 'flexFocusLight 'flexFocusNone
    '            .FixedRows = 1
    '            .FixedCols = 0
    '            .set_ColWidth(C_ColJFAMILIA, 0, 2100)
    '            .set_ColWidth(C_ColJLINEA, 0, 2100)
    '            .set_ColWidth(C_ColJSUBLINEA, 0, 1)
    '            .set_ColWidth(C_ColJARTICULO, 0, 4950)
    '            .set_ColWidth(C_ColJPRECIO, 0, 2000)
    '            .set_ColWidth(C_ColJPORCDESCTO, 0, 1600)
    '            .set_ColWidth(C_ColJCODFAMILIA, 0, 1)
    '            .set_ColWidth(C_ColJCODLINEA, 0, 1)
    '            .set_ColWidth(C_ColJCODSUBLINEA, 0, 1)
    '            .set_ColWidth(C_ColJCODARTICULO, 0, 1)
    '            .set_ColWidth(C_ColJPORCDESCTOTAG, 0, 1)
    '            .set_ColWidth(C_ColJPRECIOTAG, 0, 1)
    '            .set_ColWidth(C_ColJESNUEVO, 0, 1)
    '            .set_ColWidth(C_ColJESTATUS, 0, 1)
    '            .set_ColWidth(C_ColJESTATUSTAG, 0, 1)
    '            .set_ColWidth(C_COLJTIPO, 0, 1)

    '            .set_TextMatrix(0, C_ColJFAMILIA, "Familia")
    '            .set_TextMatrix(0, C_ColJLINEA, "Línea")
    '            .set_TextMatrix(0, C_ColJSUBLINEA, "SubLínea")
    '            .set_TextMatrix(0, C_ColJARTICULO, "Artículo")
    '            .set_TextMatrix(0, C_ColJPRECIO, "Precio")
    '            .set_TextMatrix(0, C_ColJPORCDESCTO, "% Descto")
    '            .set_TextMatrix(0, C_ColJCODFAMILIA, "CodFamilia")
    '            .set_TextMatrix(0, C_ColJCODLINEA, "CodLinea")
    '            .set_TextMatrix(0, C_ColJCODSUBLINEA, "CodSubLinea")
    '            .set_TextMatrix(0, C_ColJCODARTICULO, "CodArticulo")
    '            .set_TextMatrix(0, C_ColJPORCDESCTOTAG, "Porc")
    '            .set_TextMatrix(0, C_ColJPRECIOTAG, "PrecioTag")
    '            .set_TextMatrix(0, C_ColJESNUEVO, "Es Nuevo")
    '            .set_TextMatrix(0, C_ColJESTATUS, "Estatus")
    '            .set_TextMatrix(0, C_ColJESTATUSTAG, "Estatustag ")

    '            .set_ColAlignment(C_ColJFAMILIA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_ColJLINEA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_ColJARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_ColJPORCDESCTO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '            .set_ColAlignment(C_ColJPRECIO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)

    '            .Row = 0
    '            For LnContador = 0 To C_ColJESTATUS
    '                .Col = LnContador
    '                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
    '                .CellFontBold = True
    '            Next LnContador
    '            For LnContador = 1 To .Rows - 1
    '                .set_TextMatrix(LnContador, C_COLJTIPO, "G")
    '            Next
    '            .TopRow = 1
    '            .Row = 1
    '            .Col = C_ColJFAMILIA
    '            .WordWrap = False 'Hacer esto , para que no se puedan escribir dos o mal lineas de texto en una  sola fila, solo se usa para el encabezado
    '        End With

    '        'Encabezado X Articulo
    '        With msgXArticulo
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '            .FixedRows = 1
    '            .FixedCols = 0
    '            .set_ColWidth(C_COLXARTCODARTICULO, 0, 1600)
    '            .set_ColWidth(C_COLXARTDESCARTICULO, 0, 5935)
    '            .set_ColWidth(C_COLXARTCODANTERIOR, 0, 1600)
    '            .set_ColWidth(C_COLXARTPORCDESCTO, 0, 1600)
    '            .set_ColWidth(C_COLXARTPRECIO, 0, 2000)
    '            .set_ColWidth(C_COLXARTCODFAMILIA, 0, 0)
    '            .set_ColWidth(C_COLXARTCODGRUPO, 0, 0)
    '            .set_ColWidth(C_COLXARTCODLINEA, 0, 0)
    '            .set_ColWidth(C_COLXARTCODMARCA, 0, 0)
    '            .set_ColWidth(C_COLXARTCODMODELO, 0, 0)
    '            .set_ColWidth(C_COLXARTCODSUBLINEA, 0, 0)
    '            .set_ColWidth(C_COLXARTESNUEVO, 0, 0)
    '            .set_ColWidth(C_COLXARTESTATUS, 0, 0)
    '            .set_ColWidth(C_COLXARTESTATUSTAG, 0, 0)
    '            .set_ColWidth(C_COLXARTPORCDESCTOTAG, 0, 0)
    '            .set_ColWidth(C_COLXARTPRECIOTAG, 0, 0)
    '            .set_ColWidth(C_COLXARTTIPO, 0, 0)
    '            .set_TextMatrix(0, C_COLXARTCODARTICULO, "Código")
    '            .set_TextMatrix(0, C_COLXARTDESCARTICULO, "Artículo")
    '            .set_TextMatrix(0, C_COLXARTCODANTERIOR, "Anterior")
    '            .set_TextMatrix(0, C_COLXARTPORCDESCTO, "% Descto.")
    '            .set_TextMatrix(0, C_COLXARTPRECIO, "Precio")
    '            .Row = 0

    '            .set_ColAlignment(C_COLXARTCODARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '            .set_ColAlignment(C_COLXARTDESCARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_COLXARTCODANTERIOR, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '            .set_ColAlignment(C_COLXARTPORCDESCTO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '            .set_ColAlignment(C_COLXARTPRECIO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)

    '            For LnContador = 0 To C_COLXARTPRECIO
    '                .Col = LnContador
    '                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
    '                .CellFontBold = True
    '            Next LnContador
    '            For LnContador = 1 To .Rows - 1
    '                .set_TextMatrix(LnContador, C_COLXARTTIPO, "A")
    '            Next
    '            .TopRow = 1
    '            .Row = 1
    '            .Col = C_COLXARTCODARTICULO
    '            .WordWrap = False
    '        End With

    '        With msgArtxProv
    '            .FixedRows = 1
    '            .FixedCols = 0
    '            .set_ColWidth(C_COLXPRVCODARTICULO, 0, 1600)
    '            .set_ColWidth(C_COLXPRVDESCARTICULO, 0, 5935)
    '            .set_ColWidth(C_COLXPRVCODANTERIOR, 0, 1600)
    '            .set_ColWidth(C_COLXPRVPORCDESCTO, 0, 1600)
    '            .set_ColWidth(C_COLXPRVPRECIO, 0, 0)
    '            .set_ColWidth(C_COLXPRVCODFAMILIA, 0, 0)
    '            .set_ColWidth(C_COLXPRVCODGRUPO, 0, 0)
    '            .set_ColWidth(C_COLXPRVCODLINEA, 0, 0)
    '            .set_ColWidth(C_COLXPRVCODMARCA, 0, 0)
    '            .set_ColWidth(C_COLXPRVCODMODELO, 0, 0)
    '            .set_ColWidth(C_COLXPRVCODSUBLINEA, 0, 0)
    '            .set_ColWidth(C_COLXPRVESNUEVO, 0, 0)
    '            .set_ColWidth(C_COLXPRVESTATUS, 0, 0)
    '            .set_ColWidth(C_COLXPRVESTATUSTAG, 0, 0)
    '            .set_ColWidth(C_COLXPRVPORCDESCTOTAG, 0, 0)
    '            .set_ColWidth(C_COLXPRVPRECIOTAG, 0, 0)
    '            .set_ColWidth(C_COLXPRVTIPO, 0, 0)
    '            .set_TextMatrix(0, C_COLXPRVCODARTICULO, "Código")
    '            .set_TextMatrix(0, C_COLXPRVDESCARTICULO, "Artículo")
    '            .set_TextMatrix(0, C_COLXPRVCODANTERIOR, "Anterior")
    '            .set_TextMatrix(0, C_COLXPRVPORCDESCTO, "% Descto.")
    '            .Row = 0

    '            .set_ColAlignment(C_COLXPRVCODARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '            .set_ColAlignment(C_COLXPRVDESCARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            .set_ColAlignment(C_COLXPRVCODANTERIOR, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '            .set_ColAlignment(C_COLXPRVPORCDESCTO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)

    '            For LnContador = 0 To C_COLXPRVPORCDESCTO
    '                .Col = LnContador
    '                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
    '                .CellFontBold = True
    '            Next LnContador
    '            For LnContador = 1 To .Rows - 1
    '                .set_TextMatrix(LnContador, C_COLXPRVTIPO, "A")
    '            Next
    '            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '            .SelectionMode = MSHierarchicalFlexGridLib.SelectionModeSettings.flexSelectionFree
    '            .TopRow = 1
    '            .Row = 1
    '            .Col = C_COLXPRVCODARTICULO
    '            .WordWrap = False
    '        End With

    'Merr:
    '        If Err.Number <> 0 Then MostrarError()
    '    End Sub

    '    Sub LlenaDatosXArticulo(ByRef CodArticulo As Integer)
    '        On Error GoTo Err_Renamed
    '        If CodArticulo = 0 Then Exit Sub
    '        gStrSql = "SELECT A.CodArticulo, LTRIM(RTRIM(A.DescArticulo)) AS DescArticulo, ISNULL(A.CodFamilia, 0) AS CodFamilia, ISNULL(A.CodLinea, 0) AS CodLinea,A.CodGrupo,CASE A.CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),A.OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),A.CodigoAnt))) ,5) End as CodigoAnt, " & "ISNULL(A.CodSubLinea, 0) AS CodSubLinea, ISNULL(A.CodMarca, 0) AS CodMarca, ISNULL(A.CodModelo, 0) AS CodModelo, ISNULL(F.DescFamilia, '') " & "AS DescFamilia, ISNULL(L.DescLinea, '') AS DescLinea, ISNULL(S.DescSubLinea, '') AS DescSubLinea, ISNULL(Ma.DescMarca, '') AS DescMarca, ISNULL(Mo.DescModelo, '') AS Descmodelo " & "FROM         dbo.CatArticulos A LEFT OUTER  JOIN " & "dbo.CatFamilias F ON A.CodGrupo = F.CodGrupo AND A.CodFamilia = F.CodFamilia LEFT OUTER JOIN " & "dbo.CatLineas L ON A.CodGrupo = L.CodGrupo AND A.CodFamilia = L.CodFamilia AND A.CodLinea = L.CodLinea AND F.CodGrupo = L.CodGrupo AND " & "F.CodFamilia = L.CodFamilia LEFT OUTER JOIN " & "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca LEFT OUTER JOIN " & "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND A.CodMarca = Mo.CodMarca AND A.CodModelo = Mo.CodModelo AND " & "Ma.CodGrupo = Mo.CodGrupo AND Ma.CodMarca = Mo.CodMarca LEFT OUTER JOIN " & "dbo.CatSubLineas S ON A.CodGrupo = S.CodGrupo AND A.CodFamilia = S.CodFamilia AND A.CodLinea = S.CodLinea AND " & "A.CodSubLinea = S.CodSubLinea AND F.CodGrupo = S.CodGrupo AND F.CodFamilia = S.CodFamilia AND L.CodGrupo = S.CodGrupo AND " & "L.CodFamilia = s.CodFamilia And L.COdLinea = s.COdLinea " & "Where (A.CodArticulo = " & CodArticulo & ")"
    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
    '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '        RsGral = Cmd.Execute
    '        If RsGral.RecordCount > 0 Then
    '            With msgXArticulo
    '                .set_TextMatrix(.Row, C_COLXARTCODARTICULO, RsGral.Fields("CodArticulo").Value)
    '                .set_TextMatrix(.Row, C_COLXARTCODANTERIOR, RsGral.Fields("CodigoAnt").Value)
    '                .set_TextMatrix(.Row, C_COLXARTDESCARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
    '                .set_TextMatrix(.Row, C_COLXARTCODFAMILIA, RsGral.Fields("CodFamilia").Value)
    '                .set_TextMatrix(.Row, C_COLXARTCODGRUPO, RsGral.Fields("CodGrupo").Value)
    '                .set_TextMatrix(.Row, C_COLXARTCODLINEA, RsGral.Fields("COdLinea").Value)
    '                .set_TextMatrix(.Row, C_COLXARTCODMARCA, RsGral.Fields("CodMArca").Value)
    '                .set_TextMatrix(.Row, C_COLXARTCODMODELO, RsGral.Fields("CodModelo").Value)
    '                .set_TextMatrix(.Row, C_COLXARTCODSUBLINEA, RsGral.Fields("CodSubLinea").Value)
    '                .set_TextMatrix(.Row, C_COLXARTESTATUS, C_Aplicado)
    '                .set_TextMatrix(.Row, C_COLXARTESTATUSTAG, C_Aplicado)
    '                .set_TextMatrix(.Row, C_COLXARTESNUEVO, True)
    '                ValidarPromocionTecleadaRepetida()
    '                If ValidarPromocionGuardadaRepetida(CShort(Numerico(.get_TextMatrix(.Row, C_COLXARTCODGRUPO))), CShort(Numerico(.get_TextMatrix(.Row, C_COLXARTCODFAMILIA))), CShort(Numerico(.get_TextMatrix(.Row, C_COLXARTCODLINEA))), CShort(Numerico(.get_TextMatrix(.Row, C_COLXARTCODSUBLINEA))), CShort(Numerico(.get_TextMatrix(.Row, C_COLXARTCODMARCA))), CShort(Numerico(.get_TextMatrix(.Row, C_COLXARTCODMODELO))), CInt(Numerico(.get_TextMatrix(.Row, C_COLXARTCODARTICULO))), dtpFechaInIcioJ.Value, dtpFechaFinJ.Value, "A") = True And mblnNuevo = True Then
    '                    MsgBox("Existe una promoción registrada para este artículo." & vbNewLine & "No es posible duplicar promociones en un lapso de tiempo similar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    .set_TextMatrix(.Row, C_COLXARTCODARTICULO, "")
    '                    .set_TextMatrix(.Row, C_COLXARTDESCARTICULO, "")
    '                    .set_TextMatrix(.Row, C_COLXARTCODANTERIOR, "")
    '                    .set_TextMatrix(.Row, C_COLXARTCODFAMILIA, "")
    '                    .set_TextMatrix(.Row, C_COLXARTCODGRUPO, "")
    '                    .set_TextMatrix(.Row, C_COLXARTCODLINEA, "")
    '                    .set_TextMatrix(.Row, C_COLXARTCODMARCA, "")
    '                    .set_TextMatrix(.Row, C_COLXARTCODMODELO, "")
    '                    .set_TextMatrix(.Row, C_COLXARTCODSUBLINEA, "")
    '                    .set_TextMatrix(.Row, C_COLXARTESNUEVO, "")
    '                    .set_TextMatrix(.Row, C_COLXARTESTATUS, "")
    '                    .set_TextMatrix(.Row, C_COLXARTPRECIO, "")
    '                    .set_TextMatrix(.Row, C_COLXARTPORCDESCTO, "")
    '                    txtFlex.Text = ""
    '                    .Focus()
    '                    Exit Sub
    '                End If
    '                txtFlex.Visible = False
    '                .Col = C_COLXARTPORCDESCTO
    '                .Focus()
    '            End With
    '        Else
    '            MsgBox("Codigo de Articulo no Existe, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
    '        End If
    'Err_Renamed:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub


    '    Sub LlenaDatosArticulo(ByRef CodArticulo As Integer, ByRef CodGrupo As Integer)
    '        On Error GoTo Merr
    '        gStrSql = "SELECT A.CodArticulo, LTRIM(RTRIM(A.DescArticulo)) AS DescArticulo, ISNULL(A.CodFamilia, 0) AS CodFamilia, ISNULL(A.CodLinea, 0) AS CodLinea, " & "ISNULL(A.CodSubLinea, 0) AS CodSubLinea, ISNULL(A.CodMarca, 0) AS CodMarca, ISNULL(A.CodModelo, 0) AS CodModelo, ISNULL(F.DescFamilia, '') " & "AS DescFamilia, ISNULL(L.DescLinea, '') AS DescLinea, ISNULL(S.DescSubLinea, '') AS DescSubLinea, ISNULL(Ma.DescMarca, '') AS DescMarca, ISNULL(Mo.DescModelo, '') AS Descmodelo " & "FROM         dbo.CatArticulos A LEFT OUTER  JOIN " & "dbo.CatFamilias F ON A.CodGrupo = F.CodGrupo AND A.CodFamilia = F.CodFamilia LEFT OUTER JOIN " & "dbo.CatLineas L ON A.CodGrupo = L.CodGrupo AND A.CodFamilia = L.CodFamilia AND A.CodLinea = L.CodLinea AND F.CodGrupo = L.CodGrupo AND " & "F.CodFamilia = L.CodFamilia LEFT OUTER JOIN " & "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca LEFT OUTER JOIN " & "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND A.CodMarca = Mo.CodMarca AND A.CodModelo = Mo.CodModelo AND " & "Ma.CodGrupo = Mo.CodGrupo AND Ma.CodMarca = Mo.CodMarca LEFT OUTER JOIN " & "dbo.CatSubLineas S ON A.CodGrupo = S.CodGrupo AND A.CodFamilia = S.CodFamilia AND A.CodLinea = S.CodLinea AND " & "A.CodSubLinea = S.CodSubLinea AND F.CodGrupo = S.CodGrupo AND F.CodFamilia = S.CodFamilia AND L.CodGrupo = S.CodGrupo AND " & "L.CodFamilia = s.CodFamilia And L.COdLinea = s.COdLinea " & "Where (A.CodArticulo = " & CodArticulo & ") And (A.CodGrupo = " & CodGrupo & ")"
    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
    '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '        RsGral = Cmd.Execute
    '        If RsGral.RecordCount > 0 Then
    '            Select Case CodGrupo
    '                Case gCODJOYERIA
    '                    With msgJoyeria
    '                        .set_TextMatrix(.Row, C_ColJFAMILIA, Trim(RsGral.Fields("DescFamilia").Value))
    '                        .set_TextMatrix(.Row, C_ColJLINEA, Trim(RsGral.Fields("DescLinea").Value))
    '                        .set_TextMatrix(.Row, C_ColJSUBLINEA, Trim(RsGral.Fields("DescSubLinea").Value))
    '                        .set_TextMatrix(.Row, C_ColJCODFAMILIA, Trim(RsGral.Fields("CodFamilia").Value))
    '                        .set_TextMatrix(.Row, C_ColJCODLINEA, Trim(RsGral.Fields("COdLinea").Value))
    '                        .set_TextMatrix(.Row, C_ColJCODSUBLINEA, Trim(RsGral.Fields("CodSubLinea").Value))
    '                        .set_TextMatrix(.Row, C_ColJCODARTICULO, RsGral.Fields("CodArticulo").Value)
    '                        .set_TextMatrix(.Row, C_ColJARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
    '                        txtArticulo.Visible = False
    '                        .Col = C_ColJPORCDESCTO
    '                        msgJoyeria.Focus()
    '                    End With
    '                Case gCODRELOJERIA
    '                    With msgRelojeria
    '                        .set_TextMatrix(.Row, C_ColRMARCA, Trim(RsGral.Fields("DescMarca").Value))
    '                        .set_TextMatrix(.Row, C_ColRMODELO, Trim(RsGral.Fields("DescModelo").Value))
    '                        .set_TextMatrix(.Row, C_ColRCODMARCA, Trim(RsGral.Fields("CodMArca").Value))
    '                        .set_TextMatrix(.Row, C_ColRCODMODELO, Trim(RsGral.Fields("CodModelo").Value))
    '                        .set_TextMatrix(.Row, C_ColRCODARTICULO, RsGral.Fields("CodArticulo").Value)
    '                        .set_TextMatrix(.Row, C_ColRARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
    '                        txtArticuloR.Visible = False
    '                        .Col = C_ColRPORCDESCTO
    '                        msgRelojeria.Focus()
    '                    End With
    '                Case gCODVARIOS
    '                    With msgVarios
    '                        .set_TextMatrix(.Row, C_ColJFAMILIA, Trim(RsGral.Fields("DescFamilia").Value))
    '                        .set_TextMatrix(.Row, C_ColJLINEA, Trim(RsGral.Fields("DescLinea").Value))
    '                        .set_TextMatrix(.Row, C_ColJCODFAMILIA, Trim(RsGral.Fields("CodFamilia").Value))
    '                        .set_TextMatrix(.Row, C_ColJCODLINEA, Trim(RsGral.Fields("COdLinea").Value))
    '                        .set_TextMatrix(.Row, C_ColJCODARTICULO, RsGral.Fields("CodArticulo").Value)
    '                        .set_TextMatrix(.Row, C_ColJARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
    '                        txtArticuloV.Visible = False
    '                        .Col = C_ColJPORCDESCTO
    '                        msgVarios.Focus()
    '                    End With
    '            End Select
    '        End If
    'Merr:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    Private Sub chkAplicar_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAplicar.CheckStateChanged
    '        If mblnFueraChange Then Exit Sub

    '        If chkAplicar.CheckState = System.Windows.Forms.CheckState.Checked Then
    '            AsignaDesctoGrid(msgArtxProv, CDec(ModEstandar.Numerico(Trim(txtDesctoP.Text))))
    '        End If
    '    End Sub

    '    Private Sub chkBorrar_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkBorrar.CheckStateChanged
    '        If mblnFueraChange Then Exit Sub

    '        If chkBorrar.CheckState = System.Windows.Forms.CheckState.Checked Then
    '            BorraArtsProv(msgArtxProv)
    '        End If
    '    End Sub

    '    Private Sub chkCancelarP_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCancelarP.CheckStateChanged
    '        If mblnFueraChange Then Exit Sub

    '        If chkCancelarP.CheckState = System.Windows.Forms.CheckState.Checked Then
    '            CancelaProm(True)
    '        Else
    '            CancelaProm(False)
    '        End If
    '    End Sub

    '    Private Sub dbcJArticulo_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJArticulo.CursorChanged
    '        If FueraChange = True Then Exit Sub
    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcJArticulo.Name Then Exit Sub
    '        If sstGrupos.SelectedIndex = 0 Then
    '            gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & gCODJOYERIA & ") " & IIf((CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA))) = 0), " ", " And (CodFamilia = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA)) & ")") & IIf((CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA))) = 0), "", " And (CodLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA)) & ")") & IIf((CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA))) = 0), "", " And (CodSubLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA)) & ")") & "and DescArticulo LIKE '" & Trim(dbcJArticulo.Text) & "%'"
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & gCODVARIOS & ") " & IIf((CDbl(Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODFAMILIA))) = 0), " ", " And (CodFamilia = " & Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODFAMILIA)) & ")") & IIf((CDbl(Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODLINEA))) = 0), "", " And (CodLinea = " & Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODLINEA)) & ")") & "and DescArticulo LIKE '" & Trim(dbcJArticulo.Text) & "%'"
    '        Else
    '            Exit Sub
    '        End If
    '        ModDCombo.DCChange(gStrSql, tecla)
    '        LimpiaDatosPrecioYDescuento()
    '    End Sub

    '    Private Sub dbcJArticulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJArticulo.Enter
    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcJArticulo.Name Then Exit Sub
    '        Pon_Tool()
    '        If sstGrupos.SelectedIndex = 0 Then
    '            gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & gCODJOYERIA & ") " & IIf((CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA))) = 0), " ", " And (CodFamilia = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA)) & ")") & IIf((CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA))) = 0), "", " And (CodLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA)) & ")") & IIf((CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA))) = 0), "", " And (CodSubLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA)) & ")")
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & gCODVARIOS & ") " & IIf((CDbl(Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODFAMILIA))) = 0), " ", " And (CodFamilia = " & Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODFAMILIA)) & ")") & IIf((CDbl(Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODLINEA))) = 0), "", " And (CodLinea = " & Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODLINEA)) & ")")
    '        Else
    '            Exit Sub
    '        End If
    '        ModDCombo.DCGotFocus((gStrSql))
    '    End Sub

    '    Private Sub dbcJArticulo_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJArticulo.KeyDown
    '        tecla = eventArgs.KeyCode
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        Else
    '            Exit Sub
    '        End If
    '        With GridACtivo
    '            FueraChange = True
    '            Select Case eventArgs.KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    dbcJArticulo.Visible = False
    '                    dbcJArticulo.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    intCodSubLinea = 0
    '                    .set_TextMatrix(.Row, C_ColJCODARTICULO, 0)

    '                    gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & IIf((GridACtivo Is msgJoyeria), gCODJOYERIA, gCODVARIOS) & ") " & IIf((CDbl(Numerico(GridACtivo.get_TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA))) = 0), " ", " And (CodFamilia = " & Numerico(GridACtivo.get_TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) & ")") & IIf((CDbl(Numerico(GridACtivo.get_TextMatrix(GridACtivo.Row, C_ColJCODLINEA))) = 0), "", " And (CodLinea = " & Numerico(GridACtivo.get_TextMatrix(GridACtivo.Row, C_ColJCODLINEA)) & ")") & IIf((CDbl(Numerico(GridACtivo.get_TextMatrix(GridACtivo.Row, C_ColJCODSUBLINEA))) = 0), "", " And (CodSubLinea = " & Numerico(GridACtivo.get_TextMatrix(GridACtivo.Row, C_ColJCODSUBLINEA)) & ")") & "and DescArticulo LIKE '" & Trim(dbcJArticulo.Text) & "%'"

    '                    ModDCombo.DCLostFocus(dbcJArticulo, gStrSql, intCodArticulo)
    '                    .set_TextMatrix(.Row, .Col, Trim(dbcJArticulo.Text))
    '                    .set_TextMatrix(.Row, C_ColJCODARTICULO, intCodArticulo)
    '                    LlenaDatosArticulo(intCodArticulo, IIf((GridACtivo Is msgJoyeria), gCODJOYERIA, gCODVARIOS))
    '                    .Col = C_ColJPORCDESCTO
    '                    dbcJArticulo.Text = ""
    '                    dbcJArticulo.Visible = False
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                    .set_ColAlignment(C_ColJARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                    'Verificar si la Promoción está Siendo Repetida
    '                    '                ValidarPromocionTecleadaRepetida
    '                    '                intCodFamilia = Numerico(.TextMatrix(.Row, C_ColJCODFAMILIA))
    '                    '                intCodLinea = Numerico(.TextMatrix(.Row, C_ColJCODLINEA))
    '                    '                intCodSubLinea = Numerico(.TextMatrix(.Row, C_ColJCODSUBLINEA))
    '                    '                If ValidarPromocionGuardadaRepetida(gCODJOYERIA, intCodFamilia, intCodLinea, intCodSubLinea, 0, 0, dtpFechaInIcioJ, dtpFechaFinJ) = True Then
    '                    '                    MsgBox "Existe una Promoción registrada para este Artículo." + vbNewLine + "No es posible duplicar Promociones en un Lapso de Tiempo Similar", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    LimpiaDatosFamilia
    '                    '                    .TextMatrix(.Row, C_ColJFAMILIA) = ""
    '                    '                    .Col = C_ColJFAMILIA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '            End Select
    '            FueraChange = False
    '        End With
    '    End Sub

    '    Private Sub dbcJArticulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJArticulo.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        dbcJArticulo.Visible = False
    '    End Sub

    '    Private Sub dbcJFAmilia_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJFamilia.KeyDown
    '        tecla = eventArgs.KeyCode
    '        Dim GridACtivo As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        Else
    '            Exit Sub
    '        End If
    '        With GridACtivo
    '            FueraChange = True
    '            Select Case eventArgs.KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    dbcJFamilia.Visible = False
    '                    dbcJFamilia.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    intCodFamilia = 0
    '                    .set_TextMatrix(.Row, C_ColJCODFAMILIA, 0)
    '                    If sstGrupos.SelectedIndex = 0 Then
    '                        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " and DescFamilia LIKE '" & Trim(dbcJFamilia.Text) & "%' ORDER BY DescFamilia"
    '                    Else
    '                        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " and DescFamilia LIKE '" & Trim(dbcJFamilia.Text) & "%' ORDER BY DescFamilia"
    '                    End If
    '                    ModDCombo.DCLostFocus(dbcJFamilia, gStrSql, intCodFamilia)
    '                    .set_TextMatrix(.Row, .Col, Trim(dbcJFamilia.Text))
    '                    .set_TextMatrix(.Row, C_ColJCODFAMILIA, intCodFamilia)
    '                    .Focus()
    '                    .Col = C_ColJLINEA
    '                    dbcJFamilia.Text = ""
    '                    dbcJFamilia.Visible = False
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .set_ColAlignment(C_ColJFAMILIA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            End Select
    '            FueraChange = False
    '            .set_TextMatrix(.Row, C_ColJESTATUS, C_Aplicado)
    '            .set_TextMatrix(.Row, C_ColJESTATUSTAG, C_Aplicado)
    '            .set_TextMatrix(.Row, C_ColJESNUEVO, True)
    '        End With
    '    End Sub

    '    Private Sub dbcJFAmilia_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJFamilia.CursorChanged
    '        If FueraChange = True Then Exit Sub
    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcJFamilia.Name Then Exit Sub
    '        If sstGrupos.SelectedIndex = 0 Then
    '            gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " and DescFamilia LIKE '" & Trim(dbcJFamilia.Text) & "%' ORDER BY DescFamilia"
    '        Else
    '            gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " and DescFamilia LIKE '" & Trim(dbcJFamilia.Text) & "%' ORDER BY DescFamilia"
    '        End If
    '        ModDCombo.DCChange(gStrSql, tecla)
    '        LimpiaDatosFamilia()
    '    End Sub

    '    Private Sub dbcjFAmilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.Enter
    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcJFamilia.Name Then Exit Sub
    '        Pon_Tool()
    '        If sstGrupos.SelectedIndex = 0 Then
    '            gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " ORDER BY DescFamilia"
    '        Else
    '            gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " ORDER BY DescFamilia"
    '        End If
    '        ModDCombo.DCGotFocus((gStrSql))
    '    End Sub

    '    Private Sub dbcJFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        dbcJFamilia.Visible = False
    '    End Sub

    '    Private Sub dbcJLinea_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJLinea.KeyDown
    '        tecla = eventArgs.KeyCode
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        End If
    '        With GridACtivo
    '            FueraChange = True
    '            Select Case eventArgs.KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    dbcJLinea.Visible = False
    '                    dbcJLinea.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    intCodLinea = 0
    '                    .set_TextMatrix(.Row, C_ColJCODLINEA, 0)
    '                    If sstGrupos.SelectedIndex = 0 Then
    '                        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & Numerico(.get_TextMatrix(.Row, C_ColJCODFAMILIA)) & ") and DescLinea LIKE '" & Trim(dbcJLinea.Text) & "%' ORDER BY DescLinea"
    '                    Else
    '                        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & Numerico(.get_TextMatrix(.Row, C_ColJCODFAMILIA)) & ") and DescLinea LIKE '" & Trim(dbcJLinea.Text) & "%' ORDER BY DescLinea"
    '                    End If
    '                    ModDCombo.DCLostFocus(dbcJLinea, gStrSql, intCodLinea)
    '                    .set_TextMatrix(.Row, .Col, Trim(dbcJLinea.Text))
    '                    .set_TextMatrix(.Row, C_ColJCODLINEA, intCodLinea)
    '                    If GridACtivo Is msgJoyeria Then
    '                        .Col = C_ColJSUBLINEA
    '                    ElseIf GridACtivo Is msgVarios Then
    '                        .Col = C_ColJARTICULO
    '                    End If
    '                    dbcJLinea.Text = ""
    '                    dbcJLinea.Visible = False
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                    .set_ColAlignment(C_ColJLINEA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                    '                'Si el GridActivo es MsgVarios. Verificar si es una Promocion Repetida
    '                    '                If GridACtivo Is msgVarios Then
    '                    '                    ValidarPromocionTecleadaRepetida
    '                    '                    intCodFamilia = Numerico(.TextMatrix(.Row, C_ColJCODFAMILIA))
    '                    '                    intCodLinea = Numerico(.TextMatrix(.Row, C_ColJCODLINEA))
    '                    '                    intCodSubLinea = 0
    '                    '                    If ValidarPromocionGuardadaRepetida(gCODVARIOS, intCodFamilia, intCodLinea, intCodSubLinea, 0, 0, dtpFechaInIcioV, dtpFechaFinV) = True Then
    '                    '                        MsgBox "Existe una Promoción registrada para este Artículo." + vbNewLine + "No es posible duplicar Promociones en un Lapso de Tiempo Similar", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                        LimpiaDatosFamilia
    '                    '                        .TextMatrix(.Row, C_ColJFAMILIA) = ""
    '                    '                        .Col = C_ColJFAMILIA
    '                    '                        .SetFocus
    '                    '                        Exit Sub
    '                    '                    End If
    '                    '                End If
    '            End Select
    '            FueraChange = False
    '        End With
    '    End Sub

    '    Private Sub dbcJLinea_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJLinea.CursorChanged
    '        If FueraChange = True Then Exit Sub
    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcJLinea.Name Then Exit Sub
    '        Dim GridACtivo As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        End If
    '        If sstGrupos.SelectedIndex = 0 Then
    '            gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & Numerico(GridACtivo.get_TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) & ") and DescLinea LIKE '" & Trim(dbcJLinea.Text) & "%' ORDER BY DescLinea"
    '        Else
    '            gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & Numerico(GridACtivo.get_TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) & ") and DescLinea LIKE '" & Trim(dbcJLinea.Text) & "%' ORDER BY DescLinea"
    '        End If
    '        ModDCombo.DCChange(gStrSql, tecla)
    '        LimpiaDatosLinea()
    '        If GridACtivo Is msgVarios Then
    '            LimpiaDatosPrecioYDescuento()
    '        End If
    '    End Sub

    '    Private Sub dbcJLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.Enter
    '        If FueraChange = True Then Exit Sub
    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcJLinea.Name Then Exit Sub
    '        Dim GridACtivo As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        End If
    '        Pon_Tool()
    '        If sstGrupos.SelectedIndex = 0 Then
    '            gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & Numerico(GridACtivo.get_TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) & ")  ORDER BY DescLinea"
    '        Else
    '            gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & Numerico(GridACtivo.get_TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) & ")  ORDER BY DescLinea"
    '        End If
    '        ModDCombo.DCGotFocus((gStrSql))
    '    End Sub

    '    Private Sub dbcJLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        dbcJLinea.Visible = False
    '    End Sub

    '    Private Sub dbcJSubLinea_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJSubLinea.KeyDown
    '        tecla = eventArgs.KeyCode
    '        With msgJoyeria
    '            FueraChange = True
    '            Select Case eventArgs.KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    dbcJSubLinea.Visible = False
    '                    dbcJSubLinea.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    intCodSubLinea = 0
    '                    .set_TextMatrix(.Row, C_ColJCODSUBLINEA, 0)
    '                    gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA)) & ")  And (CodLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA)) & ") and DescSubLinea LIKE '" & Trim(dbcJSubLinea.Text) & "%' ORDER BY DescSubLinea"
    '                    ModDCombo.DCLostFocus(dbcJSubLinea, gStrSql, intCodSubLinea)
    '                    .set_TextMatrix(.Row, .Col, Trim(dbcJSubLinea.Text))
    '                    .set_TextMatrix(.Row, C_ColJCODSUBLINEA, intCodSubLinea)
    '                    .Col = C_ColJARTICULO
    '                    dbcJSubLinea.Text = ""
    '                    dbcJSubLinea.Visible = False
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                    .set_ColAlignment(C_ColJSUBLINEA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                    'Verificar si la Promoción está Siendo Repetida
    '                    '                ValidarPromocionTecleadaRepetida
    '                    '                intCodFamilia = Numerico(.TextMatrix(.Row, C_ColJCODFAMILIA))
    '                    '                intCodLinea = Numerico(.TextMatrix(.Row, C_ColJCODLINEA))
    '                    '                intCodSubLinea = Numerico(.TextMatrix(.Row, C_ColJCODSUBLINEA))
    '                    '                If ValidarPromocionGuardadaRepetida(gCODJOYERIA, intCodFamilia, intCodLinea, intCodSubLinea, 0, 0, dtpFechaInIcioJ, dtpFechaFinJ) = True Then
    '                    '                    MsgBox "Existe una promoción registrada para este artículo." + vbNewLine + "No es posible duplicar promociones en un lapso de tiempo similar", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    LimpiaDatosFamilia
    '                    '                    .TextMatrix(.Row, C_ColJFAMILIA) = ""
    '                    '                    .Col = C_ColJFAMILIA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '            End Select
    '            FueraChange = False
    '        End With
    '    End Sub

    '    Private Sub dbcJSubLinea_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcJSubLinea.CursorChanged
    '        If FueraChange = True Then Exit Sub
    '        gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA)) & ")  And (CodLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA)) & ") and DescSubLinea LIKE '" & Trim(dbcJSubLinea.Text) & "%' ORDER BY DescSubLinea"
    '        ModDCombo.DCChange(gStrSql, tecla)
    '        LimpiaDatosArticulo()
    '    End Sub

    '    Private Sub dbcJSubLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.Enter
    '        Pon_Tool()
    '        gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA)) & ")  And (CodLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA)) & ") ORDER BY DescSubLinea"
    '        ModDCombo.DCGotFocus((gStrSql))
    '    End Sub

    '    Private Sub dbcJSubLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        dbcJSubLinea.Visible = False
    '    End Sub

    '    Private Sub dbcrArticulo_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRArticulo.CursorChanged
    '        If FueraChange = True Then Exit Sub
    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcRArticulo.Name Then Exit Sub
    '        gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & gCODRELOJERIA & ") " & IIf((CDbl(Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA))) = 0), " ", " And (CodMarca = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA)) & ")") & IIf((CDbl(Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMODELO))) = 0), "", " And (Codmodelo = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMODELO)) & ")") & "and DescArticulo LIKE '" & Trim(dbcRArticulo.Text) & "%'"
    '        ModDCombo.DCChange(gStrSql, tecla)
    '        LimpiaDatosPrecioYDescuento()
    '    End Sub

    '    Private Sub dbcRArticulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRArticulo.Enter
    '        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcRArticulo.Name Then Exit Sub
    '        Pon_Tool()
    '        gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & gCODRELOJERIA & ") " & IIf((CDbl(Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA))) = 0), " ", " And (CodMarca = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA)) & ")") & IIf((CDbl(Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMODELO))) = 0), "", " And (Codmodelo = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMODELO)) & ")")
    '        ModDCombo.DCGotFocus((gStrSql))
    '    End Sub

    '    Private Sub dbcRArticulo_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRArticulo.KeyDown
    '        tecla = eventArgs.KeyCode
    '        With msgRelojeria
    '            FueraChange = True
    '            Select Case eventArgs.KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    dbcJArticulo.Visible = False
    '                    dbcJArticulo.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    intCodArticulo = 0
    '                    .set_TextMatrix(.Row, C_ColRCODARTICULO, 0)
    '                    'gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA)) & ")  And (CodLinea = " & Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODLINEA)) & ") and DescSubLinea LIKE '" & Trim(dbcJArticulo) & "%' ORDER BY DescSubLinea"
    '                    gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & gCODRELOJERIA & ") " & IIf((CDbl(Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA))) = 0), " ", " And (CodMarca = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA)) & ")") & IIf((CDbl(Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMODELO))) = 0), "", " And (Codmodelo = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMODELO)) & ")") & "and DescArticulo LIKE '" & Trim(dbcRArticulo.Text) & "%'"
    '                    ModDCombo.DCLostFocus(dbcRArticulo, gStrSql, intCodArticulo)
    '                    .set_TextMatrix(.Row, .Col, Trim(dbcRArticulo.Text))
    '                    .set_TextMatrix(.Row, C_ColRCODARTICULO, intCodArticulo)
    '                    LlenaDatosArticulo(intCodArticulo, gCODRELOJERIA)
    '                    .Col = C_ColRPORCDESCTO
    '                    dbcRArticulo.Text = ""
    '                    dbcRArticulo.Visible = False
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                    .set_ColAlignment(C_ColRARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                    'Verificar si la Promoción está Siendo Repetida
    '                    '                ValidarPromocionTecleadaRepetida
    '                    '                intCodFamilia = Numerico(.TextMatrix(.Row, C_ColJCODFAMILIA))
    '                    '                intCodLinea = Numerico(.TextMatrix(.Row, C_ColJCODLINEA))
    '                    '                intCodSubLinea = Numerico(.TextMatrix(.Row, C_ColJCODSUBLINEA))
    '                    '                If ValidarPromocionGuardadaRepetida(gCODJOYERIA, intCodFamilia, intCodLinea, intCodSubLinea, 0, 0, dtpFechaInIcioJ, dtpFechaFinJ) = True Then
    '                    '                    MsgBox "Existe una Promoción registrada para este Artículo." + vbNewLine + "No es posible duplicar Promociones en un Lapso de Tiempo Similar", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    LimpiaDatosFamilia
    '                    '                    .TextMatrix(.Row, C_ColJFAMILIA) = ""
    '                    '                    .Col = C_ColJFAMILIA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '            End Select
    '            FueraChange = False
    '        End With
    '    End Sub

    '    Private Sub dbcRArticulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRArticulo.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        dbcRArticulo.Visible = False
    '    End Sub

    '    Private Sub dbcRMarca_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRMarca.CursorChanged
    '        If FueraChange = True Then Exit Sub

    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcRMarca.Name Then Exit Sub
    '        gStrSql = "SELECT CodMarca , DescMarca =ltrim(rtrim(DescMarca))  From CatMarcas Where CodGRupo = " & gCODRELOJERIA & " and DescMarca LIKE '" & Trim(dbcRMarca.Text) & "%' ORDER BY DescMarca"
    '        ModDCombo.DCChange(gStrSql, tecla)
    '        LimpiaDatosMarca()
    '    End Sub

    '    Private Sub dbcRMarca_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.Enter
    '        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcRMarca.Name Then Exit Sub
    '        Pon_Tool()
    '        gStrSql = "SELECT CodMarca , DescMarca =ltrim(rtrim(DescMarca))  From CatMarcas Where CodGRupo = " & gCODRELOJERIA & " ORDER BY DescMarca"
    '        ModDCombo.DCGotFocus((gStrSql))
    '    End Sub

    '    Private Sub dbcRMarca_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRMarca.KeyDown
    '        tecla = eventArgs.KeyCode
    '        With msgRelojeria
    '            FueraChange = True
    '            Select Case eventArgs.KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    dbcRMarca.Visible = False
    '                    dbcRMarca.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    intCodMarca = 0
    '                    .set_TextMatrix(.Row, C_ColRCODMARCA, 0)
    '                    gStrSql = "SELECT CodMarca , DescMarca =ltrim(rtrim(DescMarca))  From CatMarcas Where CodGRupo = " & gCODRELOJERIA & " and DescMarca LIKE '" & Trim(dbcRMarca.Text) & "%' ORDER BY DescMarca"
    '                    ModDCombo.DCLostFocus(dbcRMarca, gStrSql, intCodMarca)
    '                    .set_TextMatrix(.Row, .Col, Trim(dbcRMarca.Text))
    '                    .set_TextMatrix(.Row, C_ColRCODMARCA, intCodMarca)
    '                    dbcRMarca.Text = ""
    '                    dbcRMarca.Visible = False
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                    .Col = C_ColRMODELO
    '                    .set_ColAlignment(C_ColRMARCA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '            End Select
    '            FueraChange = False
    '            .set_TextMatrix(.Row, C_ColRESTATUS, C_Aplicado)
    '            .set_TextMatrix(.Row, C_ColRESNUEVO, True)
    '        End With
    '    End Sub

    '    Private Sub dbcRMarca_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        dbcRMarca.Visible = False
    '    End Sub

    '    Private Sub dbcRmodelo_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRModelo.CursorChanged
    '        If FueraChange = True Then Exit Sub

    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcRModelo.Name Then Exit Sub
    '        gStrSql = "SELECT Codmodelo , Descmodelo =ltrim(rtrim(Descmodelo))  From Catmodelos Where CodGRupo = " & gCODRELOJERIA & " And CodMarca = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA)) & " and Descmodelo LIKE '" & Trim(dbcRModelo.Text) & "%' ORDER BY Descmodelo"
    '        ModDCombo.DCChange(gStrSql, tecla)
    '        LimpiaDatosArticulo()
    '    End Sub

    '    Private Sub dbcRmodelo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRmodelo.Enter

    '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcRModelo.Name Then Exit Sub
    '        Pon_Tool()
    '        gStrSql = "SELECT Codmodelo , Descmodelo =ltrim(rtrim(Descmodelo))  From Catmodelos Where CodGRupo = " & gCODRELOJERIA & " And CodMarca = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA)) & " ORDER BY Descmodelo"
    '        ModDCombo.DCGotFocus((gStrSql))
    '    End Sub

    '    Private Sub dbcRmodelo_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRModelo.KeyDown
    '        tecla = eventArgs.KeyCode
    '        With msgRelojeria
    '            FueraChange = True
    '            Select Case eventArgs.KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    dbcRModelo.Visible = False
    '                    dbcRModelo.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    intCodModelo = 0
    '                    .set_TextMatrix(.Row, C_ColRCODMODELO, 0)
    '                    gStrSql = "SELECT Codmodelo , Descmodelo =ltrim(rtrim(Descmodelo))  From Catmodelos Where CodGRupo = " & gCODRELOJERIA & " And CodMarca = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA)) & " and Descmodelo LIKE '" & Trim(dbcRModelo.Text) & "%' ORDER BY Descmodelo"
    '                    ModDCombo.DCLostFocus(dbcRModelo, gStrSql, intCodModelo)
    '                    .set_TextMatrix(.Row, .Col, Trim(dbcRModelo.Text))
    '                    .set_TextMatrix(.Row, C_ColRCODMODELO, intCodModelo)
    '                    .Col = C_ColRARTICULO
    '                    dbcRModelo.Text = ""
    '                    dbcRModelo.Visible = False
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                    .set_ColAlignment(C_ColRMODELO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                    'Verificar si la Promoción está Siendo Repetida
    '                    '                ValidarPromocionTecleadaRepetida
    '                    '                intCodMarca = Numerico(.TextMatrix(.Row, C_ColRCODMARCA))
    '                    '                intCodModelo = Numerico(.TextMatrix(.Row, C_ColRCODMODELO))
    '                    '                If ValidarPromocionGuardadaRepetida(gCODRELOJERIA, 0, 0, 0, intCodMarca, intCodModelo, dtpFechaInIcioR, dtpFechaFinR) = True Then
    '                    '                    MsgBox "Existe una Promoción registrada para este Artículo." + vbNewLine + "No es posible duplicar Promociones en un Lapso de Tiempo Similar", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    LimpiaDatosMarca
    '                    '                    .TextMatrix(.Row, C_ColRMARCA) = ""
    '                    '                    .Col = C_ColRMARCA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '            End Select
    '            FueraChange = False
    '        End With
    '    End Sub

    '    Private Sub dbcRModelo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        dbcRModelo.Visible = False
    '    End Sub

    '    Private Sub dtpFechaFinJ_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaFinJ.CursorChanged
    '        'sglTiempoCambio = Timer()
    '    End Sub

    '    Private Sub dtpFechaFinJ_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcProveedor.Click
    '        'sglTiempoCambio = Timer()
    '    End Sub

    '    Private Sub dtpFechaFinJ_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcProveedor.KeyPress
    '        'sglTiempoCambio = Timer()
    '    End Sub

    '    Private Sub dtpFechaFinJ_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinJ.Leave
    '        dtpFechaFinR.Value = dtpFechaFinJ.Value
    '        dtpFechaFinV.Value = dtpFechaFinJ.Value
    '    End Sub

    '    Private Sub dtpFechaInIcioJ_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInIcioJ.CursorChanged
    '        'sglTiempoCambio = Timer()
    '    End Sub

    '    Private Sub dtpFechaInIcioJ_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInIcioJ.Click
    '        'sglTiempoCambio = Timer()
    '    End Sub

    '    Private Sub dtpFechaInIcioJ_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInIcioJ.KeyDown
    '        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
    '            mblnSalir = True
    '            Me.Close()
    '        End If
    '    End Sub

    '    Private Sub dtpFechaInIcioJ_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaInIcioJ.KeyPress
    '        'sglTiempoCambio = Timer()
    '    End Sub

    '    Private Sub dtpFechaInIcioJ_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInIcioJ.Leave
    '        'Igualar las FEchas de Relojeria y VArios a la de Joyeria
    '        dtpFechaInIcioR.Value = dtpFechaInIcioJ.Value
    '        dtpFechaInIcioV.Value = dtpFechaInIcioJ.Value
    '    End Sub

    '    Private Sub frmProgramacionPromociones_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
    '        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
    '        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    '        Me.BringToFront()
    '    End Sub

    '    Private Sub frmProgramacionPromociones_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
    '        'Desactivar todas las opciones del Menu
    '        '                              Nuevo     GuCancelar      Eliminar    Buscar       Imprimir     Cerrar
    '        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    '    End Sub

    Private Sub frmProgramacionPromociones_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        'Nuevo()
        'InicializaVariables()
        'Encabezado()
        'LlenaDatos()
    End Sub

    '    Private Sub frmProgramacionPromociones_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
    '        Dim KeyCode As Integer = eventArgs.KeyCode
    '        Dim Shift As Integer = eventArgs.KeyData \ &H10000
    '        ' En este evento del formulario se valida la tecla presionada.
    '        ' Si es Enter se simula un tab(Avanza al siguiente control)
    '        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
    '        Select Case KeyCode
    '            Case System.Windows.Forms.Keys.Return
    '                'Si el control en que se presiono enter, es el Grid de Detalle de la venta que no se ejecute el avanzar tab
    '                If ActiveControl.Name <> "msgJoyeria" And ActiveControl.Name <> "msgRelojeria" And ActiveControl.Name <> "msgVarios" And ActiveControl.Name <> "msgXArticulo" And ActiveControl.Name <> "txtFlex" Then
    '                    ModEstandar.AvanzarTab(Me)
    '                End If
    '            Case System.Windows.Forms.Keys.Escape
    '                If ActiveControl.Name <> "txtJoyeria" And ActiveControl.Name <> "txtRelojeria" And ActiveControl.Name <> "sstGrupos" And ActiveControl.Name <> "msgXArticulo" Then
    '                    ModEstandar.RetrocederTab(Me)
    '                End If
    '        End Select
    '    End Sub

    '    Private Sub frmProgramacionPromociones_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
    '        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
    '        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
    '        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
    '        eventArgs.KeyChar = Chr(KeyAscii)
    '        If KeyAscii = 0 Then
    '            eventArgs.Handled = True
    '        End If
    '    End Sub

    '    Private Sub frmProgramacionPromociones_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '        Dim Cancel As Boolean = eventArgs.Cancel
    '        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
    '        If Not mblnSalir Then
    '            'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
    '            ModEstandar.RestaurarForma(Me, False)
    '        Else 'Se quiere salir con escape
    '            mblnSalir = False
    '            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
    '                Case MsgBoxResult.Yes
    '                    Cancel = 0 'Sale de la Captura, Con 1: Sigue en la captura
    '                Case MsgBoxResult.No 'No sale del formulario
    '                    Cancel = 1
    '            End Select
    '        End If
    '        eventArgs.Cancel = Cancel
    '    End Sub

    '    Private Sub frmProgramacionPromociones_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
    '        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    '        ModEstandar.LimpiaDescBarraEstado()
    '        'Me = Nothing
    '    End Sub

    '    Private Sub msgJoyeria_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgJoyeria.DblClick
    '        Dim Estatus As String
    '        Dim EstatusTag As String
    '        With msgJoyeria
    '            '        Estatus = .TextMatrix(.Row, C_ColJESTATUS)
    '            '        EstatusTag = .TextMatrix(.Row, C_ColJESTATUSTAG)
    '            '        If (Estatus = C_Aplicado And Estatus <> EstatusTag) Or (Estatus = C_Cancelado And Estatus <> EstatusTag) Then
    '            '            .TextMatrix(.Row, C_ColJESTATUS) = C_Vigente
    '            '            PonerColor (.Row)
    '            '            .SetFocus
    '            ''        ElseIf Estatus = C_Aplicado Then
    '            ''            MsgBox "No es posible modificar una Promoción Aplicada Previamente.", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '            ''            .SetFocus
    '            ''            Exit Sub
    '            '        End If
    '        End With
    '        msgJoyeria_KeyPressEvent(msgJoyeria, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    '    End Sub

    '    Private Sub msgJoyeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgJoyeria.Enter
    '        msgJoyeria.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '        '    msgJoyeria.Row = 1
    '        '    msgJoyeria.Col = 0
    '        Pon_Tool()
    '    End Sub

    '    Sub msgJoyeria_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles msgJoyeria.KeyDownEvent
    '        If eventArgs.keyCode = System.Windows.Forms.Keys.Escape Then
    '            sstGrupos.Focus()
    '            sstGrupos.SelectedIndex = 0
    '        End If
    '        With msgJoyeria
    '            If Trim(.get_TextMatrix(.Row, C_ColJFAMILIA)) = "" Then Exit Sub
    '            Select Case eventArgs.keyCode
    '                Case System.Windows.Forms.Keys.Delete
    '                    If .get_TextMatrix(.Row, C_ColJESTATUS) = C_Cancelado Then
    '                        MsgBox("Está promoción ya ha sido cancelada." & vbNewLine & "Verifique por favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                        .Focus()
    '                        Exit Sub
    '                    End If
    '                    '                If .TextMatrix(.Row, C_ColJESTATUS) = C_Aplicado And .TextMatrix(.Row, C_ColJESNUEVO) = False Then
    '                    '                    MsgBox "Está promoción ya ha sido aplicada. No es posible eliminar" + vbNewLine + "Verifique Por Favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "Mensaje")
    '                        Case MsgBoxResult.No
    '                            .Focus()
    '                            Exit Sub
    '                        Case MsgBoxResult.Cancel
    '                            .Focus()
    '                            Exit Sub
    '                    End Select
    '                    .set_TextMatrix(.Row, C_ColJESTATUS, C_Cancelado)
    '                    PonerColor((.Row))
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Insert
    '                    '                If .TextMatrix(.Row, C_ColJESTATUS) = C_Aplicado Then
    '                    '                    MsgBox "Esta promoción ya ha sido aplicada." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    '                If .TextMatrix(.Row, C_ColJESTATUS) = C_Cancelado Then
    '                    '                    MsgBox "Esta promoción ha sido cancelado. No es posible aplicarla." + vbNewLine + "Verifique Por Favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    '                If Numerico(.TextMatrix(.Row, C_ColJPORCDESCTO)) = 0 And Numerico(.TextMatrix(.Row, C_ColJPRECIO)) = 0 Then
    '                    '                    MsgBox "Información incompleta sobre la promocion." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    '                 Select Case MsgBox("¿Desea Aplicar esta Promoción para que esté disponible para Operaciones de Venta", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Mensaje")
    '                    '                    Case vbNo
    '                    '                        .SetFocus
    '                    '                        Exit Sub
    '                    '                    Case vbCancel
    '                    '                        .SetFocus
    '                    '                        Exit Sub
    '                    '                End Select
    '                    '                .TextMatrix(.Row, C_ColJESTATUS) = C_Aplicado
    '                    '                PonerColor (.Row)
    '                    '                .SetFocus
    '            End Select
    '        End With
    '    End Sub

    '    Private Sub msgJoyeria_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles msgJoyeria.KeyPressEvent
    '        Dim EsNuevo As Boolean
    '        Dim Estatus As String
    '        Dim EstatustTag As String
    '        EsNuevo = True
    '        With msgJoyeria
    '            ' Si nho se trata de un REgistro nuevo, no se podrá editar el Grid
    '            If Trim(.get_TextMatrix(.Row, C_ColJESNUEVO)) <> "" Then
    '                EsNuevo = CBool(.get_TextMatrix(.Row, C_ColJESNUEVO))
    '            End If
    '            If .get_TextMatrix(.Row, C_ColJESTATUS) = C_Cancelado Then Exit Sub
    '            FueraChange = True
    '            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then 'Para que cuando sea escape, no entre a editar el codigo,simplemente que se regrese al control anterior
    '                Select Case .Col
    '                    Case C_ColJFAMILIA ''-------------- SE EDITA LA FAMILIA ---------------------'''''
    '                        If EsNuevo = False Or mblnNuevo = False Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        '''en esta parte se validará si es el rengón, columna que le corresponde editarse
    '                        If (.Row > 1) Then
    '                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
    '                            If Trim(.get_TextMatrix(.Row - 1, C_ColJFAMILIA)) = "" Then
    '                                .Focus()
    '                                Exit Sub
    '                            End If
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgJoyeria, dbcJFamilia, eventArgs.keyAscii)
    '                    Case C_ColJLINEA ''-------------- SE EDITA LA LINEA ---------------------'''''
    '                        If EsNuevo = False Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColJFAMILIA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgJoyeria, dbcJLinea, eventArgs.keyAscii)
    '                    Case C_ColJSUBLINEA ''-------------- SE EDITA LA SUBLINEA---------------------'''''
    '                        If EsNuevo = False Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColJLINEA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgJoyeria, dbcJSubLinea, eventArgs.keyAscii)

    '                    Case C_ColJARTICULO
    '                        If EsNuevo = False Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColJFAMILIA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgJoyeria, txtArticulo, eventArgs.keyAscii)

    '                    Case C_ColJPORCDESCTO ''-------------- SE EDITA EL PORCENTAJE DE DESCTO---------------------'''''
    '                        'Or Estatus = C_Aplicado
    '                        ModEstandar.gp_CampoNumerico(eventArgs.keyAscii, ".")
    '                        If Estatus = C_Cancelado Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColJFAMILIA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, C_ColJPRECIO))) <> 0 Then
    '                            MsgBox("ya se ha Asignado un Precio. No es Posible Asignar un Porcentaje de Descuento.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgJoyeria, txtJoyeria, eventArgs.keyAscii)
    '                        .set_TextMatrix(.Row, C_ColJPRECIO, "0.00")
    '                    Case C_ColJPRECIO ''-------------- SE EDITA LA EL PRECIO ---------------------'''''
    '                        'Or Estatus = C_Aplicado
    '                        ModEstandar.gp_CampoNumerico(eventArgs.keyAscii, ".")
    '                        If Estatus = C_Cancelado Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColJFAMILIA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, C_ColJPORCDESCTO))) <> 0 Then
    '                            MsgBox("Ya se ha asignado un porcentaje de descuento. No es posible asignar el precio.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgJoyeria, txtJoyeria, eventArgs.keyAscii)
    '                        .set_TextMatrix(.Row, C_ColJPORCDESCTO, "0.00")
    '                End Select
    '            End If
    '        End With
    '        FueraChange = False
    '    End Sub

    '    Private Sub msgJoyeria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgJoyeria.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        msgJoyeria.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    '    End Sub

    '    Private Sub msgJoyeria_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgJoyeria.Scroll
    '        dbcJFamilia.Visible = False
    '        dbcJLinea.Visible = False
    '        dbcJSubLinea.Visible = False
    '        txtJoyeria.Visible = False
    '        dbcJArticulo.Visible = False
    '    End Sub

    '    Sub msgRelojeria_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles msgRelojeria.KeyDownEvent
    '        If eventArgs.keyCode = System.Windows.Forms.Keys.Escape Then
    '            sstGrupos.Focus()
    '            sstGrupos.SelectedIndex = 1
    '        End If
    '        With msgRelojeria
    '            If Trim(.get_TextMatrix(.Row, C_ColRMARCA)) = "" Then
    '                Exit Sub
    '            End If
    '            Select Case eventArgs.keyCode
    '                Case System.Windows.Forms.Keys.Delete
    '                    If .get_TextMatrix(.Row, C_ColRESTATUS) = C_Cancelado Then
    '                        MsgBox("Esta promoción ya ha sido cancelada." & vbNewLine & "Verifique por favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                        .Focus()
    '                        Exit Sub
    '                    End If
    '                    '                If .TextMatrix(.Row, C_ColRESTATUS) = C_Aplicado And .TextMatrix(.Row, C_ColRESNUEVO) = False Then
    '                    '                    MsgBox "Esta promoción ya ha sido aplicada. No es posible eliminar" + vbNewLine + "Verifique Por Favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "Mensaje")
    '                        Case MsgBoxResult.No
    '                            .Focus()
    '                            Exit Sub
    '                        Case MsgBoxResult.Cancel
    '                            .Focus()
    '                            Exit Sub
    '                    End Select
    '                    .set_TextMatrix(.Row, C_ColRESTATUS, C_Cancelado)
    '                    PonerColor((.Row))
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Insert
    '                    '                If .TextMatrix(.Row, C_ColRESTATUS) = C_Aplicado Then
    '                    '                    MsgBox "Esta promoción ya ha sido aplicada." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    '                If .TextMatrix(.Row, C_ColRESTATUS) = C_Cancelado Then
    '                    '                    MsgBox "Esta promoción ha sido cancelado. No es posible aplicarla." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    '                If Numerico(.TextMatrix(.Row, C_ColRPORCDESCTO)) = 0 And Numerico(.TextMatrix(.Row, C_ColRPRECIO)) = 0 Then
    '                    '                    MsgBox "Información incompleta sobre la promocion." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    '                Select Case MsgBox("¿Desea Aplicar esta Promoción para que esté disponible para Operaciones de Venta", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Mensaje")
    '                    '                    Case vbNo
    '                    '                        .SetFocus
    '                    '                        Exit Sub
    '                    '                    Case vbCancel
    '                    '                        .SetFocus
    '                    '                        Exit Sub
    '                    '                End Select
    '                    '                .TextMatrix(.Row, C_ColRESTATUS) = C_Aplicado
    '                    '                PonerColor (.Row)
    '                    '                .SetFocus
    '            End Select
    '        End With
    '    End Sub

    '    Private Sub msgRelojeria_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgRelojeria.Scroll
    '        txtRelojeria.Visible = False
    '        dbcRMarca.Visible = False
    '        dbcRModelo.Visible = False
    '        dbcRArticulo.Visible = False
    '    End Sub

    '    Private Sub msgVarios_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgVarios.DblClick
    '        Dim Estatus As String
    '        Dim EstatusTag As String
    '        '    With msgVarios
    '        '        Estatus = .TextMatrix(.Row, C_ColJESTATUS)
    '        '        EstatusTag = .TextMatrix(.Row, C_ColJESTATUSTAG)
    '        '        If (Estatus = C_Aplicado And Estatus <> EstatusTag) Or (Estatus = C_Cancelado And Estatus <> EstatusTag) Then
    '        '            .TextMatrix(.Row, C_ColJESTATUS) = C_Vigente
    '        '            PonerColor (.Row)
    '        '            .SetFocus
    '        ''        ElseIf Estatus = C_Aplicado Then
    '        ''            MsgBox "No es posible modificar una Promoción Aplicada Previamente.", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '        ''            .SetFocus
    '        ''            Exit Sub
    '        '        End If
    '        '    End With
    '        msgVarios_KeyPressEvent(msgVarios, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    '    End Sub

    '    Private Sub msgVarios_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgVarios.Enter
    '        msgVarios.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '        '    msgVarios.Row = 1
    '        '    msgVarios.Col = 0
    '        Pon_Tool()
    '    End Sub

    '    Private Sub msgVarios_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles msgVarios.KeyDownEvent
    '        If eventArgs.keyCode = System.Windows.Forms.Keys.Escape Then
    '            sstGrupos.Focus()
    '            sstGrupos.SelectedIndex = 2
    '        End If
    '        With msgVarios
    '            If Trim(.get_TextMatrix(.Row, C_ColJFAMILIA)) = "" Then Exit Sub
    '            Select Case eventArgs.keyCode
    '                Case System.Windows.Forms.Keys.Delete
    '                    If .get_TextMatrix(.Row, C_ColJESTATUS) = C_Cancelado Then
    '                        MsgBox("Está Promoción ya ha sido Cancelada." & vbNewLine & "Verifique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                        .Focus()
    '                        Exit Sub
    '                    End If
    '                    '                If .TextMatrix(.Row, C_ColJESTATUS) = C_Aplicado And .TextMatrix(.Row, C_ColJESNUEVO) = False Then
    '                    '                    MsgBox "Está promoción ya ha sido aplicada. No es posible eliminar" + vbNewLine + "Verifique Por Favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "Mensaje")
    '                        Case MsgBoxResult.No
    '                            .Focus()
    '                            Exit Sub
    '                        Case MsgBoxResult.Cancel
    '                            .Focus()
    '                            Exit Sub
    '                    End Select
    '                    .set_TextMatrix(.Row, C_ColJESTATUS, C_Cancelado)
    '                    PonerColor((.Row))
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Insert
    '                    '                If .TextMatrix(.Row, C_ColJESTATUS) = C_Aplicado Then
    '                    '                    MsgBox "Esta promoción ya ha sido aplicada." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    '                If .TextMatrix(.Row, C_ColJESTATUS) = C_Cancelado Then
    '                    '                    MsgBox "Esta promoción ha sido cancelado. No es posible aplicarla." + vbNewLine + "Verifique Por Favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    '                If Numerico(.TextMatrix(.Row, C_ColJPORCDESCTO)) = 0 And Numerico(.TextMatrix(.Row, C_ColJPRECIO)) = 0 Then
    '                    '                    MsgBox "Información incompleta sobre la promocion." + vbNewLine + "Verifique Por Favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '                    '                 Select Case MsgBox("¿Desea aplicar esta promoción para que esté disponible para operaciones de venta", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Mensaje")
    '                    '                    Case vbNo
    '                    '                        .SetFocus
    '                    '                        Exit Sub
    '                    '                    Case vbCancel
    '                    '                        .SetFocus
    '                    '                        Exit Sub
    '                    '                End Select
    '                    '                .TextMatrix(.Row, C_ColJESTATUS) = C_Aplicado
    '                    '                PonerColor (.Row)
    '                    '                .SetFocus
    '            End Select
    '        End With
    '    End Sub

    '    Private Sub msgVarios_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles msgVarios.KeyPressEvent
    '        Dim EsNuevo As Boolean
    '        Dim Estatus As String
    '        '    Dim EstatusTag  As String
    '        EsNuevo = True
    '        With msgVarios
    '            ' Si nho se trata de un REgistro nuevo, no se podrá editar el Grid
    '            If Trim(.get_TextMatrix(.Row, C_ColJESNUEVO)) <> "" Then
    '                EsNuevo = CBool(.get_TextMatrix(.Row, C_ColJESNUEVO))
    '            End If
    '            Estatus = .get_TextMatrix(.Row, C_ColJESTATUS)
    '            FueraChange = True
    '            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then 'Para que cuando sea escape, no entre a editar el codigo,simplemente que se regrese al control anterior
    '                Select Case .Col
    '                    Case C_ColJFAMILIA ''-------------- SE EDITA LA FAMILIA ---------------------'''''
    '                        If EsNuevo = False Or mblnNuevo = False Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        '''en esta parte se validará si es el rengón, columna que le corresponde editarse
    '                        If (.Row > 1) Then
    '                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
    '                            If Trim(.get_TextMatrix(.Row - 1, C_ColJFAMILIA)) = "" Then
    '                                .Focus()
    '                                Exit Sub
    '                            End If
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgVarios, dbcJFamilia, eventArgs.keyAscii)
    '                    Case C_ColJLINEA ''-------------- SE EDITA LA LINEA ---------------------'''''
    '                        If EsNuevo = False Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColJFAMILIA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgVarios, dbcJLinea, eventArgs.keyAscii)
    '                    Case C_ColJARTICULO
    '                        If EsNuevo = False Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColJFAMILIA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgVarios, txtArticuloV, eventArgs.keyAscii)
    '                    Case C_ColJPORCDESCTO ''-------------- SE EDITA EL PORCENTAJE DE DESCTO---------------------'''''
    '                        'Or Estatus = C_Aplicado
    '                        ModEstandar.gp_CampoNumerico(eventArgs.keyAscii, ".")
    '                        If Estatus = C_Cancelado Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColJFAMILIA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, C_ColJPRECIO))) <> 0 Then
    '                            MsgBox("ya se ha Asignado un Precio. No es Posible Asignar un Porcentaje de Descuento.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgVarios, txtJoyeria, eventArgs.keyAscii)
    '                        .set_TextMatrix(.Row, C_ColJPRECIO, "0.00")
    '                    Case C_ColJPRECIO ''-------------- SE EDITA LA EL PRECIO ---------------------'''''
    '                        'Or Estatus = C_Aplicado
    '                        ModEstandar.gp_CampoNumerico(eventArgs.keyAscii, ".")
    '                        If Estatus = C_Cancelado Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColJFAMILIA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, C_ColJPORCDESCTO))) <> 0 Then
    '                            MsgBox("Ya se ha Asignado un Porcentaje de Descuento. No es Posible Asignar el Precio.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgVarios, txtJoyeria, eventArgs.keyAscii)
    '                        .set_TextMatrix(.Row, C_ColJPORCDESCTO, "0.00")
    '                End Select
    '            End If
    '        End With
    '        FueraChange = False
    '    End Sub

    '    Private Sub msgVarios_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgVarios.Leave
    '        msgVarios.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    '    End Sub

    '    Private Sub msgVarios_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgVarios.Scroll
    '        txtJoyeria.Visible = False
    '        dbcJFamilia.Visible = False
    '        dbcJLinea.Visible = False
    '        dbcJSubLinea.Visible = False
    '        txtArticulo.Visible = False
    '    End Sub

    '    Private Sub msgXArticulo_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgXArticulo.DblClick
    '        msgXArticulo_KeyPressEvent(msgXArticulo, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    '    End Sub

    '    Private Sub msgXArticulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgXArticulo.Enter
    '        msgXArticulo.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '        Pon_Tool()
    '    End Sub

    '    Private Sub msgXArticulo_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles msgXArticulo.KeyDownEvent
    '        With msgXArticulo
    '            'If Trim(.TextMatrix(.Row, C_COLXARTCODARTICULO)) = "" Then Exit Sub
    '            Select Case eventArgs.keyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    sstGrupos.Focus()
    '                    sstGrupos.SelectedIndex = 3
    '                Case System.Windows.Forms.Keys.Delete
    '                    If .get_TextMatrix(.Row, C_COLXARTESTATUS) = C_Cancelado Then
    '                        MsgBox("Está Promoción ya ha sido Cancelada." & vbNewLine & "Verifique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                        .Focus()
    '                        Exit Sub
    '                    End If
    '                    Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "Mensaje")
    '                        Case MsgBoxResult.No
    '                            .Focus()
    '                            Exit Sub
    '                        Case MsgBoxResult.Cancel
    '                            .Focus()
    '                            Exit Sub
    '                    End Select
    '                    .set_TextMatrix(.Row, C_COLXARTESTATUS, C_Cancelado)
    '                    PonerColor((.Row))
    '                    .Focus()
    '                    '            Case vbKeyReturn
    '                    '                msgXArticulo_KeyPress vbKeyReturn
    '            End Select
    '        End With
    '    End Sub

    '    Private Sub msgXArticulo_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles msgXArticulo.KeyPressEvent
    '        Dim EsNuevo As Boolean
    '        Dim Estatus As String
    '        EsNuevo = True
    '        With msgXArticulo
    '            'Si nho se trata de un REgistro nuevo, no se podrá editar el Grid
    '            If Trim(.get_TextMatrix(.Row, C_COLXARTESNUEVO)) <> "" Then
    '                EsNuevo = CBool(.get_TextMatrix(.Row, C_COLXARTESNUEVO))
    '            End If
    '            Estatus = .get_TextMatrix(.Row, C_COLXARTESTATUS)
    '            FueraChange = True
    '            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then 'Para que cuando sea escape, no entre a editar el codigo,simplemente que se regrese al control anterior
    '                Select Case .Col
    '                    Case C_COLXARTCODARTICULO '------- SE EDITA EL CODIGO DE ARTICULO -------
    '                        If EsNuevo = False Or mblnNuevo = False Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        '''en esta parte se validará si es el rengón, columna que le corresponde editarse
    '                        If (.Row > 1) Then
    '                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
    '                            If Trim(.get_TextMatrix(.Row - 1, C_COLXARTCODARTICULO)) = "" Then
    '                                .Focus()
    '                                FueraChange = False
    '                                Exit Sub
    '                            End If
    '                        End If
    '                        txtFlex.MaxLength = 8
    '                        txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '                        ModEstandar.MSHFlexGridEdit(msgXArticulo, txtFlex, eventArgs.keyAscii)
    '                    Case C_COLXARTDESCARTICULO
    '                        If EsNuevo = False Or mblnNuevo = False Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        txtFlex.MaxLength = 150
    '                        txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
    '                        ModEstandar.MSHFlexGridEdit(msgXArticulo, txtFlex, eventArgs.keyAscii)
    '                    Case C_COLXARTCODANTERIOR
    '                        .Focus()
    '                        FueraChange = False
    '                        Exit Sub
    '                    Case C_COLXARTPORCDESCTO
    '                        ModEstandar.gp_CampoNumerico(eventArgs.keyAscii, ".")
    '                        If Estatus = C_Cancelado Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        If Trim(.get_TextMatrix(.Row, C_COLXARTCODARTICULO)) = "" Then
    '                            .Focus()
    '                            FueraChange = False
    '                            Exit Sub
    '                        End If
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, C_COLXARTPRECIO))) <> 0 Then
    '                            MsgBox("ya se ha Asignado un Precio. No es Posible Asignar un Porcentaje de Descuento.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            .Focus()
    '                            FueraChange = False
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgXArticulo, txtFlex, eventArgs.keyAscii)
    '                        .set_TextMatrix(.Row, C_COLXARTPRECIO, "0.00")
    '                        txtFlex.MaxLength = 6
    '                        txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '                    Case C_COLXARTPRECIO
    '                        ModEstandar.gp_CampoNumerico(eventArgs.keyAscii, ".")
    '                        If Estatus = C_Cancelado Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        If Trim(.get_TextMatrix(.Row, C_COLXARTCODARTICULO)) = "" Then
    '                            .Focus()
    '                            FueraChange = False
    '                            Exit Sub
    '                        End If
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, C_COLXARTPORCDESCTO))) <> 0 Then
    '                            MsgBox("Ya se ha Asignado un Porcentaje de Descuento. No es Posible Asignar el Precio.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            .Focus()
    '                            FueraChange = False
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgXArticulo, txtFlex, eventArgs.keyAscii)
    '                        .set_TextMatrix(.Row, C_COLXARTPORCDESCTO, "0.00")
    '                        txtFlex.MaxLength = 13
    '                        txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '                End Select
    '            End If
    '        End With
    '        FueraChange = False
    '    End Sub

    '    Private Sub msgXArticulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgXArticulo.Leave
    '        msgXArticulo.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    '    End Sub

    '    Private Sub sstGrupos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sstGrupos.Enter
    '        '''JOYERIA-VARIOS
    '        dbcJFamilia.Visible = False
    '        dbcJLinea.Visible = False
    '        dbcJSubLinea.Visible = False
    '        txtJoyeria.Visible = False
    '        '''RELOJERIA
    '        dbcRMarca.Visible = False
    '        dbcRModelo.Visible = False
    '        txtFlex.Visible = False
    '    End Sub

    '    Private Sub sstGrupos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles sstGrupos.KeyDown
    '        Dim KeyCode As Integer = eventArgs.KeyCode
    '        Dim Shift As Integer = eventArgs.KeyData \ &H10000
    '        '''If dtpFechaFinJ.Enabled = False Then Exit Sub
    '        If KeyCode = System.Windows.Forms.Keys.Escape Then dtpFechaFinJ.Focus()
    '        If KeyCode = System.Windows.Forms.Keys.Return Then
    '            Select Case sstGrupos.SelectedIndex
    '                Case 0 : msgJoyeria.Focus()
    '                Case 1 : msgRelojeria.Focus()
    '                Case 2 : msgVarios.Focus()
    '                Case 3 : msgXArticulo.Focus()
    '                Case 4 : dbcProveedor.Focus()
    '            End Select
    '        End If
    '    End Sub

    '    Private Sub sstGrupos_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sstGrupos.Leave
    '        '''    If FueraChange = True Then Exit Sub
    '        '''    Select Case sstGrupos.Tab
    '        '''        Case 0
    '        '''            msgJoyeria.SetFocus
    '        '''        Case 1
    '        '''            msgRelojeria.SetFocus
    '        '''        Case 2
    '        '''            msgVarios.SetFocus
    '        '''        Case 3
    '        '''            msgXArticulo.SetFocus
    '        '''        Case 4
    '        '''            dbcProveedor.SetFocus
    '        '''    End Select
    '    End Sub

    '    Private Sub txtArticulo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtArticulo.KeyDown
    '        Dim KeyCode As Integer = eventArgs.KeyCode
    '        Dim Shift As Integer = eventArgs.KeyData \ &H10000
    '        tecla = KeyCode
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 1 Then
    '            GridACtivo = msgRelojeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        Else
    '            Exit Sub
    '        End If
    '        With GridACtivo
    '            FueraChange = True
    '            Select Case KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    txtArticulo.Visible = False
    '                    txtArticulo.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    txtArticulo.Visible = False
    '                    msgJoyeria.Focus()
    '                    '                intCodSubLinea = 0
    '                    '                .TextMatrix(.Row, C_ColJCODARTICULO) = 0
    '                    '                gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & IIf((GridACtivo Is msgJoyeria), gCODJOYERIA, gCODVARIOS) & ") " & _
    '                    ''                    IIf((Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) = 0), " ", " And (CodFamilia = " & Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) & ")") & _
    '                    ''                    IIf((Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODLINEA)) = 0), "", " And (CodLinea = " & Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODLINEA)) & ")") & _
    '                    ''                    IIf((Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODSUBLINEA)) = 0), "", " And (CodSubLinea = " & Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODSUBLINEA)) & ")") & _
    '                    ''                    "and DescArticulo LIKE '" & Trim(txtArticulo) & "%'"
    '                    '
    '                    '                ModDCombo.DCLostFocus dbcJArticulo, gStrSql, intCodArticulo
    '                    '                .TextMatrix(.Row, .Col) = Trim(dbcJArticulo)
    '                    '                .TextMatrix(.Row, C_ColJCODARTICULO) = intCodArticulo
    '                    '                LlenaDatosArticulo intCodArticulo, IIf((GridACtivo Is msgJoyeria), gCODJOYERIA, gCODVARIOS)
    '                    '                .Col = C_ColJPORCDESCTO
    '                    '                dbcJArticulo = ""
    '                    '                dbcJArticulo.Visible = False
    '                    '                .FocusRect = flexFocusNone
    '                    '                .SetFocus
    '                    '                .ColAlignment(C_ColJARTICULO) = flexAlignLeftCenter
    '                    'Verificar si la Promoción está Siendo Repetida
    '                    '                ValidarPromocionTecleadaRepetida
    '                    '                intCodFamilia = Numerico(.TextMatrix(.Row, C_ColJCODFAMILIA))
    '                    '                intCodLinea = Numerico(.TextMatrix(.Row, C_ColJCODLINEA))
    '                    '                intCodSubLinea = Numerico(.TextMatrix(.Row, C_ColJCODSUBLINEA))
    '                    '                If ValidarPromocionGuardadaRepetida(gCODJOYERIA, intCodFamilia, intCodLinea, intCodSubLinea, 0, 0, dtpFechaInIcioJ, dtpFechaFinJ) = True Then
    '                    '                    MsgBox "Existe una Promoción registrada para este Artículo." + vbNewLine + "No es posible duplicar Promociones en un Lapso de Tiempo Similar", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    LimpiaDatosFamilia
    '                    '                    .TextMatrix(.Row, C_ColJFAMILIA) = ""
    '                    '                    .Col = C_ColJFAMILIA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '            End Select
    '            FueraChange = False
    '        End With
    '    End Sub

    '    Private Sub txtArticulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtArticulo.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        txtArticulo.Visible = False
    '    End Sub

    '    Private Sub txtArticuloR_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtArticuloR.KeyDown
    '        Dim KeyCode As Integer = eventArgs.KeyCode
    '        Dim Shift As Integer = eventArgs.KeyData \ &H10000
    '        tecla = KeyCode
    '        GridACtivo.Text = msgRelojeria.Text
    '        With GridACtivo
    '            FueraChange = True
    '            Select Case KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    txtArticuloR.Visible = False
    '                    txtArticuloR.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    txtArticuloR.Visible = False
    '                    msgRelojeria.Focus()
    '                    '                intCodSubLinea = 0
    '                    '                .TextMatrix(.Row, C_ColJCODARTICULO) = 0
    '                    '                gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & IIf((GridACtivo Is msgJoyeria), gCODJOYERIA, gCODVARIOS) & ") " & _
    '                    ''                    IIf((Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) = 0), " ", " And (CodFamilia = " & Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) & ")") & _
    '                    ''                    IIf((Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODLINEA)) = 0), "", " And (CodLinea = " & Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODLINEA)) & ")") & _
    '                    ''                    IIf((Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODSUBLINEA)) = 0), "", " And (CodSubLinea = " & Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODSUBLINEA)) & ")") & _
    '                    ''                    "and DescArticulo LIKE '" & Trim(txtArticulo) & "%'"
    '                    '
    '                    '                ModDCombo.DCLostFocus dbcJArticulo, gStrSql, intCodArticulo
    '                    '                .TextMatrix(.Row, .Col) = Trim(dbcJArticulo)
    '                    '                .TextMatrix(.Row, C_ColJCODARTICULO) = intCodArticulo
    '                    '                LlenaDatosArticulo intCodArticulo, IIf((GridACtivo Is msgJoyeria), gCODJOYERIA, gCODVARIOS)
    '                    '                .Col = C_ColJPORCDESCTO
    '                    '                dbcJArticulo = ""
    '                    '                dbcJArticulo.Visible = False
    '                    '                .FocusRect = flexFocusNone
    '                    '                .SetFocus
    '                    '                .ColAlignment(C_ColJARTICULO) = flexAlignLeftCenter
    '                    'Verificar si la Promoción está Siendo Repetida
    '                    '                ValidarPromocionTecleadaRepetida
    '                    '                intCodFamilia = Numerico(.TextMatrix(.Row, C_ColJCODFAMILIA))
    '                    '                intCodLinea = Numerico(.TextMatrix(.Row, C_ColJCODLINEA))
    '                    '                intCodSubLinea = Numerico(.TextMatrix(.Row, C_ColJCODSUBLINEA))
    '                    '                If ValidarPromocionGuardadaRepetida(gCODJOYERIA, intCodFamilia, intCodLinea, intCodSubLinea, 0, 0, dtpFechaInIcioJ, dtpFechaFinJ) = True Then
    '                    '                    MsgBox "Existe una Promoción registrada para este Artículo." + vbNewLine + "No es posible duplicar Promociones en un Lapso de Tiempo Similar", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    LimpiaDatosFamilia
    '                    '                    .TextMatrix(.Row, C_ColJFAMILIA) = ""
    '                    '                    .Col = C_ColJFAMILIA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '            End Select
    '            FueraChange = False
    '        End With
    '    End Sub

    '    Private Sub txtArticuloR_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtArticuloR.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        txtArticuloR.Visible = False
    '    End Sub

    '    Private Sub txtArticuloV_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtArticuloV.KeyDown
    '        Dim KeyCode As Integer = eventArgs.KeyCode
    '        Dim Shift As Integer = eventArgs.KeyData \ &H10000
    '        tecla = KeyCode
    '        GridACtivo.Text = msgVarios.Text
    '        With GridACtivo
    '            FueraChange = True
    '            Select Case KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    txtArticuloV.Visible = False
    '                    txtArticuloV.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    txtArticuloV.Visible = False
    '                    msgVarios.Focus()
    '                    '                intCodSubLinea = 0
    '                    '                .TextMatrix(.Row, C_ColJCODARTICULO) = 0
    '                    '                gStrSql = "SELECT CodArticulo,DescArticulo=Ltrim(Rtrim(DescArticulo)) From dbo.CatArticulos Where (CodGrupo = " & IIf((GridACtivo Is msgJoyeria), gCODJOYERIA, gCODVARIOS) & ") " & _
    '                    ''                    IIf((Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) = 0), " ", " And (CodFamilia = " & Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODFAMILIA)) & ")") & _
    '                    ''                    IIf((Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODLINEA)) = 0), "", " And (CodLinea = " & Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODLINEA)) & ")") & _
    '                    ''                    IIf((Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODSUBLINEA)) = 0), "", " And (CodSubLinea = " & Numerico(GridACtivo.TextMatrix(GridACtivo.Row, C_ColJCODSUBLINEA)) & ")") & _
    '                    ''                    "and DescArticulo LIKE '" & Trim(txtArticulo) & "%'"
    '                    '
    '                    '                ModDCombo.DCLostFocus dbcJArticulo, gStrSql, intCodArticulo
    '                    '                .TextMatrix(.Row, .Col) = Trim(dbcJArticulo)
    '                    '                .TextMatrix(.Row, C_ColJCODARTICULO) = intCodArticulo
    '                    '                LlenaDatosArticulo intCodArticulo, IIf((GridACtivo Is msgJoyeria), gCODJOYERIA, gCODVARIOS)
    '                    '                .Col = C_ColJPORCDESCTO
    '                    '                dbcJArticulo = ""
    '                    '                dbcJArticulo.Visible = False
    '                    '                .FocusRect = flexFocusNone
    '                    '                .SetFocus
    '                    '                .ColAlignment(C_ColJARTICULO) = flexAlignLeftCenter
    '                    'Verificar si la Promoción está Siendo Repetida
    '                    '                ValidarPromocionTecleadaRepetida
    '                    '                intCodFamilia = Numerico(.TextMatrix(.Row, C_ColJCODFAMILIA))
    '                    '                intCodLinea = Numerico(.TextMatrix(.Row, C_ColJCODLINEA))
    '                    '                intCodSubLinea = Numerico(.TextMatrix(.Row, C_ColJCODSUBLINEA))
    '                    '                If ValidarPromocionGuardadaRepetida(gCODJOYERIA, intCodFamilia, intCodLinea, intCodSubLinea, 0, 0, dtpFechaInIcioJ, dtpFechaFinJ) = True Then
    '                    '                    MsgBox "Existe una Promoción registrada para este Artículo." + vbNewLine + "No es posible duplicar Promociones en un Lapso de Tiempo Similar", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '                    '                    LimpiaDatosFamilia
    '                    '                    .TextMatrix(.Row, C_ColJFAMILIA) = ""
    '                    '                    .Col = C_ColJFAMILIA
    '                    '                    .SetFocus
    '                    '                    Exit Sub
    '                    '                End If
    '            End Select
    '            FueraChange = False
    '        End With
    '    End Sub

    '    Private Sub txtArticuloV_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtArticuloV.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        txtArticuloV.Visible = False
    '    End Sub

    '    Private Sub txtDesctoP_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesctoP.Enter
    '        ModEstandar.SelTextoTxt(txtDesctoP)
    '    End Sub

    '    Private Sub txtDesctoP_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDesctoP.KeyPress
    '        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
    '        ModEstandar.gp_CampoNumerico(KeyAscii, ".")
    '        KeyAscii = ModEstandar.MskCantidad(txtDesctoP.Text, KeyAscii, 3, 2, (txtDesctoP.SelectionStart))
    '        eventArgs.KeyChar = Chr(KeyAscii)
    '        If KeyAscii = 0 Then
    '            eventArgs.Handled = True
    '        End If
    '    End Sub

    '    Private Sub txtDesctoP_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesctoP.Leave
    '        If CDec(ModEstandar.Numerico(Trim(txtDesctoP.Text))) > 100 Then
    '            MsgBox("El porcentaje no puede ser mayor de 100", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
    '            txtDesctoP.Text = "0.00"
    '            txtDesctoP.Focus()
    '            Exit Sub
    '        End If
    '        txtDesctoP.Text = Format(CDec(ModEstandar.Numerico(Trim(txtDesctoP.Text))), "##0.00")
    '        chkAplicar.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        chkBorrar.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        chkCancelarP.CheckState = System.Windows.Forms.CheckState.Unchecked
    '    End Sub

    '    Private Sub txtFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Enter
    '        txtFlex.Text = Trim(txtFlex.Text)
    '        If Len(txtFlex.Text) > 1 Then
    '            SelTextoTxt(txtFlex)
    '        End If
    '        Pon_Tool()
    '    End Sub

    '    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
    '        Dim KeyCode As Integer = eventArgs.KeyCode
    '        Dim Shift As Integer = eventArgs.KeyData \ &H10000
    '        'Aqui se muestran los datos del control editable, en el Grid
    '        'Se deberá formatear el Valor de Acuerdo al Tipo de Dato en uso
    '        Dim rowsiguiente As Integer
    '        Dim ColSiguiente As Integer
    '        Dim ResBusquedaArt As Integer
    '        Select Case KeyCode
    '            Case System.Windows.Forms.Keys.Escape
    '                txtFlex.Focus()
    '                txtFlex.Visible = False
    '                txtFlex.Text = ""
    '                msgXArticulo.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                msgXArticulo.Focus()
    '            Case System.Windows.Forms.Keys.Return
    '                With msgXArticulo
    '                    If .Row > 1 Then
    '                        If Trim(.get_TextMatrix(.Row - 1, C_COLXARTCODARTICULO)) = "" Then Exit Sub
    '                    End If
    '                    rowsiguiente = .Row
    '                    If .Col = C_COLXARTCODARTICULO Then
    '                        ColSiguiente = C_COLXARTPORCDESCTO
    '                        If Trim(txtFlex.Text) <> "" Then
    '                            ResBusquedaArt = BuscarCodigoArticulo(Trim(txtFlex.Text))
    '                            If ResBusquedaArt > 0 Or ResBusquedaArt = -1 Then
    '                                If ResBusquedaArt > 0 Then
    '                                    LlenaDatosXArticulo(ResBusquedaArt)
    '                                ElseIf ResBusquedaArt = -1 Then
    '                                    LlenaDatosXArticulo(CInt(Numerico(txtFlex.Text)))
    '                                End If
    '                                Exit Sub
    '                            ElseIf ResBusquedaArt = -2 And CDbl(Numerico(txtFlex.Text)) <> 0 Then
    '                                ResBusquedaArt = CInt(txtFlex.Text)
    '                                .set_TextMatrix(.Row, C_COLXARTCODARTICULO, "")
    '                                BuscarArticulos(True, (New String("0", 6) & CStr(ResBusquedaArt)))
    '                                'BuscarArticulos True, Right(String(6, "0") + Trim(.TextMatrix(.Row, C_ColCODIGO)), 6)
    '                                Exit Sub
    '                            End If
    '                            LlenaDatosXArticulo(CInt(Numerico(txtFlex.Text)))
    '                        Else
    '                            .set_TextMatrix(.Row, C_COLXARTCODARTICULO, "")
    '                            .set_TextMatrix(.Row, C_COLXARTDESCARTICULO, "")
    '                            .set_TextMatrix(.Row, C_COLXARTCODANTERIOR, "")
    '                            .set_TextMatrix(.Row, C_COLXARTCODFAMILIA, "")
    '                            .set_TextMatrix(.Row, C_COLXARTCODGRUPO, "")
    '                            .set_TextMatrix(.Row, C_COLXARTCODLINEA, "")
    '                            .set_TextMatrix(.Row, C_COLXARTCODMARCA, "")
    '                            .set_TextMatrix(.Row, C_COLXARTCODMODELO, "")
    '                            .set_TextMatrix(.Row, C_COLXARTCODSUBLINEA, "")
    '                            .set_TextMatrix(.Row, C_COLXARTESNUEVO, "")
    '                            .set_TextMatrix(.Row, C_COLXARTESTATUS, "")
    '                            .set_TextMatrix(.Row, C_COLXARTPRECIO, "")
    '                            .set_TextMatrix(.Row, C_COLXARTPORCDESCTO, "")
    '                        End If
    '                    ElseIf .Col = C_COLXARTPRECIO Then
    '                        .set_TextMatrix(.Row, .Col, Format(Numerico(txtFlex.Text), gstrFormatoCantidad))
    '                        rowsiguiente = .Row + 1
    '                        ColSiguiente = C_COLXARTCODARTICULO
    '                    ElseIf .Col = C_COLXARTPORCDESCTO Then
    '                        If CDbl(Numerico(txtFlex.Text)) > 100 Then
    '                            MsgBox("El Porcentaje de Descuento no puede ser mayor de 100.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            txtFlex.Focus()
    '                            Exit Sub
    '                        End If
    '                        .set_TextMatrix(.Row, .Col, Format(Numerico(txtFlex.Text), gstrFormatoCantidad))
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, .Col))) = 0 Then
    '                            rowsiguiente = .Row
    '                            ColSiguiente = C_COLXARTPRECIO
    '                        Else
    '                            rowsiguiente = .Row + 1
    '                            ColSiguiente = C_COLXARTCODARTICULO
    '                        End If
    '                    End If
    '                    FueraChange = True
    '                    txtFlex.Text = ""
    '                    txtFlex.Visible = False
    '                    FueraChange = False
    '                    .Col = .Col
    '                    .Row = .Row
    '                    .Focus()
    '                    If (.Col = C_COLXARTCODARTICULO Or .Col = C_COLXARTDESCARTICULO) And (Trim(.get_TextMatrix(.Row, .Col)) = "") Then
    '                        Exit Sub
    '                    End If
    '                    If .Row = .Rows - 1 Then
    '                        .Rows = .Rows + 1
    '                        .Row = rowsiguiente
    '                        .Col = ColSiguiente
    '                        .set_TextMatrix(.Row, C_COLXARTTIPO, "A")
    '                    Else
    '                        .Row = rowsiguiente
    '                        .Col = ColSiguiente
    '                    End If
    '                    If .Row > 7 Then
    '                        .TopRow = .Row
    '                    End If
    '                End With
    '        End Select
    '    End Sub

    '    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
    '        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
    '        'En este Evento se validan los datos que se introduzcan al control txtjoyeria,dependiendo de la columan en que se esté editando
    '        If KeyAscii = 0 Or KeyAscii = 13 Then GoTo EventExitSub
    '        With msgXArticulo
    '            Select Case .Col
    '                Case C_COLXARTCODARTICULO
    '                    KeyAscii = ModEstandar.MskCantidad(txtFlex.Text, KeyAscii, 8, 0, (txtFlex.SelectionStart))
    '                Case C_COLXARTDESCARTICULO
    '                Case C_COLXARTPRECIO
    '                    ModEstandar.gp_CampoNumerico(KeyAscii, ".")
    '                    KeyAscii = ModEstandar.MskCantidad(txtFlex.Text, KeyAscii, 10, 2, (txtFlex.SelectionStart))
    '                Case C_COLXARTPORCDESCTO
    '                    ModEstandar.gp_CampoNumerico(KeyAscii, ".")
    '                    KeyAscii = ModEstandar.MskCantidad(txtFlex.Text, KeyAscii, 3, 2, (txtFlex.SelectionStart))
    '            End Select
    '        End With
    'EventExitSub:
    '        eventArgs.KeyChar = Chr(KeyAscii)
    '        If KeyAscii = 0 Then
    '            eventArgs.Handled = True
    '        End If
    '    End Sub

    '    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        txtFlex.Visible = False
    '    End Sub

    '    Private Sub txtjoyeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtJoyeria.Enter
    '        txtJoyeria.Text = Trim(txtJoyeria.Text)
    '        If Len(txtJoyeria.Text) > 1 Then
    '            SelTextoTxt(txtJoyeria)
    '        End If
    '        Pon_Tool()
    '    End Sub

    '    Private Sub txtjoyeria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtJoyeria.KeyDown
    '        Dim KeyCode As Integer = eventArgs.KeyCode
    '        Dim Shift As Integer = eventArgs.KeyData \ &H10000
    '        'Aqui se muestran los datos del control editable, en el Grid
    '        'Se deberá formatear el Valor de Acuerdo al Tipo de Dato en uso
    '        Dim rowsiguiente As Integer
    '        Dim ColSiguiente As Integer
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        End If
    '        With GridACtivo
    '            Select Case KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    .Focus()
    '                    txtJoyeria.Visible = False
    '                    txtJoyeria.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    ValidarPromocionTecleadaRepetida()
    '                    rowsiguiente = .Row
    '                    intCodFamilia = CShort(Numerico(.get_TextMatrix(.Row, C_ColJCODFAMILIA)))
    '                    intCodLinea = CShort(Numerico(.get_TextMatrix(.Row, C_ColJCODLINEA)))
    '                    intCodSubLinea = CShort(Numerico(.get_TextMatrix(.Row, C_ColJCODSUBLINEA)))
    '                    intCodArticulo = CInt(Numerico(.get_TextMatrix(.Row, C_ColJCODARTICULO)))
    '                    If ValidarPromocionGuardadaRepetida(IIf((GridACtivo Is msgJoyeria), gCODJOYERIA, gCODVARIOS), intCodFamilia, intCodLinea, intCodSubLinea, 0, 0, intCodArticulo, dtpFechaInIcioJ.Value, dtpFechaFinJ.Value, "G") = True And mblnNuevo = True Then
    '                        MsgBox("Existe una promoción registrada para este artículo." & vbNewLine & "No es posible duplicar promociones en un lapso de tiempo similar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                        LimpiaDatosFamilia()
    '                        .set_TextMatrix(.Row, C_ColJFAMILIA, "")
    '                        .Col = C_ColJFAMILIA
    '                        .Focus()
    '                        '                    Exit Sub
    '                    End If
    '                    'Si la Columna en que se está escribiendo es Codigo o Cantidad, Formatear el Valor par que quede numérico
    '                    If .Col = C_ColJPORCDESCTO Then
    '                        If CDbl(Numerico(txtJoyeria.Text)) > 100 Then
    '                            MsgBox("El Porcentaje de Descuento no puede ser mayor de 100.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            txtJoyeria.Focus()
    '                            Exit Sub
    '                        End If
    '                        .set_TextMatrix(.Row, .Col, VB6.Format(Numerico(txtJoyeria.Text), gstrFormatoCantidad))
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, .Col))) = 0 Then
    '                            rowsiguiente = .Row
    '                            ColSiguiente = C_ColJPRECIO
    '                        Else
    '                            rowsiguiente = .Row + 1
    '                            ColSiguiente = C_ColJFAMILIA
    '                        End If
    '                    ElseIf .Col = C_ColJPRECIO Then
    '                        .set_TextMatrix(.Row, .Col, VB6.Format(Numerico(txtJoyeria.Text), gstrFormatoCantidad))
    '                        rowsiguiente = .Row + 1
    '                        ColSiguiente = C_ColJFAMILIA
    '                    End If
    '                    FueraChange = True
    '                    txtJoyeria.Text = ""
    '                    txtJoyeria.Visible = False
    '                    dbcJFamilia.Text = ""
    '                    dbcJLinea.Text = ""
    '                    dbcJSubLinea.Text = ""
    '                    dbcJFamilia.Visible = False
    '                    dbcJLinea.Visible = False
    '                    dbcJSubLinea.Visible = False
    '                    .Col = .Col
    '                    .Row = .Row
    '                    .Focus()
    '                    If .Row = .Rows - 1 Then
    '                        .Rows = .Rows + 1
    '                        .Row = rowsiguiente
    '                        .Col = ColSiguiente
    '                        If GridACtivo.Name = "msgJoyeria" Or GridACtivo.Name = "msgVarios" Then
    '                            .set_TextMatrix(.Row, C_COLJTIPO, "G")
    '                        End If
    '                    Else
    '                        .Row = rowsiguiente
    '                        .Col = ColSiguiente
    '                    End If
    '                    If .Row > 7 Then
    '                        .TopRow = .Row
    '                    End If
    '                Case System.Windows.Forms.Keys.Delete
    '            End Select
    '        End With
    '    End Sub

    '    Private Sub txtjoyeria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtJoyeria.KeyPress
    '        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
    '        'En este Evento se validan los datos que se introduzcan al control txtjoyeria,dependiendo de la columan en que se esté editando
    '        If KeyAscii = 0 Or KeyAscii = 13 Then GoTo EventExitSub
    '        With msgJoyeria
    '            If .Col = C_ColJPORCDESCTO Then
    '                KeyAscii = ModEstandar.MskCantidad(txtJoyeria.Text, KeyAscii, 3, 2, (txtJoyeria.SelectionStart))
    '            End If
    '            If .Col = C_ColJPRECIO Then
    '                KeyAscii = ModEstandar.MskCantidad(txtJoyeria.Text, KeyAscii, 10, 2, (txtJoyeria.SelectionStart))
    '            End If
    '        End With
    'EventExitSub:
    '        eventArgs.KeyChar = Chr(KeyAscii)
    '        If KeyAscii = 0 Then
    '            eventArgs.Handled = True
    '        End If
    '    End Sub

    '    Sub InicializaVariables()
    '        FueraChange = False
    '        mblnSalir = False
    '        mblnNuevo = True
    '    End Sub

    '    Sub LimpiaDatosFamilia()
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        Else
    '            Exit Sub
    '        End If
    '        With GridACtivo
    '            FueraChange = True
    '            '        .TextMatrix(.Row, C_ColJCODFAMILIA) = ""
    '            .set_TextMatrix(.Row, C_ColJLINEA, "")
    '            '        .TextMatrix(.Row, C_ColJSUBLINEA) = ""
    '            '        .TextMatrix(.Row, C_ColJARTICULO) = ""
    '            .set_TextMatrix(.Row, C_ColJCODLINEA, "")
    '            '        .TextMatrix(.Row, C_ColJCODSUBLINEA) = ""
    '            '        .TextMatrix(.Row, C_ColJCODARTICULO) = ""
    '            FueraChange = False
    '        End With
    '        LimpiaDatosLinea()
    '        '    LimpiaDatosPrecioYDescuento
    '    End Sub

    '    Sub LimpiaDatosLinea()
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        End If
    '        With GridACtivo
    '            FueraChange = True
    '            .set_TextMatrix(.Row, C_ColJCODSUBLINEA, "")
    '            .set_TextMatrix(.Row, C_ColJSUBLINEA, "")
    '            FueraChange = False
    '        End With
    '        LimpiaDatosArticulo()
    '    End Sub

    '    Sub LimpiaDatosArticulo()
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 1 Then
    '            GridACtivo = msgRelojeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        End If
    '        If GridACtivo Is msgRelojeria Then
    '            With GridACtivo
    '                FueraChange = True
    '                .set_TextMatrix(.Row, C_ColRARTICULO, "")
    '                .set_TextMatrix(.Row, C_ColRCODARTICULO, "")
    '                FueraChange = False
    '            End With
    '        Else
    '            With GridACtivo
    '                FueraChange = True
    '                .set_TextMatrix(.Row, C_ColJARTICULO, "")
    '                .set_TextMatrix(.Row, C_ColJCODARTICULO, "")
    '                FueraChange = False
    '            End With
    '        End If
    '        LimpiaDatosPrecioYDescuento()
    '    End Sub

    '    Sub LimpiaDatosPrecioYDescuento()
    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 1 Then
    '            GridACtivo = msgRelojeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        End If
    '        If GridACtivo Is msgRelojeria Then
    '            With GridACtivo
    '                FueraChange = True
    '                '        .TextMatrix(.Row, C_ColRCODMODELO) = ""
    '                .set_TextMatrix(.Row, C_ColRPORCDESCTO, "")
    '                .set_TextMatrix(.Row, C_ColRPRECIO, "")
    '                .set_TextMatrix(.Row, C_ColRPORCDESCTOTAG, "")
    '                .set_TextMatrix(.Row, C_ColRPRECIOTAG, "")
    '                '        .TextMatrix(.Row, C_ColRCODARTICULO) = ""
    '                FueraChange = False
    '            End With
    '        Else
    '            With GridACtivo
    '                FueraChange = True
    '                '        .TextMatrix(.Row, C_ColJCODSUBLINEA) = ""
    '                .set_TextMatrix(.Row, C_ColJPORCDESCTO, "")
    '                .set_TextMatrix(.Row, C_ColJPRECIO, "")
    '                .set_TextMatrix(.Row, C_ColJPORCDESCTOTAG, "")
    '                .set_TextMatrix(.Row, C_ColJPRECIOTAG, "")
    '                '        .TextMatrix(.Row, C_ColJCODARTICULO) = ""
    '                FueraChange = False
    '            End With
    '        End If
    '    End Sub

    '    Sub LimpiaDatosMarca()
    '        With msgRelojeria
    '            FueraChange = True
    '            .set_TextMatrix(.Row, C_ColRMODELO, "")
    '            .set_TextMatrix(.Row, C_ColRCODMODELO, "")
    '            '        .TextMatrix(.Row, C_ColRCODARTICULO) = ""
    '            FueraChange = False
    '        End With
    '        LimpiaDatosArticulo()
    '    End Sub

    '    ''' SE MODIFICO CONSULTA PARA GRUPO - X ARTICULO Y SE AGREGO ART X PROV
    '    ''' 20ABR2006 - MAVF
    '    Sub Nuevo()
    '        'Este procedimiento genera un nuevo registro para una venta
    '        'Se deben Limpiar todos los controles del formulario con excepcion del Control de la Llavve principal
    '        On Error GoTo Merr

    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
    '        'FueraChange = True
    '        dtpFechaInIcioJ.Value = Today
    '        dtpFechaFinJ.Value = Today
    '        dtpFechaInIcioR.Value = Today
    '        dtpFechaFinR.Value = Today
    '        dtpFechaInIcioV.Value = Today
    '        dtpFechaFinV.Value = Today
    '        sstGrupos.SelectedIndex = 0
    '        msgJoyeria.Clear()
    '        msgRelojeria.Clear()
    '        msgVarios.Clear()
    '        msgXArticulo.Clear()
    '        msgArtxProv.Clear()

    '        mintCodProveedor = 0
    '        mintcodRenglon = 0
    '        mProveedor = 0
    '        mrenProv = 0
    '        mintTotalRen = 0
    '        mblnFueraChange = False
    '        dbcProveedor.Enabled = True
    '        dbcProveedor.Text = ""
    '        txtDesctoP.Text = "0.00"
    '        chkAplicar.Enabled = True
    '        chkAplicar.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        chkBorrar.Enabled = True
    '        chkBorrar.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        chkCancelarP.Enabled = False
    '        chkCancelarP.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        lblTotArt.Text = "0"
    '        txtDetArtxProv.Text = ""
    '        txtDetArtxProv.Visible = False

    '        Encabezado()
    '        mblnNuevo = True
    '        HabilitarFechas()
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '        Exit Sub

    'Merr:
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    Sub Limpiar()
    '        'Esta función Limpia todos los controles del formulario.
    '        'No se valida si hubo cambios, ya que no es posible modificar una venta
    '        On Error GoTo Merr
    '        Nuevo()
    '        'FueraChange = True
    '        dtpFechaInIcioJ.Focus()
    '        Exit Sub
    'Merr:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    ''-------------------Relojeria
    '    Private Sub msgRelojeria_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgRelojeria.DblClick
    '        Dim Estatus As String
    '        Dim EstatusTag As String
    '        '    With msgRelojeria
    '        '        Estatus = .TextMatrix(.Row, C_ColRESTATUS)
    '        '        EstatusTag = .TextMatrix(.Row, C_ColRESTATUSTAG)
    '        '        If (Estatus = C_Aplicado And Estatus <> EstatusTag) Or (Estatus = C_Cancelado And Estatus <> EstatusTag) Then
    '        '            .TextMatrix(.Row, C_ColRESTATUS) = C_Vigente
    '        '            PonerColor (.Row)
    '        '            .SetFocus
    '        ''        ElseIf Estatus = C_Aplicado Then
    '        ''            MsgBox "No es posible modificar una Promoción Aplicada Previamente.", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '        ''            .SetFocus
    '        ''            Exit Sub
    '        '        End If
    '        '    End With
    '        msgRelojeria_KeyPressEvent(msgRelojeria, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    '    End Sub

    '    Private Sub msgRelojeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgRelojeria.Enter
    '        msgRelojeria.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '        '    msgRelojeria.Row = 1
    '        '    msgRelojeria.Col = 0
    '        Pon_Tool()
    '    End Sub

    '    Private Sub msgRelojeria_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles msgRelojeria.KeyPressEvent
    '        Dim EsNuevo As Boolean
    '        Dim Estatus As String
    '        '    Dim Estatustag As String
    '        EsNuevo = True
    '        With msgRelojeria
    '            'Si no se trata de un REgistro nuevo, no se podrá editar el Grid
    '            If Trim(.get_TextMatrix(.Row, C_ColRESNUEVO)) <> "" Then
    '                EsNuevo = CBool(.get_TextMatrix(.Row, C_ColRESNUEVO))
    '            End If
    '            FueraChange = True
    '            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then 'Para que cuando sea escape, no entre a editar el codigo,simplemente que se regrese al control anterior
    '                Select Case .Col
    '                    Case C_ColRMARCA ''-------------- SE EDITA LA MARCA ---------------------'''''
    '                        If EsNuevo = False Or mblnNuevo = False Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        '''en esta parte se validará si es el rengón, columna que le corresponde editarse
    '                        If (.Row > 1) Then
    '                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
    '                            If Trim(.get_TextMatrix(.Row - 1, C_ColRMARCA)) = "" Then
    '                                .Focus()
    '                                Exit Sub
    '                            End If
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgRelojeria, dbcRMarca, eventArgs.keyAscii)
    '                    Case C_ColRMODELO ''-------------- SE EDITA EL MODELO ---------------------'''''
    '                        If EsNuevo = False Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColRMARCA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgRelojeria, dbcRModelo, eventArgs.keyAscii)
    '                    Case C_ColRARTICULO
    '                        If EsNuevo = False Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColRMARCA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgRelojeria, txtArticuloR, eventArgs.keyAscii)
    '                    Case C_ColRPORCDESCTO ''-------------- SE EDITA EL PORCENTAJE DE DESCTO---------------------'''''
    '                        'Or Estatus = C_Aplicado
    '                        ModEstandar.gp_CampoNumerico(eventArgs.keyAscii, ".")
    '                        If Estatus = C_Cancelado Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColRMARCA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, C_ColRPRECIO))) <> 0 Then
    '                            MsgBox("ya se ha Asignado un Precio. No es Posible Asignar un Porcentaje de Descuento.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgRelojeria, txtRelojeria, eventArgs.keyAscii)
    '                        .set_TextMatrix(.Row, C_ColRPRECIO, "0.00")
    '                    Case C_ColRPRECIO ''-------------- SE EDITA LA EL PRECIO ---------------------'''''
    '                        'Or Estatus = C_Aplicado
    '                        ModEstandar.gp_CampoNumerico(eventArgs.keyAscii, ".")
    '                        If Estatus = C_Cancelado Then Exit Sub
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        If Trim(.get_TextMatrix(.Row, C_ColRMARCA)) = "" Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, C_ColRPORCDESCTO))) <> 0 Then
    '                            MsgBox("Ya se ha Asignado un Porcentaje de Descuento. No es Posible Asignar el Precio.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgRelojeria, txtRelojeria, eventArgs.keyAscii)
    '                        .set_TextMatrix(.Row, C_ColRPORCDESCTO, "0.00")
    '                End Select
    '            End If
    '        End With
    '        FueraChange = False
    '    End Sub

    '    Private Sub msgRelojeria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgRelojeria.Leave
    '        msgRelojeria.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    '    End Sub

    '    Private Sub txtJoyeria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtJoyeria.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        txtJoyeria.Visible = False
    '    End Sub

    '    Private Sub txtrelojeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtrelojeria.Enter
    '        txtRelojeria.Text = Trim(txtRelojeria.Text)
    '        If Len(txtRelojeria.Text) > 1 Then
    '            SelTextoTxt(txtRelojeria)
    '        End If
    '        Pon_Tool()
    '    End Sub

    '    Private Sub txtrelojeria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtrelojeria.KeyDown
    '        Dim KeyCode As Integer = eventArgs.KeyCode
    '        Dim Shift As Integer = eventArgs.KeyData \ &H10000
    '        'Aqui se muestran los datos del control editable, en el Grid
    '        'Se deberá formatear el Valor de Acuerdo al Tipo de Dato en uso
    '        Dim rowsiguiente As Integer
    '        Dim ColSiguiente As Integer
    '        With msgRelojeria
    '            Select Case KeyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    .Focus()
    '                    txtRelojeria.Visible = False
    '                    txtRelojeria.Text = ""
    '                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                    .Focus()
    '                Case System.Windows.Forms.Keys.Return
    '                    ValidarPromocionTecleadaRepetida()
    '                    rowsiguiente = .Row
    '                    intCodMarca = CShort(Numerico(.get_TextMatrix(.Row, C_ColRCODMARCA)))
    '                    intCodModelo = CShort(Numerico(.get_TextMatrix(.Row, C_ColRCODMODELO)))
    '                    intCodArticulo = CInt(Numerico(.get_TextMatrix(.Row, C_ColRCODARTICULO)))
    '                    If ValidarPromocionGuardadaRepetida(gCODRELOJERIA, 0, 0, 0, intCodMarca, intCodModelo, intCodArticulo, dtpFechaInIcioJ.Value, dtpFechaFinJ.Value, "G") = True And mblnNuevo = True Then
    '                        MsgBox("Existe una promoción registrada para este artículo." & vbNewLine & "No es posible duplicar promociones en un lapso de tiempo similar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                        LimpiaDatosMarca()
    '                        .set_TextMatrix(.Row, C_ColJFAMILIA, "")
    '                        .Col = C_ColJFAMILIA
    '                        .Focus()
    '                        Exit Sub
    '                    End If
    '                    'Si la Columna en que se está escribiendo es Codigo o Cantidad, Formatear el Valor par que quede numérico
    '                    If .Col = C_ColRPORCDESCTO Then
    '                        If CDbl(Numerico(txtRelojeria.Text)) > 100 Then
    '                            MsgBox("El Porcentaje de Descuento no puede ser mayor de 100.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            txtRelojeria.Focus()
    '                            Exit Sub
    '                        End If
    '                        .set_TextMatrix(.Row, .Col, VB6.Format(Numerico(txtRelojeria.Text), gstrFormatoCantidad))
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, .Col))) = 0 Then
    '                            rowsiguiente = .Row
    '                            ColSiguiente = C_ColRPRECIO
    '                        Else
    '                            rowsiguiente = .Row + 1
    '                            ColSiguiente = C_ColRMARCA
    '                        End If
    '                    ElseIf .Col = C_ColRPRECIO Then
    '                        .set_TextMatrix(.Row, .Col, VB6.Format(Numerico(txtRelojeria.Text), gstrFormatoCantidad))
    '                        rowsiguiente = .Row + 1
    '                        ColSiguiente = C_ColRMARCA
    '                    End If
    '                    FueraChange = True
    '                    txtRelojeria.Text = ""
    '                    txtRelojeria.Visible = False
    '                    dbcJFamilia.Text = ""
    '                    dbcJLinea.Text = ""
    '                    dbcJSubLinea.Text = ""
    '                    dbcJFamilia.Visible = False
    '                    dbcJLinea.Visible = False
    '                    dbcJSubLinea.Visible = False
    '                    dbcRMarca.Visible = False
    '                    dbcRModelo.Visible = False
    '                    '.Col = .Col
    '                    '.Row = .Row
    '                    '.SetFocus
    '                    If .Row = .Rows - 1 Then
    '                        .Rows = .Rows + 1
    '                        .Row = rowsiguiente
    '                        .Col = ColSiguiente
    '                        .set_TextMatrix(.Row, C_COLRTIPO, "G")
    '                    Else
    '                        .Row = rowsiguiente
    '                        .Col = ColSiguiente
    '                    End If
    '                    If .Row > 7 Then
    '                        .TopRow = .Row
    '                    End If
    '                    .Focus()
    '            End Select
    '        End With
    '    End Sub

    '    Private Sub txtrelojeria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtrelojeria.KeyPress
    '        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
    '        'En este Evento se validan los datos que se introduzcan al control txtrelojeria,dependiendo de la columan en que se esté editando
    '        If KeyAscii = 0 Or KeyAscii = 13 Then GoTo EventExitSub
    '        With msgRelojeria
    '            If .Col = C_ColRPORCDESCTO Then
    '                KeyAscii = ModEstandar.MskCantidad(txtRelojeria.Text, KeyAscii, 3, 2, (txtRelojeria.SelectionStart))
    '            End If
    '            If .Col = C_ColRPRECIO Then
    '                KeyAscii = ModEstandar.MskCantidad(txtRelojeria.Text, KeyAscii, 10, 2, (txtRelojeria.SelectionStart))
    '            End If
    '        End With
    'EventExitSub:
    '        eventArgs.KeyChar = Chr(KeyAscii)
    '        If KeyAscii = 0 Then
    '            eventArgs.Handled = True
    '        End If
    '    End Sub

    '    ''' SE MODIFICO CONSULTA PARA GRUPO - X ARTICULO Y SE AGREGO ART X PROV
    '    ''' 20ABR2006 - MAVF
    '    Function ValidaDatos() As Boolean

    '        ValidaDatos = True
    '        If mblnNuevo = False Then Exit Function
    '        ValidaDatos = False
    '        'Do While (Timer() - sglTiempoCambio) <= 2.1
    '        'Loop
    '        System.Windows.Forms.Application.DoEvents()
    '        dtpFechaInIcioR.Value = dtpFechaInIcioJ.Value
    '        dtpFechaInIcioV.Value = dtpFechaInIcioJ.Value
    '        dtpFechaFinR.Value = dtpFechaFinJ.Value
    '        dtpFechaFinV.Value = dtpFechaFinJ.Value

    '        If CDate(dtpFechaFinJ.Value) < CDate(dtpFechaInIcioJ.Value) Then
    '            MsgBox("La fecha final debe ser mayor que la inicial. " & vbNewLine & "Verifique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '            Me.dtpFechaFinJ.Focus()
    '            Exit Function
    '        End If
    '        If CDate(dtpFechaFinR.Value) < CDate(dtpFechaInIcioR.Value) Then
    '            MsgBox("La fecha final debe ser mayor que la inicial. " & vbNewLine & "Verifique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '            Me.dtpFechaFinR.Focus()
    '            Exit Function
    '        End If
    '        If CDate(dtpFechaFinV.Value) < CDate(dtpFechaInIcioV.Value) Then
    '            MsgBox("La fecha final debe ser mayor que la inicial. " & vbNewLine & "Verifique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '            Me.dtpFechaFinV.Focus()
    '            Exit Function
    '        End If
    '        If CDate(dtpFechaInIcioJ.Value) < CDate(Today) Then
    '            MsgBox("La fecha inicial debe ser mayor que la actual. " & vbNewLine & "Verifique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '            Me.dtpFechaInIcioJ.Focus()
    '            Exit Function
    '        End If
    '        If CDate(dtpFechaInIcioR.Value) < CDate(Today) Then
    '            MsgBox("La fecha inicial debe ser mayor que la actual. " & vbNewLine & "Verifique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '            Me.dtpFechaInIcioR.Focus()
    '            Exit Function
    '        End If
    '        If CDate(dtpFechaInIcioV.Value) < CDate(Today) Then
    '            MsgBox("La fecha inicial debe ser mayor que la actual. " & vbNewLine & "Verifique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '            Me.dtpFechaInIcioV.Focus()
    '            Exit Function
    '        End If

    '        With msgJoyeria
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_ColJFAMILIA)) = "" Then Exit For
    '                If CDbl(Numerico(.get_TextMatrix(I, C_ColJPRECIO))) = 0 And CDbl(Numerico(.get_TextMatrix(I, C_ColJPORCDESCTO))) = 0 Then
    '                    MsgBox("Proporcione el importe de promoción o descto de la promoción...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    .Row = I
    '                    .Col = C_ColJPORCDESCTO
    '                    .Focus()
    '                    Exit Function
    '                End If
    '            Next
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_ColJFAMILIA)) = "" Then Exit For
    '                If ValidarPromocionGuardadaRepetida(gCODJOYERIA, CShort(Numerico(.get_TextMatrix(I, C_ColJCODFAMILIA))), CShort(Numerico(.get_TextMatrix(I, C_ColJCODLINEA))), CShort(Numerico(.get_TextMatrix(I, C_ColJCODSUBLINEA))), 0, 0, CInt(Numerico(.get_TextMatrix(I, C_ColJCODARTICULO))), dtpFechaInIcioJ.Value, dtpFechaFinJ.Value, "G") = True And mblnNuevo = True Then
    '                    MsgBox("Existe una promoción registrada para este artículo." & vbNewLine & "No es posible duplicar promociones en un lapso de tiempo similar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    '.RemoveItem I
    '                    .Col = 0
    '                    .Row = I
    '                    sstGrupos.SelectedIndex = 0
    '                    .Focus()
    '                    Exit Function
    '                End If
    '            Next
    '        End With
    '        With msgRelojeria
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_ColRMARCA)) = "" Then Exit For
    '                If CDbl(Numerico(.get_TextMatrix(I, C_ColRPRECIO))) = 0 And CDbl(Numerico(.get_TextMatrix(I, C_ColRPORCDESCTO))) = 0 Then
    '                    MsgBox("Proporcione el importe de promoción. ", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    .Col = C_ColRPORCDESCTO
    '                    .Focus()
    '                    Exit Function
    '                End If
    '            Next
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_ColRMARCA)) = "" Then Exit For
    '                If ValidarPromocionGuardadaRepetida(gCODRELOJERIA, 0, 0, 0, CShort(Numerico(.get_TextMatrix(I, C_ColRCODMARCA))), CShort(Numerico(.get_TextMatrix(I, C_ColRCODMODELO))), CInt(Numerico(.get_TextMatrix(I, C_ColRCODARTICULO))), dtpFechaInIcioJ.Value, dtpFechaFinJ.Value, "G") = True And mblnNuevo = True Then
    '                    MsgBox("Existe una promoción registrada para este artículo." & vbNewLine & "No es posible duplicar promociones en un lapso de tiempo similar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    '.RemoveItem I
    '                    .Col = 0
    '                    .Row = I
    '                    sstGrupos.SelectedIndex = 1
    '                    .Focus()
    '                    Exit Function
    '                End If
    '            Next
    '        End With
    '        With msgVarios
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_ColJFAMILIA)) = "" Then Exit For
    '                If CDbl(Numerico(.get_TextMatrix(I, C_ColJPRECIO))) = 0 And CDbl(Numerico(.get_TextMatrix(I, C_ColJPORCDESCTO))) = 0 Then
    '                    MsgBox("Proporcione el importe de promoción. ", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    .Col = C_ColJPORCDESCTO
    '                    Exit Function
    '                End If
    '            Next
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_ColJFAMILIA)) = "" Then Exit For
    '                If ValidarPromocionGuardadaRepetida(gCODVARIOS, CShort(Numerico(.get_TextMatrix(I, C_ColJCODFAMILIA))), CShort(Numerico(.get_TextMatrix(I, C_ColJCODLINEA))), CShort(Numerico(.get_TextMatrix(I, C_ColJCODSUBLINEA))), 0, 0, CInt(Numerico(.get_TextMatrix(I, C_ColJCODARTICULO))), dtpFechaInIcioJ.Value, dtpFechaFinJ.Value, "G") = True And mblnNuevo = True Then
    '                    MsgBox("Existe una promoción registrada para este artículo." & vbNewLine & "No es posible duplicar promociones en un lapso de tiempo similar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    '.RemoveItem I
    '                    .Col = 0
    '                    .Row = I
    '                    sstGrupos.SelectedIndex = 2
    '                    .Focus()
    '                    Exit Function
    '                End If
    '            Next
    '        End With

    '        With msgXArticulo
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_COLXARTCODARTICULO)) = "" Then Exit For
    '                If CDec(Numerico(.get_TextMatrix(I, C_COLXARTPRECIO))) = 0 And CDec(Numerico(.get_TextMatrix(I, C_COLXARTPORCDESCTO))) = 0 Then
    '                    MsgBox("Proporcione el Importe de la promocion.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    .Col = C_COLXARTPORCDESCTO
    '                    Exit Function
    '                End If
    '            Next
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_COLXARTCODARTICULO)) = "" Then Exit For
    '                If ValidarPromocionGuardadaRepetida(CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODGRUPO))), CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODFAMILIA))), CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODLINEA))), CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODSUBLINEA))), CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODMARCA))), CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODMODELO))), CInt(Numerico(.get_TextMatrix(I, C_COLXARTCODARTICULO))), dtpFechaInIcioJ.Value, dtpFechaFinJ.Value, "A") = True And mblnNuevo = True Then
    '                    MsgBox("Existe una promoción registrada para este artículo." & vbNewLine & "No es posible duplicar promociones en un lapso de tiempo similar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    .Col = 0
    '                    .Row = I
    '                    sstGrupos.SelectedIndex = 3
    '                    .Focus()
    '                    Exit Function
    '                End If
    '            Next
    '        End With

    '        '''la programacion de promociones incluye articulos x proveedor
    '        If CShort(ModEstandar.Numerico((lblTotArt.Text))) > 0 Then
    '            If (Trim(dbcProveedor.Text) = "" Or mintCodProveedor = 0) Then
    '                MsgBox("Proporcione el nombre del proveedor", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
    '                dbcProveedor.Focus()
    '                Exit Function
    '            End If
    '        End If
    '        With msgArtxProv
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_COLXPRVCODARTICULO)) = "" Then Exit For
    '                If CDec(Numerico(.get_TextMatrix(I, C_COLXPRVPORCDESCTO))) = 0 Then
    '                    MsgBox("Proporcione el descuento de la promoción", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    .Row = I
    '                    .Col = C_COLXPRVPORCDESCTO
    '                    .Focus()
    '                    Exit Function
    '                End If
    '            Next
    '            If I < CInt(ModEstandar.Numerico((lblTotArt.Text))) Then
    '                MsgBox("No deben existir espacios en blanco en la captura...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
    '            End If
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_COLXPRVCODARTICULO)) = "" Then Exit For
    '                If ValidarPromocionGuardadaRepetida(CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODGRUPO))), CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODFAMILIA))), CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODLINEA))), CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODSUBLINEA))), CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODMARCA))), CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODMODELO))), CInt(Numerico(.get_TextMatrix(I, C_COLXPRVCODARTICULO))), dtpFechaInIcioJ.Value, dtpFechaFinJ.Value, "A") = True And mblnNuevo = True Then
    '                    MsgBox("Existe una promoción registrada para este artículo." & vbNewLine & "No es posible duplicar promociones en un lapso de tiempo similar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    .Col = 0
    '                    .Row = I
    '                    sstGrupos.SelectedIndex = 4
    '                    .Focus()
    '                    Exit Function
    '                End If
    '            Next
    '        End With
    '        ValidaDatos = True

    '    End Function

    '    ''' SE MODIFICO CONSULTA PARA GRUPO - X ARTICULO Y SE AGREGO ART X PROV
    '    ''' 20ABR2006 - MAVF
    '    Sub Guardar()
    '        On Error GoTo Merr
    '        Dim CodGrupo As Integer
    '        Dim CodFamilia As Integer
    '        Dim COdLinea As Integer
    '        Dim CodSubLinea As Integer
    '        Dim CodMArca As Integer
    '        Dim CodModelo As Integer
    '        Dim CodArticulo As Integer
    '        Dim importe As Decimal
    '        Dim Porcentaje As Decimal
    '        Dim ImporteTag As Decimal
    '        Dim PorcentajeTag As Decimal
    '        Dim FechaInicio As Date
    '        Dim FechaFin As Date
    '        Dim Estatus As String
    '        Dim EstatusTag As String
    '        Dim EsNuevo As Boolean
    '        Dim blnTransaccion As Boolean
    '        Dim lProveedor As Integer
    '        Dim lRenglon As Integer

    '        gstrProcesoqueGeneraError = "FrmProgramacionPromociones (Guardar)"
    '        If ValidaDatos() = False Then Exit Sub
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

    '        Cnn.BeginTrans()
    '        blnTransaccion = True
    '        With msgJoyeria
    '            CodGrupo = gCODJOYERIA
    '            FechaInicio = dtpFechaInIcioJ.Value
    '            FechaFin = dtpFechaFinJ.Value
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_ColJFAMILIA)) = "" Then Exit For
    '                CodFamilia = CShort(Numerico(.get_TextMatrix(I, C_ColJCODFAMILIA)))
    '                COdLinea = CShort(Numerico(.get_TextMatrix(I, C_ColJCODLINEA)))
    '                CodSubLinea = CShort(Numerico(.get_TextMatrix(I, C_ColJCODSUBLINEA)))
    '                CodArticulo = CInt(Numerico(.get_TextMatrix(I, C_ColJCODARTICULO)))
    '                CodMArca = 0
    '                CodModelo = 0
    '                importe = CDec(Numerico(.get_TextMatrix(I, C_ColJPRECIO)))
    '                Porcentaje = CDec(Numerico(.get_TextMatrix(I, C_ColJPORCDESCTO)))
    '                ImporteTag = CDec(Numerico(.get_TextMatrix(I, C_ColJPRECIOTAG)))
    '                PorcentajeTag = CDec(Numerico(.get_TextMatrix(I, C_ColJPORCDESCTOTAG)))
    '                Estatus = .get_TextMatrix(I, C_ColJESTATUS)
    '                EstatusTag = .get_TextMatrix(I, C_ColJESTATUSTAG)
    '                EsNuevo = CBool(.get_TextMatrix(I, C_ColJESNUEVO))
    '                lProveedor = 0
    '                lRenglon = 0
    '                'Verificar si se TRata de un Registro Nuevo o no
    '                If EsNuevo = True And Estatus <> "C" Then
    '                    ModStoredProcedures.PR_IMEPromocionesVentas(CStr(CodGrupo), CStr(CodFamilia), CStr(COdLinea), CStr(CodSubLinea), CStr(CodMArca), CStr(CodModelo), CStr(CodArticulo), CStr(importe), CStr(Porcentaje), VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR), VB6.Format(FechaFin, C_FORMATFECHAGUARDAR), Estatus, "01/01/1900", .get_TextMatrix(I, C_COLJTIPO), Trim(CStr(lProveedor)), Trim(CStr(lRenglon)), C_INSERCION, CStr(0))
    '                    Cmd.Execute()
    '                Else
    '                    If importe <> ImporteTag Or Porcentaje <> PorcentajeTag Or Estatus <> EstatusTag Then
    '                        ModStoredProcedures.PR_IMEPromocionesVentas(CStr(CodGrupo), CStr(CodFamilia), CStr(COdLinea), CStr(CodSubLinea), CStr(CodMArca), CStr(CodModelo), CStr(CodArticulo), CStr(importe), CStr(Porcentaje), VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR), VB6.Format(FechaFin, C_FORMATFECHAGUARDAR), Estatus, "01/01/1900", .get_TextMatrix(I, C_COLJTIPO), Trim(CStr(lProveedor)), Trim(CStr(lRenglon)), C_MODIFICACION, CStr(0))
    '                        Cmd.Execute()
    '                    End If
    '                End If
    '            Next
    '        End With
    '        With msgRelojeria
    '            CodGrupo = gCODRELOJERIA
    '            FechaInicio = dtpFechaInIcioR.Value
    '            FechaFin = dtpFechaFinR.Value
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_ColRMARCA)) = "" Then Exit For
    '                CodFamilia = 0
    '                COdLinea = 0
    '                CodSubLinea = 0
    '                CodMArca = CShort(Numerico(.get_TextMatrix(I, C_ColRCODMARCA)))
    '                CodModelo = CShort(Numerico(.get_TextMatrix(I, C_ColRCODMODELO)))
    '                CodArticulo = CInt(Numerico(.get_TextMatrix(I, C_ColRCODARTICULO)))
    '                importe = CDec(Numerico(.get_TextMatrix(I, C_ColRPRECIO)))
    '                Porcentaje = CDec(Numerico(.get_TextMatrix(I, C_ColRPORCDESCTO)))
    '                ImporteTag = CDec(Numerico(.get_TextMatrix(I, C_ColRPRECIOTAG)))
    '                PorcentajeTag = CDec(Numerico(.get_TextMatrix(I, C_ColRPORCDESCTOTAG)))
    '                Estatus = .get_TextMatrix(I, C_ColRESTATUS)
    '                EstatusTag = .get_TextMatrix(I, C_ColRESTATUSTAG)
    '                EsNuevo = CBool(.get_TextMatrix(I, C_ColRESNUEVO))
    '                lProveedor = 0
    '                lRenglon = 0
    '                'Verificar si se TRata de un Registro Nuevo o no
    '                If EsNuevo = True And Estatus <> "C" Then
    '                    ModStoredProcedures.PR_IMEPromocionesVentas(CStr(CodGrupo), CStr(CodFamilia), CStr(COdLinea), CStr(CodSubLinea), CStr(CodMArca), CStr(CodModelo), CStr(CodArticulo), CStr(importe), CStr(Porcentaje), VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR), VB6.Format(FechaFin, C_FORMATFECHAGUARDAR), Estatus, "01/01/1900", .get_TextMatrix(I, C_COLRTIPO), Trim(CStr(lProveedor)), Trim(CStr(lRenglon)), C_INSERCION, CStr(0))
    '                    Cmd.Execute()
    '                Else
    '                    If importe <> ImporteTag Or Porcentaje <> PorcentajeTag Or Estatus <> EstatusTag Then
    '                        ModStoredProcedures.PR_IMEPromocionesVentas(CStr(CodGrupo), CStr(CodFamilia), CStr(COdLinea), CStr(CodSubLinea), CStr(CodMArca), CStr(CodModelo), CStr(CodArticulo), CStr(importe), CStr(Porcentaje), VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR), VB6.Format(FechaFin, C_FORMATFECHAGUARDAR), Estatus, "01/01/1900", .get_TextMatrix(I, C_COLRTIPO), Trim(CStr(lProveedor)), Trim(CStr(lRenglon)), C_MODIFICACION, CStr(0))
    '                        Cmd.Execute()
    '                    End If
    '                End If
    '            Next
    '        End With
    '        With msgVarios
    '            CodGrupo = gCODVARIOS
    '            FechaInicio = dtpFechaInIcioV.Value
    '            FechaFin = dtpFechaFinV.Value
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_ColJCODFAMILIA)) = "" Then Exit For
    '                CodFamilia = CShort(Numerico(.get_TextMatrix(I, C_ColJCODFAMILIA)))
    '                COdLinea = CShort(Numerico(.get_TextMatrix(I, C_ColJCODLINEA)))
    '                CodArticulo = CInt(Numerico(.get_TextMatrix(I, C_ColJCODARTICULO)))
    '                CodSubLinea = 0
    '                CodMArca = 0
    '                CodModelo = 0
    '                importe = CDec(Numerico(.get_TextMatrix(I, C_ColJPRECIO)))
    '                Porcentaje = CDec(Numerico(.get_TextMatrix(I, C_ColJPORCDESCTO)))
    '                ImporteTag = CDec(Numerico(.get_TextMatrix(I, C_ColJPRECIOTAG)))
    '                PorcentajeTag = CDec(Numerico(.get_TextMatrix(I, C_ColJPORCDESCTOTAG)))
    '                Estatus = .get_TextMatrix(I, C_ColJESTATUS)
    '                EstatusTag = .get_TextMatrix(I, C_ColJESTATUSTAG)
    '                EsNuevo = CBool(.get_TextMatrix(I, C_ColJESNUEVO))
    '                lProveedor = 0
    '                lRenglon = 0
    '                'Verificar si se TRata de un Registro Nuevo o no
    '                If EsNuevo = True And Estatus <> "C" Then
    '                    ModStoredProcedures.PR_IMEPromocionesVentas(CStr(CodGrupo), CStr(CodFamilia), CStr(COdLinea), CStr(CodSubLinea), CStr(CodMArca), CStr(CodModelo), CStr(CodArticulo), CStr(importe), CStr(Porcentaje), VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR), VB6.Format(FechaFin, C_FORMATFECHAGUARDAR), Estatus, "01/01/1900", .get_TextMatrix(I, C_COLJTIPO), Trim(CStr(lProveedor)), Trim(CStr(lRenglon)), C_INSERCION, CStr(0))
    '                    Cmd.Execute()
    '                Else
    '                    If importe <> ImporteTag Or Porcentaje <> PorcentajeTag Or Estatus <> EstatusTag Then
    '                        ModStoredProcedures.PR_IMEPromocionesVentas(CStr(CodGrupo), CStr(CodFamilia), CStr(COdLinea), CStr(CodSubLinea), CStr(CodMArca), CStr(CodModelo), CStr(CodArticulo), CStr(importe), CStr(Porcentaje), VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR), VB6.Format(FechaFin, C_FORMATFECHAGUARDAR), Estatus, "01/01/1900", .get_TextMatrix(I, C_COLJTIPO), Trim(CStr(lProveedor)), Trim(CStr(lRenglon)), C_MODIFICACION, CStr(0))
    '                        Cmd.Execute()
    '                    End If
    '                End If
    '            Next
    '        End With
    '        With msgXArticulo
    '            FechaInicio = dtpFechaInIcioV.Value
    '            FechaFin = dtpFechaFinV.Value
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_COLXARTCODARTICULO)) = "" Then Exit For
    '                CodGrupo = CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODGRUPO)))
    '                CodFamilia = CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODFAMILIA)))
    '                COdLinea = CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODLINEA)))
    '                CodArticulo = CInt(Numerico(.get_TextMatrix(I, C_COLXARTCODARTICULO)))
    '                CodSubLinea = CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODSUBLINEA)))
    '                CodMArca = CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODMARCA)))
    '                CodModelo = CShort(Numerico(.get_TextMatrix(I, C_COLXARTCODMODELO)))
    '                importe = CDec(Numerico(.get_TextMatrix(I, C_COLXARTPRECIO)))
    '                Porcentaje = CDec(Numerico(.get_TextMatrix(I, C_COLXARTPORCDESCTO)))
    '                ImporteTag = CDec(Numerico(.get_TextMatrix(I, C_COLXARTPRECIOTAG)))
    '                PorcentajeTag = CDec(Numerico(.get_TextMatrix(I, C_COLXARTPORCDESCTOTAG)))
    '                Estatus = .get_TextMatrix(I, C_COLXARTESTATUS)
    '                EstatusTag = .get_TextMatrix(I, C_COLXARTESTATUSTAG)
    '                EsNuevo = CBool(.get_TextMatrix(I, C_COLXARTESNUEVO))
    '                lProveedor = 0
    '                lRenglon = 0
    '                'Verificar si se TRata de un Registro Nuevo o no
    '                If EsNuevo = True And Estatus <> "C" Then
    '                    ModStoredProcedures.PR_IMEPromocionesVentas(CStr(CodGrupo), CStr(CodFamilia), CStr(COdLinea), CStr(CodSubLinea), CStr(CodMArca), CStr(CodModelo), CStr(CodArticulo), CStr(importe), CStr(Porcentaje), VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR), VB6.Format(FechaFin, C_FORMATFECHAGUARDAR), Estatus, "01/01/1900", .get_TextMatrix(I, C_COLXARTTIPO), Trim(CStr(lProveedor)), Trim(CStr(lRenglon)), C_INSERCION, CStr(0))
    '                    Cmd.Execute()
    '                Else
    '                    If importe <> ImporteTag Or Porcentaje <> PorcentajeTag Or Estatus <> EstatusTag Then
    '                        ModStoredProcedures.PR_IMEPromocionesVentas(CStr(CodGrupo), CStr(CodFamilia), CStr(COdLinea), CStr(CodSubLinea), CStr(CodMArca), CStr(CodModelo), CStr(CodArticulo), CStr(importe), CStr(Porcentaje), VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR), VB6.Format(FechaFin, C_FORMATFECHAGUARDAR), Estatus, "01/01/1900", .get_TextMatrix(I, C_COLXARTTIPO), Trim(CStr(lProveedor)), Trim(CStr(lRenglon)), C_MODIFICACION, CStr(0))
    '                        Cmd.Execute()
    '                    End If
    '                End If
    '            Next
    '        End With

    '        With msgArtxProv
    '            FechaInicio = dtpFechaInIcioV.Value
    '            FechaFin = dtpFechaFinV.Value
    '            If mblnNuevo Then
    '                mrenProv = CalculaRenglonProv(mintCodProveedor, FechaInicio, FechaFin)
    '            Else
    '                mrenProv = mintcodRenglon
    '            End If
    '            For I = 1 To .Rows - 1
    '                If Trim(.get_TextMatrix(I, C_COLXPRVCODARTICULO)) = "" Then Exit For
    '                CodGrupo = CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODGRUPO)))
    '                CodFamilia = CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODFAMILIA)))
    '                COdLinea = CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODLINEA)))
    '                CodArticulo = CInt(Numerico(.get_TextMatrix(I, C_COLXPRVCODARTICULO)))
    '                CodSubLinea = CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODSUBLINEA)))
    '                CodMArca = CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODMARCA)))
    '                CodModelo = CShort(Numerico(.get_TextMatrix(I, C_COLXPRVCODMODELO)))
    '                importe = CDec(Numerico(.get_TextMatrix(I, C_COLXPRVPRECIO)))
    '                Porcentaje = CDec(Numerico(.get_TextMatrix(I, C_COLXPRVPORCDESCTO)))
    '                ImporteTag = CDec(Numerico(.get_TextMatrix(I, C_COLXPRVPRECIOTAG)))
    '                PorcentajeTag = CDec(Numerico(.get_TextMatrix(I, C_COLXPRVPORCDESCTOTAG)))
    '                Estatus = Trim(.get_TextMatrix(I, C_COLXPRVESTATUS))
    '                EstatusTag = Trim(.get_TextMatrix(I, C_COLXPRVESTATUSTAG))
    '                EsNuevo = CBool(.get_TextMatrix(I, C_COLXPRVESNUEVO))
    '                lProveedor = mintCodProveedor
    '                lRenglon = mrenProv
    '                'Verificar si se Trata de un Registro Nuevo o no
    '                If EsNuevo = True And Estatus <> "C" Then
    '                    ModStoredProcedures.PR_IMEPromocionesVentas(CStr(CodGrupo), CStr(CodFamilia), CStr(COdLinea), CStr(CodSubLinea), CStr(CodMArca), CStr(CodModelo), CStr(CodArticulo), CStr(importe), CStr(Porcentaje), VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR), VB6.Format(FechaFin, C_FORMATFECHAGUARDAR), Estatus, "01/01/1900", .get_TextMatrix(I, C_COLXPRVTIPO), Trim(CStr(lProveedor)), Trim(CStr(lRenglon)), C_INSERCION, CStr(0))
    '                    Cmd.Execute()
    '                Else
    '                    If Porcentaje <> PorcentajeTag Or Estatus <> EstatusTag Then
    '                        ModStoredProcedures.PR_IMEPromocionesVentas(CStr(CodGrupo), CStr(CodFamilia), CStr(COdLinea), CStr(CodSubLinea), CStr(CodMArca), CStr(CodModelo), CStr(CodArticulo), CStr(importe), CStr(Porcentaje), VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR), VB6.Format(FechaFin, C_FORMATFECHAGUARDAR), Estatus, "01/01/1900", .get_TextMatrix(I, C_COLXPRVTIPO), Trim(CStr(lProveedor)), Trim(CStr(lRenglon)), C_MODIFICACION, CStr(0))
    '                        Cmd.Execute()
    '                    End If
    '                End If
    '            Next
    '        End With
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '        Cnn.CommitTrans()
    '        blnTransaccion = False
    '        MsgBox("Las promociones se ha guardado con éxito.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
    '        Limpiar()
    '        Exit Sub

    'Merr:
    '        If Err.Number <> 0 Then
    '            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '            If blnTransaccion = True Then Cnn.RollbackTrans()
    '            ModEstandar.MostrarError("Ocurrió un error en el formulario y proceso: " & gstrProcesoqueGeneraError)
    '        End If
    '    End Sub

    '    ''' SE MODIFICO CONSULTA PARA GRUPO - X ARTICULO Y SE AGREGO ART X PROV
    '    ''' 20ABR2006 - MAVF
    '    Sub LlenaDatos(ByRef Grid As Integer)
    '        On Error GoTo Merr
    '        Dim CodGrupo As Integer
    '        Dim CodFamilia As Integer
    '        Dim COdLinea As Integer
    '        Dim CodSubLinea As Integer
    '        Dim CodMArca As Integer
    '        Dim CodModelo As Integer
    '        Dim importe As Decimal
    '        Dim Porcentaje As Decimal
    '        Dim FechaInicio As Date
    '        Dim FechaFin As Date
    '        Dim Estatus As String
    '        Dim J As Integer
    '        Dim lCont As Integer

    '        Select Case Grid

    '            Case 4 ''' ARTICULO X PROVEEDOR

    '                msgArtxProv.Clear()
    '                Encabezado()

    '                gStrSql = "SELECT P.CodGrupo, ISNULL(P.CodFamilia, 0) as CodFamilia, ISNULL(F.DescFamilia,'') as DescFamilia, ISNULL(P.CodLinea, 0) AS CodLinea, ISNULL(L.DescLinea, '') AS DescLinea, ISNULL(P.CodSubLinea, 0) AS CodSubLinea, ISNULL(S.DescSubLinea, '') AS DescSubLinea, P.Importe, P.Porcentaje, " & "ISNULL(P.CodMarca,0) AS CodMarca, ISNULL(M.DescMarca,0) AS DescMarca, ISNULL(P.CodModelo,0) AS CodModelo, ISNULL(O.DescModelo,0) AS DescModelo, P.FechaInicio, P.FechaFin, P.Estatus, ISNULL(A.CodArticulo, 0) AS CodArticulo, ISNULL(A.DescArticulo, '') AS DescArticulo, P.TipoProm, CASE A.CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),A.OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),A.CodigoAnt))) ,5) End as CodigoAnt, IsNull(P.codProvAcreed, 0) as CodProvAcreed, IsNull(V.DescProvAcreed, '-') as DescProvAcreed, Renglon " & "FROM dbo.PromocionesVentas P LEFT OUTER JOIN dbo.CatFamilias  F ON P.CodFamilia = F.CodFamilia AND P.CodGrupo = F.CodGrupo LEFT OUTER JOIN dbo.CatLineas    L ON P.CodLinea = L.CodLinea AND P.CodFamilia = L.CodFamilia AND P.CodGrupo = L.CodGrupo LEFT OUTER JOIN dbo.CatSubLineas S ON P.CodSubLinea = S.CodSubLinea AND P.CodFamilia = S.CodFamilia AND P.CodLinea = S.CodLinea " & "LEFT OUTER JOIN dbo.CatMarcas M ON P.CodGrupo = M.CodGrupo And P.CodMarca = M.CodMarca LEFT OUTER JOIN dbo.CatModelos   O ON P.CodGrupo = O.CodGrupo And P.CodMarca = O.CodMarca And P.CodModelo = O.CodModelo LEFT OUTER JOIN dbo.CatProvAcreed  V ON P.CodProvAcreed = V.CodProvAcreed Inner JOIN dbo.CatArticulos A ON P.CodArticulo = A.CodArticulo " & "WHERE (P.FechaInicio = '" & VB6.Format(dtpFechaInIcioJ.Value, C_FORMATFECHAGUARDAR) & "') AND (P.FechaFin = '" & VB6.Format(dtpFechaFinJ.Value, C_FORMATFECHAGUARDAR) & "') and Estatus <> 'C' AND P.TipoProm = 'A' AND P.CodProvAcreed = " & mProveedor & " And Renglon = " & mrenProv & " " & "GROUP BY P.CodGrupo, P.CodFamilia, F.DescFamilia, P.CodLinea, L.DescLinea, S.DescSubLinea, P.CodSubLinea, P.Importe,  P.Porcentaje, P.FechaInicio, P.CodModelo, P.CodMarca, A.CodigoAnt, A.OrigenAnt, P.FechaFin, P.Estatus, A.CodArticulo , A.DescArticulo, P.TipoProm, M.DescMarca, O.DescModelo, P.codProvAcreed, V.DescProvAcreed, Renglon " & "Order    BY A.CodArticulo "
    '                ModEstandar.BorraCmd()
    '                Cmd.CommandText = "dbo.Up_Select_Datos"
    '                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '                RsGral = Cmd.Execute
    '                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    '                If RsGral.RecordCount > 0 Then
    '                    If RsGral.RecordCount <= 6 Then
    '                        msgArtxProv.Rows = RsGral.RecordCount + 12
    '                    Else
    '                        msgArtxProv.Rows = RsGral.RecordCount + 3
    '                    End If
    '                    lCont = 0
    '                    dbcProveedor.Enabled = False
    '                    chkBorrar.Enabled = False
    '                    mblnFueraChange = True
    '                    dbcProveedor.Text = Trim(RsGral.Fields("DescProvACreed").Value)
    '                    chkCancelarP.Enabled = True
    '                    chkCancelarP.CheckState = System.Windows.Forms.CheckState.Unchecked
    '                    mblnFueraChange = False
    '                    mintCodProveedor = RsGral.Fields("CodProvAcreed").Value
    '                    mintcodRenglon = RsGral.Fields("Renglon").Value

    '                    With msgArtxProv
    '                        sstGrupos.SelectedIndex = 4
    '                        If RsGral.RecordCount > .Rows - 1 Then .Rows = RsGral.RecordCount + 1
    '                        For I = 1 To .Rows - 1
    '                            If RsGral.EOF Then Exit For
    '                            lCont = lCont + 1
    '                            .set_TextMatrix(I, C_COLXPRVCODGRUPO, RsGral.Fields("CodGrupo").Value)
    '                            .set_TextMatrix(I, C_COLXPRVCODFAMILIA, Trim(RsGral.Fields("CodFamilia").Value))
    '                            .set_TextMatrix(I, C_COLXPRVCODLINEA, Trim(RsGral.Fields("COdLinea").Value))
    '                            .set_TextMatrix(I, C_COLXPRVCODSUBLINEA, Trim(RsGral.Fields("CodSubLinea").Value))
    '                            .set_TextMatrix(I, C_COLXPRVCODMODELO, RsGral.Fields("CodModelo").Value)
    '                            .set_TextMatrix(I, C_COLXPRVCODMARCA, RsGral.Fields("CodMArca").Value)
    '                            .set_TextMatrix(I, C_COLXPRVDESCARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
    '                            .set_TextMatrix(I, C_COLXPRVCODARTICULO, Trim(RsGral.Fields("CodArticulo").Value))
    '                            '''.TextMatrix(I, C_COLXPRVPRECIO) = Format(Trim(RsGral!importe), gstrFormatoCantidad)
    '                            .set_TextMatrix(I, C_COLXPRVPORCDESCTO, VB6.Format(Trim(RsGral.Fields("Porcentaje").Value), "0.00"))
    '                            .set_TextMatrix(I, C_COLXPRVPRECIOTAG, VB6.Format(Trim(RsGral.Fields("importe").Value), gstrFormatoCantidad))
    '                            .set_TextMatrix(I, C_COLXPRVPORCDESCTOTAG, VB6.Format(Trim(RsGral.Fields("Porcentaje").Value), "0.00"))
    '                            .set_TextMatrix(I, C_COLXPRVESTATUS, Trim(RsGral.Fields("Estatus").Value))
    '                            .set_TextMatrix(I, C_COLXPRVESTATUSTAG, Trim(RsGral.Fields("Estatus").Value))
    '                            .set_TextMatrix(I, C_COLXPRVESNUEVO, False)
    '                            .set_TextMatrix(I, C_COLXPRVTIPO, "A")
    '                            .set_TextMatrix(I, C_COLXPRVCODANTERIOR, RsGral.Fields("CodigoAnt").Value)
    '                            RsGral.MoveNext()
    '                            PonerColor((I))
    '                            .set_ColAlignment(C_COLXPRVDESCARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                        Next
    '                        lblTotArt.Text = VB6.Format(lCont, "###,##0")
    '                        .Col = C_COLXPRVCODARTICULO
    '                        .Row = 1
    '                        .TopRow = 1
    '                    End With
    '                    sstGrupos.SelectedIndex = 4
    '                    msgArtxProv.Focus()
    '                End If

    '            Case 3 '''X ARTICULO
    '                gStrSql = "SELECT P.CodGrupo, ISNULL(P.CodFamilia, 0) as CodFamilia, ISNULL(F.DescFamilia,'') as DescFamilia, ISNULL(P.CodLinea, 0) AS CodLinea, ISNULL(L.DescLinea, '') AS DescLinea, ISNULL(P.CodSubLinea, 0) AS CodSubLinea, ISNULL(S.DescSubLinea, '') AS DescSubLinea, P.Importe, P.Porcentaje, " & "ISNULL(P.CodMarca,0) AS CodMarca, ISNULL(M.DescMarca,0) AS DescMarca, ISNULL(P.CodModelo,0) AS CodModelo, ISNULL(O.DescModelo,0) AS DescModelo, P.FechaInicio, P.FechaFin, P.Estatus, ISNULL(A.CodArticulo, 0) AS CodArticulo, ISNULL(A.DescArticulo, '') AS DescArticulo, P.TipoProm, CASE A.CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),A.OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),A.CodigoAnt))) ,5) End as CodigoAnt, IsNull(P.codProvAcreed, 0) as CodProvAcreed, IsNull(V.DescProvAcreed, '-') as DescProvAcreed " & "FROM dbo.PromocionesVentas P LEFT OUTER JOIN dbo.CatFamilias  F ON P.CodFamilia = F.CodFamilia AND P.CodGrupo = F.CodGrupo LEFT OUTER JOIN dbo.CatLineas    L ON P.CodLinea = L.CodLinea AND P.CodFamilia = L.CodFamilia AND P.CodGrupo = L.CodGrupo LEFT OUTER JOIN dbo.CatSubLineas S ON P.CodSubLinea = S.CodSubLinea AND P.CodFamilia = S.CodFamilia AND P.CodLinea = S.CodLinea " & "LEFT OUTER JOIN dbo.CatMarcas M ON P.CodGrupo = M.CodGrupo And P.CodMarca = M.CodMarca LEFT OUTER JOIN dbo.CatModelos   O ON P.CodGrupo = O.CodGrupo And P.CodMarca = O.CodMarca And P.CodModelo = O.CodModelo LEFT OUTER JOIN dbo.CatProvAcreed  V ON P.CodProvAcreed = V.CodProvAcreed Inner JOIN dbo.CatArticulos A ON P.CodArticulo = A.CodArticulo " & "WHERE (P.FechaInicio = '" & VB6.Format(dtpFechaInIcioJ.Value, C_FORMATFECHAGUARDAR) & "') AND (P.FechaFin = '" & VB6.Format(dtpFechaFinJ.Value, C_FORMATFECHAGUARDAR) & "') and Estatus <> 'C' AND P.TipoProm = 'A' AND P.CodProvAcreed Is Null " & "GROUP BY P.CodGrupo, P.CodFamilia, F.DescFamilia, P.CodLinea, L.DescLinea, S.DescSubLinea, P.CodSubLinea, P.Importe,  P.Porcentaje, P.FechaInicio, P.CodModelo, P.CodMarca, A.CodigoAnt, A.OrigenAnt, P.FechaFin, P.Estatus, A.CodArticulo , A.DescArticulo, P.TipoProm, M.DescMarca, O.DescModelo, P.codProvAcreed, V.DescProvAcreed "
    '                ModEstandar.BorraCmd()
    '                Cmd.CommandText = "dbo.Up_Select_Datos"
    '                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '                RsGral = Cmd.Execute
    '                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '                If RsGral.RecordCount > 0 Then
    '                    With msgXArticulo
    '                        sstGrupos.SelectedIndex = 3
    '                        If RsGral.RecordCount > .Rows - 1 Then .Rows = RsGral.RecordCount + 1
    '                        For I = 1 To .Rows - 1
    '                            If RsGral.EOF Then Exit For
    '                            .set_TextMatrix(I, C_COLXARTCODGRUPO, RsGral.Fields("CodGrupo").Value)
    '                            .set_TextMatrix(I, C_COLXARTCODFAMILIA, Trim(RsGral.Fields("CodFamilia").Value))
    '                            .set_TextMatrix(I, C_COLXARTCODLINEA, Trim(RsGral.Fields("COdLinea").Value))
    '                            .set_TextMatrix(I, C_COLXARTCODSUBLINEA, Trim(RsGral.Fields("CodSubLinea").Value))
    '                            .set_TextMatrix(I, C_COLXARTCODMODELO, RsGral.Fields("CodModelo").Value)
    '                            .set_TextMatrix(I, C_COLXARTCODMARCA, RsGral.Fields("CodMArca").Value)
    '                            .set_TextMatrix(I, C_COLXARTDESCARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
    '                            .set_TextMatrix(I, C_COLXARTCODARTICULO, Trim(RsGral.Fields("CodArticulo").Value))
    '                            .set_TextMatrix(I, C_COLXARTPRECIO, VB6.Format(Trim(RsGral.Fields("importe").Value), gstrFormatoCantidad))
    '                            .set_TextMatrix(I, C_COLXARTPORCDESCTO, VB6.Format(Trim(RsGral.Fields("Porcentaje").Value), "0.00"))
    '                            .set_TextMatrix(I, C_COLXARTPRECIOTAG, VB6.Format(Trim(RsGral.Fields("importe").Value), gstrFormatoCantidad))
    '                            .set_TextMatrix(I, C_COLXARTPORCDESCTOTAG, VB6.Format(Trim(RsGral.Fields("Porcentaje").Value), "0.00"))
    '                            .set_TextMatrix(I, C_COLXARTESTATUS, Trim(RsGral.Fields("Estatus").Value))
    '                            .set_TextMatrix(I, C_COLXARTESTATUSTAG, Trim(RsGral.Fields("Estatus").Value))
    '                            .set_TextMatrix(I, C_COLXARTESNUEVO, False)
    '                            .set_TextMatrix(I, C_COLXARTTIPO, "A")
    '                            .set_TextMatrix(I, C_COLXARTCODANTERIOR, RsGral.Fields("CodigoAnt").Value)
    '                            RsGral.MoveNext()
    '                            PonerColor((I))
    '                            .set_ColAlignment(C_COLXARTDESCARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                        Next
    '                    End With
    '                    sstGrupos.SelectedIndex = 3
    '                    msgXArticulo.Focus()
    '                End If

    '            Case 2 ''' VARIOS
    '                gStrSql = "SELECT P.CodGrupo, Isnull(P.CodFamilia,0) as CodFamilia , Isnull(F.DescFamilia,'') as DescFamilia , Isnull(P.CodLinea,0) as CodLinea, Isnull(L.DescLinea,'') as DescLinea , P.Importe, P.Porcentaje, P.FechaInicio, P.FechaFin, P.Estatus " & ",ISNULL(A.CodArticulo, 0) AS CodArticulo, ISNULL(A.DescArticulo, '') AS DescArticulo,P.TipoProm " & "FROM dbo.PromocionesVentas P INNER JOIN " & "dbo.CatFamilias F ON P.CodFamilia = F.CodFamilia AND P.CodGrupo = F.CodGrupo left outer  JOIN " & "dbo.CatLineas L ON P.CodLinea = L.CodLinea AND P.CodFamilia = L.CodFamilia AND P.CodGrupo = L.CodGrupo  LEFT OUTER  JOIN " & "dbo.CatArticulos A ON  P.CodArticulo = A.CodArticulo " & "WHERE (P.FechaInicio = '" & VB6.Format(dtpFechaInIcioV.Value, C_FORMATFECHAGUARDAR) & "') AND (P.FechaFin = '" & VB6.Format(dtpFechaFinV.Value, C_FORMATFECHAGUARDAR) & " ') and Estatus <> 'C' AND P.TipoProm = 'G' " & "GROUP BY P.CodGrupo, P.CodFamilia, F.DescFamilia, P.CodLinea, L.DescLinea, P.Importe, P.Porcentaje, P.FechaInicio, P.FechaFin, P.Estatus , A.CodArticulo, A.DescArticulo,P.TipoProm " & "Having (P.CodGrupo =" & gCODVARIOS & ")"
    '                ModEstandar.BorraCmd()
    '                Cmd.CommandText = "dbo.Up_Select_Datos"
    '                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '                RsGral = Cmd.Execute
    '                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '                If RsGral.RecordCount > 0 Then
    '                    With msgVarios
    '                        sstGrupos.SelectedIndex = 2
    '                        If RsGral.RecordCount > .Rows - 1 Then .Rows = RsGral.RecordCount + 1
    '                        For I = 1 To .Rows - 1
    '                            If RsGral.EOF Then Exit For
    '                            .set_TextMatrix(I, C_ColJFAMILIA, Trim(RsGral.Fields("DescFamilia").Value))
    '                            .set_TextMatrix(I, C_ColJLINEA, Trim(RsGral.Fields("DescLinea").Value))
    '                            .set_TextMatrix(I, C_ColJARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
    '                            .set_TextMatrix(I, C_ColJCODFAMILIA, Trim(RsGral.Fields("CodFamilia").Value))
    '                            .set_TextMatrix(I, C_ColJCODLINEA, Trim(RsGral.Fields("COdLinea").Value))
    '                            .set_TextMatrix(I, C_ColJCODARTICULO, Trim(RsGral.Fields("CodArticulo").Value))
    '                            .set_TextMatrix(I, C_ColJPRECIO, VB6.Format(Trim(RsGral.Fields("importe").Value), gstrFormatoCantidad))
    '                            .set_TextMatrix(I, C_ColJPORCDESCTO, VB6.Format(Trim(RsGral.Fields("Porcentaje").Value), "0.00"))
    '                            .set_TextMatrix(I, C_ColJPRECIOTAG, VB6.Format(Trim(RsGral.Fields("importe").Value), gstrFormatoCantidad))
    '                            .set_TextMatrix(I, C_ColJPORCDESCTOTAG, VB6.Format(Trim(RsGral.Fields("Porcentaje").Value), "0.00"))
    '                            .set_TextMatrix(I, C_ColJESTATUS, Trim(RsGral.Fields("Estatus").Value))
    '                            .set_TextMatrix(I, C_ColJESTATUSTAG, Trim(RsGral.Fields("Estatus").Value))
    '                            .set_TextMatrix(I, C_ColJESNUEVO, False)
    '                            .set_TextMatrix(I, C_COLJTIPO, "G")
    '                            RsGral.MoveNext()
    '                            PonerColor((I))
    '                            'For J = C_ColJFAMILIA To C_ColJESNUEVO
    '                            .set_ColAlignment(C_ColJFAMILIA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                            .set_ColAlignment(C_ColJLINEA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                            .set_ColAlignment(C_ColJARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                            '.ColAlignment(J) = flexAlignLeftCenter
    '                            'Next
    '                        Next
    '                    End With
    '                    sstGrupos.SelectedIndex = 2
    '                    msgVarios.Focus()
    '                End If

    '            Case 1 '''RELOJERIA
    '                gStrSql = "SELECT P.CodGrupo, P.CodMarca, MA.DescMarca, ISNULL(P.CodModelo, 0) AS CodModelo, ISNULL(CO.DescModelo, '') AS DescModelo, P.Importe, " & "P.Porcentaje, P.FechaInicio, P.FechaFin, P.Estatus, ISNULL(Ar.CodArticulo, 0) AS CodArticulo, ISNULL(Ar.DescArticulo, '') AS DescArticulo,P.TipoProm " & "FROM dbo.PromocionesVentas P INNER JOIN " & "dbo.CatMarcas MA ON P.CodMarca = MA.CodMarca LEFT OUTER JOIN " & "dbo.CatModelos CO ON P.CodMarca = CO.CodMarca AND P.CodModelo = CO.CodModelo LEFT OUTER JOIN " & "dbo.CatArticulos Ar ON P.CodArticulo = Ar.CodArticulo " & "WHERE     (P.FechaInicio = ' " & VB6.Format(dtpFechaInIcioR.Value, C_FORMATFECHAGUARDAR) & "') AND (P.FechaFin = '" & VB6.Format(dtpFechaFinR.Value, C_FORMATFECHAGUARDAR) & "') AND (P.Estatus <> 'C') AND P.TipoProm = 'G' " & "GROUP BY P.CodGrupo, P.Estatus, P.CodMarca, MA.DescMarca, P.CodModelo, CO.DescModelo, P.Importe, P.Porcentaje, P.FechaInicio, P.FechaFin, P.Estatus, " & "Ar.CodArticulo , Ar.DescArticulo, P.Estatus,P.TipoProm " & "Having(P.CodGrupo = " & gCODRELOJERIA & ")"
    '                ModEstandar.BorraCmd()
    '                Cmd.CommandText = "dbo.Up_Select_Datos"
    '                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '                RsGral = Cmd.Execute
    '                '    If ValidaDatos = False Then Exit Sub
    '                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '                If RsGral.RecordCount > 0 Then
    '                    With msgRelojeria
    '                        sstGrupos.SelectedIndex = 1
    '                        If RsGral.RecordCount > .Rows - 1 Then .Rows = RsGral.RecordCount + 1
    '                        For I = 1 To .Rows - 1
    '                            If RsGral.EOF Then Exit For
    '                            .set_TextMatrix(I, C_ColRMARCA, Trim(RsGral.Fields("DescMarca").Value))
    '                            .set_TextMatrix(I, C_ColRMODELO, Trim(RsGral.Fields("DescModelo").Value))
    '                            .set_TextMatrix(I, C_ColRARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
    '                            .set_TextMatrix(I, C_ColRCODMARCA, Trim(RsGral.Fields("CodMArca").Value))
    '                            .set_TextMatrix(I, C_ColRCODMODELO, Trim(RsGral.Fields("CodModelo").Value))
    '                            .set_TextMatrix(I, C_ColRCODARTICULO, Trim(RsGral.Fields("CodArticulo").Value))
    '                            .set_TextMatrix(I, C_ColRPRECIO, VB6.Format(Trim(RsGral.Fields("importe").Value), gstrFormatoCantidad))
    '                            .set_TextMatrix(I, C_ColRPORCDESCTO, VB6.Format(Trim(RsGral.Fields("Porcentaje").Value), "0.00"))
    '                            .set_TextMatrix(I, C_ColRPRECIOTAG, VB6.Format(Trim(RsGral.Fields("importe").Value), gstrFormatoCantidad))
    '                            .set_TextMatrix(I, C_ColRPORCDESCTOTAG, VB6.Format(Trim(RsGral.Fields("Porcentaje").Value), "0.00"))
    '                            .set_TextMatrix(I, C_ColRESTATUS, Trim(RsGral.Fields("Estatus").Value))
    '                            .set_TextMatrix(I, C_ColRESTATUSTAG, Trim(RsGral.Fields("Estatus").Value))
    '                            .set_TextMatrix(I, C_ColRESNUEVO, False)
    '                            .set_TextMatrix(I, C_COLRTIPO, "G")
    '                            RsGral.MoveNext()
    '                            PonerColor((I))
    '                            .set_ColAlignment(C_ColRMARCA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                            .set_ColAlignment(C_ColRMODELO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                            .set_ColAlignment(C_ColRARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                        Next
    '                    End With
    '                    sstGrupos.SelectedIndex = 1
    '                    msgRelojeria.Focus()
    '                End If

    '            Case 0 '''JOYERIA
    '                gStrSql = "SELECT     P.CodGrupo, P.CodFamilia, F.DescFamilia, ISNULL(P.CodLinea, 0) AS CodLinea, ISNULL(L.DescLinea, '') AS DescLinea, ISNULL(P.CodSubLinea, 0) AS CodSubLinea, " & "ISNULL(S.DescSubLinea, '') AS DescSubLinea, P.Importe, P.Porcentaje, P.FechaInicio, " & "P.FechaFin , P.Estatus , ISNULL(A.CodArticulo, 0) AS CodArticulo, ISNULL(A.DescArticulo, '') AS DescArticulo,P.TipoProm " & "FROM dbo.PromocionesVentas P INNER JOIN  " & "dbo.CatFamilias F ON P.CodFamilia = F.CodFamilia AND P.CodGrupo = F.CodGrupo LEFT OUTER  JOIN  " & "dbo.CatLineas L ON P.CodLinea = L.CodLinea AND P.CodFamilia = L.CodFamilia AND P.CodGrupo = L.CodGrupo  LEFT OUTER   JOIN " & "dbo.CatSubLineas S ON P.CodSubLinea = S.CodSubLinea AND P.CodFamilia = S.CodFamilia AND P.CodLinea = S.CodLinea  LEFT OUTER JOIN " & "dbo.CatArticulos A ON P.CodArticulo = A.CodArticulo " & "WHERE     (P.FechaInicio = '" & VB6.Format(dtpFechaInIcioJ.Value, C_FORMATFECHAGUARDAR) & "') AND (P.FechaFin = '" & VB6.Format(dtpFechaFinJ.Value, C_FORMATFECHAGUARDAR) & " ') and Estatus <> 'C' AND P.TipoProm = 'G' " & "GROUP BY P.CodGrupo, P.CodFamilia, F.DescFamilia, P.CodLinea, L.DescLinea, S.DescSubLinea, P.CodSubLinea, P.Importe, P.Porcentaje, P.FechaInicio, " & "P.FechaFin , P.Estatus, A.CodArticulo, A.DescArticulo,P.TipoProm " & "Having (P.CodGrupo = " & gCODJOYERIA & ") "
    '                ModEstandar.BorraCmd()
    '                Cmd.CommandText = "dbo.Up_Select_Datos"
    '                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '                RsGral = Cmd.Execute
    '                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '                If RsGral.RecordCount > 0 Then
    '                    With msgJoyeria
    '                        sstGrupos.SelectedIndex = 0
    '                        If RsGral.RecordCount > .Rows - 1 Then .Rows = RsGral.RecordCount + 1
    '                        For I = 1 To .Rows - 1
    '                            If RsGral.EOF Then Exit For
    '                            .set_TextMatrix(I, C_ColJFAMILIA, Trim(RsGral.Fields("DescFamilia").Value))
    '                            .set_TextMatrix(I, C_ColJLINEA, Trim(RsGral.Fields("DescLinea").Value))
    '                            .set_TextMatrix(I, C_ColJSUBLINEA, Trim(RsGral.Fields("DescSubLinea").Value))
    '                            .set_TextMatrix(I, C_ColJARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
    '                            .set_TextMatrix(I, C_ColJCODFAMILIA, Trim(RsGral.Fields("CodFamilia").Value))
    '                            .set_TextMatrix(I, C_ColJCODLINEA, Trim(RsGral.Fields("COdLinea").Value))
    '                            .set_TextMatrix(I, C_ColJCODSUBLINEA, Trim(RsGral.Fields("CodSubLinea").Value))
    '                            .set_TextMatrix(I, C_ColJCODARTICULO, Trim(RsGral.Fields("CodArticulo").Value))
    '                            .set_TextMatrix(I, C_ColJPRECIO, VB6.Format(Trim(RsGral.Fields("importe").Value), gstrFormatoCantidad))
    '                            .set_TextMatrix(I, C_ColJPORCDESCTO, VB6.Format(Trim(RsGral.Fields("Porcentaje").Value), "0.00"))
    '                            .set_TextMatrix(I, C_ColJPRECIOTAG, VB6.Format(Trim(RsGral.Fields("importe").Value), gstrFormatoCantidad))
    '                            .set_TextMatrix(I, C_ColJPORCDESCTOTAG, VB6.Format(Trim(RsGral.Fields("Porcentaje").Value), "0.00"))
    '                            .set_TextMatrix(I, C_ColJESTATUS, Trim(RsGral.Fields("Estatus").Value))
    '                            .set_TextMatrix(I, C_ColJESTATUSTAG, Trim(RsGral.Fields("Estatus").Value))
    '                            .set_TextMatrix(I, C_ColJESNUEVO, False)
    '                            .set_TextMatrix(I, C_COLJTIPO, "G")
    '                            RsGral.MoveNext()
    '                            PonerColor((I))
    '                            'For J = C_ColJFAMILIA To C_ColJESNUEVO
    '                            .set_ColAlignment(C_ColJFAMILIA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                            .set_ColAlignment(C_ColJLINEA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                            .set_ColAlignment(C_ColJSUBLINEA, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                            .set_ColAlignment(C_ColJARTICULO, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '                            'Next
    '                        Next
    '                    End With
    '                    sstGrupos.SelectedIndex = 0
    '                    msgJoyeria.Focus()
    '                End If
    '            Case Else
    '                Exit Sub
    '        End Select
    '        mblnNuevo = False
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '        DesHabilitarFechas()
    '        Exit Sub
    'Merr:
    '        If Err.Number <> 0 Then
    '            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '            ModEstandar.MostrarError("Ocurrió un error en el formulario y proceso: " & gstrProcesoqueGeneraError)
    '        End If
    '    End Sub

    '    Function BuscarCodigoArticulo(ByRef Codigo As String) As Integer
    '        On Error GoTo Merr
    '        Dim CodigoString As String
    '        Dim CodAnterior As Integer
    '        Dim CodOrigen As Integer
    '        BuscarCodigoArticulo = 0
    '        'Esta función recibe como parámetro el código de artículo que se desea buscar.
    '        'Se buscará en la tabla de articulos, en codigo de Articulo y Código de Articulo anterior.
    '        'Es posible que se presenten tres situaciones:
    '        'el código buscado está en el campo código de Articulo de la Tabla
    '        '       --En este caso el codigo del articulo a buscar no cambia.
    '        'el código buscado está en el campo código anterior
    '        '       --En este caso, el codigo a buscar ahora será el que corresponda en el campo Codigo articulo del mismo registro.
    '        'El codigo a buscar está en los dos campos anteriores
    '        '       --En este caso, se mostrará una pantalla de ayuda para mostarle al usuario, los dos articulos encontrados. De los cuales debe seleccionar uno.

    '        ''Esta función regresa:
    '        '    -1 : Si el articulo no es encontró
    '        '    -2 : Si se encontró más de un Artículo
    '        'Codigo = IIf((Trim(Codigo) = ""), CLng(Numerico(Codigo)), Trim(Codigo))ç
    '        If Trim(Codigo) = "" Then Exit Function
    '        If Len(CStr(Codigo)) = 6 Then
    '            CodigoString = (New String("0", 6) & CStr(Codigo))
    '            CodOrigen = CDbl((CodigoString))
    '            CodAnterior = CInt((CodigoString))
    '            gStrSql = "SELECT  * From CatArticulos WHERE (CodArticulo = " & Codigo & ") " & "OR   (OrigenAnt = " & CodOrigen & " AND CodigoAnt = " & CodAnterior & ")"
    '        Else
    '            gStrSql = "SELECT  * From CatArticulos WHERE (CodArticulo = " & Codigo & ") "
    '        End If
    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
    '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '        RsGral = Cmd.Execute
    '        Select Case RsGral.RecordCount
    '            Case Is <= 0
    '                'No se encontró el código de articulo
    '                BuscarCodigoArticulo = -1
    '            Case 1
    '                BuscarCodigoArticulo = CInt(RsGral.Fields("CodArticulo").Value)
    '            Case Else
    '                'Se encontró más de un registro en el catalogo.
    '                BuscarCodigoArticulo = -2
    '        End Select
    '        Exit Function
    'Merr:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Function

    '    Sub BuscarArticulos(ByRef BusquedaEspecial As Boolean, ByRef CodArticulo As String)
    '        On Error GoTo Merr
    '        Dim strSQL As String
    '        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
    '        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
    '        Dim strControlActual As String 'Nombre del control actual
    '        Dim Columna As Integer
    '        strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
    '        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
    '        With msgXArticulo
    '            'Obtener la columna de donde se está ejecutando la consulta
    '            Columna = .Col
    '            If Columna = C_COLXARTCODARTICULO Then 'Se Busca por código
    '                strCaptionForm = "Consulta de Articulos"
    '                If BusquedaEspecial Then
    '                    strSQL = "SELECT     CodArticulo AS CODIGO, RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION, " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR], " & "dbo.FormatCantidad(A.PrecioPubDolar)  AS [PRECIO PÚBLICO] , " & "case PesosFijos WHEN 0 THEN 'DÓLARES' WHEN 1 THEN 'PESOS' END AS [MONEDA] " & "From CatArticulos A cross Join Configuraciongeneral c WHERE (CodArticulo = " & CInt(CodArticulo) & ") " & "OR   (OrigenAnt = " & CInt((CodArticulo)) & ") AND (CodigoAnt = " & CInt((CodArticulo)) & ")"
    '                End If
    '            Else
    '                'Sale del Sub si no es ninguna de estas columnas de donde se ejecuto la consulta, y no hace nada
    '                Exit Sub
    '            End If
    '        End With

    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.Up_Select_Datos"
    '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, strSQL))
    '        RsGral = Cmd.Execute

    '        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
    '        If RsGral.RecordCount = 0 Then
    '            MsgBox(C_msgSINDATOS & vbNewLine & "Verifique por favor....", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
    '            RsGral.Close()
    '            Exit Sub
    '        End If

    '        'Carga el formulario de consulta
    '        'Load(FrmConsultas)
    '        'If BusquedaEspecial = True Then
    '        '    Call ConfiguraConsultas(FrmConsultas, 10300, RsGral, strTag, strCaptionForm)
    '        'Else
    '        '    Call ConfiguraConsultas(FrmConsultas, 7400, RsGral, strTag, strCaptionForm)
    '        'End If

    '        'With FrmConsultas.Flexdet
    '        '    With msgXArticulo
    '        '        'Obtener la columna de donde se está ejecutando la consulta
    '        '        Columna = .Col
    '        '    End With
    '        '    If BusquedaEspecial = True Then
    '        '        If Columna = C_COLXARTCODARTICULO Then 'Se Busca por código
    '        '            .set_ColWidth(0,  , 900)
    '        '            .set_ColWidth(1,  , 4800)
    '        '            .set_ColWidth(2,  , 1700)
    '        '            .set_ColWidth(3,  , 1700)
    '        '            .set_ColWidth(4,  , 1200)
    '        '            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '            .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '        '            .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '            .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '            .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
    '        '        Else
    '        '            'Sale del Sub si no es ninguna de estas columnas de donde se ejecuto la consulta, y no hace nada
    '        '            Exit Sub
    '        '        End If
    '        '    Else
    '        '        If Columna = C_COLXARTCODARTICULO Then 'Se Busca por código
    '        '            .set_ColWidth(0,  , 900)
    '        '            .set_ColWidth(1,  , 4800)
    '        '            .set_ColWidth(2,  , 1700)
    '        '            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '            .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '        '            .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '        Else
    '        '            'Sale del Sub si no es ninguna de estas columnas de donde se ejecuto la consulta, y no hace nada
    '        '            Exit Sub
    '        '        End If
    '        '    End If
    '        'End With
    '        'CentrarForma(FrmConsultas)
    '        'FrmConsultas.ShowDialog()
    'Merr:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    ''' SE MODIFICO CONSULTA PARA GRUPO - X ARTICULO Y SE AGREGO ART X PROV
    '    ''' 20ABR2006 - MAVF
    '    Sub Buscar()
    '        On Error GoTo Merr

    '        Dim strSQL As String
    '        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
    '        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
    '        Dim strControlActual As String 'Nombre del control actual
    '        Dim Columna As Integer

    '        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
    '        strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
    '        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
    '        'Si el control actual es el control de Folio de Venta, se hace una búsqueda de folios de venta.
    '        'Para lo cual validar si si se requiere autorizacion para  consultar folios de venta, segun el parametro de configuracion general del PV
    '        Select Case strControlActual
    '            Case "DTPFECHAINICIOJ", "DTPFECHAINICIOR", "DTPFECHAINICIOV"
    '                If gblnAutConsultaFoliosVta = True Then
    '                    'Pedir el usuario y password para modificar el descto
    '                    'Para esto se usará la forma: frmAutorizacionConfig.
    '                    'frmAutorizacionConfig.Text = "Autorizacion para Consultar Promociones de Ventas"
    '                    'frmAutorizacionConfig.ShowDialog()
    '                    If gblnAutorizacionAceptada = False Then
    '                        'Si la Peticion no fue aceptada, es decir que el usuario que se proporciono no tiene derecho para autorizar o para modificar
    '                        'entonces no podrá ser modificado el descuento
    '                        If gblnSalioSinValidar = False Then 'Si valido el Usuari y Password y no tuvo derecho, mostrar el aviso de ke no puede hacerlo
    '                            MsgBox(C_msgSINAUTORIZACION & "Consultar Promociones de Ventas.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AVISO")
    '                        End If
    '                        Exit Sub
    '                    End If
    '                End If
    '                '''Case "TXTARTICULO", "TXTARTICULOR", "TXTARTICULOV", "TXTFLEX", "MSGJOYERIA", "MSGRELOJERIA", "MSGVARIOS", "MSGXARTICULO", "TXTDETARTXPROV", "MSGARTXPROV"
    '            Case "TXTARTICULO", "TXTARTICULOR", "TXTARTICULOV", "TXTFLEX", "TXTDETARTXPROV"
    '                If (strControlActual = "TXTFLEX" Or strControlActual = "MSGXARTICULO") Then
    '                    If msgXArticulo.Row > 1 Then
    '                        If Trim(msgXArticulo.get_TextMatrix(msgXArticulo.Row - 1, C_COLXARTCODARTICULO)) = "" Then Exit Sub
    '                    End If
    '                End If
    '                '''MODIFIC.-  SE AGREGO LA SECCION DE ASIGNAR DESCUENTOS A ARTICULOS POR PROVEEDOR
    '                '''21ABR2006 - MAVF
    '                If (strControlActual = "TXTDETARTXPROV" Or strControlActual = "MSGARTXPROV") Then
    '                    If msgArtxProv.Row > 1 Then
    '                        If Trim(msgArtxProv.get_TextMatrix(msgArtxProv.Row - 1, C_COLXPRVCODARTICULO)) = "" Then Exit Sub
    '                    End If
    '                End If
    '            Case Else
    '                'Sale de este sub para ke no ejecute ninguna opcion
    '                Exit Sub
    '        End Select
    '        Select Case strControlActual
    '            Case "DTPFECHAINICIOJ", "DTPFECHAINICIOR", "DTPFECHAINICIOV"
    '                strCaptionForm = "Consulta de Promociones de Ventas"

    '                '''MODIFIC.-  SE AGREGO EL ASIGNAR DESCUENTOS A ARTICULOS POR PROVEEDOR
    '                '''20ABR2006 - MAVF
    '                gStrSql = "Select   Ltrim(Rtrim(GRUPO)) as GRUPO, Ltrim(Rtrim(FECHAINICIO)) AS [  FECHA INICIO], Ltrim(Rtrim(FECHAFIN))AS [    FECHA FIN], " & "         FechaI , FechaF, DescProvACreed as PROVEEDOR, Case When Renglon = 0 Then '' Else ltrim(rtrim(convert(char(4), Renglon))) End as [NO. PROG], CodProvAcreed " & "From     vw_Promocionesventas " & "GROUP    BY GRUPO,FECHAINICIO,FECHAFIN, FechaI, FechaF, DescProvAcreed, Renglon, CodProvAcreed  " & "Order    by FechaI Desc, DescProvAcreed, Renglon "

    '                '''gStrSql = "select Ltrim(Rtrim(GRUPO)) as GRUPO, Ltrim(Rtrim(FECHAINICIO)) AS [FECHA INICIO], Ltrim(Rtrim(FECHAFIN))AS [FECHA FIN], FechaI, FechaF " & _
    '                    '"From vw_Promocionesventas  " & _
    '                    '"GROUP BY GRUPO,FECHAINICIO,FECHAFIN, FechaI, FechaF  " & _
    '                    '"order by FechaI desc"

    '                '''Case "TXTARTICULO", "TXTARTICULOR", "TXTARTICULOV", "TXTFLEX", "MSGJOYERIA", "MSGRELOJERIA", "MSGVARIOS", "MSGXARTICULO", "TXTDETARTXPROV", "MSGARTXPROV"
    '            Case "TXTARTICULO", "TXTARTICULOR", "TXTARTICULOV", "TXTFLEX", "TXTDETARTXPROV"
    '                strCaptionForm = "Consulta de Articulos"
    '                If sstGrupos.SelectedIndex = 0 Then
    '                    If strControlActual = "TXTARTICULO" And CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA))) <> 0 Then
    '                        If Trim(txtArticulo.Text) = "" Then
    '                            gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos WHERE CodGrupo = 1 " & "AND CodFamilia = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA)) & IIf(CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA))) = 0, " ", " AND CodLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA))) & IIf(CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA))) = 0, " ", " AND CodSubLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA))) & " ORDER BY DescArticulo"
    '                        Else
    '                            gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos where CodGrupo = 1 AND DescArticulo Like '" & Trim(txtArticulo.Text) & "%' " & "AND CodFamilia = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA)) & IIf(CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA))) = 0, " ", " AND CodLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODLINEA))) & IIf(CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA))) = 0, " ", " AND CodSubLinea = " & Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA))) & " ORDER BY DescArticulo"
    '                        End If
    '                    ElseIf strControlActual = "MSGJOYERIA" And CDbl(Numerico(msgJoyeria.get_TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA))) <> 0 Then
    '                        'If Trim(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJARTICULO)) = "" Then
    '                        '    gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos WHERE CodGrupo = 1 " & _
    '                        ''    "AND CodFamilia = " & Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA)) & _
    '                        ''    IIf(Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODLINEA)) = 0, " ", " AND CodLinea = " & Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODLINEA))) & _
    '                        ''    IIf(Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA)) = 0, " ", " AND CodSubLinea = " & Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA))) & _
    '                        ''    " ORDER BY DescArticulo"
    '                        'Else
    '                        '    gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos where CodGrupo = 1 AND DescArticulo Like '" & Trim(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJARTICULO)) & "%' " & _
    '                        ''    "AND CodFamilia = " & Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODFAMILIA)) & _
    '                        ''    IIf(Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODLINEA)) = 0, " ", " AND CodLinea = " & Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODLINEA))) & _
    '                        ''    IIf(Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA)) = 0, " ", " AND CodSubLinea = " & Numerico(msgJoyeria.TextMatrix(msgJoyeria.Row, C_ColJCODSUBLINEA))) & _
    '                        ''    " ORDER BY DescArticulo"
    '                        'End If
    '                    Else
    '                        Exit Sub
    '                    End If
    '                ElseIf sstGrupos.SelectedIndex = 1 Then
    '                    If strControlActual = "TXTARTICULOR" And CDbl(Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA))) <> 0 Then
    '                        If Trim(txtArticuloR.Text) = "" Then
    '                            gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos WHERE CodGrupo = 2 " & "AND CodMarca = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA)) & IIf(CDbl(Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMODELO))) = 0, " ", " AND CodModelo = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMODELO))) & " ORDER BY DescArticulo"
    '                        Else
    '                            gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos WHERE CodGrupo = 2 AND DescArticulo Like '" & Trim(txtArticuloR.Text) & "%' " & "AND CodMarca = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA)) & IIf(CDbl(Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMODELO))) = 0, " ", " AND CodModelo = " & Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMODELO))) & " ORDER BY DescArticulo"
    '                        End If
    '                    ElseIf strControlActual = "MSGRELOJERIA" And CDbl(Numerico(msgRelojeria.get_TextMatrix(msgRelojeria.Row, C_ColRCODMARCA))) <> 0 Then
    '                        'If Trim(msgRelojeria.TextMatrix(msgRelojeria.Row, C_ColRARTICULO)) = "" Then
    '                        '    gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos WHERE CodGrupo = 2 " & _
    '                        ''    "AND CodMarca = " & Numerico(msgRelojeria.TextMatrix(msgRelojeria.Row, C_ColRCODMARCA)) & _
    '                        ''    IIf(Numerico(msgRelojeria.TextMatrix(msgRelojeria.Row, C_ColRCODMODELO)) = 0, " ", " AND CodModelo = " & Numerico(msgRelojeria.TextMatrix(msgRelojeria.Row, C_ColRCODMODELO))) & _
    '                        ''    " ORDER BY DescArticulo"
    '                        'Else
    '                        '    gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos where CodGrupo = 2 AND DescArticulo Like '" & Trim(msgRelojeria.TextMatrix(msgRelojeria.Row, C_ColRARTICULO)) & "%' " & _
    '                        ''    "AND CodMarca = " & Numerico(msgRelojeria.TextMatrix(msgRelojeria.Row, C_ColRCODMARCA)) & _
    '                        ''    IIf(Numerico(msgRelojeria.TextMatrix(msgRelojeria.Row, C_ColRCODMODELO)) = 0, " ", " AND CodModelo = " & Numerico(msgRelojeria.TextMatrix(msgRelojeria.Row, C_ColRCODMODELO))) & _
    '                        ''    " ORDER BY DescArticulo"
    '                        'End If
    '                    Else
    '                        Exit Sub
    '                    End If
    '                ElseIf sstGrupos.SelectedIndex = 2 Then
    '                    If strControlActual = "TXTARTICULOV" And CDbl(Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODFAMILIA))) <> 0 Then
    '                        If Trim(txtArticuloV.Text) = "" Then
    '                            gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos WHERE CodGrupo = 3 " & "AND CodFamilia = " & Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODFAMILIA)) & IIf(CDbl(Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODLINEA))) = 0, " ", " AND CodLinea = " & Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODLINEA))) & " ORDER BY DescArticulo"
    '                        Else
    '                            gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos WHERE CodGrupo = 3 AND DescArticulo Like '" & Trim(txtArticuloV.Text) & "%' " & "AND CodFamilia = " & Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODFAMILIA)) & IIf(CDbl(Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODLINEA))) = 0, " ", " AND CodLinea = " & Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODLINEA))) & " ORDER BY DescArticulo"
    '                        End If
    '                    ElseIf strControlActual = "MSGVARIOS" And CDbl(Numerico(msgVarios.get_TextMatrix(msgVarios.Row, C_ColJCODFAMILIA))) <> 0 Then
    '                        'If Trim(msgVarios.TextMatrix(msgVarios.Row, C_ColJARTICULO)) = "" Then
    '                        '    gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos WHERE CodGrupo = 3 " & _
    '                        ''    "AND CodFamilia = " & Numerico(msgVarios.TextMatrix(msgVarios.Row, C_ColJCODFAMILIA)) & _
    '                        ''    IIf(Numerico(msgVarios.TextMatrix(msgVarios.Row, C_ColJCODLINEA)) = 0, " ", " AND CodLinea = " & Numerico(msgVarios.TextMatrix(msgVarios.Row, C_ColJCODLINEA))) & _
    '                        ''    " ORDER BY DescArticulo"
    '                        'Else
    '                        '    gStrSql = "SELECT Rtrim(Ltrim(DescArticulo)) AS DESCRIPCION,CodArticulo as CODIGO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] From CatArticulos where CodGrupo = 2 AND DescArticulo Like '" & Trim(msgVarios.TextMatrix(msgVarios.Row, C_ColJARTICULO)) & "%' " & _
    '                        ''    "AND CodFamilia = " & Numerico(msgVarios.TextMatrix(msgVarios.Row, C_ColJCODFAMILIA)) & _
    '                        ''    IIf(Numerico(msgVarios.TextMatrix(msgVarios.Row, C_ColJCODLINEA)) = 0, " ", " AND CodLinea = " & Numerico(msgVarios.TextMatrix(msgVarios.Row, C_ColJCODLINEA))) & _
    '                        ''    " ORDER BY DescArticulo"
    '                        'End If
    '                    Else
    '                        Exit Sub
    '                    End If
    '                ElseIf sstGrupos.SelectedIndex = 3 Then
    '                    If (strControlActual = "TXTFLEX" Or strControlActual = "MSGXARTICULO") And msgXArticulo.Col = C_COLXARTCODARTICULO Then
    '                        gStrSql = "SELECT CodArticulo AS CODIGO, RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION, " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] " & "From CatArticulos ORDER BY CodArticulo"
    '                    ElseIf (strControlActual = "TXTFLEX" Or strControlActual = "MSGXARTICULO") And msgXArticulo.Col = C_COLXARTDESCARTICULO Then
    '                        If strControlActual = "TXTFLEX" And Trim(txtFlex.Text) = "" Then
    '                            gStrSql = "SELECT RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION,CodArticulo AS CODIGO,  " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] " & "From CatArticulos ORDER BY DescArticulo"
    '                        ElseIf strControlActual = "TXTFLEX" And Trim(txtFlex.Text) <> "" Then
    '                            gStrSql = "SELECT RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION,CodArticulo AS CODIGO,  " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] " & "From CatArticulos WHERE DescArticulo LIKE '" & Trim(txtFlex.Text) & "%' ORDER BY DescArticulo"
    '                        ElseIf strControlActual = "MSGXARTICULO" And Trim(msgXArticulo.get_TextMatrix(msgXArticulo.Row, C_COLXARTDESCARTICULO)) = "" Then
    '                            'gStrSql = "SELECT RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION,CodArticulo AS CODIGO,  " & _
    '                            ''"CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] " & _
    '                            ''"From CatArticulos ORDER BY DescArticulo"
    '                        ElseIf strControlActual = "MSGXARTICULO" And Trim(msgXArticulo.get_TextMatrix(msgXArticulo.Row, C_COLXARTDESCARTICULO)) <> "" Then
    '                            'gStrSql = "SELECT RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION,CodArticulo AS CODIGO,  " & _
    '                            ''"CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] " & _
    '                            ''"From CatArticulos WHERE DescArticulo LIKE '" & Trim(msgXArticulo.TextMatrix(msgXArticulo.Row, C_COLXARTDESCARTICULO)) & "%' ORDER BY DescArticulo"
    '                        End If
    '                    Else
    '                        Exit Sub
    '                    End If

    '                    '''MODIFIC.-  SE AGREGO LA SECCION DE ASIGNAR DESCUENTOS A ARTICULOS POR PROVEEDOR
    '                    '''21ABR2006 - MAVF
    '                ElseIf sstGrupos.SelectedIndex = 4 Then
    '                    If (strControlActual = "TXTDETARTXPROV" Or strControlActual = "MSGARTXPROV") And msgArtxProv.Col = C_COLXPRVCODARTICULO Then
    '                        gStrSql = "SELECT CodArticulo AS CODIGO, RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION, " & "CASE   CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] " & "From   CatArticulos (Nolock) " & "Where  CodProveedor = " & mintCodProveedor & " ORDER BY CodArticulo "
    '                    ElseIf (strControlActual = "TXTDETARTXPROV" Or strControlActual = "MSGARTXPROV") And msgArtxProv.Col = C_COLXPRVDESCARTICULO Then
    '                        If strControlActual = "TXTDETARTXPROV" And Trim(txtDetArtxProv.Text) = "" Then
    '                            gStrSql = "SELECT RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION,CodArticulo AS CODIGO,  " & "CASE   CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] " & "From   CatArticulos (Nolock) " & "Where  CodProveedor = " & mintCodProveedor & " ORDER BY DescArticulo "
    '                        ElseIf strControlActual = "TXTDETARTXPROV" And Trim(txtDetArtxProv.Text) <> "" Then
    '                            gStrSql = "SELECT RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION,CodArticulo AS CODIGO,  " & "CASE   CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] " & "From   CatArticulos (Nolock) " & "WHERE  DescArticulo LIKE '" & Trim(txtDetArtxProv.Text) & "%' And CodProveedor = " & mintCodProveedor & " ORDER BY DescArticulo "
    '                        ElseIf strControlActual = "MSGARTXPROV" And Trim(msgArtxProv.get_TextMatrix(msgArtxProv.Row, C_COLXPRVDESCARTICULO)) = "" Then
    '                            '''gStrSql = "SELECT RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION,CodArticulo AS CODIGO,  " & _
    '                            '"CASE   CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] " & _
    '                            '"From   CatArticulos (Nolock) " & _
    '                            '"Where  CodProveedor = " & mintCodProveedor & " " & _
    '                            '"ORDER  BY DescArticulo "
    '                        ElseIf strControlActual = "MSGARTXPROV" And Trim(msgArtxProv.get_TextMatrix(msgArtxProv.Row, C_COLXPRVDESCARTICULO)) <> "" Then
    '                            '''gStrSql = "SELECT RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION,CodArticulo AS CODIGO,  " & _
    '                            '"CASE   CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR] " & _
    '                            '"From   CatArticulos (Nolock) " & _
    '                            '"WHERE  DescArticulo LIKE '" & Trim(msgArtxProv.TextMatrix(msgArtxProv.Row, C_COLXPRVDESCARTICULO)) & "%' And CodProveedor = " & mintCodProveedor & " ORDER BY DescArticulo"
    '                        End If
    '                    Else
    '                        Exit Sub
    '                    End If

    '                End If
    '            Case Else
    '                'Sale de este sub para ke no ejecute ninguna opcion
    '                Exit Sub
    '        End Select
    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.Up_Select_Datos"
    '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '        RsGral = Cmd.Execute
    '        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
    '        If RsGral.RecordCount = 0 Then
    '            MsgBox(C_msgSINDATOS & vbNewLine & "Verifique Por Favor....", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
    '            RsGral.Close()
    '            Exit Sub
    '        End If
    '        'Carga el formulario de consulta
    '        'Load(FrmConsultas)
    '        'With FrmConsultas.Flexdet
    '        '    Select Case strControlActual
    '        '        Case "DTPFECHAINICIOJ", "DTPFECHAINICIOR", "DTPFECHAINICIOV"
    '        '            Call ConfiguraConsultas(FrmConsultas, 8100, RsGral, strTag, strCaptionForm)
    '        '            .set_ColWidth(0,  , 1500) 'GRUPO
    '        '            .set_ColWidth(1,  , 1300) 'fECHAINICIO
    '        '            .set_ColWidth(2,  , 1300) 'FECHAFIN
    '        '            .set_ColWidth(3,  , 0) 'FECHAI
    '        '            .set_ColWidth(4,  , 0) 'FECHAF
    '        '            .set_ColWidth(5,  , 3000) 'PROVEEDOR
    '        '            .set_ColWidth(6,  , 950) 'NO. PROG
    '        '            .set_ColWidth(7,  , 0) 'PROVEEDOR
    '        '            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '        '            .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '            .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '            .set_ColAlignment(5, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '        '            .set_ColAlignment(6, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
    '        '        Case "TXTARTICULO", "TXTARTICULOR", "TXTARTICULOV", "TXTFLEX", "MSGJOYERIA", "MSGRELOJERIA", "MSGVARIOS", "MSGXARTICULO", "TXTDETARTXPROV", "MSGARTXPROV"
    '        '            If (strControlActual = "TXTFLEX" Or strControlActual = "MSGXARTICULO") And msgXArticulo.Col = C_COLXARTCODARTICULO Then
    '        '                Call ConfiguraConsultas(FrmConsultas, 7400, RsGral, strTag, strCaptionForm)
    '        '                .set_ColWidth(0,  , 900)
    '        '                .set_ColWidth(1,  , 4800)
    '        '                .set_ColWidth(2,  , 1700)
    '        '                .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '                .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '        '                .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '            ElseIf (strControlActual = "TXTDETARTXPROV" Or strControlActual = "MSGARTXPROV") And msgArtxProv.Col = C_COLXPRVCODARTICULO Then
    '        '                Call ConfiguraConsultas(FrmConsultas, 7400, RsGral, strTag, strCaptionForm)
    '        '                .set_ColWidth(0,  , 900)
    '        '                .set_ColWidth(1,  , 4800)
    '        '                .set_ColWidth(2,  , 1700)
    '        '                .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '                .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '        '                .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '            Else
    '        '                Call ConfiguraConsultas(FrmConsultas, 7400, RsGral, strTag, strCaptionForm)
    '        '                .set_ColWidth(0,  , 4800)
    '        '                .set_ColWidth(1,  , 900)
    '        '                .set_ColWidth(2,  , 1700)
    '        '                .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '                .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
    '        '                .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
    '        '            End If
    '        '        Case Else
    '        '            Exit Sub
    '        '    End Select
    '        'End With
    '        'FrmConsultas.Flexdet.FontFixed = VB6.FontChangeName(FrmConsultas.Flexdet.FontFixed, "Small Fonts")
    '        'FrmConsultas.Flexdet.Font = VB6.FontChangeName(FrmConsultas.Flexdet.Font, "MS Sans Serif")
    '        'FrmConsultas.Flexdet.Font = VB6.FontChangeName(FrmConsultas.Flexdet.Font, CStr(8.25))
    '        'FrmConsultas.Flexdet.FontFixed = VB6.FontChangeName(FrmConsultas.Flexdet.FontFixed, CStr(6.75))
    '        'FrmConsultas.Flexdet.Row = 0
    '        'For I = 0 To FrmConsultas.Flexdet.get_COLS() - 1
    '        '    FrmConsultas.Flexdet.Col = I
    '        '    FrmConsultas.Flexdet.CellFontBold = True
    '        'Next
    '        'FrmConsultas.Flexdet.FontFixed = VB6.FontChangeName(FrmConsultas.Flexdet.FontFixed, CStr(6.75))
    '        'CentrarForma(FrmConsultas)
    '        'FrmConsultas.ShowDialog()
    '        'Exit Sub

    'Merr:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    ''' SE MODIFICO CONSULTA PARA GRUPO - X ARTICULO Y SE AGREGO ART X PROV
    '    ''' 20ABR2006 - MAVF
    '    Sub PonerColor(ByRef Fila As Integer)
    '        Dim I As Integer
    '        Dim Ctl As System.Windows.Forms.Control
    '        Dim nCol As Integer
    '        Dim GridACtivo As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    '        Dim Estatus As String

    '        Select Case sstGrupos.SelectedIndex
    '            Case 0
    '                GridACtivo = msgJoyeria
    '            Case 1
    '                GridACtivo = msgRelojeria
    '            Case 2
    '                GridACtivo = msgVarios
    '            Case 3
    '                GridACtivo = msgXArticulo
    '            Case 4
    '                GridACtivo = msgArtxProv
    '        End Select

    '        If GridACtivo Is msgJoyeria Then
    '            Estatus = GridACtivo.get_TextMatrix(Fila, C_ColJESTATUS)
    '        ElseIf GridACtivo Is msgRelojeria Then
    '            Estatus = GridACtivo.get_TextMatrix(Fila, C_ColRESTATUS)
    '        ElseIf GridACtivo Is msgVarios Then
    '            Estatus = GridACtivo.get_TextMatrix(Fila, C_ColJESTATUS)
    '        ElseIf GridACtivo Is msgXArticulo Then
    '            Estatus = GridACtivo.get_TextMatrix(Fila, C_COLXARTESTATUS)
    '        ElseIf GridACtivo Is msgArtxProv Then
    '            Estatus = GridACtivo.get_TextMatrix(Fila, C_COLXPRVESTATUS)
    '        End If

    '        With GridACtivo
    '            Select Case Estatus
    '                Case C_Cancelado
    '                    Ctl = lblCancelada
    '                Case Else
    '                    Ctl = GridACtivo
    '            End Select
    '            .Row = Fila
    '            For I = 0 To 10
    '                .Col = I
    '                '.CellBackColor = System.Drawing.ColorTranslator.FromOle(Ctl.BackColor)
    '            Next
    '            .Col = nCol
    '        End With

    '    End Sub

    '    ''' SE MODIFICO CONSULTA PARA GRUPO - X ARTICULO Y SE AGREGO ART X PROV
    '    ''' 20ABR2006 - MAVF
    '    '''ojo - *****  pendiente detrminar si se agrega el prov - 21ABR  *****
    '    Function ValidarPromocionGuardadaRepetida(ByRef CodGrupo As Integer, ByRef CodFamilia As Integer, ByRef COdLinea As Integer, ByRef CodSubLinea As Integer, ByRef CodMArca As Integer, ByRef CodModelo As Integer, ByRef CodArticulo As Integer, ByRef FechaInicio As Date, ByRef FechaFin As Date, ByRef TipoProm As String) As Boolean
    '        'Esta Función Valida si una promoción se está repitiendo. tomando en cuenta el Articulo que se está poniendo en promoción y las fecha de Vigencia.
    '        'Si las fechas se traslapan para un mismo Artículo. Se considera que se está repitiendo una promoción.
    '        'Se Valida en relacion a promociones almacenadas en la BD
    '        gStrSql = "Select [dbo].ArticuloenPromocionRepetida(" & CodGrupo & "," & CodFamilia & "," & COdLinea & "," & CodSubLinea & " , " & CodMArca & " , " & CodModelo & " , " & CodArticulo & "  , '" & VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR) & "' , '" & VB6.Format(FechaFin, C_FORMATFECHAGUARDAR) & "','" & TipoProm & "') as Repuesta"
    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.UP_Select_Datos"
    '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '        RsGral = Cmd.Execute
    '        If RsGral.Fields(0).Value = 0 Then
    '            ValidarPromocionGuardadaRepetida = False
    '        Else
    '            ValidarPromocionGuardadaRepetida = True
    '        End If
    '    End Function

    '    '''SE AGREGO EL PROCESO PARA LA SECCION DE DESCUENTOS POR PROVEEDOR
    '    '''21ABR2006 - MAVF
    '    Sub ValidarPromocionTecleadaRepetida()
    '        'Esta Función Valida si una promoción se está repitiendo. tomando en cuenta el Articulo que se está poniendo en promoción y las fecha de Vigencia.
    '        'Si las fechas se traslapan para un mismo Artículo. Se considera que se está repitiendo una promoción.
    '        'Se Valida en relacion a promociones que se están tecleando en el GRid, y que  no se han guardado en la BD

    '        Dim CodFamilia As Integer
    '        Dim COdLinea As Integer
    '        Dim CodSubLinea As Integer
    '        Dim CodMArca As Integer
    '        Dim CodModelo As Integer
    '        Dim CodArticulo As Integer
    '        Dim CodFamilia2 As Integer
    '        Dim codLinea2 As Integer
    '        Dim CodSubLinea2 As Integer
    '        Dim CodMarca2 As Integer
    '        Dim CodModelo2 As Integer
    '        Dim CodArticulo2 As Integer
    '        Dim J As Integer

    '        If sstGrupos.SelectedIndex = 0 Then
    '            GridACtivo = msgJoyeria
    '        ElseIf sstGrupos.SelectedIndex = 1 Then
    '            GridACtivo = msgRelojeria
    '        ElseIf sstGrupos.SelectedIndex = 2 Then
    '            GridACtivo = msgVarios
    '        ElseIf sstGrupos.SelectedIndex = 3 Then
    '            GridACtivo = msgXArticulo
    '        ElseIf sstGrupos.SelectedIndex = 4 Then
    '            GridACtivo = msgArtxProv
    '        End If
    '        With GridACtivo
    '            If GridACtivo Is msgJoyeria Or GridACtivo Is msgVarios Then
    '                For I = 1 To .Rows - 1
    '                    If Trim(.get_TextMatrix(I, C_ColJFAMILIA)) = "" Then Exit For
    '                    CodFamilia = CShort(Numerico(.get_TextMatrix(I, C_ColJCODFAMILIA)))
    '                    COdLinea = CShort(Numerico(.get_TextMatrix(I, C_ColJCODLINEA)))
    '                    CodSubLinea = CShort(Numerico(.get_TextMatrix(I, C_ColJCODSUBLINEA)))
    '                    CodArticulo = CInt(Numerico(.get_TextMatrix(I, C_ColJCODARTICULO)))
    '                    CodMArca = 0
    '                    CodModelo = 0
    '                    For J = I + 1 To .Rows - 2
    '                        If Trim(.get_TextMatrix(J, C_ColJFAMILIA)) = "" Then Exit For
    '                        CodFamilia2 = CShort(Numerico(.get_TextMatrix(J, C_ColJCODFAMILIA)))
    '                        codLinea2 = CShort(Numerico(.get_TextMatrix(J, C_ColJCODLINEA)))
    '                        CodSubLinea2 = CShort(Numerico(.get_TextMatrix(J, C_ColJCODSUBLINEA)))
    '                        CodArticulo2 = CShort(Numerico(.get_TextMatrix(J, C_ColJCODARTICULO)))
    '                        CodMarca2 = 0
    '                        CodModelo2 = 0
    '                        If CodFamilia = CodFamilia2 And COdLinea = codLinea2 And CodSubLinea = CodSubLinea2 And CodArticulo = CodArticulo2 Then
    '                            MsgBox("No es posible repetir una promoción para un artículo." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            'LimpiaDatosFamilia
    '                            BorraGrid(J, GridACtivo, 0)
    '                            .set_TextMatrix(J, C_ColJFAMILIA, "")
    '                            .Focus()
    '                            .Col = C_ColJFAMILIA
    '                            Exit Sub
    '                        End If
    '                        Select Case True
    '                            Case CodFamilia = CodFamilia2 And (COdLinea = 0 Or codLinea2 = 0) And (CodSubLinea = 0 Or CodSubLinea2 = 0) And CodArticulo2 = 0 And CodArticulo = 0
    '                                MsgBox("No es posible repetir la información para un artículo." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                                LimpiaDatosFamilia()
    '                                .set_TextMatrix(J, C_ColJFAMILIA, "")
    '                                BorraGrid(J, GridACtivo, 0)
    '                                .Focus()
    '                                .Col = C_ColJFAMILIA
    '                                Exit Sub
    '                            Case CodFamilia = CodFamilia2 And COdLinea = codLinea2 And (CodSubLinea = 0 Or CodSubLinea2 = 0) And CodArticulo2 = 0 And CodArticulo = 0
    '                                MsgBox("No es posible repetir la información para un artículo." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                                LimpiaDatosFamilia()
    '                                .set_TextMatrix(J, C_ColJFAMILIA, "")
    '                                BorraGrid(J, GridACtivo, 0)
    '                                .Focus()
    '                                .Col = C_ColJFAMILIA
    '                                Exit Sub
    '                        End Select
    '                    Next
    '                Next
    '            ElseIf GridACtivo Is msgRelojeria Then
    '                For I = 1 To .Rows - 1
    '                    If Trim(.get_TextMatrix(I, C_ColRMARCA)) = "" Then Exit For
    '                    CodFamilia = 0
    '                    COdLinea = 0
    '                    CodSubLinea = 0
    '                    CodMArca = CShort(Numerico(.get_TextMatrix(I, C_ColRCODMARCA)))
    '                    CodModelo = CShort(Numerico(.get_TextMatrix(I, C_ColRCODMODELO)))
    '                    CodArticulo = CInt(Numerico(.get_TextMatrix(I, C_ColRCODARTICULO)))
    '                    For J = I + 1 To .Rows - 2
    '                        If Trim(.get_TextMatrix(J, C_ColJFAMILIA)) = "" Then Exit For
    '                        CodFamilia2 = 0
    '                        codLinea2 = 0
    '                        CodSubLinea2 = 0
    '                        CodMarca2 = CShort(Numerico(.get_TextMatrix(J, C_ColRCODMARCA)))
    '                        CodModelo2 = CShort(Numerico(.get_TextMatrix(J, C_ColRCODMODELO)))
    '                        CodArticulo2 = CShort(Numerico(.get_TextMatrix(J, C_ColRCODARTICULO)))
    '                        If CodMArca = CodMarca2 And CodModelo = CodModelo2 And CodArticulo = CodArticulo2 Then
    '                            MsgBox("No es posible repetir una promoción para un artículo." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            LimpiaDatosMarca()
    '                            .set_TextMatrix(J, C_ColRMARCA, "")
    '                            .Focus()
    '                            .Col = C_ColRMARCA
    '                            Exit Sub
    '                        End If
    '                        Select Case True
    '                            Case CodMArca = CodMarca2 And (CodModelo = 0 Or CodModelo2 = 0) And CodArticulo2 = 0 And CodArticulo = 0
    '                                MsgBox("No es Posible repetir la Información para un Artículo." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                                LimpiaDatosMarca()
    '                                BorraGrid(J, GridACtivo, 0)
    '                                LimpiaDatosMarca()
    '                                .Focus()
    '                                .Col = C_ColRMARCA
    '                                Exit Sub
    '                        End Select
    '                    Next
    '                Next
    '            ElseIf GridACtivo Is msgXArticulo Then
    '                For I = 1 To .Rows - 1
    '                    If Trim(.get_TextMatrix(I, C_COLXARTCODARTICULO)) = "" Then Exit Sub
    '                    CodArticulo = CInt(.get_TextMatrix(I, C_COLXARTCODARTICULO))
    '                    For J = I + 1 To .Rows - 2
    '                        If Trim(.get_TextMatrix(J, C_COLXARTCODARTICULO)) = "" Then Exit For
    '                        CodArticulo2 = CShort(.get_TextMatrix(J, C_COLXARTCODARTICULO))
    '                        If CodArticulo = CodArticulo2 Then
    '                            MsgBox("No es posible repetir una promoción para un artículo." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            .set_TextMatrix(J, C_COLXARTCODARTICULO, "")
    '                            .set_TextMatrix(J, C_COLXARTDESCARTICULO, "")
    '                            .set_TextMatrix(J, C_COLXARTCODANTERIOR, "")
    '                            .set_TextMatrix(J, C_COLXARTCODFAMILIA, "")
    '                            .set_TextMatrix(J, C_COLXARTCODGRUPO, "")
    '                            .set_TextMatrix(J, C_COLXARTCODLINEA, "")
    '                            .set_TextMatrix(J, C_COLXARTCODMARCA, "")
    '                            .set_TextMatrix(J, C_COLXARTCODMODELO, "")
    '                            .set_TextMatrix(J, C_COLXARTCODSUBLINEA, "")
    '                            .set_TextMatrix(J, C_COLXARTESNUEVO, "")
    '                            .set_TextMatrix(J, C_COLXARTESTATUS, "")
    '                            .set_TextMatrix(J, C_COLXARTPRECIO, "")
    '                            .set_TextMatrix(J, C_COLXARTPORCDESCTO, "")
    '                            .Col = C_COLXARTCODARTICULO
    '                            txtFlex.Text = ""
    '                            .Focus()
    '                        End If
    '                    Next
    '                Next

    '            ElseIf GridACtivo Is msgArtxProv Then
    '                For I = 1 To .Rows - 1
    '                    If Trim(.get_TextMatrix(I, C_COLXPRVCODARTICULO)) = "" Then Exit Sub
    '                    CodArticulo = CInt(.get_TextMatrix(I, C_COLXPRVCODARTICULO))
    '                    For J = I + 1 To .Rows - 2
    '                        If Trim(.get_TextMatrix(J, C_COLXPRVCODARTICULO)) = "" Then Exit For
    '                        CodArticulo2 = CShort(.get_TextMatrix(J, C_COLXPRVCODARTICULO))
    '                        If CodArticulo = CodArticulo2 Then
    '                            MsgBox("No es posible repetir una promoción para un artículo." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            txtDetArtxProv.Text = ""
    '                            txtDetArtxProv.Visible = False
    '                            .set_TextMatrix(J, C_COLXPRVCODARTICULO, "")
    '                            .set_TextMatrix(J, C_COLXPRVDESCARTICULO, "")
    '                            .set_TextMatrix(J, C_COLXPRVCODANTERIOR, "")
    '                            .set_TextMatrix(J, C_COLXPRVCODFAMILIA, "")
    '                            .set_TextMatrix(J, C_COLXPRVCODGRUPO, "")
    '                            .set_TextMatrix(J, C_COLXPRVCODLINEA, "")
    '                            .set_TextMatrix(J, C_COLXPRVCODMARCA, "")
    '                            .set_TextMatrix(J, C_COLXPRVCODMODELO, "")
    '                            .set_TextMatrix(J, C_COLXPRVCODSUBLINEA, "")
    '                            .set_TextMatrix(J, C_COLXPRVESNUEVO, "")
    '                            .set_TextMatrix(J, C_COLXPRVESTATUS, "")
    '                            .set_TextMatrix(J, C_COLXPRVPRECIO, "")
    '                            .set_TextMatrix(J, C_COLXPRVPORCDESCTO, "")
    '                            PonerColor(J)
    '                            .Col = C_COLXPRVCODARTICULO
    '                            .Row = .Row - 1
    '                            ''' ESTABA txtDetArtxProv = ""
    '                            .Focus()
    '                        End If
    '                    Next
    '                Next

    '            End If
    '        End With
    '    End Sub

    '    Sub BorraGrid(ByRef Row As Integer, ByRef GridACtivo As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid, ByRef RowSel As Integer)
    '        Dim lRen As Integer
    '        Dim lRen2 As Integer
    '        Dim I As Integer

    '        lRen = 0
    '        lRen2 = 0
    '        'Este Procediento borra un renglon del Grid
    '        'Si el Número de Filas que kedan en el grid, es menor de 8, se insertará una nueva fila al final del grid
    '        With GridACtivo
    '            '''es un solo renglon
    '            If (Row = RowSel) Or (Row > 0 And RowSel = 0) Then
    '                .RemoveItem(Row)
    '                'Si el número de filas es menor de 10 o esta posicionado en la utlima fila, entonces, agrega una fila
    '                If .Rows < 11 Or .Row = .Rows - 1 Then
    '                    .AddItem("")
    '                    .Row = .Row
    '                End If
    '            Else ''' si es un conjunto de renglones
    '                If Row < RowSel Then
    '                    lRen = Row
    '                    lRen2 = RowSel
    '                Else
    '                    lRen = RowSel
    '                    lRen2 = Row
    '                End If
    '                For I = lRen2 To lRen Step -1
    '                    .RemoveItem(I)
    '                    'Si el número de filas es menor de 10 o esta posicionado en la utlima fila, entonces, agrega una fila
    '                    'If .Rows < 11 Or .Row = .Rows - 1 Then
    '                    '   .AddItem ""
    '                    '   .Row = .Row
    '                    'End If
    '                Next I
    '                If .Rows <= 6 Then
    '                    .Rows = 11
    '                Else
    '                    .Rows = .Rows + 1
    '                End If
    '            End If
    '        End With

    '    End Sub

    '    Private Sub txtRelojeria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRelojeria.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        txtRelojeria.Visible = False
    '    End Sub

    '    '*************************************************************/
    '    '*************************************************************/
    '    ' articulos por proveedor

    '    Private Sub dbcProveedor_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ColumnWidthChangedEventArgs) Handles dbcProveedor.CursorChanged
    '        On Error GoTo Merr
    '        Dim lStrSql As String

    '        If mblnFueraChange Then Exit Sub


    '        lStrSql = "SELECT CodProvAcreed, DescProvAcreed =ltrim(rtrim(DescProvAcreed)) From CatProvAcreed Where DescProvAcreed like '" & Trim(dbcProveedor.Text) & "%' ORDER BY DescProvAcreed "
    '        ModDCombo.DCChange(lStrSql, tecla, dbcProveedor)

    '        If Trim(dbcProveedor.Text) = "" Then
    '            mintCodProveedor = 0
    '        End If

    'Merr:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
    '        Pon_Tool()
    '        gStrSql = "SELECT CodProvAcreed, DescProvAcreed =ltrim(rtrim(DescProvAcreed)) From CatProvAcreed ORDER BY DescProvAcreed "
    '        ModDCombo.DCGotFocus(gStrSql, dbcProveedor)
    '    End Sub

    '    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcProveedor.KeyDown
    '        If eventSender.keyCode = System.Windows.Forms.Keys.Escape Then
    '            sstGrupos.Focus()
    '            eventSender.keyCode = 0
    '        End If
    '        tecla = eventSender.keyCode
    '    End Sub

    '    Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
    '        On Error GoTo Merr
    '        Dim Aux As Integer

    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub

    '        gStrSql = "SELECT CodProvAcreed, DescProvAcreed =ltrim(rtrim(DescProvAcreed)) From CatProvAcreed Where DescProvAcreed = '" & Trim(dbcProveedor.Text) & "' ORDER BY DescProvAcreed "
    '        Aux = mintCodProveedor
    '        mintCodProveedor = 0
    '        '''If Trim(dbcProveedor.text) <> Trim(C_TODAS) Or Trim(dbcProveedor.text) = "" Then
    '        If Trim(dbcProveedor.Text) <> "" Then
    '            ModDCombo.DCLostFocus(dbcProveedor, gStrSql, mintCodProveedor)
    '            LlenaArticulosProveedor(CInt(mintCodProveedor))
    '        End If

    '        If Aux <> mintCodProveedor Then
    '            If mintCodProveedor = 0 Then
    '                mblnFueraChange = True
    '                dbcProveedor.Text = ""
    '                dbcProveedor.Enabled = True
    '                mblnFueraChange = False
    '            End If
    '        End If

    'Merr:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    Private Sub msgArtxProv_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgArtxProv.DblClick
    '        msgArtxProv_KeyPressEvent(msgArtxProv, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    '    End Sub

    '    Private Sub msgArtxProv_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgArtxProv.Enter
    '        mblnFueraChange = True
    '        If chkAplicar.CheckState = System.Windows.Forms.CheckState.Checked Then chkAplicar.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        If chkBorrar.CheckState = System.Windows.Forms.CheckState.Checked Then chkBorrar.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        mblnFueraChange = False

    '        With msgArtxProv
    '            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightAlways
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '            .SelectionMode = MSHierarchicalFlexGridLib.SelectionModeSettings.flexSelectionFree
    '        End With
    '        Pon_Tool()
    '    End Sub

    '    Private Sub msgArtxProv_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgArtxProv.Leave
    '        With msgArtxProv
    '            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '            .SelectionMode = MSHierarchicalFlexGridLib.SelectionModeSettings.flexSelectionFree
    '        End With
    '    End Sub

    '    Private Sub msgArtxProv_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles msgArtxProv.KeyDownEvent
    '        Dim lRow As Integer
    '        Dim lCol As Integer

    '        With msgArtxProv
    '            lCol = .Col
    '            Select Case eventArgs.keyCode
    '                Case System.Windows.Forms.Keys.Escape
    '                    chkBorrar.Focus()
    '                Case System.Windows.Forms.Keys.Delete
    '                    If .get_TextMatrix(.Row, C_COLXPRVESTATUS) = C_Cancelado Then
    '                        MsgBox("Está Promoción ya ha sido Cancelada." & vbNewLine & "Verifique Por Favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                        .Focus()
    '                        Exit Sub
    '                    End If
    '                    Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "Mensaje")
    '                        Case MsgBoxResult.No
    '                            .Focus()
    '                            Exit Sub
    '                        Case MsgBoxResult.Cancel
    '                            .Focus()
    '                            Exit Sub
    '                    End Select

    '                    If mblnNuevo Then
    '                        BorraGrid(.Row, msgArtxProv, .RowSel)
    '                        If .Row < .RowSel Then
    '                            .Row = .Row
    '                        Else
    '                            .Row = .RowSel
    '                        End If
    '                        .Col = lCol
    '                        .Focus()
    '                    Else
    '                        CancelaArticulosPromocion(msgArtxProv, .Row, .RowSel)
    '                        '''.TextMatrix(.Row, C_COLXPRVESTATUS) = C_Cancelado
    '                        '''PonerColor .Row
    '                    End If

    '                    CalculaTotales(msgArtxProv)
    '                    .Focus()
    '            End Select
    '        End With
    '    End Sub

    '    Private Sub msgArtxProv_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles msgArtxProv.KeyPressEvent
    '        Dim EsNuevo As Boolean
    '        Dim Estatus As String
    '        EsNuevo = True
    '        With msgArtxProv
    '            'Si nho se trata de un REgistro nuevo, no se podrá editar el Grid
    '            If Trim(.get_TextMatrix(.Row, C_COLXPRVESNUEVO)) <> "" Then
    '                EsNuevo = CBool(.get_TextMatrix(.Row, C_COLXPRVESNUEVO))
    '            End If
    '            Estatus = .get_TextMatrix(.Row, C_COLXPRVESTATUS)
    '            FueraChange = True
    '            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then 'Para que cuando sea escape, no entre a editar el codigo,simplemente que se regrese al control anterior
    '                Select Case .Col
    '                    Case C_COLXPRVCODARTICULO '------- SE EDITA EL CODIGO DE ARTICULO -------
    '                        If EsNuevo = False Or mblnNuevo = False Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)
    '                        '''en esta parte se validará si es el rengón, columna que le corresponde editarse
    '                        If (.Row > 1) Then
    '                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
    '                            If Trim(.get_TextMatrix(.Row - 1, C_COLXPRVCODARTICULO)) = "" Then
    '                                .Focus()
    '                                FueraChange = False
    '                                Exit Sub
    '                            End If
    '                        End If
    '                        'UPGRADE_WARNING: TextBox property txtDetArtxProv.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    '                        txtDetArtxProv.MaxLength = 8
    '                        txtDetArtxProv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '                        ModEstandar.MSHFlexGridEdit(msgArtxProv, txtDetArtxProv, eventArgs.keyAscii)
    '                    Case C_COLXPRVDESCARTICULO
    '                        If EsNuevo = False Or mblnNuevo = False Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.gp_CampoAlfanumerico(eventArgs.keyAscii)

    '                        If (.Row > 1) Then
    '                            '''de tal modo que si el renglón es mayor que 1 y si un renglón antes del renglón actual está vacío, el renglón actual no se editará
    '                            If Trim(.get_TextMatrix(.Row - 1, C_COLXPRVDESCARTICULO)) = "" Then
    '                                .Focus()
    '                                FueraChange = False
    '                                Exit Sub
    '                            End If
    '                        End If

    '                        txtDetArtxProv.MaxLength = 150
    '                        txtDetArtxProv.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
    '                        ModEstandar.MSHFlexGridEdit(msgArtxProv, txtDetArtxProv, eventArgs.keyAscii)
    '                    Case C_COLXPRVCODANTERIOR
    '                        .Focus()
    '                        FueraChange = False
    '                        Exit Sub
    '                    Case C_COLXPRVPORCDESCTO
    '                        ModEstandar.gp_CampoNumerico(eventArgs.keyAscii, ".")
    '                        If Estatus = C_Cancelado Then
    '                            .Focus()
    '                            Exit Sub
    '                        End If
    '                        If Trim(.get_TextMatrix(.Row, C_COLXPRVCODARTICULO)) = "" Then
    '                            .Focus()
    '                            FueraChange = False
    '                            Exit Sub
    '                        End If
    '                        If CDbl(Numerico(.get_TextMatrix(.Row, C_COLXPRVPRECIO))) <> 0 Then
    '                            MsgBox("ya se ha Asignado un Precio. No es Posible Asignar un Porcentaje de Descuento.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            .Focus()
    '                            FueraChange = False
    '                            Exit Sub
    '                        End If
    '                        ModEstandar.MSHFlexGridEdit(msgArtxProv, txtDetArtxProv, eventArgs.keyAscii)
    '                        .set_TextMatrix(.Row, C_COLXPRVPRECIO, "0.00")
    '                        txtDetArtxProv.MaxLength = 6
    '                        txtDetArtxProv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '                        '''Case C_COLXPRVPRECIO
    '                End Select
    '            End If
    '        End With
    '        FueraChange = False
    '    End Sub

    '    Private Sub txtDetArtxProv_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetArtxProv.Enter
    '        txtDetArtxProv.Text = Trim(txtDetArtxProv.Text)
    '        If Len(Trim(txtDetArtxProv.Text)) > 1 Then
    '            SelTextoTxt(txtDetArtxProv)
    '        End If
    '        Pon_Tool()
    '    End Sub

    '    Private Sub txtDetArtxProv_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDetArtxProv.KeyDown
    '        Dim KeyCode As Integer = eventArgs.KeyCode
    '        Dim Shift As Integer = eventArgs.KeyData \ &H10000
    '        'Aqui se muestran los datos del control editable, en el Grid
    '        'Se deberá formatear el Valor de Acuerdo al Tipo de Dato en uso
    '        Dim rowsiguiente As Integer
    '        Dim ColSiguiente As Integer
    '        Dim ResBusquedaArt As Integer

    '        Select Case KeyCode
    '            Case System.Windows.Forms.Keys.Escape
    '                txtDetArtxProv.Focus()
    '                txtDetArtxProv.Visible = False
    '                txtDetArtxProv.Text = ""
    '                msgArtxProv.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                msgArtxProv.Focus()
    '            Case System.Windows.Forms.Keys.Return
    '                With msgArtxProv
    '                    If .Row > 1 Then
    '                        If Trim(.get_TextMatrix(.Row - 1, C_COLXPRVCODARTICULO)) = "" Then Exit Sub
    '                    End If
    '                    rowsiguiente = .Row
    '                    If .Col = C_COLXPRVCODARTICULO Then
    '                        ColSiguiente = C_COLXPRVPORCDESCTO
    '                        If Trim(txtDetArtxProv.Text) <> "" Then
    '                            ResBusquedaArt = BuscarCodigoArticulo(Trim(txtDetArtxProv.Text))
    '                            If ResBusquedaArt > 0 Or ResBusquedaArt = -1 Then
    '                                If ResBusquedaArt > 0 Then
    '                                    LlenaDatosArtxProv(ResBusquedaArt)
    '                                ElseIf ResBusquedaArt = -1 Then
    '                                    LlenaDatosArtxProv(CInt(Numerico(txtDetArtxProv.Text)))
    '                                End If
    '                                Exit Sub
    '                            ElseIf ResBusquedaArt = -2 And CDbl(Numerico(txtDetArtxProv.Text)) <> 0 Then
    '                                ResBusquedaArt = CInt(txtDetArtxProv.Text)
    '                                .set_TextMatrix(.Row, C_COLXPRVCODARTICULO, "")
    '                                BuscarArticulos(True, (New String("0", 6) & CStr(ResBusquedaArt)))
    '                                Exit Sub
    '                            End If
    '                            LlenaDatosArtxProv(CInt(Numerico(txtDetArtxProv.Text)))
    '                        Else
    '                            .set_TextMatrix(.Row, C_COLXPRVCODARTICULO, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVDESCARTICULO, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVCODANTERIOR, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVCODFAMILIA, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVCODGRUPO, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVCODLINEA, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVCODMARCA, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVCODMODELO, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVCODSUBLINEA, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVESNUEVO, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVESTATUS, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVPRECIO, "")
    '                            .set_TextMatrix(.Row, C_COLXPRVPORCDESCTO, "")
    '                        End If
    '                        '''ElseIf .Col = C_COLXPRVPRECIO Then
    '                    ElseIf .Col = C_COLXPRVPORCDESCTO Then
    '                        If CShort(ModEstandar.Numerico(txtDetArtxProv.Text)) > 100 Then
    '                            MsgBox("El Porcentaje de Descuento no puede ser mayor de 100.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                            txtDetArtxProv.Focus()
    '                            Exit Sub
    '                        End If
    '                        .set_TextMatrix(.Row, .Col, VB6.Format(Numerico(txtDetArtxProv.Text), gstrFormatoCantidad))
    '                        If CShort(ModEstandar.Numerico(.get_TextMatrix(.Row, .Col))) = 0 Then
    '                            MarcaArtsNoSel(.Row, 1)
    '                            rowsiguiente = .Row + 1
    '                            ColSiguiente = C_COLXPRVCODARTICULO
    '                        Else
    '                            MarcaArtsNoSel(.Row, 2)
    '                            rowsiguiente = .Row + 1
    '                            ColSiguiente = C_COLXPRVCODARTICULO
    '                        End If
    '                    End If
    '                    FueraChange = True
    '                    txtDetArtxProv.Text = ""
    '                    txtDetArtxProv.Visible = False
    '                    FueraChange = False
    '                    .Col = .Col
    '                    .Row = .Row
    '                    .Focus()
    '                    If (.Col = C_COLXPRVCODARTICULO Or .Col = C_COLXPRVDESCARTICULO) And (Trim(.get_TextMatrix(.Row, .Col)) = "") Then
    '                        Exit Sub
    '                    End If
    '                    If .Row = .Rows - 1 Then
    '                        .Rows = .Rows + 2
    '                        .Row = rowsiguiente
    '                        .Col = ColSiguiente
    '                        .set_TextMatrix(.Row, C_COLXPRVTIPO, "A")
    '                        .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightAlways
    '                        .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                        .SelectionMode = MSHierarchicalFlexGridLib.SelectionModeSettings.flexSelectionFree
    '                        .Focus()
    '                    Else
    '                        .Row = rowsiguiente
    '                        .Col = ColSiguiente
    '                    End If
    '                    If (.Row - .TopRow) > 5 Then
    '                        .TopRow = .TopRow + 1
    '                    End If
    '                End With
    '        End Select

    '    End Sub

    '    Private Sub txtDetArtxProv_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDetArtxProv.KeyPress
    '        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
    '        'En este Evento se validan los datos que se introduzcan al control txtjoyeria,dependiendo de la columan en que se esté editando
    '        If KeyAscii = 0 Or KeyAscii = 13 Then GoTo EventExitSub
    '        With msgArtxProv
    '            Select Case .Col
    '                Case C_COLXPRVCODARTICULO
    '                    KeyAscii = ModEstandar.MskCantidad(txtDetArtxProv.Text, KeyAscii, 8, 0, (txtDetArtxProv.SelectionStart))
    '                Case C_COLXPRVDESCARTICULO
    '                Case C_COLXPRVPORCDESCTO
    '                    ModEstandar.gp_CampoNumerico(KeyAscii, ".")
    '                    KeyAscii = ModEstandar.MskCantidad(txtDetArtxProv.Text, KeyAscii, 3, 2, (txtDetArtxProv.SelectionStart))
    '            End Select
    '        End With
    'EventExitSub:
    '        eventArgs.KeyChar = Chr(KeyAscii)
    '        If KeyAscii = 0 Then
    '            eventArgs.Handled = True
    '        End If
    '    End Sub

    '    Private Sub txtDetArtxProv_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDetArtxProv.Leave
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
    '        txtDetArtxProv.Visible = False
    '    End Sub

    '    ''' SE MODIFICO CONSULTA PARA GRUPO - X ARTICULO Y SE AGREGO ART X PROV
    '    ''' 20ABR2006 - MAVF
    '    Sub LlenaDatosArtxProv(ByRef CodArticulo As Integer)
    '        Dim lDescto As Decimal
    '        On Error GoTo Err_Renamed

    '        If CodArticulo = 0 Then Exit Sub

    '        gStrSql = "SELECT A.CodArticulo, LTRIM(RTRIM(A.DescArticulo)) AS DescArticulo, ISNULL(A.CodFamilia, 0) AS CodFamilia, ISNULL(A.CodLinea, 0) AS CodLinea,A.CodGrupo,CASE A.CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),A.OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),A.CodigoAnt))) ,5) End as CodigoAnt, " & "ISNULL(A.CodSubLinea, 0) AS CodSubLinea, ISNULL(A.CodMarca, 0) AS CodMarca, ISNULL(A.CodModelo, 0) AS CodModelo, ISNULL(F.DescFamilia, '') " & "AS DescFamilia, ISNULL(L.DescLinea, '') AS DescLinea, ISNULL(S.DescSubLinea, '') AS DescSubLinea, ISNULL(Ma.DescMarca, '') AS DescMarca, ISNULL(Mo.DescModelo, '') AS Descmodelo " & "FROM         dbo.CatArticulos A LEFT OUTER  JOIN " & "dbo.CatFamilias F ON A.CodGrupo = F.CodGrupo AND A.CodFamilia = F.CodFamilia LEFT OUTER JOIN " & "dbo.CatLineas L ON A.CodGrupo = L.CodGrupo AND A.CodFamilia = L.CodFamilia AND A.CodLinea = L.CodLinea AND F.CodGrupo = L.CodGrupo AND " & "F.CodFamilia = L.CodFamilia LEFT OUTER JOIN " & "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca LEFT OUTER JOIN " & "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND A.CodMarca = Mo.CodMarca AND A.CodModelo = Mo.CodModelo AND " & "Ma.CodGrupo = Mo.CodGrupo AND Ma.CodMarca = Mo.CodMarca LEFT OUTER JOIN " & "dbo.CatSubLineas S ON A.CodGrupo = S.CodGrupo AND A.CodFamilia = S.CodFamilia AND A.CodLinea = S.CodLinea AND " & "A.CodSubLinea = S.CodSubLinea AND F.CodGrupo = S.CodGrupo AND F.CodFamilia = S.CodFamilia AND L.CodGrupo = S.CodGrupo AND " & "L.CodFamilia = s.CodFamilia And L.COdLinea = s.COdLinea " & "Where (A.CodArticulo = " & CodArticulo & ") and a.CodProveedor = " & mintCodProveedor & " "
    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
    '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '        RsGral = Cmd.Execute

    '        If RsGral.RecordCount > 0 Then
    '            lDescto = CDec(ModEstandar.Numerico(Trim(txtDesctoP.Text)))
    '            With msgArtxProv
    '                .set_TextMatrix(.Row, C_COLXPRVCODARTICULO, RsGral.Fields("CodArticulo").Value)
    '                .set_TextMatrix(.Row, C_COLXPRVCODANTERIOR, RsGral.Fields("CodigoAnt").Value)
    '                .set_TextMatrix(.Row, C_COLXPRVDESCARTICULO, Trim(RsGral.Fields("DescArticulo").Value))
    '                '''toma el descuento de manera general para los articulos ingresados manualmente del txtDesctoP
    '                .set_TextMatrix(.Row, C_COLXPRVPORCDESCTO, VB6.Format(lDescto, "###,##0.00"))
    '                .set_TextMatrix(.Row, C_COLXPRVPORCDESCTOTAG, VB6.Format(CDec(ModEstandar.Numerico(Trim(txtDesctoP.Text))), "###,##0.00"))
    '                If lDescto = 0 Then
    '                    MarcaArtsNoSel(.Row, 1)
    '                Else
    '                    MarcaArtsNoSel(.Row, 2)
    '                End If

    '                .set_TextMatrix(.Row, C_COLXPRVCODFAMILIA, RsGral.Fields("CodFamilia").Value)
    '                .set_TextMatrix(.Row, C_COLXPRVCODGRUPO, RsGral.Fields("CodGrupo").Value)
    '                .set_TextMatrix(.Row, C_COLXPRVCODLINEA, RsGral.Fields("COdLinea").Value)
    '                .set_TextMatrix(.Row, C_COLXPRVCODMARCA, RsGral.Fields("CodMArca").Value)
    '                .set_TextMatrix(.Row, C_COLXPRVCODMODELO, RsGral.Fields("CodModelo").Value)
    '                .set_TextMatrix(.Row, C_COLXPRVCODSUBLINEA, RsGral.Fields("CodSubLinea").Value)
    '                .set_TextMatrix(.Row, C_COLXPRVESTATUS, C_Aplicado)
    '                .set_TextMatrix(.Row, C_COLXPRVESTATUSTAG, C_Aplicado)
    '                .set_TextMatrix(.Row, C_COLXPRVESNUEVO, True)
    '                ValidarPromocionTecleadaRepetida()
    '                If ValidarPromocionGuardadaRepetida(CShort(Numerico(.get_TextMatrix(.Row, C_COLXPRVCODGRUPO))), CShort(Numerico(.get_TextMatrix(.Row, C_COLXPRVCODFAMILIA))), CShort(Numerico(.get_TextMatrix(.Row, C_COLXPRVCODLINEA))), CShort(Numerico(.get_TextMatrix(.Row, C_COLXPRVCODSUBLINEA))), CShort(Numerico(.get_TextMatrix(.Row, C_COLXPRVCODMARCA))), CShort(Numerico(.get_TextMatrix(.Row, C_COLXPRVCODMODELO))), CInt(Numerico(.get_TextMatrix(.Row, C_COLXPRVCODARTICULO))), dtpFechaInIcioJ.Value, dtpFechaFinJ.Value, "A") = True And mblnNuevo = True Then
    '                    MsgBox("Existe una promoción registrada para este artículo." & vbNewLine & "No es posible duplicar promociones en un lapso de tiempo similar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
    '                    .set_TextMatrix(.Row, C_COLXPRVCODARTICULO, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVDESCARTICULO, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVCODANTERIOR, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVCODFAMILIA, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVCODGRUPO, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVCODLINEA, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVCODMARCA, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVCODMODELO, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVCODSUBLINEA, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVESNUEVO, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVESTATUS, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVPRECIO, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVPORCDESCTO, "")
    '                    .set_TextMatrix(.Row, C_COLXPRVTIPO, "")

    '                    MarcaArtsNoSel(.Row, 2)
    '                    BorraGrid(.Row, msgArtxProv, 0)
    '                    txtDetArtxProv.Text = ""
    '                    txtDetArtxProv.Visible = False
    '                    .Row = .Row
    '                    .Col = C_COLXPRVCODARTICULO
    '                    .Focus()
    '                    Exit Sub
    '                End If
    '                txtDetArtxProv.Visible = False
    '                .Col = C_COLXPRVPORCDESCTO
    '                .Focus()
    '            End With
    '            CalculaTotales(msgArtxProv)
    '        Else
    '            MsgBox("Código de articulo no existe o no" & vbNewLine & "pertenece a este proveedor" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
    '            txtDetArtxProv.Text = ""
    '            txtDetArtxProv.Visible = False
    '            msgArtxProv.Row = msgArtxProv.Row
    '            msgArtxProv.Col = msgArtxProv.Col
    '            msgArtxProv.Focus()
    '            '''txtDetArtxProv.SetFocus
    '        End If

    'Err_Renamed:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    ''' SE MODIFICO CONSULTA PARA GRUPO - X ARTICULO Y SE AGREGO ART X PROV
    '    ''' 20ABR2006 - MAVF
    '    Sub LlenaArticulosProveedor(ByRef CodProveedor As Integer)
    '        On Error GoTo Err_Renamed
    '        Dim rsLocal As ADODB.Recordset
    '        Dim lRen As Integer
    '        Dim lDescto As Decimal
    '        Dim lCont As Integer

    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

    '        msgArtxProv.Clear()
    '        lblTotArt.Text = "0"
    '        lblTotArt.Refresh()
    '        Encabezado()
    '        chkAplicar.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        chkBorrar.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        mintTotalRen = 0
    '        If CodProveedor = 0 Then Exit Sub

    '        gStrSql = "SELECT A.CodArticulo, LTRIM(RTRIM(A.DescArticulo)) AS DescArticulo, ISNULL(A.CodFamilia, 0) AS CodFamilia, ISNULL(A.CodLinea, 0) AS CodLinea,A.CodGrupo,CASE A.CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),A.OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),A.CodigoAnt))) ,5) End as CodigoAnt, " & "       ISNULL(A.CodSubLinea, 0) AS CodSubLinea, ISNULL(A.CodMarca, 0) AS CodMarca, ISNULL(A.CodModelo, 0) AS CodModelo, ISNULL(F.DescFamilia, '') " & "       AS DescFamilia, ISNULL(L.DescLinea, '') AS DescLinea, ISNULL(S.DescSubLinea, '') AS DescSubLinea, ISNULL(Ma.DescMarca, '') AS DescMarca, ISNULL(Mo.DescModelo, '') AS Descmodelo " & "FROM   dbo.CatArticulos A LEFT OUTER  JOIN " & "       dbo.CatFamilias F ON A.CodGrupo = F.CodGrupo AND A.CodFamilia = F.CodFamilia LEFT OUTER JOIN " & "       dbo.CatLineas L ON A.CodGrupo = L.CodGrupo AND A.CodFamilia = L.CodFamilia AND A.CodLinea = L.CodLinea AND F.CodGrupo = L.CodGrupo AND " & "       F.CodFamilia = L.CodFamilia LEFT OUTER JOIN " & "       dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca LEFT OUTER JOIN " & "       dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND A.CodMarca = Mo.CodMarca AND A.CodModelo = Mo.CodModelo AND " & "       Ma.CodGrupo = Mo.CodGrupo AND Ma.CodMarca = Mo.CodMarca LEFT OUTER JOIN " & "       dbo.CatSubLineas S ON A.CodGrupo = S.CodGrupo AND A.CodFamilia = S.CodFamilia AND A.CodLinea = S.CodLinea AND " & "       A.CodSubLinea = S.CodSubLinea AND F.CodGrupo = S.CodGrupo AND F.CodFamilia = S.CodFamilia AND L.CodGrupo = S.CodGrupo AND " & "       L.CodFamilia = s.CodFamilia And L.COdLinea = s.COdLinea " & "       LEFT OUTER JOIN  ( " & "       Select   E.CodArticulo, E.Existencia From ( " & "                Select   CodArticulo, sum((ExistenciaInicial + Entradas) - (Salidas + Apartados)) as Existencia " & "                From     Inventario I (Nolock) Inner Join CatAlmacen A (Nolock) On I.CodAlmacen = A.CodAlmacen And A.TipoAlmacen = 'P' " & "                Group    by CodArticulo " & "       ) E Inner Join CatArticulos A (Nolock) On E.CodArticulo = A.CodArticulo " & "       Where A.CodProveedor = " & CodProveedor & " " & "       ) Ex On A.CodArticulo = Ex.CodArticulo " & "Where  (A.CodProveedor = " & CodProveedor & ") " & "And    Ex.Existencia > 0 " & "Order  by A.CodArticulo "

    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
    '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '        rsLocal = Cmd.Execute

    '        If rsLocal.RecordCount > 0 Then
    '            msgArtxProv.Rows = rsLocal.RecordCount + 10
    '            rsLocal.MoveFirst()
    '            lDescto = CDec(ModEstandar.Numerico(Trim(txtDesctoP.Text)))
    '            lCont = 0
    '            For lRen = 1 To rsLocal.RecordCount

    '                With msgArtxProv
    '                    If Not ValidarPromocionGuardadaRepetida(rsLocal.Fields("CodGrupo").Value, rsLocal.Fields("CodFamilia").Value, rsLocal.Fields("COdLinea").Value, rsLocal.Fields("CodSubLinea").Value, rsLocal.Fields("CodMArca").Value, rsLocal.Fields("CodModelo").Value, rsLocal.Fields("CodArticulo").Value, dtpFechaInIcioJ.Value, dtpFechaFinJ.Value, "A") = True And mblnNuevo = True Then
    '                        lCont = lCont + 1
    '                        .set_TextMatrix(lCont, C_COLXPRVCODARTICULO, rsLocal.Fields("CodArticulo").Value)
    '                        .set_TextMatrix(lCont, C_COLXPRVCODANTERIOR, rsLocal.Fields("CodigoAnt").Value)
    '                        .set_TextMatrix(lCont, C_COLXPRVDESCARTICULO, Trim(rsLocal.Fields("DescArticulo").Value))
    '                        .set_TextMatrix(lCont, C_COLXPRVPORCDESCTO, Format(lDescto, "###,##0.00"))
    '                        .set_TextMatrix(lCont, C_COLXPRVPORCDESCTOTAG, Format(lDescto, "###,##0.00"))
    '                        If lDescto = 0 Then
    '                            MarcaArtsNoSel(lCont, 1)
    '                        Else
    '                            MarcaArtsNoSel(lCont, 2)
    '                        End If

    '                        .set_TextMatrix(lCont, C_COLXPRVCODFAMILIA, rsLocal.Fields("CodFamilia").Value)
    '                        .set_TextMatrix(lCont, C_COLXPRVCODGRUPO, rsLocal.Fields("CodGrupo").Value)
    '                        .set_TextMatrix(lCont, C_COLXPRVCODLINEA, rsLocal.Fields("COdLinea").Value)
    '                        .set_TextMatrix(lCont, C_COLXPRVCODMARCA, rsLocal.Fields("CodMArca").Value)
    '                        .set_TextMatrix(lCont, C_COLXPRVCODMODELO, rsLocal.Fields("CodModelo").Value)
    '                        .set_TextMatrix(lCont, C_COLXPRVCODSUBLINEA, rsLocal.Fields("CodSubLinea").Value)
    '                        .set_TextMatrix(lCont, C_COLXPRVESTATUS, C_Aplicado)
    '                        .set_TextMatrix(lCont, C_COLXPRVESTATUSTAG, C_Aplicado)
    '                        .set_TextMatrix(lCont, C_COLXPRVESNUEVO, True)
    '                        .set_TextMatrix(lCont, C_COLXPRVTIPO, "A")

    '                        ValidarPromocionTecleadaRepetida()

    '                    End If
    '                End With
    '                rsLocal.MoveNext()
    '            Next
    '            txtFlex.Visible = False
    '            mintTotalRen = lCont
    '            lblTotArt.Text = VB6.Format(mintTotalRen, "##0")
    '            If lCont = 0 Then
    '                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '                MsgBox("No existen artículos disponibles para el registro" & vbNewLine & "de promociones en las fechas indicadas", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
    '            End If
    '            With msgArtxProv
    '                .TopRow = 1
    '                .Row = 1
    '                .Col = C_COLXPRVCODARTICULO
    '                .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightAlways
    '                .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '                .SelectionMode = MSHierarchicalFlexGridLib.SelectionModeSettings.flexSelectionFree
    '                .Focus()
    '            End With
    '        Else
    '            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '            MsgBox("No existen artículos disponobles para este" & vbNewLine & "proveedor en esta fecha de promoción" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
    '        End If
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    'Err_Renamed:
    '        If Err.Number <> 0 Then
    '            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '            ModEstandar.MostrarError()
    '        End If
    '    End Sub

    '    Sub AsignaDesctoGrid(ByRef msg As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid, ByRef lcurDescto As Decimal)
    '        On Error GoTo Err_Renamed
    '        Dim I As Integer

    '        With msg
    '            For I = 1 To .Rows - 1
    '                If (Trim(.get_TextMatrix(I, C_COLXPRVCODARTICULO)) <> "") And (Trim(.get_TextMatrix(I, C_COLXPRVDESCARTICULO)) <> "") Then
    '                    If (Trim(.get_TextMatrix(I, C_COLXPRVCODARTICULO)) = "") And (Trim(.get_TextMatrix(I, C_COLXPRVDESCARTICULO)) = "") Then Exit For


    '                    .set_TextMatrix(I, C_COLXPRVPORCDESCTO, VB6.Format(lcurDescto, "###,##0.00"))
    '                    If Trim(.get_TextMatrix(I, C_COLXPRVESTATUS)) <> "C" Then
    '                        If lcurDescto = 0 Then
    '                            MarcaArtsNoSel(I, 1)
    '                        Else
    '                            MarcaArtsNoSel(I, 2)
    '                        End If
    '                    End If
    '                End If
    '            Next
    '            .Row = 1
    '            .Col = C_COLXPRVCODARTICULO
    '            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '            .SelectionMode = MSHierarchicalFlexGridLib.SelectionModeSettings.flexSelectionFree
    '        End With

    'Err_Renamed:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    Sub BorraArtsProv(ByRef msg As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid)
    '        On Error GoTo Err_Renamed

    '        msg.Clear()
    '        Encabezado()
    '        mblnFueraChange = True
    '        txtDesctoP.Text = "0.00"
    '        chkAplicar.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        mblnFueraChange = False
    '        mintTotalRen = 0
    '        lblTotArt.Text = Format(mintTotalRen, "##0")

    'Err_Renamed:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    Sub MarcaArtsNoSel(ByRef Fila As Integer, ByRef lTipo As Integer)
    '        On Error GoTo Err_Renamed
    '        Dim I As Integer
    '        Dim Ctl As System.Windows.Forms.Control
    '        Dim nCol As Integer
    '        Dim GridACtivo As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid

    '        Select Case sstGrupos.SelectedIndex
    '            Case 0
    '                GridACtivo = msgJoyeria
    '            Case 1
    '                GridACtivo = msgRelojeria
    '            Case 2
    '                GridACtivo = msgVarios
    '            Case 3
    '                GridACtivo = msgXArticulo
    '            Case 4
    '                GridACtivo = msgArtxProv
    '        End Select

    '        With GridACtivo
    '            Select Case lTipo
    '                Case 1
    '                    Ctl = lblArtsNoSel
    '                Case 2
    '                    Ctl = GridACtivo
    '            End Select
    '            .Row = Fila
    '            For I = 0 To 10
    '                .Col = I
    '                '.CellBackColor = System.Drawing.ColorTranslator.FromOle(Ctl.BackColor)
    '            Next
    '            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightAlways
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '            .SelectionMode = MSHierarchicalFlexGridLib.SelectionModeSettings.flexSelectionFree
    '        End With

    'Err_Renamed:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    Sub CalculaTotales(ByRef msg As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid)
    '        On Error GoTo Err_Renamed
    '        Dim I As Integer
    '        Dim lTotR As Integer
    '        Dim lRen As Integer
    '        Dim lCol As Integer

    '        lTotR = 0
    '        With msg
    '            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '            .SelectionMode = MSHierarchicalFlexGridLib.SelectionModeSettings.flexSelectionFree
    '            lRen = .Row
    '            lCol = .Col
    '            For I = 1 To .Rows - 1
    '                If (Trim(.get_TextMatrix(I, C_COLXPRVCODARTICULO)) <> "") And (Trim(.get_TextMatrix(I, C_COLXPRVDESCARTICULO)) <> "") Then
    '                    If (Trim(.get_TextMatrix(I, C_COLXPRVCODARTICULO)) = "") And (Trim(.get_TextMatrix(I, C_COLXPRVDESCARTICULO)) = "") Then Exit For
    '                    lTotR = lTotR + 1
    '                End If
    '            Next
    '            lblTotArt.Text = Format(lTotR, "##0")
    '            .Row = lRen
    '            .Col = lCol
    '            .HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightWithFocus
    '            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    '            .SelectionMode = MSHierarchicalFlexGridLib.SelectionModeSettings.flexSelectionFree
    '        End With

    'Err_Renamed:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    Private Function CalculaRenglonProv(ByRef lCodProv As Integer, ByRef FechaIni As Date, ByRef FechaFin As Date) As Integer
    '        On Error GoTo Err_Renamed

    '        gStrSql = "SELECT ISNULL(MAX(Renglon), 0) AS Ren " & "From   PromocionesVentas (Nolock) " & "WHERE  (FechaInicio = '" & VB6.Format(FechaIni, C_FORMATFECHAGUARDAR) & "') AND (CodProvAcreed = " & mintCodProveedor & ") AND (TipoProm = 'A') AND (Estatus <> 'C') "

    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.Up_Select_Datos"
    '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '        RsGral = Cmd.Execute
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    '        If RsGral.RecordCount > 0 Then
    '            CalculaRenglonProv = RsGral.Fields("Ren").Value + 1
    '        End If

    'Err_Renamed:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Function

    '    '''Tipo = 1 CANCELAR PROMOCION
    '    '''       2 ACTIVAR PROMOCION
    '    Private Sub CancelaProm(ByRef lTipo As Boolean)
    '        On Error GoTo Err_Renamed
    '        Dim I As Object
    '        Dim J As Integer
    '        Dim Ctl As System.Windows.Forms.Control
    '        Dim nCol As Integer
    '        Dim GridACtivo As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    '        Dim lColEstatus As Integer
    '        Dim lColArticulo As Integer

    '        Select Case sstGrupos.SelectedIndex
    '            Case 0
    '                GridACtivo = msgJoyeria
    '                lColEstatus = 0
    '                lColArticulo = 0
    '            Case 1
    '                GridACtivo = msgRelojeria
    '                lColEstatus = 0
    '                lColArticulo = 0
    '            Case 2
    '                GridACtivo = msgVarios
    '                lColEstatus = 0
    '                lColArticulo = 0
    '            Case 3
    '                GridACtivo = msgXArticulo
    '                lColEstatus = 0
    '                lColArticulo = 0
    '            Case 4
    '                GridACtivo = msgArtxProv
    '                lColEstatus = C_COLXPRVESTATUS
    '                lColArticulo = C_COLXPRVCODARTICULO
    '        End Select

    '        With GridACtivo
    '            For I = 1 To GridACtivo.Rows - 1
    '                If Trim(GridACtivo.get_TextMatrix(I, lColArticulo)) = "" Then Exit For
    '                Select Case lTipo
    '                    Case True
    '                        GridACtivo.set_TextMatrix(I, lColEstatus, "C")
    '                        Ctl = lblCancelada
    '                    Case False
    '                        GridACtivo.set_TextMatrix(I, lColEstatus, "A")
    '                        Ctl = GridACtivo
    '                End Select
    '                .Row = I
    '                For J = 0 To 10
    '                    .Col = J
    '                    '.CellBackColor = System.Drawing.ColorTranslator.FromOle(Ctl.BackColor)
    '                Next
    '                .Col = nCol
    '            Next I
    '        End With

    'Err_Renamed:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Sub

    '    Private Sub CancelaArticulosPromocion(ByRef GridACtivo As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid, ByRef lFila As Integer, ByRef lFilaSel As Integer)
    '        Dim lRen As Integer
    '        Dim lRen2 As Integer
    '        Dim I As Integer

    '        lRen = 0
    '        lRen2 = 0
    '        'Este Procediento marca como cancelado un renglon del Grid

    '        With GridACtivo
    '            If (lFila = lFilaSel) Or (lFila < lFilaSel) Then
    '                lRen = lFila
    '                lRen2 = lFilaSel
    '            Else
    '                lRen = lFilaSel
    '                lRen2 = lFila
    '            End If

    '            For I = lRen2 To lRen Step -1
    '                If Trim(.get_TextMatrix(I, C_COLXPRVCODARTICULO)) <> "" Then
    '                    .set_TextMatrix(I, C_COLXPRVESTATUS, C_Cancelado)
    '                    PonerColor(I)
    '                End If
    '            Next I
    '        End With

    '    End Sub

    '    ' articulos por proveedor
    '    '*************************************************************/
    '    '*************************************************************/





End Class