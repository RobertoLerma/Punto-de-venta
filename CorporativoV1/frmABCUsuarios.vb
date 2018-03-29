'**********************************************************************************************************************'
'*PROGRAMA: ABC USUARIOS JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmABCUsuarios

    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents _fraSeg_1 As System.Windows.Forms.GroupBox
    Public WithEvents _optTipo_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optTipo_0 As System.Windows.Forms.RadioButton
    Public WithEvents _fraSeg_0 As System.Windows.Forms.GroupBox
    Public WithEvents _lblSeg_0 As System.Windows.Forms.Label
    Public WithEvents _lblSeg_1 As System.Windows.Forms.Label
    Public WithEvents _optTipoUsuario_2 As System.Windows.Forms.RadioButton
    Public WithEvents _optTipoUsuario_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optTipoUsuario_0 As System.Windows.Forms.RadioButton
    Public WithEvents fraTipoUsuario As System.Windows.Forms.GroupBox
    Public WithEvents txtConfirmar As System.Windows.Forms.TextBox
    Public WithEvents txtPassWord As System.Windows.Forms.TextBox
    Public WithEvents _lblSeg_4 As System.Windows.Forms.Label
    Public WithEvents _lblSeg_3 As System.Windows.Forms.Label
    Public WithEvents fraPassWord As System.Windows.Forms.GroupBox
    Public WithEvents __lstUsuarios_1_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
    Public WithEvents _lstUsuarios_1 As System.Windows.Forms.ListView
    Public WithEvents __lstUsuarios_0_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
    Public WithEvents _lstUsuarios_0 As System.Windows.Forms.ListView
    Public WithEvents chkGrupo As System.Windows.Forms.CheckBox
    Public WithEvents txtCodigo As System.Windows.Forms.TextBox
    Public WithEvents txtNombre As System.Windows.Forms.TextBox
    Public WithEvents _btnMoverU_0 As System.Windows.Forms.Button
    Public WithEvents _btnMoverU_1 As System.Windows.Forms.Button
    Public WithEvents _btnMoverU_2 As System.Windows.Forms.Button
    Public WithEvents _btnMoverU_3 As System.Windows.Forms.Button
    Public WithEvents _dbcGrupos_0 As System.Windows.Forms.ComboBox
    Public WithEvents _SSTabSeg_TabPage0 As System.Windows.Forms.TabPage
    Public WithEvents _lblSeg_5 As System.Windows.Forms.Label
    Public WithEvents _lblSeg_6 As System.Windows.Forms.Label
    Public WithEvents _lblSeg_8 As System.Windows.Forms.Label
    Public WithEvents _lblSeg_7 As System.Windows.Forms.Label
    Public WithEvents dbcModulo As System.Windows.Forms.ComboBox
    Public WithEvents dbcUsuarios As System.Windows.Forms.ComboBox
    Public WithEvents _dbcGrupos_1 As System.Windows.Forms.ComboBox
    Public WithEvents __lstPrivilegios_0_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
    Public WithEvents _lstPrivilegios_0 As System.Windows.Forms.ListView
    Public WithEvents __lstPrivilegios_1_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
    Public WithEvents _lstPrivilegios_1 As System.Windows.Forms.ListView
    Public WithEvents _btnMoverP_3 As System.Windows.Forms.Button
    Public WithEvents _btnMoverP_2 As System.Windows.Forms.Button
    Public WithEvents _btnMoverP_1 As System.Windows.Forms.Button
    Public WithEvents _btnMoverP_0 As System.Windows.Forms.Button
    Public WithEvents _SSTabSeg_TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents SSTabSeg As System.Windows.Forms.TabControl
    Public WithEvents btnMoverP As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    Public WithEvents btnMoverU As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    Public WithEvents dbcGrupos As System.Windows.Forms.ComboBox
    Public WithEvents fraSeg As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblSeg As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lstPrivilegios As Microsoft.VisualBasic.Compatibility.VB6.ListViewArray
    Public WithEvents lstUsuarios As Microsoft.VisualBasic.Compatibility.VB6.ListViewArray
    Public WithEvents optTipo As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optTipoUsuario As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray

    'PARA LOS OPTION
    Const nGRUPO As Integer = 0
    Const nUSUARIO As Integer = 1

    'PARA LOS TAB
    Const nGRUPOUSUARIO As Integer = 0
    Const nPRIVILEGIOS As Integer = 1

    Dim mblnSalir As Boolean

    Dim mblnNuevo As Boolean
    Dim mblnCambiosEnCodigo As Boolean

    'Variables para los combos
    Dim mintCodGrupo As Integer
    Dim mintCodUsuario As Integer
    Dim mintCodModulo As Integer
    Dim tecla As Integer
    Dim mblnFueraChange As Boolean
    Dim cTipoUsuario As String
    Dim cTipoUsuarioTag As String
    Public WithEvents btnBuscar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents btnGuardar As Button
    Dim Item As System.Windows.Forms.ListViewItem
    Public strControlActual As String 'Nombre del control actual

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._optTipo_1 = New System.Windows.Forms.RadioButton()
        Me._optTipo_0 = New System.Windows.Forms.RadioButton()
        Me._optTipoUsuario_2 = New System.Windows.Forms.RadioButton()
        Me._optTipoUsuario_1 = New System.Windows.Forms.RadioButton()
        Me._optTipoUsuario_0 = New System.Windows.Forms.RadioButton()
        Me.txtConfirmar = New System.Windows.Forms.TextBox()
        Me.txtPassWord = New System.Windows.Forms.TextBox()
        Me.chkGrupo = New System.Windows.Forms.CheckBox()
        Me._btnMoverU_0 = New System.Windows.Forms.Button()
        Me._btnMoverU_1 = New System.Windows.Forms.Button()
        Me._btnMoverU_2 = New System.Windows.Forms.Button()
        Me._btnMoverU_3 = New System.Windows.Forms.Button()
        Me._btnMoverP_3 = New System.Windows.Forms.Button()
        Me._btnMoverP_2 = New System.Windows.Forms.Button()
        Me._btnMoverP_1 = New System.Windows.Forms.Button()
        Me._btnMoverP_0 = New System.Windows.Forms.Button()
        Me._fraSeg_0 = New System.Windows.Forms.GroupBox()
        Me._fraSeg_1 = New System.Windows.Forms.GroupBox()
        Me.SSTabSeg = New System.Windows.Forms.TabControl()
        Me._SSTabSeg_TabPage0 = New System.Windows.Forms.TabPage()
        Me._lblSeg_0 = New System.Windows.Forms.Label()
        Me._lblSeg_1 = New System.Windows.Forms.Label()
        Me.fraTipoUsuario = New System.Windows.Forms.GroupBox()
        Me.fraPassWord = New System.Windows.Forms.GroupBox()
        Me._lblSeg_4 = New System.Windows.Forms.Label()
        Me._lblSeg_3 = New System.Windows.Forms.Label()
        Me._lstUsuarios_1 = New System.Windows.Forms.ListView()
        Me.__lstUsuarios_1_ColumnHeader_1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me._lstUsuarios_0 = New System.Windows.Forms.ListView()
        Me.__lstUsuarios_0_ColumnHeader_1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.txtCodigo = New System.Windows.Forms.TextBox()
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me._dbcGrupos_0 = New System.Windows.Forms.ComboBox()
        Me._SSTabSeg_TabPage1 = New System.Windows.Forms.TabPage()
        Me._lblSeg_5 = New System.Windows.Forms.Label()
        Me._lblSeg_6 = New System.Windows.Forms.Label()
        Me._lblSeg_8 = New System.Windows.Forms.Label()
        Me._lblSeg_7 = New System.Windows.Forms.Label()
        Me.dbcModulo = New System.Windows.Forms.ComboBox()
        Me.dbcUsuarios = New System.Windows.Forms.ComboBox()
        Me._dbcGrupos_1 = New System.Windows.Forms.ComboBox()
        Me._lstPrivilegios_0 = New System.Windows.Forms.ListView()
        Me.__lstPrivilegios_0_ColumnHeader_1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me._lstPrivilegios_1 = New System.Windows.Forms.ListView()
        Me.__lstPrivilegios_1_ColumnHeader_1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.btnMoverP = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.btnMoverU = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.dbcGrupos = New System.Windows.Forms.ComboBox()
        Me.fraSeg = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblSeg = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lstPrivilegios = New Microsoft.VisualBasic.Compatibility.VB6.ListViewArray(Me.components)
        Me.lstUsuarios = New Microsoft.VisualBasic.Compatibility.VB6.ListViewArray(Me.components)
        Me.optTipo = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optTipoUsuario = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me._fraSeg_0.SuspendLayout()
        Me.SSTabSeg.SuspendLayout()
        Me._SSTabSeg_TabPage0.SuspendLayout()
        Me.fraTipoUsuario.SuspendLayout()
        Me.fraPassWord.SuspendLayout()
        Me._SSTabSeg_TabPage1.SuspendLayout()
        CType(Me.btnMoverP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnMoverU, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraSeg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSeg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lstPrivilegios, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lstUsuarios, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optTipo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optTipoUsuario, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_optTipo_1
        '
        Me._optTipo_1.BackColor = System.Drawing.SystemColors.Control
        Me._optTipo_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipo_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optTipo_1.Location = New System.Drawing.Point(507, 20)
        Me._optTipo_1.Name = "_optTipo_1"
        Me._optTipo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipo_1.Size = New System.Drawing.Size(145, 17)
        Me._optTipo_1.TabIndex = 2
        Me._optTipo_1.TabStop = True
        Me._optTipo_1.Text = "Usuario Independiente"
        Me.ToolTip1.SetToolTip(Me._optTipo_1, "Usuario independiente")
        Me._optTipo_1.UseVisualStyleBackColor = False
        '
        '_optTipo_0
        '
        Me._optTipo_0.BackColor = System.Drawing.SystemColors.Control
        Me._optTipo_0.Checked = True
        Me._optTipo_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipo_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optTipo_0.Location = New System.Drawing.Point(24, 20)
        Me._optTipo_0.Name = "_optTipo_0"
        Me._optTipo_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipo_0.Size = New System.Drawing.Size(145, 17)
        Me._optTipo_0.TabIndex = 1
        Me._optTipo_0.TabStop = True
        Me._optTipo_0.Text = "Grupo de Usuarios"
        Me.ToolTip1.SetToolTip(Me._optTipo_0, "Grupos de Usuarios")
        Me._optTipo_0.UseVisualStyleBackColor = False
        '
        '_optTipoUsuario_2
        '
        Me._optTipoUsuario_2.BackColor = System.Drawing.SystemColors.Control
        Me._optTipoUsuario_2.Checked = True
        Me._optTipoUsuario_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoUsuario_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optTipoUsuario_2.Location = New System.Drawing.Point(296, 16)
        Me._optTipoUsuario_2.Name = "_optTipoUsuario_2"
        Me._optTipoUsuario_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoUsuario_2.Size = New System.Drawing.Size(80, 17)
        Me._optTipoUsuario_2.TabIndex = 14
        Me._optTipoUsuario_2.TabStop = True
        Me._optTipoUsuario_2.Text = "Empleado"
        Me.ToolTip1.SetToolTip(Me._optTipoUsuario_2, "Empleado")
        Me._optTipoUsuario_2.UseVisualStyleBackColor = False
        '
        '_optTipoUsuario_1
        '
        Me._optTipoUsuario_1.BackColor = System.Drawing.SystemColors.Control
        Me._optTipoUsuario_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoUsuario_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optTipoUsuario_1.Location = New System.Drawing.Point(160, 16)
        Me._optTipoUsuario_1.Name = "_optTipoUsuario_1"
        Me._optTipoUsuario_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoUsuario_1.Size = New System.Drawing.Size(78, 17)
        Me._optTipoUsuario_1.TabIndex = 13
        Me._optTipoUsuario_1.TabStop = True
        Me._optTipoUsuario_1.Text = "Supervisor"
        Me.ToolTip1.SetToolTip(Me._optTipoUsuario_1, "Supervisor")
        Me._optTipoUsuario_1.UseVisualStyleBackColor = False
        '
        '_optTipoUsuario_0
        '
        Me._optTipoUsuario_0.BackColor = System.Drawing.SystemColors.Control
        Me._optTipoUsuario_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoUsuario_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optTipoUsuario_0.Location = New System.Drawing.Point(24, 16)
        Me._optTipoUsuario_0.Name = "_optTipoUsuario_0"
        Me._optTipoUsuario_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoUsuario_0.Size = New System.Drawing.Size(89, 17)
        Me._optTipoUsuario_0.TabIndex = 12
        Me._optTipoUsuario_0.TabStop = True
        Me._optTipoUsuario_0.Text = "Administrador"
        Me.ToolTip1.SetToolTip(Me._optTipoUsuario_0, "Administrador del Sistema")
        Me._optTipoUsuario_0.UseVisualStyleBackColor = False
        '
        'txtConfirmar
        '
        Me.txtConfirmar.AcceptsReturn = True
        Me.txtConfirmar.BackColor = System.Drawing.SystemColors.Window
        Me.txtConfirmar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConfirmar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConfirmar.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtConfirmar.Location = New System.Drawing.Point(104, 72)
        Me.txtConfirmar.MaxLength = 0
        Me.txtConfirmar.Name = "txtConfirmar"
        Me.txtConfirmar.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtConfirmar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConfirmar.Size = New System.Drawing.Size(145, 20)
        Me.txtConfirmar.TabIndex = 19
        Me.ToolTip1.SetToolTip(Me.txtConfirmar, "Confirmar Password")
        '
        'txtPassWord
        '
        Me.txtPassWord.AcceptsReturn = True
        Me.txtPassWord.BackColor = System.Drawing.SystemColors.Window
        Me.txtPassWord.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPassWord.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPassWord.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtPassWord.Location = New System.Drawing.Point(104, 32)
        Me.txtPassWord.MaxLength = 0
        Me.txtPassWord.Name = "txtPassWord"
        Me.txtPassWord.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassWord.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPassWord.Size = New System.Drawing.Size(145, 20)
        Me.txtPassWord.TabIndex = 17
        Me.ToolTip1.SetToolTip(Me.txtPassWord, "Establecer Password")
        '
        'chkGrupo
        '
        Me.chkGrupo.BackColor = System.Drawing.SystemColors.Control
        Me.chkGrupo.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkGrupo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkGrupo.Location = New System.Drawing.Point(88, 106)
        Me.chkGrupo.Name = "chkGrupo"
        Me.chkGrupo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkGrupo.Size = New System.Drawing.Size(122, 17)
        Me.chkGrupo.TabIndex = 9
        Me.chkGrupo.Text = "Pertenece al Grupo ..."
        Me.ToolTip1.SetToolTip(Me.chkGrupo, "¿Pertenece a un grupo el Usuario?")
        Me.chkGrupo.UseVisualStyleBackColor = False
        '
        '_btnMoverU_0
        '
        Me._btnMoverU_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnMoverU_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnMoverU_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnMoverU_0.Location = New System.Drawing.Point(456, 168)
        Me._btnMoverU_0.Name = "_btnMoverU_0"
        Me._btnMoverU_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnMoverU_0.Size = New System.Drawing.Size(33, 21)
        Me._btnMoverU_0.TabIndex = 21
        Me._btnMoverU_0.Text = "<<"
        Me.ToolTip1.SetToolTip(Me._btnMoverU_0, "Mover Usuarios")
        Me._btnMoverU_0.UseVisualStyleBackColor = False
        '
        '_btnMoverU_1
        '
        Me._btnMoverU_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnMoverU_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnMoverU_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnMoverU_1.Location = New System.Drawing.Point(456, 200)
        Me._btnMoverU_1.Name = "_btnMoverU_1"
        Me._btnMoverU_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnMoverU_1.Size = New System.Drawing.Size(33, 21)
        Me._btnMoverU_1.TabIndex = 22
        Me._btnMoverU_1.Text = "<"
        Me.ToolTip1.SetToolTip(Me._btnMoverU_1, "Mover Usuarios")
        Me._btnMoverU_1.UseVisualStyleBackColor = False
        '
        '_btnMoverU_2
        '
        Me._btnMoverU_2.BackColor = System.Drawing.SystemColors.Control
        Me._btnMoverU_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnMoverU_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnMoverU_2.Location = New System.Drawing.Point(456, 232)
        Me._btnMoverU_2.Name = "_btnMoverU_2"
        Me._btnMoverU_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnMoverU_2.Size = New System.Drawing.Size(33, 21)
        Me._btnMoverU_2.TabIndex = 23
        Me._btnMoverU_2.Text = ">"
        Me.ToolTip1.SetToolTip(Me._btnMoverU_2, "Mover Usuarios")
        Me._btnMoverU_2.UseVisualStyleBackColor = False
        '
        '_btnMoverU_3
        '
        Me._btnMoverU_3.BackColor = System.Drawing.SystemColors.Control
        Me._btnMoverU_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnMoverU_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnMoverU_3.Location = New System.Drawing.Point(456, 264)
        Me._btnMoverU_3.Name = "_btnMoverU_3"
        Me._btnMoverU_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnMoverU_3.Size = New System.Drawing.Size(33, 21)
        Me._btnMoverU_3.TabIndex = 24
        Me._btnMoverU_3.Text = ">>"
        Me.ToolTip1.SetToolTip(Me._btnMoverU_3, "Mover Usuarios")
        Me._btnMoverU_3.UseVisualStyleBackColor = False
        '
        '_btnMoverP_3
        '
        Me._btnMoverP_3.BackColor = System.Drawing.SystemColors.Control
        Me._btnMoverP_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnMoverP_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnMoverP_3.Location = New System.Drawing.Point(456, 264)
        Me._btnMoverP_3.Name = "_btnMoverP_3"
        Me._btnMoverP_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnMoverP_3.Size = New System.Drawing.Size(33, 21)
        Me._btnMoverP_3.TabIndex = 36
        Me._btnMoverP_3.Text = ">>"
        Me.ToolTip1.SetToolTip(Me._btnMoverP_3, "Mover Funciones")
        Me._btnMoverP_3.UseVisualStyleBackColor = False
        '
        '_btnMoverP_2
        '
        Me._btnMoverP_2.BackColor = System.Drawing.SystemColors.Control
        Me._btnMoverP_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnMoverP_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnMoverP_2.Location = New System.Drawing.Point(456, 232)
        Me._btnMoverP_2.Name = "_btnMoverP_2"
        Me._btnMoverP_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnMoverP_2.Size = New System.Drawing.Size(33, 21)
        Me._btnMoverP_2.TabIndex = 35
        Me._btnMoverP_2.Text = ">"
        Me.ToolTip1.SetToolTip(Me._btnMoverP_2, "Mover Funciones")
        Me._btnMoverP_2.UseVisualStyleBackColor = False
        '
        '_btnMoverP_1
        '
        Me._btnMoverP_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnMoverP_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnMoverP_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnMoverP_1.Location = New System.Drawing.Point(456, 200)
        Me._btnMoverP_1.Name = "_btnMoverP_1"
        Me._btnMoverP_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnMoverP_1.Size = New System.Drawing.Size(33, 21)
        Me._btnMoverP_1.TabIndex = 34
        Me._btnMoverP_1.Text = "<"
        Me.ToolTip1.SetToolTip(Me._btnMoverP_1, "Mover Funciones")
        Me._btnMoverP_1.UseVisualStyleBackColor = False
        '
        '_btnMoverP_0
        '
        Me._btnMoverP_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnMoverP_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnMoverP_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnMoverP_0.Location = New System.Drawing.Point(456, 168)
        Me._btnMoverP_0.Name = "_btnMoverP_0"
        Me._btnMoverP_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnMoverP_0.Size = New System.Drawing.Size(33, 21)
        Me._btnMoverP_0.TabIndex = 33
        Me._btnMoverP_0.Text = "<<"
        Me.ToolTip1.SetToolTip(Me._btnMoverP_0, "Mover Funciones")
        Me._btnMoverP_0.UseVisualStyleBackColor = False
        '
        '_fraSeg_0
        '
        Me._fraSeg_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraSeg_0.Controls.Add(Me._fraSeg_1)
        Me._fraSeg_0.Controls.Add(Me._optTipo_1)
        Me._fraSeg_0.Controls.Add(Me._optTipo_0)
        Me._fraSeg_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraSeg_0.Location = New System.Drawing.Point(8, 8)
        Me._fraSeg_0.Name = "_fraSeg_0"
        Me._fraSeg_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraSeg_0.Size = New System.Drawing.Size(945, 49)
        Me._fraSeg_0.TabIndex = 0
        Me._fraSeg_0.TabStop = False
        Me._fraSeg_0.Text = "Tipo de Usuario ..."
        '
        '_fraSeg_1
        '
        Me._fraSeg_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraSeg_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraSeg_1.Location = New System.Drawing.Point(472, 8)
        Me._fraSeg_1.Name = "_fraSeg_1"
        Me._fraSeg_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraSeg_1.Size = New System.Drawing.Size(2, 33)
        Me._fraSeg_1.TabIndex = 3
        Me._fraSeg_1.TabStop = False
        '
        'SSTabSeg
        '
        Me.SSTabSeg.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.SSTabSeg.Controls.Add(Me._SSTabSeg_TabPage0)
        Me.SSTabSeg.Controls.Add(Me._SSTabSeg_TabPage1)
        Me.SSTabSeg.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTabSeg.Location = New System.Drawing.Point(8, 64)
        Me.SSTabSeg.Name = "SSTabSeg"
        Me.SSTabSeg.SelectedIndex = 0
        Me.SSTabSeg.Size = New System.Drawing.Size(945, 329)
        Me.SSTabSeg.TabIndex = 4
        '
        '_SSTabSeg_TabPage0
        '
        Me._SSTabSeg_TabPage0.Controls.Add(Me._lblSeg_0)
        Me._SSTabSeg_TabPage0.Controls.Add(Me._lblSeg_1)
        Me._SSTabSeg_TabPage0.Controls.Add(Me.fraTipoUsuario)
        Me._SSTabSeg_TabPage0.Controls.Add(Me.fraPassWord)
        Me._SSTabSeg_TabPage0.Controls.Add(Me._lstUsuarios_1)
        Me._SSTabSeg_TabPage0.Controls.Add(Me._lstUsuarios_0)
        Me._SSTabSeg_TabPage0.Controls.Add(Me.chkGrupo)
        Me._SSTabSeg_TabPage0.Controls.Add(Me.txtCodigo)
        Me._SSTabSeg_TabPage0.Controls.Add(Me.txtNombre)
        Me._SSTabSeg_TabPage0.Controls.Add(Me._btnMoverU_0)
        Me._SSTabSeg_TabPage0.Controls.Add(Me._btnMoverU_1)
        Me._SSTabSeg_TabPage0.Controls.Add(Me._btnMoverU_2)
        Me._SSTabSeg_TabPage0.Controls.Add(Me._btnMoverU_3)
        Me._SSTabSeg_TabPage0.Controls.Add(Me._dbcGrupos_0)
        Me._SSTabSeg_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._SSTabSeg_TabPage0.Name = "_SSTabSeg_TabPage0"
        Me._SSTabSeg_TabPage0.Size = New System.Drawing.Size(937, 303)
        Me._SSTabSeg_TabPage0.TabIndex = 0
        Me._SSTabSeg_TabPage0.Text = "Grupos y Usuarios"
        '
        '_lblSeg_0
        '
        Me._lblSeg_0.AutoSize = True
        Me._lblSeg_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblSeg_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSeg_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSeg_0.Location = New System.Drawing.Point(32, 44)
        Me._lblSeg_0.Name = "_lblSeg_0"
        Me._lblSeg_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSeg_0.Size = New System.Drawing.Size(40, 13)
        Me._lblSeg_0.TabIndex = 5
        Me._lblSeg_0.Text = "Código"
        '
        '_lblSeg_1
        '
        Me._lblSeg_1.AutoSize = True
        Me._lblSeg_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblSeg_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSeg_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSeg_1.Location = New System.Drawing.Point(32, 76)
        Me._lblSeg_1.Name = "_lblSeg_1"
        Me._lblSeg_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSeg_1.Size = New System.Drawing.Size(44, 13)
        Me._lblSeg_1.TabIndex = 7
        Me._lblSeg_1.Text = "Nombre"
        '
        'fraTipoUsuario
        '
        Me.fraTipoUsuario.BackColor = System.Drawing.SystemColors.Control
        Me.fraTipoUsuario.Controls.Add(Me._optTipoUsuario_2)
        Me.fraTipoUsuario.Controls.Add(Me._optTipoUsuario_1)
        Me.fraTipoUsuario.Controls.Add(Me._optTipoUsuario_0)
        Me.fraTipoUsuario.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraTipoUsuario.Location = New System.Drawing.Point(16, 144)
        Me.fraTipoUsuario.Name = "fraTipoUsuario"
        Me.fraTipoUsuario.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTipoUsuario.Size = New System.Drawing.Size(418, 41)
        Me.fraTipoUsuario.TabIndex = 11
        Me.fraTipoUsuario.TabStop = False
        Me.fraTipoUsuario.Text = "Tipo de Usuario"
        '
        'fraPassWord
        '
        Me.fraPassWord.BackColor = System.Drawing.SystemColors.Control
        Me.fraPassWord.Controls.Add(Me.txtConfirmar)
        Me.fraPassWord.Controls.Add(Me.txtPassWord)
        Me.fraPassWord.Controls.Add(Me._lblSeg_4)
        Me.fraPassWord.Controls.Add(Me._lblSeg_3)
        Me.fraPassWord.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraPassWord.Location = New System.Drawing.Point(16, 187)
        Me.fraPassWord.Name = "fraPassWord"
        Me.fraPassWord.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraPassWord.Size = New System.Drawing.Size(273, 113)
        Me.fraPassWord.TabIndex = 15
        Me.fraPassWord.TabStop = False
        Me.fraPassWord.Text = "Proporcione la clave de acceso ..."
        '
        '_lblSeg_4
        '
        Me._lblSeg_4.AutoSize = True
        Me._lblSeg_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblSeg_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSeg_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSeg_4.Location = New System.Drawing.Point(24, 76)
        Me._lblSeg_4.Name = "_lblSeg_4"
        Me._lblSeg_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSeg_4.Size = New System.Drawing.Size(51, 13)
        Me._lblSeg_4.TabIndex = 18
        Me._lblSeg_4.Text = "Confirmar"
        '
        '_lblSeg_3
        '
        Me._lblSeg_3.AutoSize = True
        Me._lblSeg_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblSeg_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSeg_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSeg_3.Location = New System.Drawing.Point(24, 36)
        Me._lblSeg_3.Name = "_lblSeg_3"
        Me._lblSeg_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSeg_3.Size = New System.Drawing.Size(61, 13)
        Me._lblSeg_3.TabIndex = 16
        Me._lblSeg_3.Text = "Contraseña"
        '
        '_lstUsuarios_1
        '
        Me._lstUsuarios_1.BackColor = System.Drawing.SystemColors.Window
        Me._lstUsuarios_1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.__lstUsuarios_1_ColumnHeader_1})
        Me._lstUsuarios_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lstUsuarios_1.FullRowSelect = True
        Me._lstUsuarios_1.GridLines = True
        Me._lstUsuarios_1.LabelWrap = False
        Me._lstUsuarios_1.Location = New System.Drawing.Point(497, 144)
        Me._lstUsuarios_1.Name = "_lstUsuarios_1"
        Me._lstUsuarios_1.Size = New System.Drawing.Size(433, 169)
        Me._lstUsuarios_1.TabIndex = 25
        Me._lstUsuarios_1.UseCompatibleStateImageBehavior = False
        Me._lstUsuarios_1.View = System.Windows.Forms.View.Details
        '
        '__lstUsuarios_1_ColumnHeader_1
        '
        Me.__lstUsuarios_1_ColumnHeader_1.Text = "Usuarios Independientes"
        Me.__lstUsuarios_1_ColumnHeader_1.Width = 753
        '
        '_lstUsuarios_0
        '
        Me._lstUsuarios_0.BackColor = System.Drawing.SystemColors.Window
        Me._lstUsuarios_0.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.__lstUsuarios_0_ColumnHeader_1})
        Me._lstUsuarios_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lstUsuarios_0.FullRowSelect = True
        Me._lstUsuarios_0.GridLines = True
        Me._lstUsuarios_0.LabelWrap = False
        Me._lstUsuarios_0.Location = New System.Drawing.Point(16, 144)
        Me._lstUsuarios_0.Name = "_lstUsuarios_0"
        Me._lstUsuarios_0.Size = New System.Drawing.Size(433, 169)
        Me._lstUsuarios_0.TabIndex = 20
        Me._lstUsuarios_0.UseCompatibleStateImageBehavior = False
        Me._lstUsuarios_0.View = System.Windows.Forms.View.Details
        '
        '__lstUsuarios_0_ColumnHeader_1
        '
        Me.__lstUsuarios_0_ColumnHeader_1.Text = "Usuarios del Grupo"
        Me.__lstUsuarios_0_ColumnHeader_1.Width = 753
        '
        'txtCodigo
        '
        Me.txtCodigo.AcceptsReturn = True
        Me.txtCodigo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodigo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodigo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodigo.Location = New System.Drawing.Point(88, 40)
        Me.txtCodigo.MaxLength = 0
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodigo.Size = New System.Drawing.Size(57, 20)
        Me.txtCodigo.TabIndex = 6
        '
        'txtNombre
        '
        Me.txtNombre.AcceptsReturn = True
        Me.txtNombre.BackColor = System.Drawing.SystemColors.Window
        Me.txtNombre.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombre.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombre.Location = New System.Drawing.Point(88, 72)
        Me.txtNombre.MaxLength = 30
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombre.Size = New System.Drawing.Size(361, 20)
        Me.txtNombre.TabIndex = 8
        '
        '_dbcGrupos_0
        '
        Me._dbcGrupos_0.Location = New System.Drawing.Point(216, 104)
        Me._dbcGrupos_0.Name = "_dbcGrupos_0"
        Me._dbcGrupos_0.Size = New System.Drawing.Size(233, 21)
        Me._dbcGrupos_0.TabIndex = 10
        '
        '_SSTabSeg_TabPage1
        '
        Me._SSTabSeg_TabPage1.Controls.Add(Me._lblSeg_5)
        Me._SSTabSeg_TabPage1.Controls.Add(Me._lblSeg_6)
        Me._SSTabSeg_TabPage1.Controls.Add(Me._lblSeg_8)
        Me._SSTabSeg_TabPage1.Controls.Add(Me._lblSeg_7)
        Me._SSTabSeg_TabPage1.Controls.Add(Me.dbcModulo)
        Me._SSTabSeg_TabPage1.Controls.Add(Me.dbcUsuarios)
        Me._SSTabSeg_TabPage1.Controls.Add(Me._dbcGrupos_1)
        Me._SSTabSeg_TabPage1.Controls.Add(Me._lstPrivilegios_0)
        Me._SSTabSeg_TabPage1.Controls.Add(Me._lstPrivilegios_1)
        Me._SSTabSeg_TabPage1.Controls.Add(Me._btnMoverP_3)
        Me._SSTabSeg_TabPage1.Controls.Add(Me._btnMoverP_2)
        Me._SSTabSeg_TabPage1.Controls.Add(Me._btnMoverP_1)
        Me._SSTabSeg_TabPage1.Controls.Add(Me._btnMoverP_0)
        Me._SSTabSeg_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._SSTabSeg_TabPage1.Name = "_SSTabSeg_TabPage1"
        Me._SSTabSeg_TabPage1.Size = New System.Drawing.Size(937, 303)
        Me._SSTabSeg_TabPage1.TabIndex = 1
        Me._SSTabSeg_TabPage1.Text = "Configuración de Privilegios"
        '
        '_lblSeg_5
        '
        Me._lblSeg_5.AutoSize = True
        Me._lblSeg_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblSeg_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSeg_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSeg_5.Location = New System.Drawing.Point(32, 44)
        Me._lblSeg_5.Name = "_lblSeg_5"
        Me._lblSeg_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSeg_5.Size = New System.Drawing.Size(36, 13)
        Me._lblSeg_5.TabIndex = 26
        Me._lblSeg_5.Text = "Grupo"
        '
        '_lblSeg_6
        '
        Me._lblSeg_6.AutoSize = True
        Me._lblSeg_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblSeg_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSeg_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSeg_6.Location = New System.Drawing.Point(32, 44)
        Me._lblSeg_6.Name = "_lblSeg_6"
        Me._lblSeg_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSeg_6.Size = New System.Drawing.Size(43, 13)
        Me._lblSeg_6.TabIndex = 28
        Me._lblSeg_6.Text = "Usuario"
        '
        '_lblSeg_8
        '
        Me._lblSeg_8.AutoSize = True
        Me._lblSeg_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblSeg_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSeg_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSeg_8.Location = New System.Drawing.Point(32, 76)
        Me._lblSeg_8.Name = "_lblSeg_8"
        Me._lblSeg_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSeg_8.Size = New System.Drawing.Size(42, 13)
        Me._lblSeg_8.TabIndex = 30
        Me._lblSeg_8.Text = "Módulo"
        '
        '_lblSeg_7
        '
        Me._lblSeg_7.AutoSize = True
        Me._lblSeg_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblSeg_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSeg_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSeg_7.Location = New System.Drawing.Point(354, 44)
        Me._lblSeg_7.Name = "_lblSeg_7"
        Me._lblSeg_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSeg_7.Size = New System.Drawing.Size(167, 13)
        Me._lblSeg_7.TabIndex = 38
        Me._lblSeg_7.Text = "Grupo al que pertenece el usuario"
        '
        'dbcModulo
        '
        Me.dbcModulo.Location = New System.Drawing.Point(80, 72)
        Me.dbcModulo.Name = "dbcModulo"
        Me.dbcModulo.Size = New System.Drawing.Size(259, 21)
        Me.dbcModulo.TabIndex = 31
        '
        'dbcUsuarios
        '
        Me.dbcUsuarios.Location = New System.Drawing.Point(80, 40)
        Me.dbcUsuarios.Name = "dbcUsuarios"
        Me.dbcUsuarios.Size = New System.Drawing.Size(193, 21)
        Me.dbcUsuarios.TabIndex = 29
        '
        '_dbcGrupos_1
        '
        Me._dbcGrupos_1.Location = New System.Drawing.Point(80, 40)
        Me._dbcGrupos_1.Name = "_dbcGrupos_1"
        Me._dbcGrupos_1.Size = New System.Drawing.Size(259, 21)
        Me._dbcGrupos_1.TabIndex = 27
        '
        '_lstPrivilegios_0
        '
        Me._lstPrivilegios_0.BackColor = System.Drawing.SystemColors.Window
        Me._lstPrivilegios_0.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.__lstPrivilegios_0_ColumnHeader_1})
        Me._lstPrivilegios_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lstPrivilegios_0.FullRowSelect = True
        Me._lstPrivilegios_0.GridLines = True
        Me._lstPrivilegios_0.LabelWrap = False
        Me._lstPrivilegios_0.Location = New System.Drawing.Point(16, 144)
        Me._lstPrivilegios_0.Name = "_lstPrivilegios_0"
        Me._lstPrivilegios_0.Size = New System.Drawing.Size(434, 141)
        Me._lstPrivilegios_0.TabIndex = 32
        Me._lstPrivilegios_0.UseCompatibleStateImageBehavior = False
        Me._lstPrivilegios_0.View = System.Windows.Forms.View.Details
        '
        '__lstPrivilegios_0_ColumnHeader_1
        '
        Me.__lstPrivilegios_0_ColumnHeader_1.Text = "Funciones Habilitadas"
        Me.__lstPrivilegios_0_ColumnHeader_1.Width = 753
        '
        '_lstPrivilegios_1
        '
        Me._lstPrivilegios_1.BackColor = System.Drawing.SystemColors.Window
        Me._lstPrivilegios_1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.__lstPrivilegios_1_ColumnHeader_1})
        Me._lstPrivilegios_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lstPrivilegios_1.FullRowSelect = True
        Me._lstPrivilegios_1.GridLines = True
        Me._lstPrivilegios_1.LabelWrap = False
        Me._lstPrivilegios_1.Location = New System.Drawing.Point(497, 144)
        Me._lstPrivilegios_1.Name = "_lstPrivilegios_1"
        Me._lstPrivilegios_1.Size = New System.Drawing.Size(433, 141)
        Me._lstPrivilegios_1.TabIndex = 37
        Me._lstPrivilegios_1.UseCompatibleStateImageBehavior = False
        Me._lstPrivilegios_1.View = System.Windows.Forms.View.Details
        '
        '__lstPrivilegios_1_ColumnHeader_1
        '
        Me.__lstPrivilegios_1_ColumnHeader_1.Text = "Funciones Denegadas"
        Me.__lstPrivilegios_1_ColumnHeader_1.Width = 753
        '
        'btnMoverP
        '
        '
        'btnMoverU
        '
        '
        'dbcGrupos
        '
        Me.dbcGrupos.Location = New System.Drawing.Point(0, 0)
        Me.dbcGrupos.Name = "dbcGrupos"
        Me.dbcGrupos.Size = New System.Drawing.Size(121, 21)
        Me.dbcGrupos.TabIndex = 0
        '
        'lstPrivilegios
        '
        '
        'lstUsuarios
        '
        '
        'optTipo
        '
        '
        'optTipoUsuario
        '
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(588, 444)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(93, 35)
        Me.btnBuscar.TabIndex = 71
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Location = New System.Drawing.Point(488, 444)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(93, 35)
        Me.btnLimpiar.TabIndex = 70
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(389, 444)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(93, 35)
        Me.btnEliminar.TabIndex = 69
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.Location = New System.Drawing.Point(290, 444)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(93, 35)
        Me.btnGuardar.TabIndex = 68
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'frmABCUsuarios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(963, 491)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me._fraSeg_0)
        Me.Controls.Add(Me.SSTabSeg)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(15, 228)
        Me.MaximizeBox = False
        Me.Name = "frmABCUsuarios"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Configuración de Grupos, Usuarios y Privilegios"
        Me._fraSeg_0.ResumeLayout(False)
        Me.SSTabSeg.ResumeLayout(False)
        Me._SSTabSeg_TabPage0.ResumeLayout(False)
        Me._SSTabSeg_TabPage0.PerformLayout()
        Me.fraTipoUsuario.ResumeLayout(False)
        Me.fraPassWord.ResumeLayout(False)
        Me.fraPassWord.PerformLayout()
        Me._SSTabSeg_TabPage1.ResumeLayout(False)
        Me._SSTabSeg_TabPage1.PerformLayout()
        CType(Me.btnMoverP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnMoverU, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraSeg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSeg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lstPrivilegios, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lstUsuarios, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optTipo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optTipoUsuario, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


    Public Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendrá el string del tag que se le mandara al fromulario de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas

        'strControlActual = UCase(txtNombre.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        'If UCase(Me.txtCodigo.Name) <> "TXTCODIGO" And UCase(Me.txtNombre.Name) <> "TXTNOMBRE" Then
        '    Exit Sub
        'End If

        'If (txtCodigo.Text = "") Then
        '    strControlActual = UCase(txtCodigo.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        '    strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
        'ElseIf (txtnombre.Text = "") Then
        '    strControlActual = UCase(txtNombre.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        '    strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
        'End If


        If Me._optTipo_0.Checked Then
            strCaptionForm = "Consulta de Grupos de Usuarios"
        ElseIf Me._optTipo_1.Checked Then
            strCaptionForm = "Consulta de Usuarios"
        End If

        Select Case strControlActual
            Case "TXTCODIGO"
                If Me._optTipo_0.Checked Then
                    strCaptionForm = "Consulta de Grupos de Usuarios"
                    gStrSql = "SELECT RIGHT('00'+LTRIM(CodUsuario),2) AS CODIGO, Nombre AS DESCRIPCION FROM CatUsuarios WHERE Grupo = 1 ORDER BY CodUsuario"
                Else
                    strCaptionForm = "Consulta de Usuarios"
                    gStrSql = "SELECT RIGHT('000'+LTRIM(CodUsuario),3) AS CODIGO, Nombre AS NOMBRE FROM CatUsuarios WHERE Grupo = 0 ORDER BY CodUsuario"
                End If
            Case "TXTNOMBRE"
                If Me._optTipo_0.Checked Then
                    strCaptionForm = "Consulta de Grupos de Usuarios"
                    gStrSql = "SELECT Nombre AS DESCRIPCION, RIGHT('00'+LTRIM(CodUsuario),2) AS CODIGO FROM CatUsuarios WHERE Grupo = 1 ORDER BY Nombre"
                Else
                    strCaptionForm = "Consulta de Usuarios"
                    gStrSql = "SELECT Nombre AS NOMBRE, RIGHT('000'+LTRIM(CodUsuario),3) AS CODIGO FROM CatUsuarios WHERE Grupo = 0 ORDER BY Nombre"
                End If
            Case Else
                'Sale de este sub para que no ejecute ninguna opción
                Exit Sub
        End Select

        strSQL = gStrSql 'Se hace uso de una variable temporal para el query

        'Si hubo cambios y es una modificacion entonces preguntará si desea grabar los cambios
        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se cargue la consulta
                Case MsgBoxResult.Cancel 'Cancela la consulta
                    Exit Sub
            End Select
        End If

        gStrSql = strSQL 'Se regresa el valor de la variable temporal a la variable original

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsjNoExiste(C_msgSINDATOS, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta
        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODIGO"
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                Case "TXTNOMBRE"
                    .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
                    .set_ColWidth(1, 0, 900) 'Columna del Código
            End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function Cambios() As Boolean
        If Me._optTipo_0.Checked Then
            Select Case Me.SSTabSeg.SelectedIndex
                Case nGRUPOUSUARIO
                    If Trim(Me.txtNombre.Text) <> Trim(Me.txtNombre.Tag) Then
                        Cambios = True
                        Exit Function
                    End If
            End Select
        ElseIf Me._optTipo_1.Checked Then
            Select Case Me.SSTabSeg.SelectedIndex
                Case nGRUPOUSUARIO
                    If Trim(Me.txtNombre.Text) <> Trim(Me.txtNombre.Tag) Then
                        Cambios = True
                        Exit Function
                    End If
                    If Trim(Me._dbcGrupos_0.Text) <> Trim(Me._dbcGrupos_0.Tag) Then
                        Cambios = True
                        Exit Function
                    End If
            End Select
        End If
    End Function

    Public Function Guardar() As Boolean
        On Error GoTo Merr
        Dim I As Integer
        Dim rsLocal As ADODB.Recordset
        Dim blnTransaction As Boolean
        'Valida si todos los datos han sido llenados correctamente para poder ser guardados
        If Not ValidaDatos() Then
            If Me.SSTabSeg.SelectedIndex = nGRUPOUSUARIO Then
                If Trim(UCase(Me.ActiveControl.Name)) = "TXTNOMBRE" Then
                    mblnNuevo = True
                End If
            End If
            Exit Function
        End If

        If mblnNuevo Then
            'Hace una consulta para ver si ya existe el nombre de usuario o grupo
            gStrSql = "SELECT * FROM CatUsuarios WHERE Nombre = '" & Trim(Me.txtNombre.Text) & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            rsLocal = Cmd.Execute
            If rsLocal.RecordCount > 0 Then
                If Me._optTipo_0.Checked Then
                    MsgBox("Debe especificar un Nombre de Grupo distinto", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Me.txtNombre.Focus()
                    ModEstandar.SelTextoTxt((Me.txtNombre))
                    Guardar = False
                    Exit Function
                Else
                    MsgBox("Debe especificar un nombre de usuario distinto", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Me.txtNombre.Focus()
                    ModEstandar.SelTextoTxt((Me.txtNombre))
                    Guardar = False
                    Exit Function
                End If
            End If
        End If

        Cnn.BeginTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        blnTransaction = True
        If mblnNuevo Then
            'Determina si es un grupo o un usuario
            If Me._optTipo_0.Checked Then
                ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), "", CStr(True), CStr(0), cTipoUsuario, CStr(0), C_INSERCION, CStr(0))
                Cmd.Execute()
                Me.txtCodigo.Text = Format(Cmd.Parameters("ID").Value, "00")
            Else
                ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), Trim(ModEncriptacion.Encriptar((Me.txtPassWord.Text))), CStr(False), Str(mintCodGrupo), cTipoUsuario, CStr(0), C_INSERCION, CStr(0))
                Cmd.Execute()
                Me.txtCodigo.Text = Format(Cmd.Parameters("ID").Value, "000")
                If mintCodGrupo <> 0 Then
                    ' Asignar los derechos heredados
                    ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), "", CStr(False), CStr(mintCodGrupo), cTipoUsuario, CStr(0), C_MODIFICACION, CStr(5))
                    Cmd.Execute()
                End If
            End If
        Else
            If Me._optTipo_0.Checked Then
                If Me.SSTabSeg.SelectedIndex = nGRUPOUSUARIO Then
                    '0 - Modificar los datos Generales del Grupo
                    '1 - Borrar todos los derechos de los usuarios que hayan pertencecido al grupo
                    '    y que se encuentren ahora en la parte derecha del lstUsuarios
                    '2 - Borrar todos los derechos de los usuarios que pertenecen al Grupo y que estén en la
                    '    parte izquierda del lstUsuarios
                    '3 - Poner en cero el CodGrupo de los usuarios de la parte derecha del lstUsuarios
                    '4 - Poner el número de Grupo a los usuarios de la parte izquierda del lstUsuarios
                    '5 - Asignarle los derechos del grupo a los usuarios que pertenezcan al Grupo
                    ' Paso 0
                    ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), CStr(ModEncriptacion.Encriptar("")), CStr(True), CStr(0), cTipoUsuario, CStr(0), C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                    ' Pasos 1 y 2
                    ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), CStr(ModEncriptacion.Encriptar("")), CStr(True), CStr(0), cTipoUsuario, CStr(0), C_MODIFICACION, CStr(2))
                    Cmd.Execute()
                    ' Borrar los derechos de los usuarios que se encuentren en la parte izquierda y que no pertenezcan a ningún grupo
                    For I = 1 To Me._lstUsuarios_0.Items.Count - 1
                        gStrSql = "Select CodUsuario, Nombre From CatUsuarios Where LTrim(RTrim(Nombre)) = '" & Trim(Me._lstUsuarios_0.Items.Item(I).Text) & "' and CodGrupo = 0"
                        ModEstandar.BorraCmd()
                        Cmd.CommandText = "dbo.UP_Select_Datos"
                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                        rsLocal = Cmd.Execute
                        If rsLocal.RecordCount > 0 Then
                            ModStoredProcedures.PR_IEAccesos(CStr(rsLocal.Fields("CodUsuario").Value), "", "", C_ELIMINACION, CStr(0))
                            Cmd.Execute()
                        End If
                    Next
                    ' Paso 3 - Poner en cero el CodGrupo de los usuarios de la parte derecha del lstUsuarios
                    ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), CStr(ModEncriptacion.Encriptar("")), CStr(True), CStr(0), cTipoUsuario, CStr(0), C_MODIFICACION, CStr(3))
                    Cmd.Execute()
                    ' Paso 4 - Poner el número de Grupo a los usuarios de la parte izquierda del lstUsuarios
                    For I = 1 To Me._lstUsuarios_0.Items.Count - 1
                        gStrSql = "Select CodUsuario, Nombre From CatUsuarios Where LTrim(RTrim(Nombre)) = '" & Trim(Me._lstUsuarios_0.Items.Item(I).Text) & "'"
                        ModEstandar.BorraCmd()
                        Cmd.CommandText = "dbo.UP_Select_Datos"
                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                        rsLocal = Cmd.Execute
                        If rsLocal.RecordCount > 0 Then
                            ModStoredProcedures.PR_IMECatUsuarios(CStr(rsLocal.Fields("CodUsuario").Value), Trim(Me.txtNombre.Text), CStr(ModEncriptacion.Encriptar("")), CStr(False), Trim(Me.txtCodigo.Text), cTipoUsuario, CStr(0), C_MODIFICACION, CStr(4))
                            Cmd.Execute()
                            'Paso 5 - Asignarle los derechos del grupo al usuario que pertenezca al Grupo
                            ModStoredProcedures.PR_IMECatUsuarios(CStr(rsLocal.Fields("CodUsuario").Value), Trim(Me.txtNombre.Text), CStr(ModEncriptacion.Encriptar("")), CStr(False), Trim(Me.txtCodigo.Text), cTipoUsuario, CStr(0), C_MODIFICACION, CStr(5))
                            Cmd.Execute()
                        End If
                    Next
                Else 'Se trata de configurar un grupo pero en la página de privilegios
                    '1 - Borrar todos los derechos del grupo
                    '2 - Tomar a todos los usuarios que pertenezcan al grupo les borra todos los derechos, después
                    '3 - Configurar los derechos del grupo y
                    '4 - Configurar a los usuarios del grupo, los mismos derechos de éste

                    'Pasos 1 y 2 - Procurar que sean solo los privilegios del módulo seleccionado
                    ModStoredProcedures.PR_IMECatUsuarios(CStr(mintCodGrupo), Trim(Me._dbcGrupos_1.Text), "", CStr(True), CStr(0), cTipoUsuario, CStr(mintCodModulo), C_MODIFICACION, CStr(1))
                    Cmd.Execute()
                    'Paso 3  - Configurar los derechos del grupo (Aquí se utiliza el procedimiento almacenado UP_IE_Accesos)
                    For I = 1 To Me._lstPrivilegios_0.Items.Count - 1
                        gStrSql = "Select CodModulo, CodFuncion, DescFuncion, Forma From CatFunciones Where CodModulo = " & mintCodModulo & " and LTrim(RTrim(DescFuncion)) = '" & Trim(Me._lstPrivilegios_0.Items.Item(I).Text) & "'"
                        ModEstandar.BorraCmd()
                        Cmd.CommandText = "dbo.UP_Select_Datos"
                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                        rsLocal = Cmd.Execute
                        If rsLocal.RecordCount > 0 Then
                            ModStoredProcedures.PR_IEAccesos(CStr(mintCodGrupo), Trim(rsLocal.Fields("Forma").Value), CStr(rsLocal.Fields("CodModulo").Value), C_INSERCION, CStr(0))
                            Cmd.Execute()
                        End If
                    Next
                    'Paso 4 - Asignar a los usuarios del grupo, los mismos derechos de éste
                    gStrSql = "Select CodUsuario, Nombre From CatUsuarios Where CodGrupo = " & mintCodGrupo
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.UP_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                    rsLocal = Cmd.Execute
                    If rsLocal.RecordCount > 0 Then
                        rsLocal.MoveFirst()
                        For I = 1 To rsLocal.RecordCount
                            ModStoredProcedures.PR_IMECatUsuarios(CStr(rsLocal.Fields("CodUsuario").Value), Trim(rsLocal.Fields("Nombre").Value), CStr(ModEncriptacion.Encriptar("")), CStr(False), Trim(CStr(mintCodGrupo)), cTipoUsuario, CStr(mintCodModulo), C_MODIFICACION, CStr(6))
                            Cmd.Execute()
                            rsLocal.MoveNext()
                        Next
                    End If
                End If
            Else 'Se trata de configurar a un usuario en particular
                If Me.SSTabSeg.SelectedIndex = nGRUPOUSUARIO Then
                    'Sólo se modifican los datos generales del usuario, tipo y contraseña
                    'Pero si el usuario queda sin grupo o se le asigna a un grupo diferente, se le quitan los
                    'privilegios que había heredado del grupo
                    If Trim(Me._dbcGrupos_0.Text) = Trim(Me._dbcGrupos_0.Tag) Then
                        ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), CStr(ModEncriptacion.Encriptar((Me.txtPassWord.Text))), CStr(False), CStr(mintCodGrupo), cTipoUsuario, CStr(0), C_MODIFICACION, CStr(0))
                        Cmd.Execute()
                    Else 'El usuario queda sin grupo o se le asigna a un grupo diferente, por lo que pierde la herencia
                        If mintCodGrupo <> 0 Then
                            'Se le ha asignado a un grupo diferente y debe heredar los derechos correspondientes
                            '1 - Asignar al usuario el código de grupo
                            '2 - Borrar los Accesos del Usuario
                            '3 - y los derechos también

                            'Paso 1 - Asignar al usuario el código de grupo
                            ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), CStr(ModEncriptacion.Encriptar((Me.txtPassWord.Text))), CStr(False), CStr(mintCodGrupo), cTipoUsuario, CStr(0), C_MODIFICACION, CStr(0))
                            Cmd.Execute()

                            'Paso 2 - Debe borrar los Accesos del Usuario
                            ModStoredProcedures.PR_IEAccesos(Trim(Me.txtCodigo.Text), "", "", C_ELIMINACION, CStr(0))
                            Cmd.Execute()

                            'Paso 3 - Asignar los derechos heredados
                            ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), "", CStr(False), CStr(mintCodGrupo), cTipoUsuario, CStr(0), C_MODIFICACION, CStr(5))
                            Cmd.Execute()
                        Else
                            'Asignar 0 al Código de Grupo del Usuario
                            ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), CStr(ModEncriptacion.Encriptar((Me.txtPassWord.Text))), CStr(False), CStr(mintCodGrupo), cTipoUsuario, CStr(0), C_MODIFICACION, CStr(0))
                            Cmd.Execute()
                            '1 - Borrar los Accesos del Usuario
                            ModStoredProcedures.PR_IEAccesos(Trim(Me.txtCodigo.Text), "", "", C_ELIMINACION, CStr(0))
                            Cmd.Execute()
                        End If
                    End If
                Else 'Se trata de configurar un Usuario pero en la página de privilegios
                    'Debe borrar los Accesos del usuario correspondientes al módulo
                    ModStoredProcedures.PR_IEAccesos(CStr(mintCodUsuario), "", CStr(mintCodModulo), C_ELIMINACION, CStr(1))
                    Cmd.Execute()
                    'Paso 3  - Configurar los derechos del Usuario (Aquí se utiliza el procedimiento almacenado UP_IE_Accesos)
                    For I = 1 To Me._lstPrivilegios_0.Items.Count - 1
                        gStrSql = "Select CodModulo, CodFuncion, DescFuncion, Forma From CatFunciones Where CodModulo = " & mintCodModulo & " and LTrim(RTrim(DescFuncion)) = '" & Trim(Me._lstPrivilegios_0.Items.Item(I).Text) & "'"
                        ModEstandar.BorraCmd()
                        Cmd.CommandText = "dbo.UP_Select_Datos"
                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                        rsLocal = Cmd.Execute
                        If rsLocal.RecordCount > 0 Then
                            ModStoredProcedures.PR_IEAccesos(CStr(mintCodUsuario), Trim(rsLocal.Fields("Forma").Value), CStr(rsLocal.Fields("CodModulo").Value), C_INSERCION, CStr(0))
                            Cmd.Execute()
                        End If
                    Next
                End If
            End If
        End If
        Cnn.CommitTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        blnTransaction = False
        If mblnNuevo Then
            If Me._optTipo_1.Checked Then
                MsgBox("El Usuario ha sido grabado correctamente con el código " & Me.txtCodigo.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Else
                MsgBox("El Grupo ha sido grabado correctamente con el código " & Me.txtCodigo.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            End If
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        End If
        Nuevo()
        Guardar = True
        Limpiar()
Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub Eliminar()
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        gStrSql = "select codUsuario, Nombre, Grupo from CatUsuarios where codUsuario = " & Trim(Me.txtCodigo.Text)
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            MsgBox("Proporcione un código válido para Eliminar el Grupo o Usuario", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Me.SSTabSeg.SelectedIndex = nGRUPOUSUARIO
            Limpiar()
            RsGral.Close()
            Exit Sub
        End If
        If RsGral.Fields("Grupo").Value Then
            'Preguntar si desea borrar el Grupo
            If MsgBox("¿Desea Borrar el Grupo : " & Trim(RsGral.Fields("Nombre").Value) & " y dejar sin derechos a todos los usuarios que dependan de él ?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
                Exit Sub
            End If
            Cnn.BeginTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            blnTransaction = True
            'Borra el Grupo y los accesos de los usuarios que pertenecen al grupo
            ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), Trim(ModEncriptacion.Encriptar((Me.txtPassWord.Text))), CStr(True), Str(mintCodGrupo), cTipoUsuario, CStr(0), C_ELIMINACION, CStr(1))
            Cmd.Execute()
            Cnn.CommitTrans()
            blnTransaction = False
            Limpiar()
        Else
            'Preguntar si desea borrar al Usuario
            If MsgBox("¿Desea Borrar al Usuario : " & Trim(RsGral.Fields("Nombre").Value) & " ?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
                Exit Sub
            End If
            Cnn.BeginTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            blnTransaction = True
            ModStoredProcedures.PR_IMECatUsuarios(Trim(Me.txtCodigo.Text), Trim(Me.txtNombre.Text), Trim(ModEncriptacion.Encriptar((Me.txtPassWord.Text))), Str(mintCodGrupo), Str(mintCodGrupo), cTipoUsuario, CStr(0), C_ELIMINACION, CStr(0))
            Cmd.Execute()
            Cnn.CommitTrans()
            blnTransaction = False
            Limpiar()
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
Merr:
        If Err.Number <> 0 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function ValidaDatos() As Boolean
        On Error GoTo Merr
        If Me._optTipo_0.Checked Then
            Select Case Me.SSTabSeg.SelectedIndex
                Case nGRUPOUSUARIO
                    If Trim(Me.txtNombre.Text) = "" Then
                        MsgBox("Debe especificar el Nombre o Descripción del Grupo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        Me.txtNombre.Focus()
                        ValidaDatos = False
                        Exit Function
                    End If
                Case nPRIVILEGIOS
                    If mintCodGrupo = 0 Then
                        MsgBox("Debe indicar el Nombre o Descripción del Grupo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        Me._dbcGrupos_1.Focus()
                        ValidaDatos = False
                        Exit Function
                    End If
                    If mintCodModulo = 0 Then
                        MsgBox("Debe indicar el Módulo de Funciones", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        Me.dbcModulo.Focus()
                        ValidaDatos = False
                        Exit Function
                    End If
            End Select
        ElseIf Me._optTipo_1.Checked Then
            Select Case Me.SSTabSeg.SelectedIndex
                Case nGRUPOUSUARIO
                    If Trim(Me.txtNombre.Text) = "" Then
                        MsgBox("Debe especificar el Nombre del Usuario", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        Me.txtNombre.Focus()
                        ValidaDatos = False
                        Exit Function
                    End If
                    If Me.chkGrupo.CheckState = System.Windows.Forms.CheckState.Checked Then
                        If mintCodGrupo = 0 Then
                            MsgBox("Seleccione el Grupo al que pertenecerá el Usuario o Inhabilite la Casilla de verificación", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            Me.chkGrupo.Focus()
                            ValidaDatos = False
                            Exit Function
                        End If
                    End If
                    If Trim(Me.txtPassWord.Text) = "" Then
                        MsgBox("Debe introducir una contraseña de acceso", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        Me.txtPassWord.Focus()
                        ModEstandar.SelTextoTxt(txtPassWord)
                        ValidaDatos = False
                        Exit Function
                    End If
                    If Trim(Me.txtPassWord.Text) <> Trim(Me.txtConfirmar.Text) Then
                        MsgBox("La contraseña de confirmación debe ser idéntica al password", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        Me.txtConfirmar.Focus()
                        ModEstandar.SelTextoTxt(txtConfirmar)
                        ValidaDatos = False
                        Exit Function
                    End If
                Case nPRIVILEGIOS
                    If mintCodUsuario = 0 Then
                        MsgBox("Debe indicar el Nombre del Usuario", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        '''_dbcGrupos_1.SetFocus
                        ValidaDatos = False
                        Exit Function
                    End If
                    If mintCodModulo = 0 Then
                        MsgBox("Debe indicar el Módulo de Funciones", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        Me.dbcModulo.Focus()
                        ValidaDatos = False
                        Exit Function
                    End If
            End Select
        End If
        ValidaDatos = True
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function


    Public Sub Limpiar()
        On Error Resume Next
        'Validar si hubo cambios que desee guardar
        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                Case MsgBoxResult.Cancel 'Cancela la acción de limpiar pantalla
                    Exit Sub
            End Select
        End If
        Me.txtCodigo.Text = ""
        Nuevo()
        mblnCambiosEnCodigo = False
        Select Case Me.SSTabSeg.SelectedIndex
            Case nGRUPOUSUARIO
                Me.txtCodigo.Focus()
                mblnNuevo = True
            Case nPRIVILEGIOS
                If Me._optTipo_0.Checked Then
                    Me._dbcGrupos_1.Focus()
                Else
                    Me.dbcUsuarios.Focus()
                End If
                mblnNuevo = False
        End Select
    End Sub

    Public Sub Nuevo()
        Call ActivaCtl()
        If Not mblnNuevo Then
            Me.txtCodigo.Text = ""
            Me.txtCodigo.Tag = ""
        End If
        Me.txtNombre.Text = ""
        Me.txtNombre.Tag = ""
        Me.chkGrupo.CheckState = System.Windows.Forms.CheckState.Unchecked
        mblnFueraChange = True
        mintCodGrupo = 0
        Me._dbcGrupos_0.Text = ""
        Me._dbcGrupos_0.Tag = ""
        Me._dbcGrupos_1.Text = ""
        Me._dbcGrupos_1.Tag = ""
        mintCodUsuario = 0
        Me.dbcUsuarios.Text = ""
        Me.dbcUsuarios.Tag = ""
        mintCodModulo = 0
        Me.dbcModulo.Text = ""
        Me.dbcModulo.Tag = ""
        mblnFueraChange = False

        _lblSeg_7.Text = ""

        Me.txtPassWord.Text = ""
        Me.txtPassWord.Tag = ""
        Me.txtConfirmar.Text = ""
        Me.txtConfirmar.Tag = ""

        Me._lstUsuarios_0.Items.Clear()
        Me._lstUsuarios_1.Items.Clear()
        Me._btnMoverU_0.Enabled = False
        Me._btnMoverU_1.Enabled = False
        Me._btnMoverU_2.Enabled = False
        Me._btnMoverU_3.Enabled = False

        Me._lstPrivilegios_0.Items.Clear()
        Me._lstPrivilegios_1.Items.Clear()
        Me._btnMoverP_0.Enabled = False
        Me._btnMoverP_1.Enabled = False
        Me._btnMoverP_2.Enabled = False
        Me._btnMoverP_3.Enabled = False
    End Sub

    Public Sub LlenaDatos()
        On Error GoTo Merr
        If Me._optTipo_0.Checked Then
            'SE ESTÁ CONFIGURANDO UN GRUPO
            If Me.SSTabSeg.SelectedIndex = nGRUPOUSUARIO Then
                If LlenaUsuarios() Then
                Else
                    Exit Sub
                End If
            ElseIf Me.SSTabSeg.SelectedIndex = nPRIVILEGIOS Then
                If LlenaFunciones() Then
                Else
                    Exit Sub
                End If
            End If
        ElseIf Me._optTipo_1.Checked Then
            'SE ESTÁ CONFIGURANDO UN USUARIO
            If Me.SSTabSeg.SelectedIndex = nGRUPOUSUARIO Then
                If LlenaUsuario() Then
                Else
                    Exit Sub
                End If
            ElseIf Me.SSTabSeg.SelectedIndex = nPRIVILEGIOS Then
                If LlenaFunciones() Then
                Else
                    Exit Sub
                End If
            End If
        End If
        mblnCambiosEnCodigo = False
        mblnNuevo = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function LlenaUsuario() As Boolean
        'On Error GoTo Merr
        Try
            If CDbl(ModEstandar.Numerico((Me.txtCodigo.Text))) = 0 Then
                Nuevo()
                LlenaUsuario = False
                Exit Function
            End If
            'Me.txtCodigo.Text = Format(Me.txtCodigo.Text, "000")

            For i = 1 To 3 - txtCodigo.TextLength
                txtCodigo.Text = String.Concat("0", txtCodigo.Text)
            Next i

            gStrSql = "select * from CatUsuarios where Grupo = 0 and CodUsuario = " & ModEstandar.Numerico((Me.txtCodigo.Text))
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                Me.txtNombre.Text = Trim(RsGral.Fields("Nombre").Value)
                Me.txtNombre.Tag = Me.txtNombre.Text
                If RsGral.Fields("CodGrupo").Value > 0 Then
                    Me.chkGrupo.CheckState = System.Windows.Forms.CheckState.Checked
                    Me._dbcGrupos_0.Enabled = True
                    mblnFueraChange = True
                    mintCodGrupo = RsGral.Fields("CodGrupo").Value
                    Me._dbcGrupos_0.Text = Me.BuscaGrupo(mintCodGrupo)
                    Me._dbcGrupos_0.Tag = Me._dbcGrupos_0.Text
                    mblnFueraChange = False
                Else
                    Me.chkGrupo.CheckState = System.Windows.Forms.CheckState.Unchecked
                    mblnFueraChange = True
                    mintCodGrupo = 0
                    Me._dbcGrupos_0.Text = ""
                    Me._dbcGrupos_0.Tag = ""
                    mblnFueraChange = False
                End If
                Select Case RsGral.Fields("Tipo").Value
                    Case C_TADMIN
                        Me._optTipoUsuario_0.Checked = True
                        Me._optTipoUsuario_1.Checked = False
                        Me._optTipoUsuario_2.Checked = False
                        cTipoUsuario = C_TADMIN
                    Case C_TSUPERVISOR
                        Me._optTipoUsuario_0.Checked = False
                        Me._optTipoUsuario_1.Checked = True
                        Me._optTipoUsuario_2.Checked = False
                        cTipoUsuario = C_TSUPERVISOR
                    Case C_TEMPLEADO
                        Me._optTipoUsuario_0.Checked = False
                        Me._optTipoUsuario_1.Checked = False
                        Me._optTipoUsuario_2.Checked = True
                        cTipoUsuario = C_TEMPLEADO
                End Select

                Me.txtPassWord.Text = ModEncriptacion.Desencriptar(RsGral.Fields("Password").Value.ToString())
                Me.txtPassWord.Tag = Me.txtPassWord.Text
                Me.txtConfirmar.Text = ModEncriptacion.Desencriptar(RsGral.Fields("Password").Value.ToString())
                Me.txtConfirmar.Tag = Me.txtConfirmar.Text
            Else
                MsjNoExiste("Usuario", gstrNombCortoEmpresa)
                Limpiar()
                LlenaUsuario = False
                Exit Function
            End If
            LlenaUsuario = True
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
        Return LlenaUsuario
    End Function

    Public Function BuscaGrupo(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        Dim rsLocal As ADODB.Recordset
        'Selecciona el Nombre del Grupo
        gStrSql = "select Nombre from CatUsuarios where CodUsuario = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaGrupo = Trim(rsLocal.Fields("Nombre").Value)
        Else
            BuscaGrupo = ""
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Function BuscaGrupodeUsuario(ByRef CodUsuario As Integer) As String
        On Error GoTo Merr
        Dim rsLocal As ADODB.Recordset
        'Selecciona el Nombre del Grupo
        gStrSql = "select Nombre, CodGrupo from CatUsuarios where CodUsuario = " & CodUsuario
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaGrupodeUsuario = BuscaGrupo(rsLocal.Fields("CodGrupo").Value)
        Else
            BuscaGrupodeUsuario = ""
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()

    End Function

    Public Function LlenaUsuarios() As Boolean
        On Error GoTo Merr
        Dim I As Integer
        If CDbl(ModEstandar.Numerico((Me.txtCodigo.Text))) = 0 Then
            Nuevo()
            LlenaUsuarios = False
            Exit Function
        End If
        'Me.txtCodigo.Text = Format(Me.txtCodigo.Text, "00")

        For I = 1 To 2 - txtCodigo.TextLength
            txtCodigo.Text = String.Concat("0", txtCodigo.Text)
        Next I

        'Selecciona el Nombre del Grupo
        gStrSql = "select * from CatUsuarios where Grupo = 1 and CodUsuario = " & ModEstandar.Numerico((Me.txtCodigo.Text))
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Me.txtNombre.Text = Trim(RsGral.Fields("Nombre").Value)
            Me.txtNombre.Tag = Me.txtNombre.Text
        Else
            MsjNoExiste("El Grupo de Usuarios", gstrNombCortoEmpresa)
            Limpiar()
            LlenaUsuarios = False
            Exit Function
        End If
        'Selecciona los usuarios que pertenecen al grupo
        gStrSql = "select * from CatUsuarios where CodGrupo = " & ModEstandar.Numerico((Me.txtCodigo.Text))
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            RsGral.MoveFirst()
            For I = 1 To RsGral.RecordCount
                'Al asignar el valor a KEY le agrego una letra porque no acepta solamente numeros
                Item = Me._lstUsuarios_0.Items.Add("C" & CStr(RsGral.Fields("CodUsuario").Value), Trim(RsGral.Fields("Nombre").Value), "")
                RsGral.MoveNext()
            Next I
        End If
        'Selecciona los usuarios que no pertenecen a ningún grupo
        gStrSql = "select * from CatUsuarios where Grupo = 0 and CodGrupo = 0"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            RsGral.MoveFirst()
            For I = 1 To RsGral.RecordCount
                Item = Me._lstUsuarios_1.Items.Add("C" & CStr(RsGral.Fields("CodUsuario").Value), Trim(RsGral.Fields("Nombre").Value), "")
                RsGral.MoveNext()
            Next I
        End If

        If Me._lstUsuarios_0.Items.Count > 0 Then
            Me._btnMoverU_2.Enabled = True
            Me._btnMoverU_3.Enabled = True
        Else
            Me._btnMoverU_2.Enabled = False
            Me._btnMoverU_3.Enabled = False
        End If
        If Me._lstUsuarios_1.Items.Count > 0 Then
            Me._btnMoverU_0.Enabled = True
            Me._btnMoverU_1.Enabled = True
        Else
            Me._btnMoverU_0.Enabled = False
            Me._btnMoverU_1.Enabled = False
        End If

        LlenaUsuarios = True
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function LlenaFunciones() As Boolean
        On Error GoTo Merr
        Dim I As Integer
        '''Derechos ya asignados a los usuarios
        ModEstandar.BorraCmd()
        If Me._optTipo_0.Checked Then
            '''gStrSql = "Select a.Forma, a.DescFuncion From CatFunciones a, Accesos b Where b.CodUsuario = " & mintCodGrupo & " and ltrim(rtrim(a.Forma)) = ltrim(rtrim(b.Forma)) and a.CodModulo = " & mintCodModulo & " ORDER BY a.DescFuncion"
            gStrSql = "Select Upper(LTRIM(RTRIM(Func.Forma))) as Forma, Upper(LTRIM(RTRIM(Func.DescFuncion))) as DescFuncion From ( " & "Select * From CatFunciones Where CodModulo = " & mintCodModulo & ") Func Inner Join " & "(Select * From Accesos Where CodModulo = " & mintCodModulo & " And CodUsuario = " & mintCodGrupo & ") Acc On Func.CodModulo = Acc.CodModulo And Upper(LTRIM(RTRIM(Func.Forma))) = Upper(LTRIM(RTRIM(Acc.FORMA))) "
        Else
            '''gStrSql = "Select a.Forma as Forma, a.DescFuncion as DescFuncion From CatFunciones a, Accesos b Where b.CodUsuario = " & mintCodUsuario & " and ltrim(rtrim(a.Forma)) = ltrim(rtrim(b.Forma)) and a.CodModulo = " & mintCodModulo & " ORDER BY a.DescFuncion"
            gStrSql = "Select Upper(LTRIM(RTRIM(Func.Forma))) as Forma, Upper(LTRIM(RTRIM(Func.DescFuncion))) as DescFuncion From ( " & "Select * From CatFunciones Where CodModulo = " & mintCodModulo & ") Func Inner Join " & "(Select * From Accesos Where CodModulo = " & mintCodModulo & " And CodUsuario = " & mintCodUsuario & ") Acc On Func.CodModulo = Acc.CodModulo And Upper(LTRIM(RTRIM(Func.Forma))) = Upper(LTRIM(RTRIM(Acc.FORMA))) "
        End If
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        _lstPrivilegios_0.Items.Clear()
        If RsGral.RecordCount > 0 Then
            With Me
                RsGral.MoveFirst()
                For I = 1 To RsGral.RecordCount
                    'Aquí no es necesario agregar una letra a Key porque RsGral!Forma comienza con Letra
                    Item = ._lstPrivilegios_0.Items.Add(RsGral.Fields("Forma").Value, Trim(RsGral.Fields("DescFuncion").Value), "")
                    RsGral.MoveNext()
                Next I
            End With
        End If

        '''Derechos pendientes de los usuarios
        ModEstandar.BorraCmd()
        If Me._optTipo_0.Checked Then
            '''gStrSql = "Select Forma, DescFuncion From CatFunciones Where Forma Not In (Select Distinct Forma From Accesos Where CodUsuario = " & mintCodGrupo & ") And CodModulo = " & mintCodModulo & " Order By DescFuncion"
            gStrSql = "Select Upper(LTRIM(RTRIM(Forma))) as Forma, Upper(LTRIM(RTRIM(DescFuncion))) as DescFuncion From CatFunciones Where CodModulo = " & mintCodModulo & " and Forma Not In (Select Forma From Accesos Where CodUsuario = " & mintCodGrupo & " And CodModulo = " & mintCodModulo & ") Order By DescFuncion "
        Else
            '''gStrSql = "Select Forma, DescFuncion From CatFunciones Where Forma Not In (Select Distinct Forma From Accesos Where CodUsuario = " & mintCodUsuario & ") And CodModulo = " & mintCodModulo & " Order By DescFuncion"
            gStrSql = "Select Upper(LTRIM(RTRIM(Forma))) as Forma, Upper(LTRIM(RTRIM(DescFuncion))) as DescFuncion From CatFunciones Where CodModulo = " & mintCodModulo & " and Forma Not In (Select Forma From Accesos Where CodUsuario = " & mintCodUsuario & " And CodModulo = " & mintCodModulo & ") Order By DescFuncion "
        End If
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        _lstPrivilegios_1.Items.Clear()
        If RsGral.RecordCount > 0 Then
            RsGral.MoveFirst()
            For I = 1 To RsGral.RecordCount
                Item = _lstPrivilegios_1.Items.Add(RsGral.Fields("Forma").Value, Trim(RsGral.Fields("DescFuncion").Value), "")
                RsGral.MoveNext()
            Next I
        End If
        If Me._lstPrivilegios_0.Items.Count > 0 Then
            Me._btnMoverP_2.Enabled = True
            Me._btnMoverP_3.Enabled = True
        Else
            Me._btnMoverP_2.Enabled = False
            Me._btnMoverP_3.Enabled = False
        End If
        If Me._lstPrivilegios_1.Items.Count > 0 Then
            Me._btnMoverP_0.Enabled = True
            Me._btnMoverP_1.Enabled = True
        Else
            Me._btnMoverP_0.Enabled = False
            Me._btnMoverP_1.Enabled = False
        End If
        LlenaFunciones = True
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub ActivaCtl()
        If Me._optTipo_0.Checked Then
            Me.ToolTip1.SetToolTip(Me.txtCodigo, "Código del Grupo")
            Me.ToolTip1.SetToolTip(Me.txtNombre, "Nombre o Descripción del Grupo")
            'Página nGRUPOUSUARIO
            Me.chkGrupo.Visible = False
            Me._dbcGrupos_0.Visible = False
            Me._lstUsuarios_0.Visible = True
            Me._lstUsuarios_1.Visible = True
            Me.fraPassWord.Visible = False
            Me.fraTipoUsuario.Visible = False
            cTipoUsuario = "G"
            Me._btnMoverU_0.Visible = True
            Me._btnMoverU_1.Visible = True
            Me._btnMoverU_2.Visible = True
            Me._btnMoverU_3.Visible = True
            'Página nPRIVILEGIOS
            Me._dbcGrupos_1.Visible = True
            Me._lblSeg_5.Visible = True
            Me._lblSeg_6.Visible = False
            Me.dbcUsuarios.Visible = False
            Me._lblSeg_7.Visible = False
        ElseIf Me._optTipo_1.Checked Then
            'Página nGRUPOUSUARIO
            Me.ToolTip1.SetToolTip(Me.txtCodigo, "Código del Usuario")
            Me.ToolTip1.SetToolTip(Me.txtNombre, "Nombre del Usuario")
            Me.chkGrupo.Visible = True
            Me._dbcGrupos_0.Visible = True
            If Me.chkGrupo.CheckState = System.Windows.Forms.CheckState.Checked Then
                Me._dbcGrupos_0.Enabled = True
            Else
                Me._dbcGrupos_0.Enabled = False
            End If
            Me._lstUsuarios_0.Visible = False
            Me._lstUsuarios_1.Visible = False
            Me.fraPassWord.Visible = True
            Me.fraTipoUsuario.Visible = True
            Me._optTipoUsuario_2.Checked = True
            cTipoUsuario = C_TEMPLEADO
            Me._btnMoverU_0.Visible = False
            Me._btnMoverU_1.Visible = False
            Me._btnMoverU_2.Visible = False
            Me._btnMoverU_3.Visible = False
            'Página nPRIVILEGIOS
            Me._dbcGrupos_1.Visible = False
            Me._lblSeg_5.Visible = False
            Me._lblSeg_6.Visible = True
            Me.dbcUsuarios.Visible = True
            Me._lblSeg_7.Visible = True
            Me._lblSeg_7.Text = ""
        End If
    End Sub

    'Private Sub btnMoverP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnMoverP.Click
    '    Dim Index As Integer = btnMoverP.GetIndex(eventSender)
    '    Dim I As Integer
    '    Select Case Index
    '        Case 0
    '            For I = 1 To _lstPrivilegios_1.Items.Count
    '                Item = _lstPrivilegios_0.Items.Add(_lstPrivilegios_1.Items.Item(I).Name, _lstPrivilegios_1.Items.Item(I).Text, "")
    '            Next I
    '            _lstPrivilegios_1.Items.Clear()
    '            _btnMoverP_0.Enabled = False
    '            _btnMoverP_1.Enabled = False
    '            _btnMoverP_2.Enabled = True
    '            _btnMoverP_3.Enabled = True
    '        Case 1
    '            If _lstPrivilegios_1.Items.Count > 0 Then
    '                Item = _lstPrivilegios_0.Items.Add(_lstPrivilegios_1.FocusedItem.Name, _lstPrivilegios_1.FocusedItem.Text, "")
    '                _lstPrivilegios_1.Items.RemoveAt((_lstPrivilegios_1.FocusedItem.Name))
    '            End If
    '            If _lstPrivilegios_1.Items.Count > 0 Then
    '                _btnMoverP_0.Enabled = True
    '                _btnMoverP_1.Enabled = True
    '            Else
    '                _btnMoverP_0.Enabled = False
    '                _btnMoverP_1.Enabled = False
    '            End If
    '            _btnMoverP_2.Enabled = True
    '            _btnMoverP_3.Enabled = True
    '        Case 2
    '            If _lstPrivilegios_0.Items.Count > 0 Then
    '                Item = _lstPrivilegios_1.Items.Add(_lstPrivilegios_0.FocusedItem.Name, _lstPrivilegios_0.FocusedItem.Text, "")
    '                _lstPrivilegios_0.Items.RemoveAt((_lstPrivilegios_0.FocusedItem.Name))
    '            End If
    '            If _lstPrivilegios_0.Items.Count > 0 Then
    '                _btnMoverP_2.Enabled = True
    '                _btnMoverP_3.Enabled = True
    '            Else
    '                _btnMoverP_2.Enabled = False
    '                _btnMoverP_3.Enabled = False
    '            End If
    '            _btnMoverP_0.Enabled = True
    '            _btnMoverP_1.Enabled = True
    '        Case 3
    '            For I = 1 To _lstPrivilegios_0.Items.Count
    '                Item = _lstPrivilegios_1.Items.Add(_lstPrivilegios_0.Items.Item(I).Name, _lstPrivilegios_0.Items.Item(I).Text, "")
    '            Next I
    '            _lstPrivilegios_0.Items.Clear()
    '            _btnMoverP_0.Enabled = True
    '            _btnMoverP_1.Enabled = True
    '            _btnMoverP_2.Enabled = False
    '            _btnMoverP_3.Enabled = False
    '    End Select
    'End Sub

    Private Sub chkGrupo_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkGrupo.CheckStateChanged
        mblnFueraChange = True
        Me._dbcGrupos_0.Text = ""
        mintCodGrupo = 0
        If Me.chkGrupo.CheckState = System.Windows.Forms.CheckState.Checked Then
            Me._dbcGrupos_0.Enabled = True
        Else
            Me._dbcGrupos_0.Enabled = False
        End If
        mblnFueraChange = False
    End Sub

    Private Sub chkGrupo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkGrupo.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcModulo_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcModulo.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        lStrSql = "SELECT codModulo, rtrim(ltrim(descModulo)) as descModulo FROM catModulos WHERE LTrim(RTrim(descModulo)) LIKE '" & Trim(Me.dbcModulo.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcModulo)

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcModulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcModulo.Enter
        Pon_Tool()
        gStrSql = "SELECT codModulo, rtrim(ltrim(descModulo)) as descModulo FROM catModulos ORDER BY descModulo"
        ModDCombo.DCGotFocus(gStrSql, dbcModulo)
    End Sub

    Private Sub dbcModulo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcModulo.KeyDown
        Dim Aux As String
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
            Case System.Windows.Forms.Keys.Return
                Aux = Trim(Me.dbcModulo.Text)
                If Me.dbcModulo.SelectedValue <> 0 Then
                    dbcModulo_Leave(dbcModulo, New System.EventArgs())
                End If
                Me.dbcModulo.Text = Aux
                Exit Sub
            Case System.Windows.Forms.Keys.Tab
                Aux = Trim(Me.dbcModulo.Text)
                If Me.dbcModulo.SelectedValue <> 0 Then
                    dbcModulo_Leave(dbcModulo, New System.EventArgs())
                End If
                Me.dbcModulo.Text = Aux
                Exit Sub
        End Select
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcModulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcModulo.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codModulo, rtrim(ltrim(descModulo)) as descModulo FROM catModulos WHERE LTrim(RTrim(descModulo)) LIKE '" & Trim(Me.dbcModulo.Text) & "%'"
        Aux = mintCodModulo
        mintCodModulo = 0
        ModDCombo.DCLostFocus(dbcModulo, gStrSql, mintCodModulo)
        If mintCodModulo <> Aux Then
            'Llenar los controles lstPrivilegios
            LlenaDatos()
        End If
    End Sub

    Private Sub dbcModulo_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcModulo.MouseUp
        'Dim Aux As String
        'Aux = Trim(Me.dbcModulo.Text)
        'If Me.dbcModulo.SelectedItem <> 0 Then
        '    dbcModulo_Leave(dbcModulo, New System.EventArgs())
        'End If
        'Me.dbcModulo.Text = Aux
    End Sub

    Private Sub dbcUsuarios_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcUsuarios.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        lStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 0 and Nombre LIKE '" & Trim(Me.dbcUsuarios.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcUsuarios)

        If Me.dbcUsuarios.Text = "" Then
            Me._lblSeg_7.Text = ""
            mblnFueraChange = True
            Me.dbcModulo.Text = ""
            Me.dbcModulo.Tag = ""
            mintCodModulo = 0
            mblnFueraChange = False
            Me._lstPrivilegios_0.Items.Clear()
            Me._lstPrivilegios_1.Items.Clear()
        Else
            Me._lblSeg_7.Text = ""
        End If

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcUsuarios_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcUsuarios.Enter
        Pon_Tool()
        gStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 0 ORDER BY Nombre"
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcUsuarios))
    End Sub

    Private Sub dbcUsuarios_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcUsuarios.KeyDown
        Dim Aux As String
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Me.SSTabSeg.Focus()
            Case System.Windows.Forms.Keys.Return
                Aux = Trim(Me.dbcUsuarios.Text)
                If Me.dbcUsuarios.SelectedItem <> 0 Then
                    dbcUsuarios_Leave(dbcUsuarios, New System.EventArgs())
                End If
                Me.dbcUsuarios.Text = Aux
                Exit Sub
            Case System.Windows.Forms.Keys.Tab
                Aux = Trim(Me.dbcUsuarios.Text)
                If Me.dbcUsuarios.SelectedItem <> 0 Then
                    dbcUsuarios_Leave(dbcUsuarios, New System.EventArgs())
                End If
                Me.dbcUsuarios.Text = Aux
                Exit Sub
        End Select
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcUsuarios_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcUsuarios.Leave
        Dim I As Integer
        Dim Aux As Integer
        Dim cGrupodeUsuario As String
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 0 and Nombre LIKE '" & Trim(Me.dbcUsuarios.Text) & "%'"
        Aux = mintCodUsuario
        mintCodUsuario = 0
        ModDCombo.DCLostFocus((Me.dbcUsuarios), gStrSql, mintCodUsuario)
        If Aux <> mintCodUsuario Then
            If mintCodUsuario <> 0 Then
                cGrupodeUsuario = BuscaGrupodeUsuario(mintCodUsuario)
                If Trim(cGrupodeUsuario) <> "" Then
                    Me._lblSeg_7.Text = cGrupodeUsuario
                Else
                    Me._lblSeg_7.Text = "(Usuario Independiente)"
                End If
            Else
                Me._lblSeg_7.Text = ""
            End If
            mblnFueraChange = True
            Me.dbcModulo.Text = ""
            Me.dbcModulo.Tag = ""
            mintCodModulo = 0
            mblnFueraChange = False
            Me._lstPrivilegios_0.Items.Clear()
            Me._lstPrivilegios_1.Items.Clear()
        End If
    End Sub

    Private Sub dbcUsuarios_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcUsuarios.MouseUp
        'Dim Aux As String
        'Aux = Trim(Me.dbcUsuarios.Text)
        'If Me.dbcUsuarios.SelectedItem <> 0 Then
        '    dbcUsuarios_Leave(dbcUsuarios, New System.EventArgs())
        'End If
        'Me.dbcUsuarios.Text = Aux
    End Sub

    Private Sub frmABCUsuarios_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmABCUsuarios_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmABCUsuarios_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If UCase(Me.ActiveControl.Name) = "LSTUSUARIOS" Then
                    If Me.ActiveControl.Text = 1 Then
                        Me.SSTabSeg.Focus()
                    Else
                        ModEstandar.AvanzarTab(Me)
                    End If
                ElseIf UCase(Me.ActiveControl.Name) = "LSTPRIVILEGIOS" Then
                    If Me.ActiveControl.Text = 1 Then
                        Me.SSTabSeg.Focus()
                    Else
                        ModEstandar.AvanzarTab(Me)
                    End If
                ElseIf UCase(Me.ActiveControl.Name) = "TXTCONFIRMAR" Then
                    Me.SSTabSeg.Focus()
                Else
                    ModEstandar.AvanzarTab(Me)
                End If
            Case System.Windows.Forms.Keys.Escape
                If Trim(UCase(Me.ActiveControl.Name)) = "OPTTIPO" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmABCUsuarios_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayúscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmABCUsuarios_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Call Me.ActivaCtl()
        Me.SSTabSeg.SelectedIndex = nGRUPOUSUARIO
        mblnNuevo = True
        mblnCambiosEnCodigo = False
    End Sub

    Private Sub frmABCUsuarios_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si desea cerrar la forma y esta se encuentra minimizada, esta se restaura
        'If Not mblnSalir Then
        '    ModEstandar.RestaurarForma(Me, False)
        '    If Cambios() Then 'And Not (mblnNuevo)
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes
        '                If Not (Guardar()) Then
        '                    Cancel = 1
        '                End If
        '            Case MsgBoxResult.No 'No hace nada y permite que se cierre el formulario
        '                Cancel = 0
        '            Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin Guardar
        '                Cancel = 1
        '        End Select
        '    End If
        'Else 'Se quiere salir con escape
        '    mblnSalir = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Me.SSTabSeg.SelectedIndex = nGRUPOUSUARIO
        '            Me.txtCodigo.Focus()
        '            ModEstandar.SelTxt()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmABCUsuarios_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    'Private Sub lstPrivilegios_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstPrivilegios.DoubleClick
    '    Dim Index As Integer
    '    '= lstPrivilegios.GetIndex(eventSender)
    '    Select Case Index
    '        Case 0
    '            btnMoverP_Click(btnMoverP.Item(2), New System.EventArgs())
    '        Case 1
    '            btnMoverP_Click(btnMoverP.Item(1), New System.EventArgs())
    '    End Select
    'End Sub

    'Private Sub lstPrivilegios_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstPrivilegios.Enter
    '    Dim Index As Integer
    '    '= lstPrivilegios.GetIndex(eventSender)
    '    Pon_Tool()
    'End Sub


    Private Sub _lstPrivilegios_0_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _lstPrivilegios_0.DoubleClick
        Dim Index As Integer
        '= _lstPrivilegios_0.SelectedItems(0).Index
        Select Case Index
            Case 0
                _btnMoverP_3_Click(_btnMoverP_3, New System.EventArgs())
            Case 1
                _btnMoverP_2_Click(_btnMoverP_2, New System.EventArgs())
        End Select
    End Sub

    Private Sub _lstPrivilegios_0_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _lstPrivilegios_0.Enter
        Dim Index As Integer
        '= _lstPrivilegios_0.GetIndex(eventSender)
        Pon_Tool()
    End Sub


    Private Sub _lstPrivilegios_1_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _lstPrivilegios_1.DoubleClick
        Dim Index As Integer = _lstPrivilegios_1.SelectedItems(0).Index
        Select Case Index
            Case 0
                _btnMoverP_1_Click(_btnMoverP_1, New System.EventArgs())
            Case 1
                _btnMoverP_0_Click(_btnMoverP_0, New System.EventArgs())
        End Select
    End Sub

    Private Sub _lstPrivilegios_1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _lstPrivilegios_1.Enter
        Dim Index As Integer
        '= _lstPrivilegios_1.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub _btnMoverP_0_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverP_0.Click
        Dim Index As Integer = 0
        '= _btnMoverP_2.Text
        '= _btnMoverP_0.GetIndex(eventSender)
        Dim I As Integer
        Select Case Index
            Case 0
                For I = 0 To _lstPrivilegios_1.Items.Count - 1
                    Item = _lstPrivilegios_0.Items.Add(_lstPrivilegios_1.Items.Item(I).Name, _lstPrivilegios_1.Items.Item(I).Text, "")
                Next
                _lstPrivilegios_1.Items.Clear()
                _btnMoverP_0.Enabled = False
                _btnMoverP_1.Enabled = False
                _btnMoverP_2.Enabled = True
                _btnMoverP_3.Enabled = True
            Case 1
                If _lstPrivilegios_1.Items.Count > 0 Then
                    Item = _lstPrivilegios_0.Items.Add(_lstPrivilegios_1.FocusedItem.Name, _lstPrivilegios_1.FocusedItem.Text, "")
                    _lstPrivilegios_1.Items.RemoveAt((_lstPrivilegios_1.FocusedItem.Index))
                End If
                If _lstPrivilegios_1.Items.Count > 0 Then
                    _btnMoverP_0.Enabled = True
                    _btnMoverP_1.Enabled = True
                Else
                    _btnMoverP_0.Enabled = False
                    _btnMoverP_1.Enabled = False
                End If
                _btnMoverP_2.Enabled = True
                _btnMoverP_3.Enabled = True
            Case 2
                If _lstPrivilegios_0.Items.Count > 0 Then
                    Item = _lstPrivilegios_1.Items.Add(_lstPrivilegios_0.FocusedItem.Name, _lstPrivilegios_0.FocusedItem.Text, "")
                    _lstPrivilegios_0.Items.RemoveAt((_lstPrivilegios_0.FocusedItem.Index))
                End If
                If _lstPrivilegios_0.Items.Count > 0 Then
                    _btnMoverP_2.Enabled = True
                    _btnMoverP_3.Enabled = True
                Else
                    _btnMoverP_2.Enabled = False
                    _btnMoverP_3.Enabled = False
                End If
                _btnMoverP_0.Enabled = True
                _btnMoverP_1.Enabled = True
            Case 3
                For I = 0 To _lstPrivilegios_0.Items.Count - 1
                    Item = _lstPrivilegios_1.Items.Add(_lstPrivilegios_0.Items.Item(I).Name, _lstPrivilegios_0.Items.Item(I).Text, "")
                Next
                _lstPrivilegios_0.Items.Clear()
                _btnMoverP_0.Enabled = True
                _btnMoverP_1.Enabled = True
                _btnMoverP_2.Enabled = False
                _btnMoverP_3.Enabled = False
        End Select
    End Sub

    Private Sub _btnMoverP_1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverP_1.Click
        Dim Index As Integer = 1
        '= _btnMoverP_2.Text
        '= _btnMoverP_0.GetIndex(eventSender)
        Dim I As Integer
        Select Case Index
            Case 0
                For I = 0 To _lstPrivilegios_1.Items.Count - 1
                    Item = _lstPrivilegios_0.Items.Add(_lstPrivilegios_1.Items.Item(I).Name, _lstPrivilegios_1.Items.Item(I).Text, "")
                Next
                _lstPrivilegios_1.Items.Clear()
                _btnMoverP_0.Enabled = False
                _btnMoverP_1.Enabled = False
                _btnMoverP_2.Enabled = True
                _btnMoverP_3.Enabled = True
            Case 1
                If _lstPrivilegios_1.Items.Count > 0 Then
                    Item = _lstPrivilegios_0.Items.Add(_lstPrivilegios_1.FocusedItem.Name, _lstPrivilegios_1.FocusedItem.Text, "")
                    _lstPrivilegios_1.Items.RemoveAt((_lstPrivilegios_1.FocusedItem.Index))
                End If
                If _lstPrivilegios_1.Items.Count > 0 Then
                    _btnMoverP_0.Enabled = True
                    _btnMoverP_1.Enabled = True
                Else
                    _btnMoverP_0.Enabled = False
                    _btnMoverP_1.Enabled = False
                End If
                _btnMoverP_2.Enabled = True
                _btnMoverP_3.Enabled = True
            Case 2
                If _lstPrivilegios_0.Items.Count > 0 Then
                    Item = _lstPrivilegios_1.Items.Add(_lstPrivilegios_0.FocusedItem.Name, _lstPrivilegios_0.FocusedItem.Text, "")
                    _lstPrivilegios_0.Items.RemoveAt((_lstPrivilegios_0.FocusedItem.Index))
                End If
                If _lstPrivilegios_0.Items.Count > 0 Then
                    _btnMoverP_2.Enabled = True
                    _btnMoverP_3.Enabled = True
                Else
                    _btnMoverP_2.Enabled = False
                    _btnMoverP_3.Enabled = False
                End If
                _btnMoverP_0.Enabled = True
                _btnMoverP_1.Enabled = True
            Case 3
                For I = 0 To _lstPrivilegios_0.Items.Count - 1
                    Item = _lstPrivilegios_1.Items.Add(_lstPrivilegios_0.Items.Item(I).Name, _lstPrivilegios_0.Items.Item(I).Text, "")
                Next
                _lstPrivilegios_0.Items.Clear()
                _btnMoverP_0.Enabled = True
                _btnMoverP_1.Enabled = True
                _btnMoverP_2.Enabled = False
                _btnMoverP_3.Enabled = False
        End Select
    End Sub

    Private Sub _btnMoverP_2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverP_2.Click
        Dim Index As Integer = 2
        Dim I As Integer
        Select Case Index
            Case 0
                For I = 0 To _lstPrivilegios_1.Items.Count - 1
                    Item = _lstPrivilegios_0.Items.Add(_lstPrivilegios_1.Items.Item(I).Name, _lstPrivilegios_1.Items.Item(I).Text, "")
                Next
                _lstPrivilegios_1.Items.Clear()
                _btnMoverP_0.Enabled = False
                _btnMoverP_1.Enabled = False
                _btnMoverP_2.Enabled = True
                _btnMoverP_3.Enabled = True
            Case 1
                If _lstPrivilegios_1.Items.Count > 0 Then
                    Item = _lstPrivilegios_0.Items.Add(_lstPrivilegios_1.FocusedItem.Name, _lstPrivilegios_1.FocusedItem.Text, "")
                    _lstPrivilegios_1.Items.RemoveAt((_lstPrivilegios_1.FocusedItem.Index))
                End If
                If _lstPrivilegios_1.Items.Count > 0 Then
                    _btnMoverP_0.Enabled = True
                    _btnMoverP_1.Enabled = True
                Else
                    _btnMoverP_0.Enabled = False
                    _btnMoverP_1.Enabled = False
                End If
                _btnMoverP_2.Enabled = True
                _btnMoverP_3.Enabled = True
            Case 2
                If _lstPrivilegios_0.Items.Count > 0 Then
                    Item = _lstPrivilegios_1.Items.Add(_lstPrivilegios_0.Items.Item(I).Name, _lstPrivilegios_0.Items.Item(I).Text, "")
                    _lstPrivilegios_0.Items.RemoveAt((_lstPrivilegios_0.Items.Item(I).Index))
                End If
                If _lstPrivilegios_0.Items.Count > 0 Then
                    _btnMoverP_2.Enabled = True
                    _btnMoverP_3.Enabled = True
                Else
                    _btnMoverP_2.Enabled = False
                    _btnMoverP_3.Enabled = False
                End If
                _btnMoverP_0.Enabled = True
                _btnMoverP_1.Enabled = True
            Case 3
                For I = 0 To _lstPrivilegios_0.Items.Count - 1
                    Item = _lstPrivilegios_1.Items.Add(_lstPrivilegios_0.Items.Item(I).Name, _lstPrivilegios_0.Items.Item(I).Text, "")
                Next
                _lstPrivilegios_0.Items.Clear()
                _btnMoverP_0.Enabled = True
                _btnMoverP_1.Enabled = True
                _btnMoverP_2.Enabled = False
                _btnMoverP_3.Enabled = False
        End Select
    End Sub

    Private Sub _btnMoverP_3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverP_3.Click
        Dim Index As Integer = 3
        '= _btnMoverP_2.Text
        '= _btnMoverP_0.GetIndex(eventSender)
        Dim I As Integer
        Select Case Index
            Case 0
                For I = 0 To _lstPrivilegios_1.Items.Count - 1
                    Item = _lstPrivilegios_0.Items.Add(_lstPrivilegios_1.Items.Item(I).Name, _lstPrivilegios_1.Items.Item(I).Text, "")
                Next
                _lstPrivilegios_1.Items.Clear()
                _btnMoverP_0.Enabled = False
                _btnMoverP_1.Enabled = False
                _btnMoverP_2.Enabled = True
                _btnMoverP_3.Enabled = True
            Case 1
                If _lstPrivilegios_1.Items.Count > 0 Then
                    Item = _lstPrivilegios_0.Items.Add(_lstPrivilegios_1.FocusedItem.Name, _lstPrivilegios_1.FocusedItem.Text, "")
                    _lstPrivilegios_1.Items.RemoveAt((_lstPrivilegios_1.FocusedItem.Index))
                End If
                If _lstPrivilegios_1.Items.Count > 0 Then
                    _btnMoverP_0.Enabled = True
                    _btnMoverP_1.Enabled = True
                Else
                    _btnMoverP_0.Enabled = False
                    _btnMoverP_1.Enabled = False
                End If
                _btnMoverP_2.Enabled = True
                _btnMoverP_3.Enabled = True
            Case 2
                If _lstPrivilegios_0.Items.Count > 0 Then
                    Item = _lstPrivilegios_1.Items.Add(_lstPrivilegios_0.FocusedItem.Name, _lstPrivilegios_0.FocusedItem.Text, "")
                    _lstPrivilegios_0.Items.RemoveAt((_lstPrivilegios_0.FocusedItem.Index))
                End If
                If _lstPrivilegios_0.Items.Count > 0 Then
                    _btnMoverP_2.Enabled = True
                    _btnMoverP_3.Enabled = True
                Else
                    _btnMoverP_2.Enabled = False
                    _btnMoverP_3.Enabled = False
                End If
                _btnMoverP_0.Enabled = True
                _btnMoverP_1.Enabled = True
            Case 3
                For I = 0 To _lstPrivilegios_0.Items.Count - 1
                    Item = _lstPrivilegios_1.Items.Add(_lstPrivilegios_0.Items.Item(I).Name, _lstPrivilegios_0.Items.Item(I).Text, "")
                Next
                _lstPrivilegios_0.Items.Clear()
                _btnMoverP_0.Enabled = True
                _btnMoverP_1.Enabled = True
                _btnMoverP_2.Enabled = False
                _btnMoverP_3.Enabled = False
        End Select
    End Sub

    'Private Sub lstUsuarios_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstUsuarios.DoubleClick
    '    Dim Index As Integer
    '    '= lstUsuarios.GetIndex(eventSender)
    '    Select Case Index
    '        Case 0
    '            btnMoverU_Click(btnMoverU.Item(2), New System.EventArgs())
    '        Case 1
    '            btnMoverU_Click(btnMoverU.Item(1), New System.EventArgs())
    '    End Select
    'End Sub

    'Private Sub lstUsuarios_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstUsuarios.Enter
    '    Dim Index As Integer
    '    '= lstUsuarios.GetIndex(eventSender)
    '    Pon_Tool()
    'End Sub


    Private Sub _lstUsuarios_0_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _lstUsuarios_0.DoubleClick
        Dim Index As Integer
        '= _lstUsuarios_0.GetIndex(eventSender)
        Select Case Index
            Case 0
                _btnMoverU_3_Click(_btnMoverU_3, New System.EventArgs())
            Case 1
                _btnMoverU_2_Click(_btnMoverU_2, New System.EventArgs())
        End Select
    End Sub

    Private Sub _lstUsuarios_0_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _lstUsuarios_0.Enter
        Dim Index As Integer
        '= _lstUsuarios_0.GetIndex(eventSender)
        Pon_Tool()
    End Sub


    Private Sub _lstUsuarios_1_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _lstUsuarios_1.DoubleClick
        Dim Index As Integer
        '= _lstUsuarios_1.GetIndex(eventSender)
        Select Case Index
            Case 0
                _btnMoverU_1_Click(_btnMoverU_1, New System.EventArgs())
            Case 1
                _btnMoverU_0_Click(_btnMoverU_0, New System.EventArgs())
        End Select
    End Sub

    Private Sub _lstUsuarios_1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _lstUsuarios_1.Enter
        Dim Index As Integer
        '= _lstUsuarios_1.GetIndex(eventSender)
        Pon_Tool()
    End Sub


    'Private Sub btnMoverU_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnMoverU.Click
    '    Dim Index As Integer
    '    '= btnMoverU.GetIndex(eventSender)
    '    Dim I As Integer
    '    Select Case Index
    '        Case 0
    '            ModEstandar.BorraCmd()
    '            gStrSql = "Select CodUsuario, Nombre From CatUsuarios Where Grupo = 0 And (CodGrupo = 0 Or CodGrupo = " & Numerico((Me.txtCodigo.Text)) & ")"
    '            Cmd.CommandText = "dbo.UP_Select_Datos"
    '            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
    '            RsGral = Cmd.Execute
    '            _lstUsuarios_0.Items.Clear()
    '            With Me
    '                For I = 1 To RsGral.RecordCount
    '                    ' Al asignar el valor a KEY le agrego una letra porque no acepta solamente numeros
    '                    Item = ._lstUsuarios_0.Items.Add("C" & CStr(RsGral.Fields("CodUsuario").Value), Trim(RsGral.Fields("Nombre").Value), "")
    '                    RsGral.MoveNext()
    '                Next I
    '            End With
    '            _lstUsuarios_1.Items.Clear()
    '            _btnMoverU_0.Enabled = False
    '            _btnMoverU_1.Enabled = False
    '            _btnMoverU_2.Enabled = True
    '            _btnMoverU_3.Enabled = True
    '        Case 1
    '            If _lstUsuarios_1.Items.Count > 0 Then
    '                Item = _lstUsuarios_0.Items.Add(_lstUsuarios_1.FocusedItem.Name, _lstUsuarios_1.FocusedItem.Text, "")
    '                _lstUsuarios_1.Items.RemoveAt((_lstUsuarios_1.FocusedItem.Name))
    '            End If
    '            If _lstUsuarios_1.Items.Count > 0 Then
    '                _btnMoverU_0.Enabled = True
    '                _btnMoverU_1.Enabled = True
    '            Else
    '                _btnMoverU_0.Enabled = False
    '                _btnMoverU_1.Enabled = False
    '            End If
    '            _btnMoverU_2.Enabled = True
    '            _btnMoverU_3.Enabled = True
    '        Case 2
    '            If _lstUsuarios_0.Items.Count > 0 Then
    '                Item = _lstUsuarios_1.Items.Add(_lstUsuarios_0.FocusedItem.Name, _lstUsuarios_0.FocusedItem.Text, "")
    '                _lstUsuarios_0.Items.RemoveAt((_lstUsuarios_0.FocusedItem.Name))
    '            End If
    '            If _lstUsuarios_0.Items.Count > 0 Then
    '                _btnMoverU_2.Enabled = True
    '                _btnMoverU_3.Enabled = True
    '            Else
    '                _btnMoverU_2.Enabled = False
    '                _btnMoverU_3.Enabled = False
    '            End If
    '            _btnMoverU_0.Enabled = True
    '            _btnMoverU_1.Enabled = True
    '        Case 3
    '            For I = 1 To _lstUsuarios_0.Items.Count
    '                Item = _lstUsuarios_1.Items.Add(_lstUsuarios_0.Items.Item(I).Name, _lstUsuarios_0.Items.Item(I).Text, "")
    '            Next I
    '            _lstUsuarios_0.Items.Clear()
    '            _btnMoverU_0.Enabled = True
    '            _btnMoverU_1.Enabled = True
    '            _btnMoverU_2.Enabled = False
    '            _btnMoverU_3.Enabled = False
    '    End Select
    'End Sub

    'Private Sub btnMoverU_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnMoverU.Enter
    '    Dim Index As Integer
    '    '= btnMoverU.GetIndex(eventSender)
    '    Pon_Tool()
    'End Sub



    'Private Sub optTipo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTipo.CheckedChanged
    '    If eventSender.Checked Then
    '        Dim Index As Integer
    '        '= optTipo.GetIndex(eventSender)
    '        Call ActivaCtl()
    '        Limpiar()
    '        Me.optTipo(Index).Focus()
    '    End If
    'End Sub

    'Private Sub optTipo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTipo.Enter
    '    Dim Index As Integer
    '    '= optTipo.GetIndex(eventSender)
    '    Pon_Tool()
    'End Sub

    Private Sub _btnMoverU_0_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverU_0.Click
        Dim Index As Integer = 0
        '= _btnMoverU_0.GetIndex(eventSender)
        Dim I As Integer
        Select Case Index
            Case 0
                ModEstandar.BorraCmd()
                gStrSql = "Select CodUsuario, Nombre From CatUsuarios Where Grupo = 0 And (CodGrupo = 0 Or CodGrupo = " & Numerico((Me.txtCodigo.Text)) & ")"
                Cmd.CommandText = "dbo.UP_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute
                _lstUsuarios_0.Items.Clear()
                With Me
                    For I = 1 To RsGral.RecordCount - 1
                        ' Al asignar el valor a KEY le agrego una letra porque no acepta solamente numeros
                        Item = ._lstUsuarios_0.Items.Add("C" & CStr(RsGral.Fields("CodUsuario").Value), Trim(RsGral.Fields("Nombre").Value), "")
                        RsGral.MoveNext()
                    Next
                End With
                _lstUsuarios_1.Items.Clear()
                _btnMoverU_0.Enabled = False
                _btnMoverU_1.Enabled = False
                _btnMoverU_2.Enabled = True
                _btnMoverU_3.Enabled = True
            Case 1
                If _lstUsuarios_1.Items.Count > 0 Then
                    Item = _lstUsuarios_0.Items.Add(_lstUsuarios_1.FocusedItem.Name, _lstUsuarios_1.FocusedItem.Text, "")
                    _lstUsuarios_1.Items.RemoveAt((_lstUsuarios_1.FocusedItem.Index))
                End If
                If _lstUsuarios_1.Items.Count > 0 Then
                    _btnMoverU_0.Enabled = True
                    _btnMoverU_1.Enabled = True
                Else
                    _btnMoverU_0.Enabled = False
                    _btnMoverU_1.Enabled = False
                End If
                _btnMoverU_2.Enabled = True
                _btnMoverU_3.Enabled = True
            Case 2
                If _lstUsuarios_0.Items.Count > 0 Then
                    Item = _lstUsuarios_1.Items.Add(_lstUsuarios_0.FocusedItem.Name, _lstUsuarios_0.FocusedItem.Text, "")
                    _lstUsuarios_0.Items.RemoveAt((_lstUsuarios_0.FocusedItem.Index))
                End If
                If _lstUsuarios_0.Items.Count > 0 Then
                    _btnMoverU_2.Enabled = True
                    _btnMoverU_3.Enabled = True
                Else
                    _btnMoverU_2.Enabled = False
                    _btnMoverU_3.Enabled = False
                End If
                _btnMoverU_0.Enabled = True
                _btnMoverU_1.Enabled = True
            Case 3
                For I = 1 To _lstUsuarios_0.Items.Count - 1
                    Item = _lstUsuarios_1.Items.Add(_lstUsuarios_0.Items.Item(I).Name, _lstUsuarios_0.Items.Item(I).Text, "")
                Next
                _lstUsuarios_0.Items.Clear()
                _btnMoverU_0.Enabled = True
                _btnMoverU_1.Enabled = True
                _btnMoverU_2.Enabled = False
                _btnMoverU_3.Enabled = False
        End Select
    End Sub

    Private Sub _btnMoverU_0_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverU_0.Enter
        Dim Index As Integer
        '= _btnMoverU_0.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub _btnMoverU_1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverU_1.Click
        Dim Index As Integer = 1
        '= _btnMoverU_1.GetIndex(eventSender)
        Dim I As Integer
        Select Case Index
            Case 0
                ModEstandar.BorraCmd()
                gStrSql = "Select CodUsuario, Nombre From CatUsuarios Where Grupo = 0 And (CodGrupo = 0 Or CodGrupo = " & Numerico((Me.txtCodigo.Text)) & ")"
                Cmd.CommandText = "dbo.UP_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute
                _lstUsuarios_0.Items.Clear()
                With Me
                    For I = 1 To RsGral.RecordCount - 1
                        ' Al asignar el valor a KEY le agrego una letra porque no acepta solamente numeros
                        Item = ._lstUsuarios_0.Items.Add("C" & CStr(RsGral.Fields("CodUsuario").Value), Trim(RsGral.Fields("Nombre").Value), "")
                        RsGral.MoveNext()
                    Next
                End With
                _lstUsuarios_1.Items.Clear()
                _btnMoverU_0.Enabled = False
                _btnMoverU_1.Enabled = False
                _btnMoverU_2.Enabled = True
                _btnMoverU_3.Enabled = True
            Case 1
                If _lstUsuarios_1.Items.Count > 0 Then
                    Item = _lstUsuarios_0.Items.Add(_lstUsuarios_1.FocusedItem.Name, _lstUsuarios_1.FocusedItem.Text, "")
                    _lstUsuarios_1.Items.RemoveAt((_lstUsuarios_1.FocusedItem.Index))
                End If
                If _lstUsuarios_1.Items.Count > 0 Then
                    _btnMoverU_0.Enabled = True
                    _btnMoverU_1.Enabled = True
                Else
                    _btnMoverU_0.Enabled = False
                    _btnMoverU_1.Enabled = False
                End If
                _btnMoverU_2.Enabled = True
                _btnMoverU_3.Enabled = True
            Case 2
                If _lstUsuarios_0.Items.Count > 0 Then
                    Item = _lstUsuarios_1.Items.Add(_lstUsuarios_0.FocusedItem.Name, _lstUsuarios_0.FocusedItem.Text, "")
                    _lstUsuarios_0.Items.RemoveAt((_lstUsuarios_0.FocusedItem.Index))
                End If
                If _lstUsuarios_0.Items.Count > 0 Then
                    _btnMoverU_2.Enabled = True
                    _btnMoverU_3.Enabled = True
                Else
                    _btnMoverU_2.Enabled = False
                    _btnMoverU_3.Enabled = False
                End If
                _btnMoverU_0.Enabled = True
                _btnMoverU_1.Enabled = True
            Case 3
                For I = 1 To _lstUsuarios_0.Items.Count - 1
                    Item = _lstUsuarios_1.Items.Add(_lstUsuarios_0.Items.Item(I).Name, _lstUsuarios_0.Items.Item(I).Text, "")
                Next
                _lstUsuarios_0.Items.Clear()
                _btnMoverU_0.Enabled = True
                _btnMoverU_1.Enabled = True
                _btnMoverU_2.Enabled = False
                _btnMoverU_3.Enabled = False
        End Select
    End Sub

    Private Sub _btnMoverU_1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverU_1.Enter
        Dim Index As Integer
        '= _btnMoverU_1.GetIndex(eventSender)
        Pon_Tool()
    End Sub



    Private Sub _btnMoverU_2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverU_2.Click
        Dim Index As Integer = 2
        '= _btnMoverU_2.GetIndex(eventSender)
        Dim I As Integer
        Select Case Index
            Case 0
                ModEstandar.BorraCmd()
                gStrSql = "Select CodUsuario, Nombre From CatUsuarios Where Grupo = 0 And (CodGrupo = 0 Or CodGrupo = " & Numerico((Me.txtCodigo.Text)) & ")"
                Cmd.CommandText = "dbo.UP_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute
                _lstUsuarios_0.Items.Clear()
                With Me
                    For I = 1 To RsGral.RecordCount - 1
                        ' Al asignar el valor a KEY le agrego una letra porque no acepta solamente numeros
                        Item = ._lstUsuarios_0.Items.Add("C" & CStr(RsGral.Fields("CodUsuario").Value), Trim(RsGral.Fields("Nombre").Value), "")
                        RsGral.MoveNext()
                    Next
                End With
                _lstUsuarios_1.Items.Clear()
                _btnMoverU_0.Enabled = False
                _btnMoverU_1.Enabled = False
                _btnMoverU_2.Enabled = True
                _btnMoverU_3.Enabled = True
            Case 1
                If _lstUsuarios_1.Items.Count > 0 Then
                    Item = _lstUsuarios_0.Items.Add(_lstUsuarios_1.FocusedItem.Name, _lstUsuarios_1.FocusedItem.Text, "")
                    _lstUsuarios_1.Items.RemoveAt((_lstUsuarios_1.FocusedItem.Index))
                End If
                If _lstUsuarios_1.Items.Count > 0 Then
                    _btnMoverU_0.Enabled = True
                    _btnMoverU_1.Enabled = True
                Else
                    _btnMoverU_0.Enabled = False
                    _btnMoverU_1.Enabled = False
                End If
                _btnMoverU_2.Enabled = True
                _btnMoverU_3.Enabled = True
            Case 2
                If _lstUsuarios_0.Items.Count > 0 Then
                    Item = _lstUsuarios_1.Items.Add(_lstUsuarios_0.FocusedItem.Name, _lstUsuarios_0.FocusedItem.Text, "")
                    _lstUsuarios_0.Items.RemoveAt((_lstUsuarios_0.FocusedItem.Index))
                End If
                If _lstUsuarios_0.Items.Count > 0 Then
                    _btnMoverU_2.Enabled = True
                    _btnMoverU_3.Enabled = True
                Else
                    _btnMoverU_2.Enabled = False
                    _btnMoverU_3.Enabled = False
                End If
                _btnMoverU_0.Enabled = True
                _btnMoverU_1.Enabled = True
            Case 3
                For I = 1 To _lstUsuarios_0.Items.Count - 1
                    Item = _lstUsuarios_1.Items.Add(_lstUsuarios_0.Items.Item(I).Name, _lstUsuarios_0.Items.Item(I).Text, "")
                Next
                _lstUsuarios_0.Items.Clear()
                _btnMoverU_0.Enabled = True
                _btnMoverU_1.Enabled = True
                _btnMoverU_2.Enabled = False
                _btnMoverU_3.Enabled = False
        End Select
    End Sub

    Private Sub _btnMoverU_2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverU_2.Enter
        Dim Index As Integer
        '= _btnMoverU_2.GetIndex(eventSender)
        Pon_Tool()
    End Sub


    Private Sub _btnMoverU_3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverU_3.Click
        Dim Index As Integer = 3
        '= _btnMoverU_3.GetIndex(eventSender)
        Dim I As Integer
        Select Case Index
            Case 0
                ModEstandar.BorraCmd()
                gStrSql = "Select CodUsuario, Nombre From CatUsuarios Where Grupo = 0 And (CodGrupo = 0 Or CodGrupo = " & Numerico((Me.txtCodigo.Text)) & ")"
                Cmd.CommandText = "dbo.UP_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute
                _lstUsuarios_0.Items.Clear()
                With Me
                    For I = 1 To RsGral.RecordCount - 1
                        ' Al asignar el valor a KEY le agrego una letra porque no acepta solamente numeros
                        Item = ._lstUsuarios_0.Items.Add("C" & CStr(RsGral.Fields("CodUsuario").Value), Trim(RsGral.Fields("Nombre").Value), "")
                        RsGral.MoveNext()
                    Next
                End With
                _lstUsuarios_1.Items.Clear()
                _btnMoverU_0.Enabled = False
                _btnMoverU_1.Enabled = False
                _btnMoverU_2.Enabled = True
                _btnMoverU_3.Enabled = True
            Case 1
                If _lstUsuarios_1.Items.Count > 0 Then
                    Item = _lstUsuarios_0.Items.Add(_lstUsuarios_1.FocusedItem.Name, _lstUsuarios_1.FocusedItem.Text, "")
                    _lstUsuarios_1.Items.RemoveAt((_lstUsuarios_1.FocusedItem.Index))
                End If
                If _lstUsuarios_1.Items.Count > 0 Then
                    _btnMoverU_0.Enabled = True
                    _btnMoverU_1.Enabled = True
                Else
                    _btnMoverU_0.Enabled = False
                    _btnMoverU_1.Enabled = False
                End If
                _btnMoverU_2.Enabled = True
                _btnMoverU_3.Enabled = True
            Case 2
                If _lstUsuarios_0.Items.Count > 0 Then
                    Item = _lstUsuarios_1.Items.Add(_lstUsuarios_0.FocusedItem.Name, _lstUsuarios_0.FocusedItem.Text, "")
                    _lstUsuarios_0.Items.RemoveAt((_lstUsuarios_0.FocusedItem.Index))
                End If
                If _lstUsuarios_0.Items.Count > 0 Then
                    _btnMoverU_2.Enabled = True
                    _btnMoverU_3.Enabled = True
                Else
                    _btnMoverU_2.Enabled = False
                    _btnMoverU_3.Enabled = False
                End If
                _btnMoverU_0.Enabled = True
                _btnMoverU_1.Enabled = True
            Case 3
                For I = 1 To _lstUsuarios_0.Items.Count - 1
                    Item = _lstUsuarios_1.Items.Add(_lstUsuarios_0.Items.Item(I).Name, _lstUsuarios_0.Items.Item(I).Text, "")
                Next
                _lstUsuarios_0.Items.Clear()
                _btnMoverU_0.Enabled = True
                _btnMoverU_1.Enabled = True
                _btnMoverU_2.Enabled = False
                _btnMoverU_3.Enabled = False
        End Select
    End Sub

    Private Sub _btnMoverU_3_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _btnMoverU_3.Enter
        Dim Index As Integer
        '= _btnMoverU_3.GetIndex(eventSender)
        Pon_Tool()
    End Sub



    Private Sub _optTipo_0_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optTipo_0.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer
            '=  _optTipo_0.GetIndex(eventSender)
            ActivaCtl()
            Limpiar()
            Me._optTipo_0.Focus()
        End If
    End Sub

    Private Sub _optTipo_0_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optTipo_0.Enter
        Dim Index As Integer
        '=  _optTipo_0.GetIndex(eventSender)
        Pon_Tool()
    End Sub


    Private Sub _optTipo_1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optTipo_1.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer
            '= _optTipo_1.GetIndex(eventSender)
            ActivaCtl()
            Limpiar()
            Me._optTipo_1.Focus()
        End If
    End Sub

    Private Sub _optTipo_1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optTipo_1.Enter
        Dim Index As Integer
        '= _optTipo_1.GetIndex(eventSender)
        Pon_Tool()
    End Sub



    Private Sub optTipoUsuario_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTipoUsuario.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer
            '= optTipoUsuario.GetIndex(eventSender)
            If Me._optTipoUsuario_0.Checked Then
                cTipoUsuario = C_TADMIN
            ElseIf Me._optTipoUsuario_1.Checked Then
                cTipoUsuario = C_TSUPERVISOR
            Else
                cTipoUsuario = C_TEMPLEADO
            End If
        End If
    End Sub

    Private Sub SSTabSeg_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SSTabSeg.SelectedIndexChanged
        Static PreviousTab As Integer
        '= SSTabSeg.SelectedIndex()
        Select Case Me.SSTabSeg.SelectedIndex
            Case nGRUPOUSUARIO
                Me.ToolTip1.SetToolTip(Me.SSTabSeg, "Grupos y Usuarios")
            Case nPRIVILEGIOS
                Me.ToolTip1.SetToolTip(Me.SSTabSeg, "Configuración de Privilegios")
        End Select
        Limpiar()
        PreviousTab = SSTabSeg.SelectedIndex()
    End Sub

    Private Sub SSTabSeg_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SSTabSeg.Enter
        Pon_Tool()
    End Sub

    Private Sub SSTabSeg_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles SSTabSeg.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Or KeyCode = System.Windows.Forms.Keys.Tab Then
            Select Case Me.SSTabSeg.SelectedIndex
                Case nGRUPOUSUARIO
                    Me.txtCodigo.Focus()
                Case nPRIVILEGIOS
                    If Me._optTipo_0.Checked Then
                        Me._dbcGrupos_1.Focus()
                    Else
                        Me.dbcUsuarios.Focus()
                    End If
            End Select
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
            If Me._optTipo_0.Checked Then
                Me._optTipo_0.Focus()
            Else
                Me._optTipo_1.Focus()
            End If
        End If
    End Sub

    Private Sub txtCodigo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodigo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.Enter
        strControlActual = UCase("txtCodigo")
        SelTextoTxt((Me.txtCodigo))
        Pon_Tool()
    End Sub

    Private Sub txtCodigo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodigo.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Pregunta sólo en caso de que existan cambios en la clave (esto es, cuando se teclea una clave diferente a la actual)
        If Cambios() And KeyCode = System.Windows.Forms.Keys.Delete Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        KeyCode = 0
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se borre el contenido del text
                Case MsgBoxResult.Cancel
                    KeyCode = 0
                    Me.txtCodigo.Focus()
            End Select
        End If
    End Sub

    Private Sub txtCodigo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back Then
            KeyAscii = 0
        Else
            'Pregunta sólo si ha habido cambios
            If Cambios() And Not mblnNuevo Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        If Not Guardar() Then
                            KeyAscii = 0
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        KeyAscii = 0
                        Me.txtCodigo.Focus()
                End Select
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodigo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.Leave
        'If System.Windows.Forms.Form.ActiveForm.Text = Me.Text Then
        If mblnCambiosEnCodigo = True Then 'Si hubo cambios en el código hace la consulta
            LlenaDatos()
        End If
        'End If
    End Sub

    Private Sub txtConfirmar_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConfirmar.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtConfirmar)
    End Sub

    Private Sub txtNombre_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.Enter
        strControlActual = UCase("txtNombre")
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtNombre)
    End Sub

    Private Sub txtPassWord_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPassWord.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtPassWord)
    End Sub



    '    Private Sub dbcGrupos_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcGrupos.CursorChanged
    '        Dim Index As Integer
    '        '= dbcGrupos.GetIndex(eventSender)
    '        On Error GoTo MError
    '        Dim lStrSql As String

    '        If mblnFueraChange Then Exit Sub
    '        lStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 1 and Nombre LIKE '" & Trim(Me.dbcGrupos.Text) & "%'"
    '        ModDCombo.DCChange(lStrSql, tecla, dbcGrupos)

    '        If dbcGrupos.Text = "" Then
    '            If Me.SSTabSeg.SelectedIndex = nPRIVILEGIOS Then
    '                mblnFueraChange = True
    '                Me.dbcModulo.Text = ""
    '                Me.dbcModulo.Tag = ""
    '                mintCodModulo = 0
    '                mblnFueraChange = False
    '                Me._lstPrivilegios_0.Items.Clear()
    '                Me._lstPrivilegios_1.Items.Clear()
    '            End If
    '        End If
    'MError:
    '        If Err.Number <> 0 Then
    '            ModEstandar.MostrarError()
    '        End If
    '    End Sub

    '    Private Sub dbcGrupos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcGrupos.Enter
    '        Dim Index As Integer
    '        '= dbcGrupos.GetIndex(eventSender)
    '        Pon_Tool()
    '        gStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 1 ORDER BY Nombre"
    '        ModDCombo.DCGotFocus(gStrSql, dbcGrupos)
    '    End Sub

    '    Private Sub dbcGrupos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcGrupos.KeyDown
    '        Dim Index As Integer
    '        '= dbcGrupos.GetIndex(eventSender)
    '        Dim Aux As String
    '        Select Case eventArgs.KeyCode
    '            Case System.Windows.Forms.Keys.Escape
    '                If Index = nGRUPOUSUARIO Then
    '                    Me.chkGrupo.Focus()
    '                Else
    '                    Me.SSTabSeg.Focus()
    '                End If
    '            Case System.Windows.Forms.Keys.Return
    '                Aux = Trim(Me.dbcGrupos.Text)
    '                If Me.dbcGrupos.SelectedItem <> 0 Then
    '                    dbcGrupos_Leave(dbcGrupos, New System.EventArgs())
    '                End If
    '                Me.dbcGrupos.Text = Aux
    '                Exit Sub
    '            Case System.Windows.Forms.Keys.Tab
    '                Aux = Trim(Me.dbcGrupos.Text)
    '                If Me.dbcGrupos.SelectedItem <> 0 Then
    '                    dbcGrupos_Leave(dbcGrupos, New System.EventArgs())
    '                End If
    '                Me.dbcGrupos.Text = Aux
    '                Exit Sub
    '        End Select
    '        tecla = eventArgs.KeyCode
    '    End Sub

    '    Private Sub dbcGrupos_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcGrupos.Leave
    '        Dim Index As Integer
    '        '= dbcGrupos.GetIndex(eventSender)
    '        Dim I As Integer
    '        Dim Aux As Integer
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
    '            Exit Sub
    '        End If
    '        gStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 1 and Nombre LIKE '" & Trim(Me.dbcGrupos.Text) & "%'"
    '        Aux = mintCodGrupo
    '        mintCodGrupo = 0
    '        ModDCombo.DCLostFocus(dbcGrupos, gStrSql, mintCodGrupo)
    '        If Aux <> mintCodGrupo Then
    '            If Me.SSTabSeg.SelectedIndex = nPRIVILEGIOS Then
    '                mblnFueraChange = True
    '                Me.dbcModulo.Text = ""
    '                Me.dbcModulo.Tag = ""
    '                mintCodModulo = 0
    '                mblnFueraChange = False
    '                Me._lstPrivilegios_0.Items.Clear()
    '                Me._lstPrivilegios_1.Items.Clear()
    '            End If
    '        End If
    '    End Sub

    '    Private Sub dbcGrupos_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcGrupos.MouseUp
    '        Dim Index As Integer
    '        '= dbcGrupos.GetIndex(eventSender)
    '        Dim Aux As String
    '        Aux = Trim(Me.dbcGrupos.Text)
    '        If Me.dbcGrupos.SelectedItem <> 0 Then
    '            dbcGrupos_Leave(dbcGrupos, New System.EventArgs())
    '        End If
    '        Me.dbcGrupos.Text = Aux
    '    End Sub



    Private Sub _dbcGrupos_0_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _dbcGrupos_0.CursorChanged
        Dim Index As Integer
        '= _dbcGrupos_0.GetIndex(eventSender)
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        lStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 1 and Nombre LIKE '" & Trim(Me._dbcGrupos_0.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, _dbcGrupos_0)

        If _dbcGrupos_0.Text = "" Then
            If Me.SSTabSeg.SelectedIndex = nPRIVILEGIOS Then
                mblnFueraChange = True
                Me.dbcModulo.Text = ""
                Me.dbcModulo.Tag = ""
                mintCodModulo = 0
                mblnFueraChange = False
                Me._lstPrivilegios_0.Items.Clear()
                Me._lstPrivilegios_1.Items.Clear()
            End If
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub _dbcGrupos_0_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _dbcGrupos_0.Enter
        Dim Index As Integer
        '= _dbcGrupos_0.GetIndex(eventSender)
        Pon_Tool()
        gStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 1 ORDER BY Nombre"
        ModDCombo.DCGotFocus(gStrSql, _dbcGrupos_0)
    End Sub

    Private Sub _dbcGrupos_0_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles _dbcGrupos_0.KeyDown
        Dim Index As Integer
        '= _dbcGrupos_0.GetIndex(eventSender)
        Dim Aux As String
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                If Index = nGRUPOUSUARIO Then
                    Me.chkGrupo.Focus()
                Else
                    Me.SSTabSeg.Focus()
                End If
            Case System.Windows.Forms.Keys.Return
                Aux = Trim(Me._dbcGrupos_0.Text)
                If Me._dbcGrupos_0.SelectedItem <> 0 Then
                    _dbcGrupos_0_Leave(_dbcGrupos_0, New System.EventArgs())
                End If
                Me._dbcGrupos_0.Text = Aux
                Exit Sub
            Case System.Windows.Forms.Keys.Tab
                Aux = Trim(Me._dbcGrupos_0.Text)
                If Me._dbcGrupos_0.SelectedItem <> 0 Then
                    _dbcGrupos_0_Leave(_dbcGrupos_0, New System.EventArgs())
                End If
                Me._dbcGrupos_0.Text = Aux
                Exit Sub
        End Select
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub _dbcGrupos_0_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _dbcGrupos_0.Leave
        Dim Index As Integer
        '= _dbcGrupos_0.GetIndex(eventSender)
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 1 and Nombre LIKE '" & Trim(Me._dbcGrupos_0.Text) & "%'"
        Aux = mintCodGrupo
        mintCodGrupo = 0
        ModDCombo.DCLostFocus(_dbcGrupos_0, gStrSql, mintCodGrupo)
        If Aux <> mintCodGrupo Then
            If Me.SSTabSeg.SelectedIndex = nPRIVILEGIOS Then
                mblnFueraChange = True
                Me.dbcModulo.Text = ""
                Me.dbcModulo.Tag = ""
                mintCodModulo = 0
                mblnFueraChange = False
                Me._lstPrivilegios_0.Items.Clear()
                Me._lstPrivilegios_1.Items.Clear()
            End If
        End If
    End Sub

    Private Sub _dbcGrupos_0_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles _dbcGrupos_0.MouseUp
        Dim Index As Integer
        '= _dbcGrupos_0.GetIndex(eventSender)
        Dim Aux As String
        Aux = Trim(Me._dbcGrupos_0.Text)
        'If Me._dbcGrupos_0.SelectedItem <> 0 Then
        '    _dbcGrupos_0_Leave(_dbcGrupos_0, New System.EventArgs())
        'End If
        Me._dbcGrupos_0.Text = Aux
    End Sub


    Private Sub _dbcGrupos_1_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _dbcGrupos_1.CursorChanged
        Dim Index As Integer
        '= _dbcGrupos_1.GetIndex(eventSender)
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        lStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 1 and Nombre LIKE '" & Trim(Me._dbcGrupos_1.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, _dbcGrupos_1)

        If _dbcGrupos_1.Text = "" Then
            If Me.SSTabSeg.SelectedIndex = nPRIVILEGIOS Then
                mblnFueraChange = True
                Me.dbcModulo.Text = ""
                Me.dbcModulo.Tag = ""
                mintCodModulo = 0
                mblnFueraChange = False
                Me._lstPrivilegios_0.Items.Clear()
                Me._lstPrivilegios_1.Items.Clear()
            End If
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub _dbcGrupos_1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _dbcGrupos_1.Enter
        Dim Index As Integer
        '= _dbcGrupos_1.GetIndex(eventSender)
        Pon_Tool()
        gStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 1 ORDER BY Nombre"
        ModDCombo.DCGotFocus(gStrSql, _dbcGrupos_1)
    End Sub

    Private Sub _dbcGrupos_1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles _dbcGrupos_1.KeyDown
        Dim Index As Integer
        '= _dbcGrupos_1.GetIndex(eventSender)
        Dim Aux As String
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                If Index = nGRUPOUSUARIO Then
                    Me.chkGrupo.Focus()
                Else
                    Me.SSTabSeg.Focus()
                End If
            Case System.Windows.Forms.Keys.Return
                Aux = Trim(Me._dbcGrupos_1.Text)
                If Me._dbcGrupos_1.SelectedItem <> 0 Then
                    _dbcGrupos_1_Leave(_dbcGrupos_1, New System.EventArgs())
                End If
                Me._dbcGrupos_1.Text = Aux
                Exit Sub
            Case System.Windows.Forms.Keys.Tab
                Aux = Trim(Me._dbcGrupos_1.Text)
                If Me._dbcGrupos_1.SelectedItem <> 0 Then
                    _dbcGrupos_1_Leave(_dbcGrupos_1, New System.EventArgs())
                End If
                Me._dbcGrupos_1.Text = Aux
                Exit Sub
        End Select
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub _dbcGrupos_1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _dbcGrupos_1.Leave
        Dim Index As Integer
        '= _dbcGrupos_1.GetIndex(eventSender)
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codUsuario, rtrim(ltrim(Nombre)) as Nombre FROM catUsuarios WHERE Grupo = 1 and Nombre LIKE '" & Trim(Me._dbcGrupos_1.Text) & "%'"
        Aux = mintCodGrupo
        mintCodGrupo = 0
        ModDCombo.DCLostFocus(_dbcGrupos_1, gStrSql, mintCodGrupo)
        If Aux <> mintCodGrupo Then
            If Me.SSTabSeg.SelectedIndex = nPRIVILEGIOS Then
                mblnFueraChange = True
                Me.dbcModulo.Text = ""
                Me.dbcModulo.Tag = ""
                mintCodModulo = 0
                mblnFueraChange = False
                Me._lstPrivilegios_0.Items.Clear()
                Me._lstPrivilegios_1.Items.Clear()
            End If
        End If
    End Sub

    Private Sub _dbcGrupos_1_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles _dbcGrupos_1.MouseUp
        Dim Index As Integer
        '= _dbcGrupos_1.GetIndex(eventSender)
        Dim Aux As String
        Aux = Trim(Me._dbcGrupos_1.Text)
        'If Me._dbcGrupos_1.SelectedItem <> 0 Then
        '    _dbcGrupos_1_Leave(_dbcGrupos_1, New System.EventArgs())
        'End If
        Me._dbcGrupos_1.Text = Aux
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        Eliminar()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

End Class