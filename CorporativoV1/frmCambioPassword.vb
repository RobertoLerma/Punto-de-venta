Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCambioPassword
    Inherits System.Windows.Forms.Form

    Dim esNuevo As Boolean = True
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtAnterior As System.Windows.Forms.TextBox
    Public WithEvents txtConfirmacion As System.Windows.Forms.TextBox
    Public WithEvents txtNuevoPassword As System.Windows.Forms.TextBox
    Public WithEvents txtCodUsuario As System.Windows.Forms.TextBox
    Public WithEvents _label1_3 As System.Windows.Forms.Label
    Public WithEvents Image1 As System.Windows.Forms.PictureBox
    Public WithEvents _label1_2 As System.Windows.Forms.Label
    Public WithEvents _label1_1 As System.Windows.Forms.Label
    Public WithEvents _label1_0 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Dim Validar As Boolean
    Dim cambio As Boolean
    Dim Cancelar As Boolean
    Dim Psw As String
    Public WithEvents Panel3 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Dim BC As Boolean

    Private Sub frmCambioPassword_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Error_Renamed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        BringToFront()
        txtCodUsuario.Text = Trim(ModVariables.gStrNomUsuario)
Error_Renamed:
        If Err.Number <> 0 Then ModErrores.Errores()
    End Sub

    Private Sub frmCambioPassword_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case Keys.Enter = 13
                txtAnterior_Leave(New Object, New EventArgs)
            Case System.Windows.Forms.Keys.Return
                If ActiveControl.Name <> "msgEtiquetas" Then
                    ModEstandar.AvanzarTab(Me)
                End If
            Case System.Windows.Forms.Keys.Escape
                If ActiveControl.Name <> "txtDetalle" Then
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCambioPassword_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        CentrarForma(Me)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
    End Sub

    Private Sub frmCambioPassword_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '        On Error GoTo Error_Renamed
        '        If cambio = True Then
        '            Select Case MsgBox("¿Desea Grabar la Información?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "Aviso")
        '                Case MsgBoxResult.Cancel 'Cancelar
        '                    Cancelar = 1
        '                Case MsgBoxResult.Yes 'Yes
        '                    Guardar()
        '                Case MsgBoxResult.No 'No
        '                    cambio = False
        '            End Select
        '        End If
        '        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        '        ModEstandar.LimpiaDescBarraEstado()
        '        'Me = Nothing
        '        IsNothing(Me)
        'Error_Renamed:
        '        If Err.Number <> 0 Then
        '            ModErrores.Errores()
        '        End If
    End Sub

    Private Sub txtAnterior_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnterior.Enter
        txtAnterior.SelectionStart = 0
        txtAnterior.SelectionLength = Len(txtAnterior.Text)
        'MenuPrincipal.status.Items.Item(2).Text = ToolTip1.GetToolTip(txtAnterior)
    End Sub

    Private Sub txtAnterior_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAnterior.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        On Error GoTo Error_Renamed
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 39 Then
            KeyAscii = 180
        End If
        If KeyAscii = 13 Then
        End If

Error_Renamed:
        If Err.Number <> 0 Then
            ModErrores.Errores()
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtAnterior_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnterior.Leave
        If txtAnterior.Text <> "" Then
            Psw = ""
            Psw = ModEncriptacion.Encriptar(Trim(txtAnterior.Text))
            ModEstandar.BorraCmd()
            gStrSql = "select password from catusuarios" & " where nombre = '" & Trim(gStrNomUsuario) & "'" & " and password ='" & Trim(Psw) & "'"
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount <= 0 Then
                esNuevo = False
                MsgBox("Password Incorrecto ..." & vbNewLine & "Verifique Por Favor.", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Alerta")
                txtAnterior.SelectionStart = 0
                txtAnterior.SelectionLength = Len(txtAnterior.Text)
                txtAnterior.Focus()
            End If
        Else
            MsgBox("Debe de Proporcionar el Password Anterior", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Alerta")
            txtCodUsuario.Focus()
        End If
    End Sub

    Private Sub txtCodUsiario_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodUsuario.Enter
        ModEstandar.SelTextoTxt(txtCodUsuario)
        'MenuPrincipal.status.Items.Item(2).Text = ToolTip1.GetToolTip(txtCodUsiario)
    End Sub

    Private Sub txtCodUsiario_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodUsuario.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then Me.Close()
    End Sub

    Private Sub txtConfirmacion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConfirmacion.TextChanged
        cambio = True
    End Sub

    Private Sub txtConfirmacion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConfirmacion.Enter
        txtConfirmacion.SelectionStart = 0
        txtConfirmacion.SelectionLength = Len(txtConfirmacion.Text)
        'MenuPrincipal.status.Items.Item(2).Text = ToolTip1.GetToolTip(txtConfirmacion)
    End Sub

    Private Sub txtConfirmacion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConfirmacion.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Select Case KeyAscii
            Case 39
                KeyAscii = 180
            Case 13
                Guardar()
            Case 39, 42, 43, 44, 45, 46, 47 ' No permito las siguientes caracteres
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtNuevoPassword_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNuevoPassword.TextChanged
        cambio = True
    End Sub

    Private Sub txtNuevoPassword_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNuevoPassword.Enter
        txtNuevoPassword.SelectionStart = 0
        txtNuevoPassword.SelectionLength = Len(txtNuevoPassword.Text)
        'MenuPrincipal.status.Items.Item(2).Text = ToolTip1.GetToolTip(txtNuevoPassword)
    End Sub


    Private Sub txtNuevoPassword_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNuevoPassword.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Select Case KeyAscii
            Case 39
                KeyAscii = 180
            Case 13
                txtConfirmacion.Focus()
            Case 39, 42, 43, 44, 45, 46, 47 ' No permito las siguientes caracteres
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Public Sub Guardar()
        Validar = False
        Cancelar = False
        On Error GoTo Error_Renamed
        BC = False
        If cambio = True Then
            valida()
            If Validar = True Then
                Exit Sub
            End If
            Cnn.BeginTrans()
            BC = True
            ModStoredProcedures.PR_IMECatUsuarios(CStr(gIntCodUsuario), "", Trim(ModEncriptacion.Encriptar(txtConfirmacion.Text)), "False", "0", "", "0", C_MODIFICACION, CStr(7))
            Cmd.Execute()
            Cnn.CommitTrans()
            MsgBox(C_msgACTUALIZADO, vbInformation + vbOKOnly, ModVariables.gstrNombCortoEmpresa)
            BC = False
            cambio = False
            Me.Close()
        End If
Error_Renamed:
        If Err.Number <> 0 Then
            If BC = True Then Cnn.RollbackTrans()
            ModErrores.Errores()
        End If
    End Sub


    Public Sub Nuevo()
        txtCodUsuario.Text = ""
        txtAnterior.Text = ""
        txtNuevoPassword.Text = ""
        txtConfirmacion.Text = ""
    End Sub

    Public Sub valida()
        On Error GoTo Error_Renamed
        Validar = False

        If txtAnterior.Text = "" Then
            MsgBox("Debe de Proporcionar su Password Actual", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Alerta")
            txtAnterior.Focus()
            Validar = True
            Exit Sub
        Else
            Psw = ""
            Psw = ModEncriptacion.Encriptar(Trim(txtAnterior.Text))
            ModEstandar.BorraCmd()
            gStrSql = "select password from catusuarios where nombre = '" & Trim(gStrNomUsuario) & "' and password ='" & Trim(Psw) & "'"
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount = 0 Then
                MsgBox("Password Incorrecto ..." & Chr(13) & "Verifique Por Favor.", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Alerta")
                txtAnterior.SelectionStart = 0
                txtAnterior.SelectionLength = Len(txtAnterior.Text)
                txtAnterior.Focus()
                Validar = True
                Exit Sub
            End If
        End If

        If txtNuevoPassword.Text = "" Then
            MsgBox("Debe de Proporcionar el Nuevo Password", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Alerta")
            txtNuevoPassword.Focus()
            Validar = True
            Exit Sub
        End If

        If txtConfirmacion.Text = "" Then
            MsgBox("Debe de Proporcionar la Confirmación del Password", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Alerta")
            txtConfirmacion.Focus()
            Validar = True
            Exit Sub
        End If

        If Trim(txtNuevoPassword.Text) <> Trim(txtConfirmacion.Text) Then
            MsgBox("La Confirmación No Coincide ..." & Chr(13) & "Verifique Pot Favor.", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Alerta")
            txtConfirmacion.Focus()
            Validar = True
        End If
        Psw = ""
        Psw = ModEncriptacion.Encriptar(Trim(txtNuevoPassword.Text))
Error_Renamed:
        If Err.Number <> 0 Then
            ModErrores.Errores()
        End If
    End Sub

    Public Sub Cerrar()
        Me.Close()
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtAnterior = New System.Windows.Forms.TextBox()
        Me.txtConfirmacion = New System.Windows.Forms.TextBox()
        Me.txtNuevoPassword = New System.Windows.Forms.TextBox()
        Me.txtCodUsuario = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me._label1_3 = New System.Windows.Forms.Label()
        Me.Image1 = New System.Windows.Forms.PictureBox()
        Me._label1_2 = New System.Windows.Forms.Label()
        Me._label1_1 = New System.Windows.Forms.Label()
        Me._label1_0 = New System.Windows.Forms.Label()
        Me.label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.label1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtAnterior
        '
        Me.txtAnterior.AcceptsReturn = True
        Me.txtAnterior.BackColor = System.Drawing.SystemColors.Window
        Me.txtAnterior.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAnterior.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAnterior.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAnterior.Location = New System.Drawing.Point(162, 40)
        Me.txtAnterior.MaxLength = 20
        Me.txtAnterior.Name = "txtAnterior"
        Me.txtAnterior.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAnterior.Size = New System.Drawing.Size(154, 20)
        Me.txtAnterior.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtAnterior, "Password Actual del Usuario")
        '
        'txtConfirmacion
        '
        Me.txtConfirmacion.AcceptsReturn = True
        Me.txtConfirmacion.BackColor = System.Drawing.SystemColors.Window
        Me.txtConfirmacion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConfirmacion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConfirmacion.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtConfirmacion.Location = New System.Drawing.Point(162, 90)
        Me.txtConfirmacion.MaxLength = 20
        Me.txtConfirmacion.Name = "txtConfirmacion"
        Me.txtConfirmacion.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtConfirmacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConfirmacion.Size = New System.Drawing.Size(154, 20)
        Me.txtConfirmacion.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtConfirmacion, "Confirmación del  Nuevo Password")
        '
        'txtNuevoPassword
        '
        Me.txtNuevoPassword.AcceptsReturn = True
        Me.txtNuevoPassword.BackColor = System.Drawing.SystemColors.Window
        Me.txtNuevoPassword.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNuevoPassword.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNuevoPassword.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtNuevoPassword.Location = New System.Drawing.Point(162, 65)
        Me.txtNuevoPassword.MaxLength = 20
        Me.txtNuevoPassword.Name = "txtNuevoPassword"
        Me.txtNuevoPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNuevoPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNuevoPassword.Size = New System.Drawing.Size(154, 20)
        Me.txtNuevoPassword.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtNuevoPassword, "Nuevo Password del Usuario")
        '
        'txtCodUsuario
        '
        Me.txtCodUsuario.AcceptsReturn = True
        Me.txtCodUsuario.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodUsuario.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodUsuario.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodUsuario.Location = New System.Drawing.Point(162, 14)
        Me.txtCodUsuario.MaxLength = 0
        Me.txtCodUsuario.Name = "txtCodUsuario"
        Me.txtCodUsuario.ReadOnly = True
        Me.txtCodUsuario.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodUsuario.Size = New System.Drawing.Size(154, 20)
        Me.txtCodUsuario.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtCodUsuario, "Nombre del Usuario")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtAnterior)
        Me.Frame1.Controls.Add(Me.txtConfirmacion)
        Me.Frame1.Controls.Add(Me.txtNuevoPassword)
        Me.Frame1.Controls.Add(Me.txtCodUsuario)
        Me.Frame1.Controls.Add(Me._label1_3)
        Me.Frame1.Controls.Add(Me.Image1)
        Me.Frame1.Controls.Add(Me._label1_2)
        Me.Frame1.Controls.Add(Me._label1_1)
        Me.Frame1.Controls.Add(Me._label1_0)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(405, 119)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        '_label1_3
        '
        Me._label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._label1_3.Location = New System.Drawing.Point(46, 42)
        Me._label1_3.Name = "_label1_3"
        Me._label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._label1_3.Size = New System.Drawing.Size(108, 15)
        Me._label1_3.TabIndex = 2
        Me._label1_3.Text = "Password Actual :"
        '
        'Image1
        '
        Me.Image1.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.llave1
        Me.Image1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image1.ErrorImage = Nothing
        Me.Image1.Location = New System.Drawing.Point(340, 65)
        Me.Image1.Name = "Image1"
        Me.Image1.Size = New System.Drawing.Size(47, 44)
        Me.Image1.TabIndex = 9
        Me.Image1.TabStop = False
        '
        '_label1_2
        '
        Me._label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._label1_2.Location = New System.Drawing.Point(65, 94)
        Me._label1_2.Name = "_label1_2"
        Me._label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._label1_2.Size = New System.Drawing.Size(91, 14)
        Me._label1_2.TabIndex = 4
        Me._label1_2.Text = "Confirmación :"
        '
        '_label1_1
        '
        Me._label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._label1_1.Location = New System.Drawing.Point(44, 69)
        Me._label1_1.Name = "_label1_1"
        Me._label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._label1_1.Size = New System.Drawing.Size(106, 14)
        Me._label1_1.TabIndex = 3
        Me._label1_1.Text = "Nuevo Password :"
        '
        '_label1_0
        '
        Me._label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._label1_0.Location = New System.Drawing.Point(93, 18)
        Me._label1_0.Name = "_label1_0"
        Me._label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._label1_0.Size = New System.Drawing.Size(65, 14)
        Me._label1_0.TabIndex = 1
        Me._label1_0.Text = "Usuario  :"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(8, 127)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(377, 74)
        Me.Panel3.TabIndex = 69
        '
        'btnSalir
        '
        Me.btnSalir.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.salir
        Me.btnSalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSalir.Location = New System.Drawing.Point(208, 14)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(50, 42)
        Me.btnSalir.TabIndex = 70
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.buscar
        Me.btnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBuscar.Location = New System.Drawing.Point(160, 14)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(50, 42)
        Me.btnBuscar.TabIndex = 67
        Me.btnBuscar.Text = " "
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.grabar
        Me.btnGuardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnGuardar.Location = New System.Drawing.Point(11, 14)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(50, 42)
        Me.btnGuardar.TabIndex = 64
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.nuevo
        Me.btnLimpiar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnLimpiar.Location = New System.Drawing.Point(110, 14)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(50, 42)
        Me.btnLimpiar.TabIndex = 66
        Me.btnLimpiar.Text = " "
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.Eliminar
        Me.btnEliminar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnEliminar.Location = New System.Drawing.Point(61, 14)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(50, 42)
        Me.btnEliminar.TabIndex = 65
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'frmCambioPassword
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(406, 210)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(299, 279)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCambioPassword"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Cambio de Password"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.label1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub
End Class