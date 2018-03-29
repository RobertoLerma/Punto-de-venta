Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmAutorizacionConfig
    Inherits System.Windows.Forms.Form
    'Este Formulario se usará cuando se requiera alguna autorizacion para modificar un parametro de la venta.
    'Los cuales fueron especificados en la configuracion tanto del corporativo, como del punto de venta

    Public components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents TxtClave As System.Windows.Forms.TextBox
    Public WithEvents TxtUsuario As System.Windows.Forms.TextBox
    Public WithEvents _LblEtiqueta_1 As System.Windows.Forms.Label
    Public WithEvents _LblEtiqueta_0 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents LblEtiqueta As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Dim mIntUsuarioAutorio As Integer
    Dim mIntContFallos As Integer
    Dim mstrNomUsuario As String

    Private Sub frmAutorizacionConfig_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '    ModEstandar.ActivaMenu C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO
        '    Icono Me, MenuPrincipal
        Me.BringToFront()
    End Sub

    Private Sub Form_Initialize_Renamed()
        mIntContFallos = 0
    End Sub
    Private Sub frmAutorizacionConfig_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Si es usuario presiona Escape , la forma se cierra
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            gblnSalioSinValidar = True
            gblnAutorizacionAceptada = False
            Me.Close()
        End If
    End Sub

    Private Sub frmAutorizacionConfig_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        '    Dim obj As Form
        '    For Each obj In Forms
        '        If obj.Name <> "MenuPrincipal" And obj.Name <> "frmautorizacionconfig" Then
        '            Unload obj
        '            Set obj = Nothing
        '        End If
        '    Next
        '    ModPuntoVenta.DesHabilitaMenuPrincipal
    End Sub

    Private Sub frmAutorizacionConfig_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If UnloadMode = 0 Then 'Si Cero significa que el usuario presiono cerrar en ,la parte superior del formulario
        '    gblnAutorizacionAceptada = False
        '    gblnSalioSinValidar = True
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmAutorizacionConfig_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '    ModEstandar.ActivaMenu C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO
        '    ModEstandar.LimpiaDescBarraEstado
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtClave_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtClave.Enter
        ' Asigno el valor del ToolTipText de la Descripción a Barra de Estado
        'MDIMenuPrincipalCorpo.status.Items.Item(2).Text = ToolTip1.GetToolTip(TxtClave)
        ' Selecciono la información del control
        TxtClave.SelectionStart = 0
        TxtClave.SelectionLength = Len(TxtClave.Text)
    End Sub

    Private Sub txtClave_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtClave.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        On Error GoTo Merr
        Dim StrPassW As String
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            ModEstandar.BorraCmd()
            gStrSql = "Select * From CatUsuarios Where Nombre = '" & Trim(TxtUsuario.Text) & "'"
            Cmd.CommandText = "Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount < 1 Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                MsgBox("Usuario no registrado", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
                TxtUsuario.Text = ""
                TxtClave.Text = ""
                TxtUsuario.Focus()
                GoTo EventExitSub
            Else
                mIntUsuarioAutorio = RsGral.Fields("CodUsuario").Value
                StrPassW = ModEncriptacion.Encriptar(Trim(TxtClave.Text))
                If Trim(RsGral.Fields("Password").Value) <> Trim(StrPassW) Then
                    mIntContFallos = mIntContFallos + 1
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    MsgBox("Clave de usuario incorrecta... " & vbNewLine & "Intento " & CStr(mIntContFallos), MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
                    TxtClave.Text = ""
                    If mIntContFallos = 3 Then
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                        mIntContFallos = 0
                        MsgBox("Acceso denegado", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)
                        gblnAutorizacionAceptada = False
                        gblnSalioSinValidar = False
                        Me.Close()
                    End If
                Else
                    If TieneDerechodeAutorizar() = True Then
                        gblnAutorizacionAceptada = True
                    Else
                        gblnAutorizacionAceptada = False
                    End If
                    gblnSalioSinValidar = False
                    Me.Close()
                End If
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        GoTo EventExitSub
Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        ModErrores.Errores()
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtUsuario_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUsuario.Enter
        On Error GoTo Errores
        ' Asigno el valor del ToolTipText de la Descripción a Barra de Estado
        'MDIMenuPrincipalCorpo.status.Items.Item(2).Text = ToolTip1.GetToolTip(TxtUsuario)

        ' Selecciono la información del control
        TxtUsuario.SelectionStart = 0
        TxtUsuario.SelectionLength = Len(TxtUsuario.Text)
Errores:
        If Err.Number <> 0 Then ModErrores.Errores()
    End Sub

    Private Sub txtUsuario_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUsuario.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send(vbTab)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub TxtUsuario_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles TxtUsuario.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo Merr
        If TxtUsuario.Text = "" Then
            MsgBox("Proporcione un usuario", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Aviso")
            Cancel = False
        Else
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            ModEstandar.BorraCmd()
            gStrSql = "Select * From CatUsuarios Where Nombre = '" & Trim(TxtUsuario.Text) & "'"
            Cmd.CommandText = "Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If RsGral.RecordCount < 1 Then
                MsgBox("Usuario no registrado", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
                TxtUsuario.Text = ""
                Cancel = True
            ElseIf RsGral.Fields("Grupo").Value Then
                MsgBox("El usuario es grupo..." & Chr(13) & "Proporcione un Nombre de Usuario Válido", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Aviso")
                TxtUsuario.Text = ""
                Cancel = True
            Else
                mstrNomUsuario = RsGral.Fields("Nombre").Value
                '           MenuPrincipal.Status.Panels(1) = mstrNomUsuario
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        GoTo EventExitSub
Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        ''MostrarError "Ocurrió un error al intentar validar el usuario"
        If Err.Number <> 0 Then ModErrores.Errores()
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Function TieneDerechodeAutorizar() As Boolean
        ''    'Esta función determina si el usuario especificado tiene derecho de autorizar algun evento a realizar en una Venta.
        ''    'Para que pueda Hacerlo el usuario deberá ser de Tipo Gerente en la Tabla de Usuarios.
        ''    'De lo contrario no podrá hacerlo
        ''    TieneDerechodeAutorizar = False
        ''    ModEstandar.BorraCmd
        ''    gStrSql = "Select CodUsuario , Nombre , Tipo from catUsuarios Where CodUsuario = " & mIntUsuarioAutorio & ""
        ''    Cmd.CommandText = "Up_Select_Datos"
        ''    Cmd.CommandType = adCmdStoredProc
        ''    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, gStrSql)
        ''    Set RsGral = Cmd.Execute
        ''    If RsGral.RecordCount > 0 Then
        ''        If RsGral!Tipo = "A" Then
        ''            TieneDerechodeAutorizar = True
        ''        End If
        ''    End If
        'Esta función determina si el usuario especificado tiene derecho de autorizar algun evento a realizar en una Venta.
        'Para que pueda Hacerlo el usuario deberá ser de Tipo Gerente en la Tabla de Usuarios.
        'De lo contrario no podrá hacerlo
        TieneDerechodeAutorizar = False
        ModEstandar.BorraCmd()
        gStrSql = "SELECT     U.CodUsuario, U.Nombre, Isnull(U.Tipo,'') as Tipo, Isnull(U1.Tipo,'') as TipoGrupo " & "FROM         dbo.CatUsuarios U LEFT OUTER JOIN " & "dbo.CatUsuarios U1 ON U.CodGrupo = U1.CodUsuario " & "Where U.CodUsuario = " & mIntUsuarioAutorio
        Cmd.CommandText = "Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If (Trim(RsGral.Fields("Tipo").Value) = "A" Or Trim(RsGral.Fields("Tipo").Value) = "S") Or (Trim(RsGral.Fields("TipoGrupo").Value) = "A" Or Trim(RsGral.Fields("TipoGrupo").Value) = "S") Then
                TieneDerechodeAutorizar = True
            End If
        End If
    End Function


    Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAutorizacionConfig))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.TxtClave = New System.Windows.Forms.TextBox
        Me.TxtUsuario = New System.Windows.Forms.TextBox
        Me._LblEtiqueta_1 = New System.Windows.Forms.Label
        Me._LblEtiqueta_0 = New System.Windows.Forms.Label
        Me.LblEtiqueta = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        CType(Me.LblEtiqueta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.ClientSize = New System.Drawing.Size(306, 87)
        Me.Location = New System.Drawing.Point(3, 15)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ControlBox = True
        Me.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmAutorizacionConfig"
        Me.Frame1.Size = New System.Drawing.Size(289, 79)
        Me.Frame1.Location = New System.Drawing.Point(8, 0)
        Me.Frame1.TabIndex = 2
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Enabled = True
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Visible = True
        Me.Frame1.Name = "Frame1"
        Me.TxtClave.AutoSize = False
        Me.TxtClave.Size = New System.Drawing.Size(154, 20)
        Me.TxtClave.IMEMode = System.Windows.Forms.ImeMode.Disable
        Me.TxtClave.Location = New System.Drawing.Point(61, 50)
        Me.TxtClave.Maxlength = 20
        Me.TxtClave.PasswordChar = ChrW(42)
        Me.TxtClave.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.TxtClave, "Clave de acceso del usuario")
        Me.TxtClave.AcceptsReturn = True
        Me.TxtClave.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.TxtClave.BackColor = System.Drawing.SystemColors.Window
        Me.TxtClave.CausesValidation = True
        Me.TxtClave.Enabled = True
        Me.TxtClave.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtClave.HideSelection = True
        Me.TxtClave.ReadOnly = False
        Me.TxtClave.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtClave.MultiLine = False
        Me.TxtClave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtClave.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.TxtClave.TabStop = True
        Me.TxtClave.Visible = True
        Me.TxtClave.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TxtClave.Name = "TxtClave"
        Me.TxtUsuario.AutoSize = False
        Me.TxtUsuario.Size = New System.Drawing.Size(200, 20)
        Me.TxtUsuario.Location = New System.Drawing.Point(61, 18)
        Me.TxtUsuario.Maxlength = 20
        Me.TxtUsuario.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.TxtUsuario, "Nombre del usuario")
        Me.TxtUsuario.AcceptsReturn = True
        Me.TxtUsuario.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.TxtUsuario.BackColor = System.Drawing.SystemColors.Window
        Me.TxtUsuario.CausesValidation = True
        Me.TxtUsuario.Enabled = True
        Me.TxtUsuario.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtUsuario.HideSelection = True
        Me.TxtUsuario.ReadOnly = False
        Me.TxtUsuario.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtUsuario.MultiLine = False
        Me.TxtUsuario.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtUsuario.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.TxtUsuario.TabStop = True
        Me.TxtUsuario.Visible = True
        Me.TxtUsuario.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TxtUsuario.Name = "TxtUsuario"
        Me._LblEtiqueta_1.Text = "Clave:"
        Me._LblEtiqueta_1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me._LblEtiqueta_1.Size = New System.Drawing.Size(65, 17)
        Me._LblEtiqueta_1.Location = New System.Drawing.Point(8, 51)
        Me._LblEtiqueta_1.TabIndex = 4
        Me._LblEtiqueta_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._LblEtiqueta_1.BackColor = System.Drawing.SystemColors.Control
        Me._LblEtiqueta_1.Enabled = True
        Me._LblEtiqueta_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._LblEtiqueta_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._LblEtiqueta_1.UseMnemonic = True
        Me._LblEtiqueta_1.Visible = True
        Me._LblEtiqueta_1.AutoSize = False
        Me._LblEtiqueta_1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._LblEtiqueta_1.Name = "_LblEtiqueta_1"
        Me._LblEtiqueta_0.Text = "Usuario:"
        Me._LblEtiqueta_0.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me._LblEtiqueta_0.Size = New System.Drawing.Size(65, 17)
        Me._LblEtiqueta_0.Location = New System.Drawing.Point(8, 18)
        Me._LblEtiqueta_0.TabIndex = 3
        Me._LblEtiqueta_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._LblEtiqueta_0.BackColor = System.Drawing.SystemColors.Control
        Me._LblEtiqueta_0.Enabled = True
        Me._LblEtiqueta_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._LblEtiqueta_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._LblEtiqueta_0.UseMnemonic = True
        Me._LblEtiqueta_0.Visible = True
        Me._LblEtiqueta_0.AutoSize = False
        Me._LblEtiqueta_0.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._LblEtiqueta_0.Name = "_LblEtiqueta_0"
        Me.Controls.Add(Frame1)
        Me.Frame1.Controls.Add(TxtClave)
        Me.Frame1.Controls.Add(TxtUsuario)
        Me.Frame1.Controls.Add(_LblEtiqueta_1)
        Me.Frame1.Controls.Add(_LblEtiqueta_0)
        Me.LblEtiqueta.SetIndex(_LblEtiqueta_1, CType(1, Short))
        Me.LblEtiqueta.SetIndex(_LblEtiqueta_0, CType(0, Short))
        CType(Me.LblEtiqueta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub


End Class