'**********************************************************************************************************************'
'*PROGRAMA: AUTORIZA Y RECIBE TRANSFERENCIAS JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB

Friend Class frmAutorizaRecibe_Transf
    Inherits System.Windows.Forms.Form
    'Este Formulario se usará cuando se requiera el dato RECIBIR en la entrada por transferencia para poder
    'confirmar dicha entrada.
    'Sólo validará que este registrado en el Abc de Usuarios y que su passWord sea correcto.

    Dim mIntUsuarioAutorio As Integer
    Dim mIntContFallos As Integer
    Dim mstrNomUsuario As String
    Public WithEvents Frame1 As GroupBox
    Public WithEvents TxtClave As TextBox
    Public WithEvents TxtUsuario As TextBox
    Public WithEvents _LblEtiqueta_1 As Label
    Public WithEvents _LblEtiqueta_0 As Label
    Dim cmd As Command

    Private Sub frmAutorizaRecibe_Transf_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        BringToFront()
        'TxtUsuario.Text = Trim(frmInvEntradaPorTransferencia.mstrNombreR)
    End Sub

    Private Sub Form_Initialize_Renamed()
        mIntContFallos = 0
    End Sub

    Private Sub frmAutorizaRecibe_Transf_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Si es usuario presiona Escape , la forma se cierra
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            gblnSalioSinValidar = True
            gblnAutorizacionAceptada = False
            Me.Close()
        End If
    End Sub

    Private Sub frmAutorizaRecibe_Transf_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        ModEstandar.CentrarForma(Me)
    End Sub

    Private Sub frmAutorizaRecibe_Transf_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If UnloadMode = 0 Then 'Si Cero significa que el usuario presiono cerrar en ,la parte superior del formulario
            gblnAutorizacionAceptada = False
            gblnSalioSinValidar = True
        Else
            'frmInvEntradaPorTransferencia.txtRecibe.Text = ""
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmAutorizaRecibe_Transf_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Me = Nothing
    End Sub

    Private Sub txtClave_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtClave.Enter
        ' Asigno el valor del ToolTipText de la Descripción a Barra de Estado
        'UPGRADE_WARNING: Lower bound of collection MenuPrincipal.Status.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        'MenuPrincipal.Status.Items.Item(2).Text = ToolTip1.GetToolTip(TxtClave)
        ModEstandar.SelTextoTxt(TxtClave)
    End Sub

    Private Sub txtClave_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtClave.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'On Error GoTo Merr
        Try
            Dim StrPassW As String
            KeyAscii = Asc(UCase(Chr(KeyAscii)))

            If KeyAscii = 13 Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                ModEstandar.BorraCmd()
                gStrSql = "Select * From CatUsuarios Where Nombre = '" & Trim(TxtUsuario.Text) & "'"
                cmd.CommandText = "Up_Select_Datos"
                cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                cmd.Parameters.Append(cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                cmd.Parameters.Append(cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = cmd.Execute
                If RsGral.RecordCount < 1 Then
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    MsgBox("Usuario no registrado", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
                    TxtUsuario.Text = ""
                    TxtClave.Text = ""
                    TxtUsuario.Focus()
                    'GoTo EventExitSub
                Else
                    mIntUsuarioAutorio = RsGral.Fields("CodUsuario").Value
                    StrPassW = ModEncriptacion.Encriptar(Trim(TxtClave.Text))
                    If Trim(RsGral.Fields("Password").Value) <> Trim(StrPassW) Then
                        mIntContFallos = mIntContFallos + 1
                        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                        MsgBox("Clave de usuario incorrecta... " & vbNewLine & "Intento " & CStr(mIntContFallos), MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
                        TxtClave.Text = ""
                        If mIntContFallos = 3 Then
                            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                            mIntContFallos = 0
                            MsgBox("Acceso denegado", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)
                            gblnAutorizacionAceptada = False
                            gblnSalioSinValidar = False
                            Me.Close()
                        End If
                    Else
                        '''If TieneDerechodeAutorizar = True Then
                        '''    gblnAutorizacionAceptada = True
                        '''Else
                        '''    gblnAutorizacionAceptada = False
                        '''End If
                        'frmInvEntradaPorTransferencia.mstrUsuarioRecibe = Trim(TxtUsuario.Text)
                        gblnAutorizacionAceptada = True
                        gblnSalioSinValidar = False
                        Me.Close()
                    End If
                End If
            End If
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            'GoTo EventExitSub
            'Merr:
        Catch ex As Exception
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            ModErrores.Errores()
EventExitSub:
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        End Try
    End Sub

    Private Sub txtUsuario_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtUsuario.Enter
        On Error GoTo Errores
        ' Asigno el valor del ToolTipText de la Descripción a Barra de Estado
        'UPGRADE_WARNING: Lower bound of collection MenuPrincipal.Status.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        'MenuPrincipal.Status.Items.Item(2).Text = ToolTip1.GetToolTip(TxtUsuario)
        ' Selecciono la información del control
        TxtUsuario.SelectionStart = 0
        TxtUsuario.SelectionLength = Len(TxtUsuario.Text)
Errores:
        If Err.Number <> 0 Then ModErrores.Errores()
    End Sub

    Private Sub txtUsuario_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtUsuario.KeyPress
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
            TxtUsuario.Focus()
            Cancel = False
        Else
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            ModEstandar.BorraCmd()
            gStrSql = "Select * From CatUsuarios Where Nombre = '" & Trim(TxtUsuario.Text) & "'"
            cmd.CommandText = "Up_Select_Datos"
            cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            cmd.Parameters.Append(cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            cmd.Parameters.Append(cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = cmd.Execute
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
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
            End If
        End If
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        GoTo EventExitSub
Merr:
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModErrores.Errores()
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Function TieneDerechodeAutorizar() As Boolean
        'Esta función determina si el usuario especificado tiene derecho de autorizar algun evento a realizar en una Venta.
        'Para que pueda Hacerlo el usuario deberá ser de Tipo Gerente en la Tabla de Usuarios.
        'De lo contrario no podrá hacerlo
        TieneDerechodeAutorizar = False
        ModEstandar.BorraCmd()
        gStrSql = "SELECT     U.CodUsuario, U.Nombre, Isnull(U.Tipo,'') as Tipo, Isnull(U1.Tipo,'') as TipoGrupo " & "FROM         dbo.CatUsuarios U LEFT OUTER JOIN " & "dbo.CatUsuarios U1 ON U.CodGrupo = U1.CodUsuario " & "Where U.CodUsuario = " & mIntUsuarioAutorio
        cmd.CommandText = "Up_Select_Datos"
        cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        cmd.Parameters.Append(cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        cmd.Parameters.Append(cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = cmd.Execute
        If RsGral.RecordCount > 0 Then
            If (Trim(RsGral.Fields("Tipo").Value) = "A" Or Trim(RsGral.Fields("Tipo").Value) = "S") Or (Trim(RsGral.Fields("TipoGrupo").Value) = "A" Or Trim(RsGral.Fields("TipoGrupo").Value) = "S") Then
                TieneDerechodeAutorizar = True
            End If
        End If
    End Function

    Private Sub InitializeComponent()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.TxtClave = New System.Windows.Forms.TextBox()
        Me.TxtUsuario = New System.Windows.Forms.TextBox()
        Me._LblEtiqueta_1 = New System.Windows.Forms.Label()
        Me._LblEtiqueta_0 = New System.Windows.Forms.Label()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.TxtClave)
        Me.Frame1.Controls.Add(Me.TxtUsuario)
        Me.Frame1.Controls.Add(Me._LblEtiqueta_1)
        Me.Frame1.Controls.Add(Me._LblEtiqueta_0)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(57, 27)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(289, 79)
        Me.Frame1.TabIndex = 3
        Me.Frame1.TabStop = False
        '
        'TxtClave
        '
        Me.TxtClave.AcceptsReturn = True
        Me.TxtClave.BackColor = System.Drawing.SystemColors.Window
        Me.TxtClave.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtClave.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtClave.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TxtClave.Location = New System.Drawing.Point(61, 50)
        Me.TxtClave.MaxLength = 20
        Me.TxtClave.Name = "TxtClave"
        Me.TxtClave.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TxtClave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtClave.Size = New System.Drawing.Size(154, 20)
        Me.TxtClave.TabIndex = 1
        '
        'TxtUsuario
        '
        Me.TxtUsuario.AcceptsReturn = True
        Me.TxtUsuario.BackColor = System.Drawing.SystemColors.Window
        Me.TxtUsuario.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtUsuario.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtUsuario.Location = New System.Drawing.Point(61, 18)
        Me.TxtUsuario.MaxLength = 20
        Me.TxtUsuario.Name = "TxtUsuario"
        Me.TxtUsuario.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtUsuario.Size = New System.Drawing.Size(200, 20)
        Me.TxtUsuario.TabIndex = 0
        '
        '_LblEtiqueta_1
        '
        Me._LblEtiqueta_1.BackColor = System.Drawing.SystemColors.Control
        Me._LblEtiqueta_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._LblEtiqueta_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._LblEtiqueta_1.Location = New System.Drawing.Point(8, 51)
        Me._LblEtiqueta_1.Name = "_LblEtiqueta_1"
        Me._LblEtiqueta_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._LblEtiqueta_1.Size = New System.Drawing.Size(65, 17)
        Me._LblEtiqueta_1.TabIndex = 4
        Me._LblEtiqueta_1.Text = "Clave:"
        '
        '_LblEtiqueta_0
        '
        Me._LblEtiqueta_0.BackColor = System.Drawing.SystemColors.Control
        Me._LblEtiqueta_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._LblEtiqueta_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._LblEtiqueta_0.Location = New System.Drawing.Point(8, 18)
        Me._LblEtiqueta_0.Name = "_LblEtiqueta_0"
        Me._LblEtiqueta_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._LblEtiqueta_0.Size = New System.Drawing.Size(65, 17)
        Me._LblEtiqueta_0.TabIndex = 3
        Me._LblEtiqueta_0.Text = "Usuario:"
        '
        'frmAutorizaRecibe_Transf
        '
        Me.ClientSize = New System.Drawing.Size(378, 134)
        Me.Controls.Add(Me.Frame1)
        Me.Name = "frmAutorizaRecibe_Transf"
        Me.Text = "frmAutorizaRecibe_Transf"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
End Class