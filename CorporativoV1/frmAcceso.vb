'**********************************************************************************************************************'
'*PROGRAMA: ACCESO JOYERIA RAMOS 
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018      
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB

Public Class FrmAcceso

    Inherits System.Windows.Forms.Form

    Public components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents btnExit As Button
    Public WithEvents PictureBox1 As PictureBox
    Public WithEvents TxtUsuario As TextBox
    Public WithEvents TxtClave As TextBox
    Friend WithEvents lblProyecto As Label
    Public WithEvents btnLogin As PictureBox


    Private Sub FrmAcceso_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub Form_Initialize_Renamed()
        'gIntContFallos = 0
    End Sub

    Private Sub FrmAcceso_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then ModEstandar.RetrocederTab(Me)
    End Sub

    Private Sub FrmAcceso_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModConexion.Abrir(Servidor, Bd)
        'btnLogin.Visible = False
        'ModConexion.Abrir(ServidorS, BdS)
        'Icono(Me, MDIMenuPrincipalCorpo)
        'Dim obj As System.Windows.Forms.Form
        'For Each obj In My.Application.OpenForms
        '    If obj.Name <> "MenuPrincipal" And obj.Name <> "FrmAcceso" Then
        '        obj.Close()
        '        obj = Nothing
        '    End If
        'Next obj
        'ModCorporativo.DesHabilitaMenuPrincipal()

    End Sub

    Private Sub FrmAcceso_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If UnloadMode = 0 Then
        '    Cancel = 1
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub FrmAcceso_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        'ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        'ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        'Dim frm As New FrmAcceso
        'Me.Close()
    End Sub

    Private Sub txtClave_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtClave.Enter
        ' Asigno el valor del ToolTipText de la Descripción a Barra de Estado
        'MDIMenuPrincipalCorpoCorpo.status.Items.Item(2).Text = Convert.ToInt32(ToolTip1.GetToolTip(TxtClave))
        TxtClave.SelectionStart = 0
        TxtClave.SelectionLength = Len(TxtClave.Text)
    End Sub

    Private Sub TxtClave_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtClave.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then TxtUsuario.Focus()
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
                Cmd.CommandText = "Up_Select_Datos"

                Cmd.CommandType = CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", DataTypeEnum.adChar, ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute()

                If RsGral.RecordCount < 1 Then
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    MsgBox("Usuario no registrado", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
                    TxtUsuario.Text = ""
                    TxtClave.Text = ""
                    TxtUsuario.Focus()
                    'GoTo EventExitSub
                Else
                    gIntCodUsuario = RsGral.Fields("CodUsuario").Value.ToString()
                    StrPassW = ModEncriptacion.Encriptar(Trim(TxtClave.Text))
                    If Trim(RsGral.Fields("Password").Value.ToString()) <> Trim(StrPassW) Then
                        gIntContFallos = gIntContFallos + 1
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                        MsgBox("Clave de acceso incorrecta... " + Chr(13) & "Intento " + Str(gIntContFallos), MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
                        TxtClave.Text = ""
                    Else
                        'Me.Close()
                        gStrNomUsuario = TxtUsuario.Text

                        gIntContFallos = 0
                        Me.Hide()
                        Dim frmMenu As MDIMenuPrincipalCorpo = New MDIMenuPrincipalCorpo()
                        frmMenu.Show()

                        'With MDIMenuPrincipalCorpo
                        '    gIntCodUsuario = Trim(RsGral.Fields("CodUsuario").Value.ToString())
                        '    gStrNomUsuario = Trim(RsGral.Fields("Nombre").Value.ToString())
                        '    ModCorporativo.Acceso(gIntCodUsuario)
                        '    '.status.Items.Item(1).Text = "Usuario:   " & gStrNomUsuario
                        'End With
                    End If
                End If
                If gIntContFallos = 3 Then
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    MsgBox(" ACCESO NEGADO ", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
                    MDIMenuPrincipalCorpo.Salir()
                End If
            End If
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            'GoTo EventExitSub

            'con.Close()
            'Merr:
        Catch ex As Exception
            MessageBox.Show("" + ex.Message)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            ModErrores.Errores()
            'EventExitSub:
            'eventArgs.KeyChar = Chr(KeyAscii)
            'If KeyAscii = 0 Then
            '    eventArgs.Handled = True
            'End If
        End Try
    End Sub

    Private Sub txtUsuario_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtUsuario.Enter
        'On Error GoTo Errores
        Try
            ' Asigno el valor del ToolTipText de la Descripción a Barra de Estado
            'MDIMenuPrincipalCorpo.status.Items.Item(2).Text = ToolTip1.GetToolTip(TxtUsuario)

            ' Selecciono la información del control
            TxtUsuario.SelectionStart = 0
            TxtUsuario.SelectionLength = Len(TxtUsuario.Text)
            'Errores:
        Catch ex As Exception
            If Err.Number <> 0 Then ModErrores.Errores()
        End Try
    End Sub

    Private Sub TxtUsuario_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtUsuario.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If ModEstandar.Salir("Desea salir de la aplicación?", "CORPORATIVO  -  " & gstrCorpoNOMBREEMPRESA) Then End
        End If
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


    'Private Sub TxtUsuario_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtUsuario.Leave
    '    'On Error GoTo Merr
    '    Try
    '        If TxtUsuario.Text = "" Then
    '            MsgBox("Proporcione un usuario", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Aviso")
    '            TxtUsuario.Focus()
    '        Else
    '            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
    '            ModEstandar.BorraCmd()
    '            gStrSql = "Select * From CatUsuarios Where Nombre = '" & Trim(TxtUsuario.Text) & "'"
    '            cmd.CommandText = "UP_Select_DatosSql"
    '            cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '            cmd.Parameters.Append(cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '            cmd.Parameters.Append(cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
    '            RsGral = cmd.Execute
    '            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '            If RsGral.RecordCount < 1 Then
    '                MsgBox("Usuario no registrado", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
    '                TxtUsuario.Text = ""
    '                TxtUsuario.Focus()
    '            ElseIf RsGral.Fields("Grupo").Value Then
    '                MsgBox("El usuario es grupo..." & Chr(13) & "Proporcione un nombre de usuario válido", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Aviso")
    '                TxtUsuario.Text = ""
    '            Else
    '                gStrNomUsuario = RsGral.Fields("Nombre").Value = ""
    '            End If
    '        End If
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '        Exit Sub

    '        'Merr:
    '    Catch ex As Exception

    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '        If Err.Number <> 0 Then ModErrores.Errores()
    '    End Try
    'End Sub


    Private Sub TxtUsuario_LostFocus()
        'On Error GoTo Merr
        Try
            If TxtUsuario.Text = "" Then
                MsgBox("Proporcione un usuario", vbInformation + vbOKOnly, "Aviso")
                TxtUsuario.Focus()
            Else
                'Screen.MousePointer = vbHourglass
                ModEstandar.BorraCmd()
                gStrSql = "Select * From CatUsuarios Where Nombre = '" & Trim(TxtUsuario.Text) & "'"
                Cmd.CommandText = "Up_Select_Datos"
                Cmd.CommandType = CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", DataTypeEnum.adChar, ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute()
                'Screen.MousePointer = vbDefault
                If RsGral.RecordCount < 1 Then
                    MsgBox("Usuario no registrado", vbCritical + vbOKOnly, "Error")
                    TxtUsuario.Text = ""
                    TxtUsuario.Focus()
                ElseIf Trim(RsGral.Fields("Grupo").Value.ToString()) Then
                    MsgBox("El usuario es grupo..." + Chr(13) + "Proporcione un nombre de usuario válido", vbOKOnly + vbInformation, "Aviso")
                    TxtUsuario.Text = ""
                Else
                    gStrNomUsuario = RsGral.Fields("Nombre").Value.ToString()
                End If
            End If
            'Screen.MousePointer = vbDefault
            Exit Sub

            'Merr:
        Catch ex As Exception
            'Screen.MousePointer = vbDefault
            If Err.Number <> 0 Then ModErrores.Errores()
        End Try
    End Sub


    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        If (TxtUsuario.Text = "" And TxtClave.Text = "") Then
            MsgBox("Usuario no registrado", vbCritical + vbOKOnly, "Error")
        Else
            txtClave_KeyPress(TxtClave.Text, New KeyPressEventArgs(vbCr))
        End If
    End Sub
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
        System.Windows.Forms.Application.Exit()
    End Sub

    Sub InitializeComponent()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.TxtUsuario = New System.Windows.Forms.TextBox()
        Me.TxtClave = New System.Windows.Forms.TextBox()
        Me.btnLogin = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.lblProyecto = New System.Windows.Forms.Label()
        CType(Me.btnLogin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(154, Byte), Integer), CType(CType(133, Byte), Integer))
        Me.btnExit.Font = New System.Drawing.Font("Arial", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btnExit.Location = New System.Drawing.Point(664, 422)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(69, 32)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "Salir"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'TxtUsuario
        '
        Me.TxtUsuario.BackColor = System.Drawing.Color.White
        Me.TxtUsuario.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtUsuario.Font = New System.Drawing.Font("Microsoft JhengHei UI Light", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtUsuario.ForeColor = System.Drawing.Color.Black
        Me.TxtUsuario.Location = New System.Drawing.Point(117, 149)
        Me.TxtUsuario.Multiline = True
        Me.TxtUsuario.Name = "TxtUsuario"
        Me.TxtUsuario.Size = New System.Drawing.Size(115, 26)
        Me.TxtUsuario.TabIndex = 6
        Me.TxtUsuario.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TxtClave
        '
        Me.TxtClave.BackColor = System.Drawing.Color.White
        Me.TxtClave.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtClave.Font = New System.Drawing.Font("Microsoft JhengHei UI Light", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtClave.ForeColor = System.Drawing.Color.Black
        Me.TxtClave.Location = New System.Drawing.Point(114, 232)
        Me.TxtClave.Name = "TxtClave"
        Me.TxtClave.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TxtClave.Size = New System.Drawing.Size(123, 21)
        Me.TxtClave.TabIndex = 7
        Me.TxtClave.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnLogin
        '
        Me.btnLogin.BackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(154, Byte), Integer), CType(CType(133, Byte), Integer))
        Me.btnLogin.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.Entrar1
        Me.btnLogin.Location = New System.Drawing.Point(76, 340)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(160, 42)
        Me.btnLogin.TabIndex = 8
        Me.btnLogin.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.inicio
        Me.PictureBox1.Location = New System.Drawing.Point(0, -1)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(737, 441)
        Me.PictureBox1.TabIndex = 5
        Me.PictureBox1.TabStop = False
        '
        'lblProyecto
        '
        Me.lblProyecto.AutoSize = True
        Me.lblProyecto.BackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(154, Byte), Integer), CType(CType(133, Byte), Integer))
        Me.lblProyecto.Font = New System.Drawing.Font("Arial Rounded MT Bold", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProyecto.ForeColor = System.Drawing.Color.Black
        Me.lblProyecto.Location = New System.Drawing.Point(89, 286)
        Me.lblProyecto.Name = "lblProyecto"
        Me.lblProyecto.Size = New System.Drawing.Size(132, 24)
        Me.lblProyecto.TabIndex = 9
        Me.lblProyecto.Text = "Corporativo"
        '
        'FrmAcceso
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit
        Me.BackColor = System.Drawing.Color.Gray
        Me.ClientSize = New System.Drawing.Size(770, 508)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblProyecto)
        Me.Controls.Add(Me.TxtUsuario)
        Me.Controls.Add(Me.TxtClave)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnLogin)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "FrmAcceso"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.TopMost = True
        Me.TransparencyKey = System.Drawing.Color.Gray
        CType(Me.btnLogin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub


End Class