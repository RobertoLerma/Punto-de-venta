Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmFactDatosFiscales
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             DATOS FISCALES DEL CLIENTE                                                                   *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      SABADO 6 DE MARZO DE 2004                                                                    *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtCodRFC As System.Windows.Forms.TextBox
    Public WithEvents cmdAceptar As System.Windows.Forms.Button
    Public WithEvents txtColonia As System.Windows.Forms.TextBox
    Public WithEvents txtCP As System.Windows.Forms.TextBox
    Public WithEvents txtCiudad As System.Windows.Forms.TextBox
    Public WithEvents txtDomicilio As System.Windows.Forms.TextBox
    Public WithEvents txtCliente As System.Windows.Forms.TextBox
    Public WithEvents txtRFC As System.Windows.Forms.TextBox
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents cmdDatosFiscales As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnBuscar As Button
    Dim FueraChange As Boolean
    Public strControlActual As String 'Nombre del control actual
    'Public FrmConsultasClientes As FrmConsultasClientes = New FrmConsultasClientes()

    Sub Buscar()
        'On Error GoTo Err_Renamed
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "TXTRFC", "TXTCLIENTE"
                FueraChange = True
                FrmConsultasClientes.InitializeComponent()
                FrmConsultasClientes.strControlActual = strControlActual
                FrmConsultasClientes.strFormaActual = UCase(Me.Name)
                FrmConsultasClientes.ShowDialog()
                FueraChange = False
            Case Else
                'Sale de este sub para ke no ejecute ninguna opcion
                Exit Sub
        End Select

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatosRFC()
        On Error GoTo Merr
        'Este Proc. Muestra los datos de un Cliente al seleccionarlo del DataCombo
        gStrSql = "SELECT Rfc , DescClienteRFC,Ltrim(Rtrim(Domicilio)) as Domicilio,Ciudad,CP,colonia FROM CATRFC WHERE CodRFC= " & Numerico(txtCodRFC.Text) & " "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            FueraChange = True
            txtRFC.Text = Trim(RsGral.Fields("Rfc").Value)
            txtRFC.Tag = Trim(RsGral.Fields("Rfc").Value)
            txtDomicilio.Text = Trim(RsGral.Fields("Domicilio").Value)
            txtDomicilio.Tag = Trim(RsGral.Fields("Domicilio").Value)
            txtCliente.Text = Trim(RsGral.Fields("DescClienteRFC").Value)
            txtCliente.Tag = Trim(RsGral.Fields("DescClienteRFC").Value)
            txtColonia.Text = Trim(RsGral.Fields("Colonia").Value)
            txtColonia.Tag = Trim(RsGral.Fields("Colonia").Value)
            txtCiudad.Text = Trim(RsGral.Fields("Ciudad").Value)
            txtCiudad.Tag = Trim(RsGral.Fields("Ciudad").Value)
            txtCP.Text = Trim(RsGral.Fields("CP").Value)
            txtCP.Tag = Trim(RsGral.Fields("CP").Value)
            FueraChange = False
            If CDbl(Numerico(txtCodRFC.Text)) = 1 Then
                MsgBox("Al cliente publico en general no se le puede generar una factura del punto de venta.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Nuevo()
            End If
        Else
            'LimpiaDatosCliente
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Nuevo()
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "txtCliente" Then
            txtRFC.Text = ""
            txtRFC.Tag = ""
        ElseIf System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "txtRFC" Then
            txtCliente.Text = ""
            txtCliente.Tag = ""
        End If
        txtDomicilio.Text = ""
        txtDomicilio.Tag = ""
        txtColonia.Text = ""
        txtColonia.Tag = ""
        txtCiudad.Text = ""
        txtCiudad.Tag = ""
        txtCP.Text = ""
        txtCP.Tag = ""
    End Sub

    Private Sub cmdAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAceptar.Click
        Me.Close()
        frmFactAnalisisVentas.Enabled = True
    End Sub

    Private Sub frmFactDatosFiscales_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmFactDatosFiscales_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmFactDatosFiscales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If System.Windows.Forms.Form.ActiveForm.Name <> "frmFactDatosFiscales" Then Exit Sub
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmFactDatosFiscales_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmFactDatosFiscales_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        'ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        'Me.Icon = MDIMenuPrincipalCorpo.Icon
        CentrarForma(Me)
    End Sub

    Private Sub frmFactDatosFiscales_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        gintCodRFC = CInt(Numerico(txtCodRFC.Text))
        gstrNombreCliente = txtCliente.Text
        gstrRFCCliente = txtRFC.Text
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtCiudad_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCiudad.Enter
        SelTextoTxt(txtRFC)
        Pon_Tool()
    End Sub

    Private Sub txtCliente_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCliente.TextChanged
        If FueraChange Then Exit Sub
        Nuevo()
    End Sub

    Private Sub txtCliente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCliente.Enter
        strControlActual = UCase("txtCliente")
        SelTextoTxt(txtCliente)
        Pon_Tool()
    End Sub

    Private Sub TxtCliente_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCliente.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoLetras(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCliente_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCliente.Leave
        On Error GoTo Err_Renamed
        If Trim(txtCliente.Text) <> "" Then
            gStrSql = "SELECT CodRfc FROM CatRFC WHERE DescClienteRFC LIKE '" & Trim(txtCliente.Text) & "%'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                txtCodRFC.Text = RsGral.Fields("codrfc").Value
                LlenaDatosRFC()
            Else
                MsgBox("No existe este nombre de cliente, Favor de verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtCliente.Focus()
            End If
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub txtColonia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtColonia.Enter
        SelTextoTxt(txtColonia)
        Pon_Tool()
    End Sub

    Private Sub txtCP_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCP.Enter
        SelTextoTxt(txtCP)
        Pon_Tool()
    End Sub

    Private Sub txtDomicilio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.Enter
        SelTextoTxt(txtRFC)
        Pon_Tool()
    End Sub

    Private Sub txtRFC_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRFC.TextChanged
        If FueraChange Then Exit Sub
        Nuevo()
    End Sub

    Private Sub txtRFC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRFC.Enter
        strControlActual = UCase("txtRFC")
        SelTextoTxt(txtRFC)
        Pon_Tool()
    End Sub

    Private Sub txtRFC_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRFC.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Back Then GoTo EventExitSub
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        KeyAscii = ModEstandar.Valida_RFC(txtRFC.Text, KeyAscii, Len(txtRFC.Text) + 1)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtRFC_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRFC.Leave
        On Error GoTo Err_Renamed
        If Trim(txtRFC.Text) <> "" Then
            gStrSql = "SELECT CodRfc FROM CatRFC WHERE Rfc LIKE '" & Trim(txtRFC.Text) & "%'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                txtCodRFC.Text = RsGral.Fields("codrfc").Value
                LlenaDatosRFC()
            Else
                MsgBox("No existe este RFC de cliente, Favor de verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtRFC.Focus()
            End If
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub



    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCodRFC = New System.Windows.Forms.TextBox()
        Me.cmdAceptar = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtColonia = New System.Windows.Forms.TextBox()
        Me.txtCP = New System.Windows.Forms.TextBox()
        Me.txtCiudad = New System.Windows.Forms.TextBox()
        Me.txtDomicilio = New System.Windows.Forms.TextBox()
        Me.txtCliente = New System.Windows.Forms.TextBox()
        Me.txtRFC = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.cmdDatosFiscales = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtCodRFC
        '
        Me.txtCodRFC.AcceptsReturn = True
        Me.txtCodRFC.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodRFC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodRFC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodRFC.Location = New System.Drawing.Point(51, 213)
        Me.txtCodRFC.MaxLength = 0
        Me.txtCodRFC.Name = "txtCodRFC"
        Me.txtCodRFC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodRFC.Size = New System.Drawing.Size(35, 20)
        Me.txtCodRFC.TabIndex = 10
        Me.txtCodRFC.Visible = False
        '
        'cmdAceptar
        '
        Me.cmdAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAceptar.Location = New System.Drawing.Point(421, 211)
        Me.cmdAceptar.Name = "cmdAceptar"
        Me.cmdAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAceptar.Size = New System.Drawing.Size(109, 37)
        Me.cmdAceptar.TabIndex = 4
        Me.cmdAceptar.Text = "Aceptar"
        Me.cmdAceptar.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtColonia)
        Me.Frame1.Controls.Add(Me.cmdDatosFiscales)
        Me.Frame1.Controls.Add(Me.txtCP)
        Me.Frame1.Controls.Add(Me.txtCiudad)
        Me.Frame1.Controls.Add(Me.txtDomicilio)
        Me.Frame1.Controls.Add(Me.txtCliente)
        Me.Frame1.Controls.Add(Me.txtRFC)
        Me.Frame1.Controls.Add(Me.Label6)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(12, 9)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(518, 190)
        Me.Frame1.TabIndex = 5
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Datos del Cliente ..."
        '
        'txtColonia
        '
        Me.txtColonia.AcceptsReturn = True
        Me.txtColonia.BackColor = System.Drawing.SystemColors.Window
        Me.txtColonia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtColonia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtColonia.Location = New System.Drawing.Point(71, 102)
        Me.txtColonia.MaxLength = 30
        Me.txtColonia.Name = "txtColonia"
        Me.txtColonia.ReadOnly = True
        Me.txtColonia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtColonia.Size = New System.Drawing.Size(429, 21)
        Me.txtColonia.TabIndex = 13
        '
        'txtCP
        '
        Me.txtCP.AcceptsReturn = True
        Me.txtCP.BackColor = System.Drawing.SystemColors.Window
        Me.txtCP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCP.Location = New System.Drawing.Point(71, 156)
        Me.txtCP.MaxLength = 10
        Me.txtCP.Name = "txtCP"
        Me.txtCP.ReadOnly = True
        Me.txtCP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCP.Size = New System.Drawing.Size(104, 21)
        Me.txtCP.TabIndex = 12
        '
        'txtCiudad
        '
        Me.txtCiudad.AcceptsReturn = True
        Me.txtCiudad.BackColor = System.Drawing.SystemColors.Window
        Me.txtCiudad.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCiudad.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCiudad.Location = New System.Drawing.Point(71, 129)
        Me.txtCiudad.MaxLength = 30
        Me.txtCiudad.Name = "txtCiudad"
        Me.txtCiudad.ReadOnly = True
        Me.txtCiudad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCiudad.Size = New System.Drawing.Size(429, 21)
        Me.txtCiudad.TabIndex = 3
        '
        'txtDomicilio
        '
        Me.txtDomicilio.AcceptsReturn = True
        Me.txtDomicilio.BackColor = System.Drawing.SystemColors.Window
        Me.txtDomicilio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDomicilio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDomicilio.Location = New System.Drawing.Point(71, 76)
        Me.txtDomicilio.MaxLength = 65
        Me.txtDomicilio.Name = "txtDomicilio"
        Me.txtDomicilio.ReadOnly = True
        Me.txtDomicilio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDomicilio.Size = New System.Drawing.Size(429, 21)
        Me.txtDomicilio.TabIndex = 2
        '
        'txtCliente
        '
        Me.txtCliente.AcceptsReturn = True
        Me.txtCliente.BackColor = System.Drawing.SystemColors.Window
        Me.txtCliente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCliente.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCliente.Location = New System.Drawing.Point(71, 23)
        Me.txtCliente.MaxLength = 40
        Me.txtCliente.Name = "txtCliente"
        Me.txtCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCliente.Size = New System.Drawing.Size(429, 21)
        Me.txtCliente.TabIndex = 0
        '
        'txtRFC
        '
        Me.txtRFC.AcceptsReturn = True
        Me.txtRFC.BackColor = System.Drawing.SystemColors.Window
        Me.txtRFC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRFC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRFC.Location = New System.Drawing.Point(71, 49)
        Me.txtRFC.MaxLength = 15
        Me.txtRFC.Name = "txtRFC"
        Me.txtRFC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRFC.Size = New System.Drawing.Size(143, 21)
        Me.txtRFC.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(18, 103)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(51, 21)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Colonia"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(17, 158)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(31, 21)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "C. P."
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(17, 131)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(51, 21)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Ciudad"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(17, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(51, 21)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Domicilio"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(17, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(51, 21)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Cliente"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(17, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(47, 21)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "RFC"
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackColor = System.Drawing.SystemColors.Control
        Me.btnLimpiar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLimpiar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLimpiar.Location = New System.Drawing.Point(231, 211)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLimpiar.Size = New System.Drawing.Size(109, 37)
        Me.btnLimpiar.TabIndex = 94
        Me.btnLimpiar.Text = "Limpiar"
        Me.ToolTip1.SetToolTip(Me.btnLimpiar, "Registro de Clientes")
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.BackColor = System.Drawing.SystemColors.Control
        Me.btnBuscar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBuscar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBuscar.Location = New System.Drawing.Point(116, 211)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBuscar.Size = New System.Drawing.Size(109, 37)
        Me.btnBuscar.TabIndex = 93
        Me.btnBuscar.Text = "Buscar"
        Me.ToolTip1.SetToolTip(Me.btnBuscar, "Registro de Clientes")
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'cmdDatosFiscales
        '
        Me.cmdDatosFiscales.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDatosFiscales.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDatosFiscales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDatosFiscales.Location = New System.Drawing.Point(-9, -464)
        Me.cmdDatosFiscales.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdDatosFiscales.Name = "cmdDatosFiscales"
        Me.cmdDatosFiscales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDatosFiscales.Size = New System.Drawing.Size(99, 39)
        Me.cmdDatosFiscales.TabIndex = 92
        Me.cmdDatosFiscales.Text = "Datos Fiscales"
        Me.cmdDatosFiscales.UseVisualStyleBackColor = False
        '
        'frmFactDatosFiscales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(547, 262)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.txtCodRFC)
        Me.Controls.Add(Me.cmdAceptar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmFactDatosFiscales"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Datos Fiscales"
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub
End Class