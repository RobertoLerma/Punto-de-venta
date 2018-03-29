Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmConfiguracion
    Inherits System.Windows.Forms.Form

    '+-----------------------------------------------------------------------
    '|Programa: Pantalla de Configuración del sistema
    '|Fecha:    9/Dic/2002
    '+-----------------------------------------------------------------------
    Public components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents BtnBuscaTicketPrinter As System.Windows.Forms.Button
    Public WithEvents txtTicketPrinter As System.Windows.Forms.TextBox
    Public WithEvents btnBuscarImpresora As System.Windows.Forms.Button
    Public WithEvents txtImpresoraSistema As System.Windows.Forms.TextBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents lblImp As System.Windows.Forms.Label
    Public WithEvents _fraMarcos_0 As System.Windows.Forms.GroupBox
    Public WithEvents btnCancelar As System.Windows.Forms.Button
    Public WithEvents btnAceptar As System.Windows.Forms.Button
    Public WithEvents tabOpciones As AxMSComctlLib.AxTabStrip
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents fraMarcos As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray


    'Dim RsAux As ADODB.Recordset
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim mblnNuevo As Boolean
    'Dim cnn As ADODB.Connection
    'Dim cmd As ADODB.Command

    Function Cambios() As Boolean
        On Error GoTo Merr
        Cambios = True
        If Trim(txtTicketPrinter.Text) <> Trim(txtTicketPrinter.Tag) Then Exit Function
        If Trim(txtImpresoraSistema.Text) <> Trim(txtImpresoraSistema.Tag) Then Exit Function
        Cambios = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Private Sub RecuperarInfo()
        Dim rec As ADODB.Recordset
        On Error GoTo Merr
        'Recupera la información para mostrarla
        If tabOpciones.SelectedItem.Tag = "" Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'tabOpciones.SelectedItem.Tag = "V"
            Select Case tabOpciones.SelectedItem.Key
                Case "Impresion"

                    gStrSql = "SELECT * FROM ConfiguracionImpresora Where CodAlmacen = '" & gintCodAlmacen & "' "
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                    RsGral = Cmd.Execute
                    'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
                    If RsGral.RecordCount > 0 Then
                        txtTicketPrinter.Text = Trim(RsGral.Fields("TicketPrinter").Value)
                        txtTicketPrinter.Tag = Trim(RsGral.Fields("TicketPrinter").Value)
                        txtImpresoraSistema.Text = Trim(RsGral.Fields("RutaImpresora").Value)
                        txtImpresoraSistema.Tag = Trim(RsGral.Fields("RutaImpresora").Value)
                    End If

            End Select
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub

Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MostrarError("Ocurrio un error al intentar recuperar las configuraciones.")
        ''
    End Sub

    Private Function ValidarInfo() As Boolean
        On Error GoTo Merr
        Dim intI As Integer
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        For intI = 1 To tabOpciones.Tabs.Count Step 1
            Select Case tabOpciones.Tabs(intI).Key
                Case "Impresion"
                    If Trim(txtTicketPrinter.Text) <> "" Then
                        If Not BuscarImpresora(Trim(txtTicketPrinter.Text), CStr(False)) Then
                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                            Exit Function

                        End If
                        If Trim(txtImpresoraSistema.Text) <> "" Then
                            '''If Not S_BuscarImpresora(Trim(txtImpresoraSistema), True) Then
                            If Not BuscarImpresora(Trim(txtImpresoraSistema.Text), CStr(False)) Then
                                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                                Exit Function
                            End If
                            ValidarInfo = True
                        Else
                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                            MsgBox("Se requiere de una ruta válida para impresora", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                            txtImpresoraSistema.Focus()
                        End If
                    Else
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                        MsgBox("Se requiere de una ruta válida para la impresión de tickets", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                        txtTicketPrinter.Focus()
                    End If
            End Select
        Next intI
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Function

Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MostrarError("Ocurrió un error al validar la información.")
    End Function

    Private Function Guardar() As Boolean
        On Error GoTo Merr
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Cambios() = False Then
            Me.Close()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Function
        End If
        Cnn.BeginTrans()

        ModStoredProcedures.PR_IMConfiguracionImpresora(CStr(gintCodAlmacen), Trim(txtTicketPrinter.Text), Trim(txtImpresoraSistema.Text))
        Cmd.Execute()
        Guardar = True
        Cnn.CommitTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("La Información se ha Guardado Correctamente", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
        Exit Function

Merr:
        Cnn.RollbackTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Private Sub btnAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAceptar.Click
        If ValidarInfo() Then
            If Guardar() Then
                CargarRutaImpresoras()
                Me.Close()
            End If
        End If
    End Sub

    Private Sub btnBuscarImpresora_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnBuscarImpresora.Click
        Dim shlBuscaImp As New Shell32.Shell
        Dim fldRecurso As Shell32.Folder
        Dim lngOpciones As Integer
        lngOpciones = BIF_BROWSEFORPRINTER Or BIF_VALIDATE
        'fldRecurso = shlBuscaImp.BrowseForFolder(Me.Handle.ToInt32, "Seleccione una impresora local o de red", lngOpciones, Shell32.ShellSpecialFolderConstants.ssfNETWORK)
        fldRecurso = shlBuscaImp.BrowseForFolder(0, "Seleccione una impresora local o de red", lngOpciones, Shell32.ShellSpecialFolderConstants.ssfNETWORK)
        If Not fldRecurso Is Nothing Then
            If Trim(fldRecurso.Items.Item.Path) <> "" Then
                txtImpresoraSistema.Text = fldRecurso.Items.Item.Path
            End If
        End If
        txtImpresoraSistema.Focus()
    End Sub

    Private Sub btnBuscarImpresora_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnBuscarImpresora.Enter
        btnBuscarImpresora.Text = ".."
    End Sub

    Private Sub btnBuscarImpresora_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnBuscarImpresora.Leave
        btnBuscarImpresora.Text = "..."
    End Sub

    Private Sub BtnBuscaTicketPrinter_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BtnBuscaTicketPrinter.Click
        Dim shlBuscaImp As New Shell32.Shell
        Dim fldRecurso As Shell32.Folder
        Dim lngOpciones As Integer
        lngOpciones = BIF_BROWSEFORPRINTER Or BIF_VALIDATE
        'fldRecurso = shlBuscaImp.BrowseForFolder(Me.Handle.ToInt32, "Seleccione una impresora local o de red", lngOpciones, Shell32.ShellSpecialFolderConstants.ssfNETWORK)
        fldRecurso = shlBuscaImp.BrowseForFolder(0, "Seleccione una impresora local o de red", lngOpciones, Shell32.ShellSpecialFolderConstants.ssfNETWORK)
        If Not fldRecurso Is Nothing Then
            If Trim(fldRecurso.Items.Item.Path) <> "" Then
                txtTicketPrinter.Text = fldRecurso.Items.Item.Path
            End If
        End If
        txtTicketPrinter.Focus()
    End Sub

    Private Sub BtnBuscaTicketPrinter_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BtnBuscaTicketPrinter.Enter
        btnBuscarImpresora.Text = ".."
    End Sub

    Private Sub BtnBuscaTicketPrinter_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BtnBuscaTicketPrinter.Leave
        btnBuscarImpresora.Text = "..."
    End Sub

    Private Sub btnCancelar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancelar.Click
        Me.Close()
        '   MenuPrincipal.Salir
    End Sub

    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged

        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        mblnNuevo = True
        Nuevo()
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        '    If Screen.ActiveForm.ActiveControl.Name <> dbcSucursales.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen Where TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    'Private Sub dbcSucursales_KeyDown(KeyCode As Integer, Shift As Integer)
    '    'Pregunta solo si existieron cambios
    '    If Cambios = True Then 'And KeyCode = vbKeyDelete Then
    '        Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrCorpoNombreEmpresa)
    '            Case vbYes: 'Guardar el registro
    '                If Guardar = False Then
    '                    KeyCode = 0
    '                    Exit Sub
    '                End If
    '            Case vbNo: 'No hace nada y permite que se borre el contenido del text
    '            Case vbCancel: 'Cancela la captura
    '                txtNumCaja.SetFocus
    '                KeyCode = 0
    '                Exit Sub
    '        End Select
    '    End If
    '
    '    tecla = KeyCode
    '    If KeyCode = vbKeyEscape Then
    '        mblnSalir = True
    '        Unload Me
    '    End If
    'End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
    End Sub

    Private Sub frmConfiguracion_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        gStrSql = "SELECT  CodAlmacen, DescAlmacen FROM CatAlmacen  " & "Where codAlmacen in(" & gintCodAlmacen & ")"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            dbcSucursales.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            dbcSucursales_Leave(dbcSucursales, New System.EventArgs())
        End If
        RecuperarInfo()
        txtTicketPrinter.Focus()
    End Sub

    Private Sub frmConfiguracion_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        'fraMarcos(0).BorderStyle = 0
        dbcSucursales.Enabled = False
    End Sub

    Private Sub frmConfiguracion_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Me = Nothing
    End Sub

    Private Sub tabOpciones_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tabOpciones.ClickEvent
        fraMarcos(0).Visible = False
        'fraMarcos(tabOpciones.SelectedItem.Index - 1).Visible = True
        RecuperarInfo()
    End Sub

    Private Sub txtImpresoraSistema_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImpresoraSistema.TextChanged
        fraMarcos(tabOpciones.SelectedItem.Index - 1).Tag = "M"
    End Sub

    Private Sub txtImpresoraSistema_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImpresoraSistema.Enter
        ModEstandar.SelTextoTxt(txtImpresoraSistema)
    End Sub

    Private Sub txtImpresoraSistema_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtImpresoraSistema.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then btnAceptar.Focus()
    End Sub

    Private Sub txtTicketPrinter_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTicketPrinter.TextChanged
        fraMarcos(tabOpciones.SelectedItem.Index - 1).Tag = "M"
    End Sub

    Private Sub txtTicketPrinter_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTicketPrinter.Enter
        ModEstandar.SelTextoTxt(txtTicketPrinter)
    End Sub

    Private Sub txtTicketPrinter_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtTicketPrinter.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then txtImpresoraSistema.Focus()
    End Sub

    Private Sub Nuevo()

    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfiguracion))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._fraMarcos_0 = New System.Windows.Forms.GroupBox()
        Me.BtnBuscaTicketPrinter = New System.Windows.Forms.Button()
        Me.txtTicketPrinter = New System.Windows.Forms.TextBox()
        Me.btnBuscarImpresora = New System.Windows.Forms.Button()
        Me.txtImpresoraSistema = New System.Windows.Forms.TextBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.lblImp = New System.Windows.Forms.Label()
        Me.btnCancelar = New System.Windows.Forms.Button()
        Me.btnAceptar = New System.Windows.Forms.Button()
        Me.tabOpciones = New AxMSComctlLib.AxTabStrip()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.fraMarcos = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me._fraMarcos_0.SuspendLayout()
        CType(Me.tabOpciones, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraMarcos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.Color.Black
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(132, 17)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(54, 15)
        Me._Label1_0.TabIndex = 10
        Me._Label1_0.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_0, "Nombre de la Farmacia Actual")
        '
        '_fraMarcos_0
        '
        Me._fraMarcos_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraMarcos_0.Controls.Add(Me.BtnBuscaTicketPrinter)
        Me._fraMarcos_0.Controls.Add(Me.txtTicketPrinter)
        Me._fraMarcos_0.Controls.Add(Me.btnBuscarImpresora)
        Me._fraMarcos_0.Controls.Add(Me.txtImpresoraSistema)
        Me._fraMarcos_0.Controls.Add(Me._Label1_1)
        Me._fraMarcos_0.Controls.Add(Me.lblImp)
        Me._fraMarcos_0.ForeColor = System.Drawing.Color.Black
        Me.fraMarcos.SetIndex(Me._fraMarcos_0, CType(0, Short))
        Me._fraMarcos_0.Location = New System.Drawing.Point(16, 65)
        Me._fraMarcos_0.Name = "_fraMarcos_0"
        Me._fraMarcos_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraMarcos_0.Size = New System.Drawing.Size(329, 114)
        Me._fraMarcos_0.TabIndex = 1
        Me._fraMarcos_0.TabStop = False
        Me._fraMarcos_0.Text = " Impresoras del Sistema "
        '
        'BtnBuscaTicketPrinter
        '
        Me.BtnBuscaTicketPrinter.BackColor = System.Drawing.SystemColors.Control
        Me.BtnBuscaTicketPrinter.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnBuscaTicketPrinter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnBuscaTicketPrinter.Location = New System.Drawing.Point(304, 41)
        Me.BtnBuscaTicketPrinter.Name = "BtnBuscaTicketPrinter"
        Me.BtnBuscaTicketPrinter.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnBuscaTicketPrinter.Size = New System.Drawing.Size(17, 20)
        Me.BtnBuscaTicketPrinter.TabIndex = 11
        Me.BtnBuscaTicketPrinter.Text = "..."
        Me.BtnBuscaTicketPrinter.UseVisualStyleBackColor = False
        '
        'txtTicketPrinter
        '
        Me.txtTicketPrinter.AcceptsReturn = True
        Me.txtTicketPrinter.BackColor = System.Drawing.SystemColors.Window
        Me.txtTicketPrinter.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTicketPrinter.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTicketPrinter.Location = New System.Drawing.Point(8, 41)
        Me.txtTicketPrinter.MaxLength = 0
        Me.txtTicketPrinter.Name = "txtTicketPrinter"
        Me.txtTicketPrinter.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTicketPrinter.Size = New System.Drawing.Size(289, 20)
        Me.txtTicketPrinter.TabIndex = 3
        '
        'btnBuscarImpresora
        '
        Me.btnBuscarImpresora.BackColor = System.Drawing.SystemColors.Control
        Me.btnBuscarImpresora.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBuscarImpresora.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBuscarImpresora.Location = New System.Drawing.Point(304, 85)
        Me.btnBuscarImpresora.Name = "btnBuscarImpresora"
        Me.btnBuscarImpresora.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBuscarImpresora.Size = New System.Drawing.Size(17, 20)
        Me.btnBuscarImpresora.TabIndex = 6
        Me.btnBuscarImpresora.Text = "..."
        Me.btnBuscarImpresora.UseVisualStyleBackColor = False
        '
        'txtImpresoraSistema
        '
        Me.txtImpresoraSistema.AcceptsReturn = True
        Me.txtImpresoraSistema.BackColor = System.Drawing.SystemColors.Window
        Me.txtImpresoraSistema.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImpresoraSistema.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImpresoraSistema.Location = New System.Drawing.Point(8, 85)
        Me.txtImpresoraSistema.MaxLength = 0
        Me.txtImpresoraSistema.Name = "txtImpresoraSistema"
        Me.txtImpresoraSistema.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImpresoraSistema.Size = New System.Drawing.Size(289, 20)
        Me.txtImpresoraSistema.TabIndex = 5
        '
        '_Label1_1
        '
        Me._Label1_1.AutoSize = True
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(8, 25)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(105, 13)
        Me._Label1_1.TabIndex = 2
        Me._Label1_1.Text = "Impresión de Tickets"
        '
        'lblImp
        '
        Me.lblImp.AutoSize = True
        Me.lblImp.BackColor = System.Drawing.SystemColors.Control
        Me.lblImp.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblImp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblImp.Location = New System.Drawing.Point(8, 69)
        Me.lblImp.Name = "lblImp"
        Me.lblImp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblImp.Size = New System.Drawing.Size(136, 13)
        Me.lblImp.TabIndex = 4
        Me.lblImp.Text = "Impresora para Facturación"
        '
        'btnCancelar
        '
        Me.btnCancelar.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancelar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancelar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancelar.Location = New System.Drawing.Point(289, 202)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancelar.Size = New System.Drawing.Size(65, 25)
        Me.btnCancelar.TabIndex = 8
        Me.btnCancelar.Text = "&Cancelar"
        Me.btnCancelar.UseVisualStyleBackColor = False
        '
        'btnAceptar
        '
        Me.btnAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.btnAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAceptar.Location = New System.Drawing.Point(209, 202)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAceptar.Size = New System.Drawing.Size(65, 25)
        Me.btnAceptar.TabIndex = 7
        Me.btnAceptar.Text = "&Aceptar"
        Me.btnAceptar.UseVisualStyleBackColor = False
        '
        'tabOpciones
        '
        Me.tabOpciones.Location = New System.Drawing.Point(9, 38)
        Me.tabOpciones.Name = "tabOpciones"
        Me.tabOpciones.OcxState = CType(resources.GetObject("tabOpciones.OcxState"), System.Windows.Forms.AxHost.State)
        Me.tabOpciones.Size = New System.Drawing.Size(345, 153)
        Me.tabOpciones.TabIndex = 0
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(190, 12)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(155, 21)
        Me.dbcSucursales.TabIndex = 9
        '
        'frmConfiguracion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.btnCancelar
        Me.ClientSize = New System.Drawing.Size(362, 235)
        Me.Controls.Add(Me._fraMarcos_0)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.tabOpciones)
        Me.Controls.Add(Me.dbcSucursales)
        Me.Controls.Add(Me._Label1_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(254, 196)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmConfiguracion"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Configuración"
        Me._fraMarcos_0.ResumeLayout(False)
        Me._fraMarcos_0.PerformLayout()
        CType(Me.tabOpciones, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraMarcos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


End Class