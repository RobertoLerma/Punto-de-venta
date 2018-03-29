Option Strict Off
Option Explicit On
Imports ADODB
Imports VB6 = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility

Public Class frmBancosProcesoDiarioDesglosedeDepositos
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdAceptar As System.Windows.Forms.Button
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents lblTotal As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    'Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents lblImporte As System.Windows.Forms.Label
    Public WithEvents lblMoneda As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    'Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Panel1 As Panel
    Public WithEvents Label4 As Label
    Dim mblnNuevo As Boolean
    'Dim PierdeFoco As Boolean

    Function GuardarMovimientosDepositos() As Boolean
        Dim I As Integer
        Dim NumPartida As Integer
        On Error GoTo Err_Renamed
        GuardarMovimientosDepositos = True
        NumPartida = 1
        With flexDetalle
            Select Case Me.Tag
                Case "frmDesgloseDepositos"
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                            ModStoredProcedures.PR_IMEMovimientosReferencias((frmBancosProcesoDiarioRegistrodeDepositos.txtFolioIngreso).Text, CStr(NumPartida), lblImporte.Text, .get_TextMatrix(I, 0), .get_TextMatrix(I, 1), "V", "D", C_INSERCION, CStr(0))
                            Cmd.Execute()
                            NumPartida = NumPartida + 1
                        End If
                    Next
                Case "frmDesgloseCargosDiversos"
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                            ModStoredProcedures.PR_IMEMovimientosReferencias((frmBancosProcesoDiarioCargosDiversos.txtFolioEgreso).Text, CStr(NumPartida), lblImporte.Text, .get_TextMatrix(I, 0), .get_TextMatrix(I, 1), "V", "D", C_INSERCION, CStr(0))
                            Cmd.Execute()
                            NumPartida = NumPartida + 1
                        End If
                    Next
                Case "frmDesgloseOtrosIngresos"
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                            ModStoredProcedures.PR_IMEMovimientosReferencias((frmBancosProcesoDiarioRegistrodeOtrosIngresos.txtFolioIngreso).Text, CStr(NumPartida), lblImporte.Text, .get_TextMatrix(I, 0), .get_TextMatrix(I, 1), "V", "D", C_INSERCION, CStr(0))
                            Cmd.Execute()
                            NumPartida = NumPartida + 1
                        End If
                    Next
            End Select
        End With
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            GuardarMovimientosDepositos = False
        End If
    End Function

    Private Sub CambiarFormatoTxtenCaptura()
        With txtFlex
            Select Case flexDetalle.Col
                Case 0 'Referencia del Banco
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Left
                    .MaxLength = 15
                Case 1 'Importe
                    .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                    .MaxLength = 18
            End Select
        End With
    End Sub

    Function ChecarPartidas() As Boolean
        Dim I As Integer
        Dim J As Integer
        ChecarPartidas = True
        With flexDetalle
            For I = 1 To .Rows - 1
                If I = 1 Then
                    If Trim(.get_TextMatrix(I, 0)) = "" And CDbl(Numerico(.get_TextMatrix(I, 1))) = 0 Then
                        ChecarPartidas = True
                        Exit Function
                    ElseIf Trim(.get_TextMatrix(I, 0)) <> "" And CDbl(Numerico(.get_TextMatrix(I, 1))) > 0 Then
                        ChecarPartidas = True
                    Else
                        MsgBox("No ha Capturado Toda la Información de la Ultima Patida, Favor de Verificar..", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        ChecarPartidas = False
                        Exit Function
                    End If
                Else
                    If Trim(.get_TextMatrix(I, 0)) = "" And CDbl(Numerico(.get_TextMatrix(I, 1))) = 0 Then
                        ChecarPartidas = True
                    ElseIf Trim(.get_TextMatrix(I, 0)) <> "" And CDbl(Numerico(.get_TextMatrix(I, 1))) > 0 Then
                        ChecarPartidas = True
                        If I = .Rows - 1 Then
                            Exit Function
                        End If
                    Else
                        MsgBox("No ha Capturado Toda la Información de la Ultima Patida, Favor de Verificar..", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        ChecarPartidas = False
                        Exit Function
                    End If
                End If
            Next
        End With
    End Function

    Sub EliminarLinea()
        Dim Ren As Integer
        Ren = flexDetalle.Rows
        flexDetalle.RemoveItem(flexDetalle.Row)
        flexDetalle.Rows = Ren
        flexDetalle.set_TextMatrix(flexDetalle.Rows - 1, 1, "0.00")
        CalculoImporte()
    End Sub

    Function EstaVacia() As Boolean
        Dim I As Integer
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                    EstaVacia = False
                    Exit Function
                End If
            Next
            EstaVacia = True
        End With
    End Function

    Sub InsertarLinea()
        flexDetalle.AddItem("", flexDetalle.Row)
        flexDetalle.set_TextMatrix(flexDetalle.Row, 1, "0.00")
    End Sub

    Sub CalculoImporte()
        Dim I As Integer
        lblTotal.Text = ""
        With flexDetalle
            For I = 1 To .Rows - 1
                If CDbl(Numerico(.get_TextMatrix(I, 1))) <> 0 Then
                    lblTotal.Text = CStr(CDbl(VB6.Format(Numerico(lblTotal.Text), "#####0.00")) + CDbl(VB6.Format(Numerico(.get_TextMatrix(I, 1)), "#####0.00")))
                End If
            Next
        End With
        If Trim(lblTotal.Text) = "" Then
            lblTotal.Text = "0.00"
        Else
            lblTotal.Text = VB6.Format(lblTotal.Text, "###,##0.00")
        End If
        If CDbl(Numerico(lblTotal.Text)) < CDbl(Numerico(lblImporte.Text)) Then
            lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0)
        ElseIf CDbl(Numerico(lblTotal.Text)) = CDbl(Numerico(lblImporte.Text)) Then
            lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
        End If
    End Sub

    Sub Encabezado()
        Dim I As Integer
        With flexDetalle
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 2000)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Referencia Banco"
            .Col = 1
            .set_ColWidth(1, 0, 1500)
            .CellFontBold = True
            .CellAlignment = 5
            .Text = "Importe"
            .Col = 0
            .Row = 1
            .set_ColAlignment(0, 1)
            For I = 1 To .Rows - 1
                .set_TextMatrix(I, 1, "0.00")
            Next
        End With
    End Sub

    Private Sub cmdAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAceptar.Click
        If ChecarPartidas() Then
            If Me.Tag = "frmDesgloseDepositos" Then
                If Not EstaVacia() Then
                    frmBancosProcesoDiarioRegistrodeDepositos.cmdReferencias.Enabled = False
                Else
                    frmBancosProcesoDiarioRegistrodeDepositos.cmdDesglose.Enabled = True
                    frmBancosProcesoDiarioRegistrodeDepositos.cmdReferencias.Enabled = True
                End If
                frmDesgloseDepositos.Hide()
            ElseIf Me.Tag = "frmDesgloseCargosDiversos" Then
                frmDesgloseCargosDiversos.Hide()
            ElseIf Me.Tag = "frmDesgloseOtrosIngresos" Then
                frmDesgloseOtrosIngresos.Hide()
            End If
        End If
    End Sub

    Private Sub cmdAceptar_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAceptar.Enter
        txtFlex.Visible = False
    End Sub

    Private Sub cmdAceptar_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmdAceptar.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            flexDetalle.Focus()
        End If
    End Sub

    Private Sub flexDetalle_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.ClickEvent
        txtFlex.Visible = False
    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.DblClick
        flexDetalle_KeyPressEvent(flexDetalle, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub flexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        txtFlex.Visible = False
        Pon_Tool()
    End Sub

    Private Sub flexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexDetalle.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete And mblnNuevo Then
            EliminarLinea()
        ElseIf eventArgs.keyCode = System.Windows.Forms.Keys.Insert And mblnNuevo Then
            InsertarLinea()
            'ElseIf KeyCode = vbKeyEscape Then
            '    cmdAceptar_Click
        End If
    End Sub

    Private Sub flexDetalle_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexDetalle.KeyPressEvent
        Dim lonR, lonI As Integer
        If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape And mblnNuevo Then
            'Verifica si se puede capturar la fila
            If flexDetalle.Row > 1 Then
                If flexDetalle.get_TextMatrix(flexDetalle.Row - 1, 0) <> "" Then
                    For lonR = 1 To flexDetalle.Row - 1 Step 1
                        For lonI = 0 To 1 Step 1
                            If flexDetalle.get_TextMatrix(lonR, lonI) = "" Then
                                'MsgBox "Hace falta información en la captura", vbExclamation, cNomEmp
                                flexDetalle.Row = lonR
                                flexDetalle.Col = lonI
                                If flexDetalle.Col = 1 Then
                                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                                End If
                                CambiarFormatoTxtenCaptura()
                                MSHFlexGridEdit(flexDetalle, txtFlex, eventArgs.keyAscii)
                                If Len(Trim(txtFlex.Text)) = 1 Then
                                    'System.Windows.Forms.SendKeys.Send("{right}")
                                End If
                                Exit Sub
                            End If
                        Next lonI
                    Next lonR
                Else
                    'flexDetalle.SetFocus
                    Exit Sub
                End If
            End If
            'Edita el campo sólo si es Editable
            If flexDetalle.Row >= 1 And flexDetalle.Col < 5 Then
                If flexDetalle.Col = 1 And Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) = "" Then
                    MsgBox("Para Capturar el Importe Primero Debe Capturar la Referencia, Favor de Verificar..", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Sub
                Else
                    If flexDetalle.Col = 1 Then
                        If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                    End If
                    CambiarFormatoTxtenCaptura()
                    MSHFlexGridEdit(flexDetalle, txtFlex, eventArgs.keyAscii)
                    If Len(Trim(txtFlex.Text)) = 1 Then
                        'System.Windows.Forms.SendKeys.Send("{right}")
                    End If
                End If
                '        ElseIf flexDetalle.Col = 4 Then
                '            flexDetalle.SetFocus
                '            If flexDetalle.Row < flexDetalle.Rows - 1 Then
                '                flexDetalle.Row = flexDetalle.Row + 1
                '                flexDetalle.Col = 0
                '            Else
                '                flexDetalle.Rows = flexDetalle.Rows + 1
                '                flexDetalle.Row = flexDetalle.Row + 1
                '                flexDetalle.TopRow = flexDetalle.Row
                '                flexDetalle.Col = 0
                '            End If
            End If
            '    Else
            '        blnBuscar = False
            '        If flexDetalle.Col = 0 Or flexDetalle.Col = 1 Then
            '            flexDetalle.Col = 2
            '        ElseIf flexDetalle.Col = 2 Or flexDetalle.Col = 3 Then
            '            If Trim(flexDetalle.TextMatrix(flexDetalle.Row, 2)) <> "" Then
            '                flexDetalle.Col = 4
            '            End If
            '        End If
            '
        Else
            If Not mblnNuevo Then
                'System.Windows.Forms.SendKeys.SendWait("{tab}")
                Exit Sub
            End If
        End If
    End Sub
    Private Sub frmBancosProcesoDiarioDesglosedeDepositos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        flexDetalle.Enabled = True
        cmdAceptar.Enabled = True
    End Sub

    Private Sub frmBancosProcesoDiarioDesglosedeDepositos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If Not mblnNuevo Then
                    ModEstandar.AvanzarTab(Me)
                Else
                    flexDetalle.Focus()
                End If
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmBancosProcesoDiarioDesglosedeDepositos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Public Sub frmBancosProcesoDiarioDesglosedeDepositos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        ModEstandar.CentrarForma(Me)
        Icono(Me, MDIMenuPrincipalCorpo)
        Encabezado()
        lblTotal.Text = "0.00"
    End Sub

    Private Sub frmBancosProcesoDiarioDesglosedeDepositos_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        If flexDetalle.TabIndex = 0 Then
            mblnNuevo = True
        ElseIf flexDetalle.TabIndex = 1 Then
            mblnNuevo = False
        End If
    End Sub

    Private Sub txtFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Enter
        SelTextoTxt(txtFlex)
        Pon_Tool()
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        With flexDetalle
            If KeyCode = System.Windows.Forms.Keys.Return Then
                Select Case .Col
                    Case 0, 1
                        If .Col = 0 Then
                            .Text = Trim(txtFlex.Text)
                            .Col = .Col + 1
                            txtFlex.Visible = False
                            Exit Sub
                        ElseIf .Col = 1 Then
                            If CDbl(Numerico(txtFlex.Text)) = 0 Then
                                MsgBox("Debe Teclear una Cantidad Mayor que Cero...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                                'txtFlex.Text = ""
                                'txtFlex.Focus()
                                Exit Sub
                            End If
                            .Text = Trim(txtFlex.Text)
                            .set_TextMatrix(.Row, 1, VB6.Format(Numerico(.get_TextMatrix(.Row, 1)), "###,##0.00"))
                            CalculoImporte()
                            If .Row = .Rows - 1 Then
                                .Rows = .Rows + 1
                                .Row = .Row + 1
                                .TopRow = .Row
                            Else
                                .Row = .Row + 1
                            End If
                            .Col = 0
                        End If
                        txtFlex.Visible = False
                End Select
            ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
                'If ActiveControl.Name = "txtFlex" Then Exit Sub
                If flexDetalle.Col = 1 And CDbl(Numerico(txtFlex.Text)) = 0 Then
                    MsgBox("Debe Teclear una Cantidad Mayor que Cero...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    'txtFlex.Text = ""
                    'txtFlex.Focus()
                    Exit Sub
                End If
                .Focus()
                txtFlex.Visible = False
            End If
        End With
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
            Case Else
                Select Case flexDetalle.Col
                    Case 0
                        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "/-_")
                    Case 1
                        ModEstandar.MskCantidad(txtFlex.Text, KeyAscii, 15, 2, (txtFlex.SelectionStart))
                End Select
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> txtFlex.Name Then
        '    Exit Sub
        'End If
    End Sub

    Private Sub txtFlex_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtFlex.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If flexDetalle.Col = 1 And CDbl(Numerico(txtFlex.Text)) = 0 Then
            MsgBox("Debe Teclear una Cantidad Mayor que Cero...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFlex.Text = ""
            Cancel = True
        Else
            Cancel = False
        End If
        eventArgs.Cancel = Cancel
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBancosProcesoDiarioDesglosedeDepositos))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdAceptar = New System.Windows.Forms.Button()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblImporte = New System.Windows.Forms.Label()
        Me.lblMoneda = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdAceptar
        '
        Me.cmdAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAceptar.Location = New System.Drawing.Point(12, 265)
        Me.cmdAceptar.Name = "cmdAceptar"
        Me.cmdAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAceptar.Size = New System.Drawing.Size(80, 24)
        Me.cmdAceptar.TabIndex = 1
        Me.cmdAceptar.Text = "&Aceptar"
        Me.cmdAceptar.UseVisualStyleBackColor = False
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(16, 47)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(63, 20)
        Me.txtFlex.TabIndex = 7
        Me.txtFlex.Visible = False
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(12, 50)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(285, 173)
        Me.flexDetalle.TabIndex = 0
        '
        'lblTotal
        '
        Me.lblTotal.BackColor = System.Drawing.SystemColors.Window
        Me.lblTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblTotal.Location = New System.Drawing.Point(188, 238)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotal.Size = New System.Drawing.Size(100, 21)
        Me.lblTotal.TabIndex = 10
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(145, 243)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(32, 15)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Total :"
        '
        'lblImporte
        '
        Me.lblImporte.BackColor = System.Drawing.SystemColors.Window
        Me.lblImporte.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImporte.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblImporte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblImporte.Location = New System.Drawing.Point(194, 6)
        Me.lblImporte.Name = "lblImporte"
        Me.lblImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblImporte.Size = New System.Drawing.Size(100, 21)
        Me.lblImporte.TabIndex = 9
        Me.lblImporte.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblMoneda
        '
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMoneda.Location = New System.Drawing.Point(197, 271)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.Size = New System.Drawing.Size(91, 17)
        Me.lblMoneda.TabIndex = 6
        Me.lblMoneda.Text = "PESOS"
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(80, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(108, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Importe del Depósito :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(9, 331)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(311, 16)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Supr = Eliminar Renglón  Insert = Insertar Renglón"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.cmdAceptar)
        Me.Panel1.Controls.Add(Me.lblMoneda)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.lblTotal)
        Me.Panel1.Controls.Add(Me.lblImporte)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.flexDetalle)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(311, 307)
        Me.Panel1.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(15, 34)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(113, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Desglose del Deposito"
        '
        'frmBancosProcesoDiarioDesglosedeDepositos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(335, 361)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(247, 104)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoDiarioDesglosedeDepositos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Registro de Depósitos Bancarios"
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

End Class