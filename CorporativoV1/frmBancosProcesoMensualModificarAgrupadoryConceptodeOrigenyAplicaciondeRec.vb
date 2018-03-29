Option Strict Off
Option Explicit On

Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports System
Imports System.Windows.Forms
Imports System.Data
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports System.Data.SqlClient

Public Class frmBancosProcesoMensualModificarAgrupadoryConceptodeOrigenyAplicaciondeRec
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             MODIFICACIÓN DE AGRUPADOR Y CONCEPTO DE RUBRO                                                *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      SABADO 09 DE AGOSTO DE 2003                                                                  *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdAceptar As System.Windows.Forms.Button
    Public WithEvents txtRubro As System.Windows.Forms.TextBox
    Public WithEvents txtAgrupador As System.Windows.Forms.TextBox
    Public WithEvents dbcAgrupador As System.Windows.Forms.ComboBox
    Public WithEvents dbcRubro As System.Windows.Forms.ComboBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents lblFolio As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Public Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmBancosProcesoMensualModificarAgrupadoryConceptodeOrigenyAplicaciondeRec))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me.cmdAceptar = New System.Windows.Forms.Button
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.txtRubro = New System.Windows.Forms.TextBox
        Me.txtAgrupador = New System.Windows.Forms.TextBox
        Me.dbcAgrupador = New System.Windows.Forms.ComboBox
        Me.dbcRubro = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblFolio = New System.Windows.Forms.Label
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        ' CType(Me.dbcAgrupador, System.ComponentModel.ISupportInitialize).BeginInit()
        ' CType(Me.dbcRubro, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Text = "Modificar Agrupador y Concepto de Origen y Aplicación de Recursos"
        Me.ClientSize = New System.Drawing.Size(460, 126)
        Me.Location = New System.Drawing.Point(273, 171)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.MinimizeBox = False
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ControlBox = True
        Me.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = True
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmBancosProcesoMensualModificarAgrupadoryConceptodeOrigenyAplicaciondeRec"
        Me.cmdAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdAceptar.Text = "&Aceptar"
        Me.cmdAceptar.Size = New System.Drawing.Size(83, 25)
        Me.cmdAceptar.Location = New System.Drawing.Point(343, 95)
        Me.cmdAceptar.TabIndex = 4
        Me.cmdAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAceptar.CausesValidation = True
        Me.cmdAceptar.Enabled = True
        Me.cmdAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAceptar.TabStop = True
        Me.cmdAceptar.Name = "cmdAceptar"
        Me.Frame1.Size = New System.Drawing.Size(445, 84)
        Me.Frame1.Location = New System.Drawing.Point(9, 4)
        Me.Frame1.TabIndex = 5
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Enabled = True
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Visible = True
        Me.Frame1.Name = "Frame1"
        Me.txtRubro.AutoSize = False
        Me.txtRubro.Enabled = False
        Me.txtRubro.Size = New System.Drawing.Size(57, 21)
        Me.txtRubro.Location = New System.Drawing.Point(76, 48)
        Me.txtRubro.MaxLength = 6
        Me.txtRubro.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtRubro, "Codigo del Rubro.")
        Me.txtRubro.AcceptsReturn = True
        Me.txtRubro.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtRubro.BackColor = System.Drawing.SystemColors.Window
        Me.txtRubro.CausesValidation = True
        Me.txtRubro.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRubro.HideSelection = True
        Me.txtRubro.ReadOnly = False
        Me.txtRubro.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRubro.Multiline = False
        Me.txtRubro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRubro.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtRubro.TabStop = True
        Me.txtRubro.Visible = True
        Me.txtRubro.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtRubro.Name = "txtRubro"
        Me.txtAgrupador.AutoSize = False
        Me.txtAgrupador.Enabled = False
        Me.txtAgrupador.Size = New System.Drawing.Size(57, 21)
        Me.txtAgrupador.Location = New System.Drawing.Point(76, 21)
        Me.txtAgrupador.MaxLength = 4
        Me.txtAgrupador.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtAgrupador, "Codigo del Agrupador.")
        Me.txtAgrupador.AcceptsReturn = True
        Me.txtAgrupador.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtAgrupador.BackColor = System.Drawing.SystemColors.Window
        Me.txtAgrupador.CausesValidation = True
        Me.txtAgrupador.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAgrupador.HideSelection = True
        Me.txtAgrupador.ReadOnly = False
        Me.txtAgrupador.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAgrupador.Multiline = False
        Me.txtAgrupador.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAgrupador.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtAgrupador.TabStop = True
        Me.txtAgrupador.Visible = True
        Me.txtAgrupador.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtAgrupador.Name = "txtAgrupador"
        'dbcAgrupador.OcxState = CType(resources.GetObject("dbcAgrupador.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dbcAgrupador.Size = New System.Drawing.Size(297, 21)
        Me.dbcAgrupador.Location = New System.Drawing.Point(134, 21)
        Me.dbcAgrupador.TabIndex = 1
        Me.dbcAgrupador.Name = "dbcAgrupador"
        'dbcRubro.OcxState = CType(resources.GetObject("dbcRubro.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dbcRubro.Size = New System.Drawing.Size(297, 21)
        Me.dbcRubro.Location = New System.Drawing.Point(134, 48)
        Me.dbcRubro.TabIndex = 3
        Me.dbcRubro.Name = "dbcRubro"
        Me.Label2.Text = "Rubro :"
        Me.Label2.Size = New System.Drawing.Size(49, 21)
        Me.Label2.Location = New System.Drawing.Point(11, 50)
        Me.Label2.TabIndex = 7
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Enabled = True
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.UseMnemonic = True
        Me.Label2.Visible = True
        Me.Label2.AutoSize = False
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label2.Name = "Label2"
        Me.Label1.Text = "Agrupador :"
        Me.Label1.Size = New System.Drawing.Size(65, 21)
        Me.Label1.Location = New System.Drawing.Point(11, 23)
        Me.Label1.TabIndex = 6
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Enabled = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.UseMnemonic = True
        Me.Label1.Visible = True
        Me.Label1.AutoSize = False
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label1.Name = "Label1"
        Me.lblFolio.Size = New System.Drawing.Size(113, 17)
        Me.lblFolio.Location = New System.Drawing.Point(16, 96)
        Me.lblFolio.TabIndex = 8
        Me.lblFolio.Visible = False
        Me.lblFolio.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblFolio.BackColor = System.Drawing.SystemColors.Control
        Me.lblFolio.Enabled = True
        Me.lblFolio.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFolio.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFolio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFolio.UseMnemonic = True
        Me.lblFolio.AutoSize = False
        Me.lblFolio.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lblFolio.Name = "lblFolio"
        'CType(Me.dbcRubro, System.ComponentModel.ISupportInitialize).EndInit()
        'CType(Me.dbcAgrupador, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Controls.Add(cmdAceptar)
        Me.Controls.Add(Frame1)
        Me.Controls.Add(lblFolio)
        Me.Frame1.Controls.Add(txtRubro)
        Me.Frame1.Controls.Add(txtAgrupador)
        'Me.Frame1.Controls.Add(dbcAgrupador)
        'Me.Frame1.Controls.Add(dbcRubro)
        Me.Frame1.Controls.Add(Label2)
        Me.Frame1.Controls.Add(Label1)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub


    Dim tecla As Short
    Dim FueraChange As Boolean
    Dim intCodAgrupador As Short
    Dim intCodRubro As Short

    Function ChecaSiExisteLlave() As Boolean
        On Error GoTo MErr
        Dim I As Short
        gStrSql = "SELECT * FROM MovimientosBancarios MB INNER JOIN MovimientosOrigenAplic MOA " & "ON MB.FolioMovto = MOA.FolioMovto WHERE MB.FolioMovto = '" & lblFolio.Text & "' AND " & "MOA.CodOrigenAplicR = " & txtAgrupador.Text & " AND MOA.CodRubro = " & txtRubro.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            ChecaSiExisteLlave = True
            Exit Function
        Else
            ChecaSiExisteLlave = False
        End If
        With frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle
            For I = 1 To .Rows - 1
                If I <> .Row Then
                    If .get_TextMatrix(I, 1) = lblFolio.Text And Numerico(txtAgrupador.Text) = .get_TextMatrix(I, 7) And Numerico(txtRubro.Text) = .get_TextMatrix(I, 8) Then
                        ChecaSiExisteLlave = True
                        Exit Function
                    End If
                End If
            Next
        End With
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Private Sub cmdAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAceptar.Click
        Dim I As Integer
        If CDbl(Numerico(txtAgrupador.Text)) = 0 Then
            MsgBox("Proporcione un Agrupador, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcAgrupador.Focus()
            Exit Sub
        End If
        If CDbl(Numerico(txtRubro.Text)) = 0 Then
            MsgBox("Proporcione un Rubro, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcRubro.Focus()
            Exit Sub
        End If
        If (VB6.Format(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.get_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 5), "0000") <> txtAgrupador.Text Or VB6.Format(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.get_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 6), "000000") <> txtRubro.Text) Then
            'frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 8) = "M" Then
            If ChecaSiExisteLlave() Then
                MsgBox("Este Combinación de Agrupador y Rubro ya esta Asignada en Otro Movimiento. " & Chr(13) & "               Favor de Seleccionar un Agrupador o un Rubro Distinto.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Sub
            End If
            frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.set_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 7, Numerico(txtAgrupador.Text))
            frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.set_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 8, Numerico(txtRubro.Text))
            frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.set_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 2, dbcRubro.Text)
            frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.set_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 9, "M")
            frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.set_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 10, dbcAgrupador.Text)
            For I = 0 To 4
                frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Col = I
                frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.CellBackColor = frmBancosProcesoMensualConsultaOrigenAplicRec.lblModificados.BackColor
            Next
        Else
            frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.set_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 7, Numerico(txtAgrupador.Text))
            frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.set_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 8, Numerico(txtRubro.Text))
            frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.set_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 2, dbcRubro.Text)
            frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.set_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 9, "")
            For I = 0 To 4
                frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Col = I
                frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.CellBackColor = frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.BackColor
            Next
        End If
        frmBancosProcesoMensualConsultaOrigenAplicRec.Enabled = True
        Me.Hide()
    End Sub

    Private Sub dbcAgrupador_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcAgrupador.CursorChanged
        If FueraChange Then Exit Sub
        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcAgrupador.Name Then
            Exit Sub
        End If
        'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR LIKE '" & Trim(dbcAgrupador.Text) & "%' ORDER BY CodOrigenAplicR"
        DCChange(gStrSql, tecla, dbcAgrupador)
        intCodAgrupador = 0
        FueraChange = True
        'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        dbcRubro.Text = ""
        txtRubro.Text = "000000"
        FueraChange = False
    End Sub

    Private Sub dbcAgrupador_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcAgrupador.Enter
        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcAgrupador.Name Then Exit Sub
        gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR FROM CatOrigenAplicRecursos ORDER BY CodOrigenAplicR"
        DCGotFocus(gStrSql, dbcAgrupador)
        FueraChange = False
    End Sub

    Private Sub dbcAgrupador_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcAgrupador.KeyDown
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcAgrupador_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcAgrupador.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcAgrupador_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcAgrupador.KeyUp
        Dim Aux As String
        Aux = dbcAgrupador.Text
        If dbcAgrupador.SelectedItem <> "" Then
            intCodAgrupador = 0
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR LIKE '" & Trim(dbcAgrupador.Text) & "%' ORDER BY CodOrigenAplicR"
            FueraChange = True
            DCLostFocus(dbcAgrupador, gStrSql, intCodAgrupador)
            If intCodAgrupador <> 0 Then
                txtAgrupador.Text = VB6.Format(intCodAgrupador, "0000")
            End If
            FueraChange = False
        End If
        FueraChange = True
        dbcAgrupador.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcAgrupador_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcAgrupador.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        intCodAgrupador = 0
        gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR LIKE '" & Trim(dbcAgrupador.Text) & "%' ORDER BY CodOrigenAplicR"
        FueraChange = True
        DCLostFocus(dbcAgrupador, gStrSql, intCodAgrupador)
        If intCodAgrupador <> 0 Then
            txtAgrupador.Text = VB6.Format(intCodAgrupador, "0000")
        End If
        FueraChange = False
    End Sub

    Private Sub dbcAgrupador_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcAgrupador.MouseUp
        Dim Aux As String
        Aux = dbcAgrupador.Text
        If dbcAgrupador.SelectedItem <> "" Then
            intCodAgrupador = 0
            gStrSql = "SELECT CodOrigenAplicR,DescOrigenAplicR FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR LIKE '" & Trim(dbcAgrupador.Text) & "%' ORDER BY CodOrigenAplicR"
            FueraChange = True
            DCLostFocus(dbcAgrupador, gStrSql, intCodAgrupador)
            If intCodAgrupador <> 0 Then
                txtAgrupador.Text = VB6.Format(intCodAgrupador, "0000")
            End If
            FueraChange = False
        End If
        FueraChange = True
        dbcAgrupador.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcRubro_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRubro.CursorChanged
        If FueraChange = True Then Exit Sub
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcRubro.Name Then
            Exit Sub
        End If
        gStrSql = "SELECT CodRubro,DescRubro FROM CatRubrosOrigenAplicRecursos WHERE DescRubro LIKE '" & Trim(dbcRubro.Text) & "%' AND CodOrigAplicR = " & txtAgrupador.Text & " ORDER BY DescRubro"
        DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcRubro_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRubro.Enter
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcRubro.Name Then
            Exit Sub
        End If
        Dim Aux As String
        Aux = dbcRubro.Text
        dbcRubro.Text = ""
        gStrSql = "SELECT CodRubro,DescRubro FROM CatRubrosOrigenAplicRecursos WHERE CodOrigAplicR = " & txtAgrupador.Text & " ORDER BY DescRubro"
        DCGotFocus(gStrSql)
        dbcRubro.Text = Aux
        dbcRubro.SelectionStart = 0
        dbcRubro.SelectionLength = Len(Trim(dbcRubro.Text))
        FueraChange = False
    End Sub

    Private Sub dbcRubro_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRubro.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            dbcAgrupador.Focus()
        End If
    End Sub

    Private Sub dbcRubro_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcRubro.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcRubro_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcRubro.KeyUp
        Dim Aux As String
        Aux = dbcRubro.Text
        If dbcRubro.SelectedItem <> "" Then
            gStrSql = "SELECT CodRubro,DescRubro FROM CatRubrosOrigenAplicRecursos WHERE DescRubro LIKE '" & Trim(dbcRubro.Text) & "' ORDER BY CodRubro"
            FueraChange = True
            DCLostFocus(dbcRubro, gStrSql, intCodRubro)
            txtRubro.Text = VB6.Format(intCodRubro, "000000")
            FueraChange = False
        End If
        FueraChange = True
        'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        dbcRubro.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcRubro_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRubro.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        gStrSql = "SELECT CodRubro,DescRubro FROM CatRubrosOrigenAplicRecursos WHERE DescRubro LIKE '" & Trim(dbcRubro.Text) & "' ORDER BY CodRubro"
        FueraChange = True
        DCLostFocus(dbcRubro, gStrSql, intCodRubro)
        txtRubro.Text = VB6.Format(intCodRubro, "000000")
        FueraChange = False
    End Sub

    Private Sub dbcRubro_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcRubro.MouseUp
        Dim Aux As String
        Aux = dbcRubro.Text
        If dbcRubro.SelectedItem <> "" Then
            gStrSql = "SELECT CodRubro,DescRubro FROM CatRubrosOrigenAplicRecursos WHERE DescRubro LIKE '" & Trim(dbcRubro.Text) & "' ORDER BY CodRubro"
            FueraChange = True
            DCLostFocus(dbcRubro, gStrSql, intCodRubro)
            txtRubro.Text = VB6.Format(intCodRubro, "000000")
            FueraChange = False
        End If
        FueraChange = True
        dbcRubro.Text = Aux
        FueraChange = False
    End Sub

    Private Sub frmBancosProcesoMensualModificarAgrupadoryConceptodeOrigenyAplicaciondeRec_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmBancosProcesoMensualModificarAgrupadoryConceptodeOrigenyAplicaciondeRec_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.gp_CampoMayusculas(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoMensualModificarAgrupadoryConceptodeOrigenyAplicaciondeRec_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'InitializeComponent()
        frmBancosProcesoMensualConsultaOrigenAplicRec.Enabled = False
        ModEstandar.CentrarForma(Me)
    End Sub

    Private Sub frmBancosProcesoMensualModificarAgrupadoryConceptodeOrigenyAplicaciondeRec_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Dim I As Short
        frmBancosProcesoMensualConsultaOrigenAplicRec.Enabled = True
        If frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.get_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 9) = "" Then
            For I = 0 To 4
                frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Col = I
                frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.CellBackColor = frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.BackColor
            Next
        ElseIf frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.get_TextMatrix(frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Row, 9) = "M" Then
            For I = 0 To 4
                frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.Col = I
                frmBancosProcesoMensualConsultaOrigenAplicRec.flexDetalle.CellBackColor = frmBancosProcesoMensualConsultaOrigenAplicRec.lblModificados.BackColor
            Next
        End If
        'UPGRADE_NOTE: Object frmBancosProcesoMensualModificarAgrupadoryConceptodeOrigenyAplicaciondeRec may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'Me = Nothing
    End Sub
End Class