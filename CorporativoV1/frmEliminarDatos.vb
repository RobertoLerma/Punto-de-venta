Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmEliminarDatos
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents btnAceptar As System.Windows.Forms.Button
    Public WithEvents lblNoExisteProm As System.Windows.Forms.Label
    Public WithEvents fraProcesoRealizado As System.Windows.Forms.GroupBox
    Public WithEvents btnNo As System.Windows.Forms.Button
    Public WithEvents btnSi As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox

    Private Sub btnAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAceptar.Click
        Me.Close()
    End Sub

    Private Sub btnNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnNo.Click
        Me.Close()
    End Sub

    Private Sub btnSi_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSi.Click
        InicializaInformacion()
    End Sub

    Private Sub Command1_Click()
        Me.Close()
    End Sub

    Private Sub frmEliminarDatos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '    Dim ShiftDown, AltDown, CtrlDown, Txt
        '    ShiftDown = (Shift And vbShiftMask) > 0
        '    AltDown = (Shift And vbAltMask) > 0
        '    CtrlDown = (Shift And vbCtrlMask) > 0
        '    If ShiftDown And vbKeyW Then
        '        Select Case (MsgBox("¿Desea Ejecutar este Proceso?", vbCritical + vbQuestion + vbYesNo, "Aviso"))
        '            Case vbYes
        '                ModStoredProcedures.PR_InicializaInformacion
        '                Cmd.Execute
        '                MsgBox "Proceso Realizado"
        '            Case esle
        '                Exit Sub
        '        End Select
        '    End If
    End Sub

    Function InicializaInformacion() As Boolean
        Dim Shift As Object
        On Error GoTo Merr
        Dim CtrlDown, ShiftDown, AltDown, Txt As Object
        ShiftDown = (Shift And VB6.ShiftConstants.ShiftMask) > 0
        AltDown = (Shift And VB6.ShiftConstants.AltMask) > 0
        CtrlDown = (Shift And VB6.ShiftConstants.CtrlMask) > 0
        Cnn.BeginTrans()
        ModStoredProcedures.PR_InicializaInformacion(CStr(gintCodAlmacen))
        Cmd.Execute()
        If Cmd.Parameters("Estatus").Value = 0 Then ' 0 = Significa que hubo un Fallo  en el Proceso  de Borrar
            Cnn.RollbackTrans()
            MsgBox("Fallo en el Proceso." & vbNewLine & "Ejecute el Proceso de Nuevo.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Aviso")
            Me.Close()
            Exit Function
        End If
        fraProcesoRealizado.Visible = True
        btnAceptar.Focus()
        Cnn.CommitTrans()
        Exit Function
Merr:
        Cnn.RollbackTrans()
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraProcesoRealizado = New System.Windows.Forms.GroupBox()
        Me.btnAceptar = New System.Windows.Forms.Button()
        Me.lblNoExisteProm = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.btnNo = New System.Windows.Forms.Button()
        Me.btnSi = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.fraProcesoRealizado.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraProcesoRealizado
        '
        Me.fraProcesoRealizado.BackColor = System.Drawing.SystemColors.Control
        Me.fraProcesoRealizado.Controls.Add(Me.btnAceptar)
        Me.fraProcesoRealizado.Controls.Add(Me.lblNoExisteProm)
        Me.fraProcesoRealizado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraProcesoRealizado.Location = New System.Drawing.Point(8, 0)
        Me.fraProcesoRealizado.Name = "fraProcesoRealizado"
        Me.fraProcesoRealizado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraProcesoRealizado.Size = New System.Drawing.Size(193, 81)
        Me.fraProcesoRealizado.TabIndex = 5
        Me.fraProcesoRealizado.TabStop = False
        Me.fraProcesoRealizado.Visible = False
        '
        'btnAceptar
        '
        Me.btnAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.btnAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAceptar.Location = New System.Drawing.Point(48, 51)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAceptar.Size = New System.Drawing.Size(97, 25)
        Me.btnAceptar.TabIndex = 1
        Me.btnAceptar.Text = "&Aceptar"
        Me.btnAceptar.UseVisualStyleBackColor = False
        '
        'lblNoExisteProm
        '
        Me.lblNoExisteProm.BackColor = System.Drawing.Color.Transparent
        Me.lblNoExisteProm.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNoExisteProm.Font = New System.Drawing.Font("Trebuchet MS", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNoExisteProm.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNoExisteProm.Location = New System.Drawing.Point(8, 11)
        Me.lblNoExisteProm.Name = "lblNoExisteProm"
        Me.lblNoExisteProm.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNoExisteProm.Size = New System.Drawing.Size(175, 37)
        Me.lblNoExisteProm.TabIndex = 6
        Me.lblNoExisteProm.Text = "Porceso Realizado Correctamente !!!"
        Me.lblNoExisteProm.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.btnNo)
        Me.Frame1.Controls.Add(Me.btnSi)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(193, 81)
        Me.Frame1.TabIndex = 3
        Me.Frame1.TabStop = False
        '
        'btnNo
        '
        Me.btnNo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNo.Location = New System.Drawing.Point(112, 40)
        Me.btnNo.Name = "btnNo"
        Me.btnNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNo.Size = New System.Drawing.Size(65, 25)
        Me.btnNo.TabIndex = 0
        Me.btnNo.Text = "&No"
        Me.btnNo.UseVisualStyleBackColor = False
        '
        'btnSi
        '
        Me.btnSi.BackColor = System.Drawing.SystemColors.Control
        Me.btnSi.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSi.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSi.Location = New System.Drawing.Point(16, 40)
        Me.btnSi.Name = "btnSi"
        Me.btnSi.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSi.Size = New System.Drawing.Size(65, 25)
        Me.btnSi.TabIndex = 2
        Me.btnSi.Text = "&Si"
        Me.btnSi.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(180, 25)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "¿Desea Realizar este Proceso?"
        '
        'frmEliminarDatos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(205, 87)
        Me.ControlBox = False
        Me.Controls.Add(Me.fraProcesoRealizado)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(3, 628)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEliminarDatos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Confirmación"
        Me.fraProcesoRealizado.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub


End Class