Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosProcesoMensualCierreConciliacion
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             CIERRE DE CONCILIACIÓN MENSUAL                                                               *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      VIERNES 15 DE AGOSTO DE 2003                                                                 *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtMesACerrar As System.Windows.Forms.TextBox
    Public WithEvents txtAño As System.Windows.Forms.TextBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents txtUltimoAño As System.Windows.Forms.TextBox
    Public WithEvents txtUltimoMes As System.Windows.Forms.TextBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents prgCierre As System.Windows.Forms.ProgressBar
    Public WithEvents btnNuevo As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents lblAvance As System.Windows.Forms.Label

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.txtMesACerrar = New System.Windows.Forms.TextBox()
        Me.txtAño = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtUltimoAño = New System.Windows.Forms.TextBox()
        Me.txtUltimoMes = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.prgCierre = New System.Windows.Forms.ProgressBar()
        Me.lblAvance = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtMesACerrar)
        Me.Frame2.Controls.Add(Me.txtAño)
        Me.Frame2.Controls.Add(Me.Label5)
        Me.Frame2.Controls.Add(Me.Label4)
        Me.Frame2.Location = New System.Drawing.Point(16, 96)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(329, 65)
        Me.Frame2.TabIndex = 8
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Periodo a Cerrar"
        '
        'txtMesACerrar
        '
        Me.txtMesACerrar.AcceptsReturn = True
        Me.txtMesACerrar.BackColor = System.Drawing.SystemColors.Window
        Me.txtMesACerrar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMesACerrar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMesACerrar.Location = New System.Drawing.Point(56, 24)
        Me.txtMesACerrar.Name = "txtMesACerrar"
        Me.txtMesACerrar.ReadOnly = True
        Me.txtMesACerrar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMesACerrar.Size = New System.Drawing.Size(114, 21)
        Me.txtMesACerrar.TabIndex = 2
        '
        'txtAño
        '
        Me.txtAño.AcceptsReturn = True
        Me.txtAño.BackColor = System.Drawing.SystemColors.Window
        Me.txtAño.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAño.Location = New System.Drawing.Point(224, 24)
        Me.txtAño.Name = "txtAño"
        Me.txtAño.ReadOnly = True
        Me.txtAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAño.Size = New System.Drawing.Size(73, 21)
        Me.txtAño.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(192, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(33, 21)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Año :"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(24, 26)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(33, 21)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Mes :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtUltimoAño)
        Me.Frame1.Controls.Add(Me.txtUltimoMes)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Location = New System.Drawing.Point(16, 16)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(329, 65)
        Me.Frame1.TabIndex = 5
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Ultimo Periodo Cerrado"
        '
        'txtUltimoAño
        '
        Me.txtUltimoAño.AcceptsReturn = True
        Me.txtUltimoAño.BackColor = System.Drawing.SystemColors.Window
        Me.txtUltimoAño.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUltimoAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUltimoAño.Location = New System.Drawing.Point(224, 24)
        Me.txtUltimoAño.MaxLength = 0
        Me.txtUltimoAño.Name = "txtUltimoAño"
        Me.txtUltimoAño.ReadOnly = True
        Me.txtUltimoAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUltimoAño.Size = New System.Drawing.Size(73, 21)
        Me.txtUltimoAño.TabIndex = 1
        '
        'txtUltimoMes
        '
        Me.txtUltimoMes.AcceptsReturn = True
        Me.txtUltimoMes.BackColor = System.Drawing.SystemColors.Window
        Me.txtUltimoMes.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUltimoMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUltimoMes.Location = New System.Drawing.Point(56, 24)
        Me.txtUltimoMes.MaxLength = 0
        Me.txtUltimoMes.Name = "txtUltimoMes"
        Me.txtUltimoMes.ReadOnly = True
        Me.txtUltimoMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUltimoMes.Size = New System.Drawing.Size(114, 21)
        Me.txtUltimoMes.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(192, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(33, 21)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Año :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(24, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(33, 21)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Mes :"
        '
        'prgCierre
        '
        Me.prgCierre.Location = New System.Drawing.Point(16, 168)
        Me.prgCierre.Name = "prgCierre"
        Me.prgCierre.Size = New System.Drawing.Size(332, 25)
        Me.prgCierre.TabIndex = 4
        '
        'lblAvance
        '
        Me.lblAvance.BackColor = System.Drawing.Color.Transparent
        Me.lblAvance.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAvance.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblAvance.Location = New System.Drawing.Point(16, 194)
        Me.lblAvance.Name = "lblAvance"
        Me.lblAvance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAvance.Size = New System.Drawing.Size(332, 21)
        Me.lblAvance.TabIndex = 11
        Me.lblAvance.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(132, 209)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 42
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(17, 209)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 41
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoMensualCierreConciliacion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(366, 253)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.prgCierre)
        Me.Controls.Add(Me.lblAvance)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoMensualCierreConciliacion"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Cierre Mensual de Conciliación"
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub


    Dim mblnSalir As Boolean

    Function Guardar() As Boolean
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        System.Windows.Forms.Application.DoEvents()
        On Error GoTo MErr
        Dim blnTransaccion As Boolean
        Dim sglTiempo As Single
        Dim FechaInicial As String
        Dim FechaFinal As String
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        'System.Windows.Forms.Application.DoEvents()
        'blnTransaccion = True

        ObtenerLimitedeFechas(CInt(VB.Left(Trim(txtMesACerrar.Text), 2)), CInt(Trim(txtAño.Text)), FechaInicial, FechaFinal)
        'Guardar la Fecha del Corte de Conciliacion
        ModStoredProcedures.PR_IME_ConfiguracionBancos("01/01/1900", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), C_MODIFICACION, CStr(1))
        Cmd.Execute()
        Do While prgCierre.Value < 100
            sglTiempo = VB.Timer()
            Do While (VB.Timer() - sglTiempo) < 2
            Loop
            prgCierre.Value = prgCierre.Value + 25
            lblAvance.Text = prgCierre.Value & "%"
            System.Windows.Forms.Application.DoEvents()
        Loop
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        sglTiempo = VB.Timer()
        Do While (VB.Timer() - sglTiempo) < 1.5
        Loop
        lblAvance.Text = "Proceso Completado Exitosamente"
        System.Windows.Forms.Application.DoEvents()
        sglTiempo = VB.Timer()
        Do While (VB.Timer() - sglTiempo) < 3
        Loop
        prgCierre.Value = 0
        lblAvance.Text = ""
        System.Windows.Forms.Application.DoEvents()
        ObtenerUltimoCierre()
MErr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Function

    Sub ChecaUltimoCierre()
        On Error GoTo MErr
        Dim blnTransaccion As Boolean
        Cnn.BeginTrans()
        blnTransaccion = True
        gStrSql = "SELECT * FROM ConfiguracionBancos"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            ModStoredProcedures.PR_IME_ConfiguracionBancos("01/01/1900", "01/01/1900", C_INSERCION, CStr(0))
            Cmd.Execute()
        End If
        Cnn.CommitTrans()
        blnTransaccion = False
MErr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Sub ObtenerProximoPeriodo(ByRef MesAnte As Integer, ByRef Año As Integer, ByRef txtmes As System.Windows.Forms.TextBox, ByRef txtAño As System.Windows.Forms.TextBox)
        MesAnte = MesAnte + 1
        If MesAnte > 12 Then
            MesAnte = 1
            Año = Año + 1
        End If
        txtAño.Text = CStr(Año)
        txtmes.Text = VB6.Format(MesAnte, "00") & " - " & MesLetra(CDate("01" & "/" & MesAnte & "/" & Año), False)
    End Sub

    Sub ObtenerUltimoCierre()
        On Error GoTo MErr
        gStrSql = "SELECT UltCierreConciliacion FROM ConfiguracionBancos"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtUltimoAño.Text = VB6.Format(Year(RsGral.Fields("UltCierreConciliacion").Value), "0000")
            txtUltimoMes.Text = VB6.Format(Month(RsGral.Fields("UltCierreConciliacion").Value.ToString()), "00") & " - " & MesLetra(RsGral.Fields("UltCierreConciliacion").Value.ToString(), False)
        End If
        If CDbl(VB6.Format(Year(RsGral.Fields("UltCierreConciliacion").Value), "0000")) = 1900 Then
            txtAño.Text = VB6.Format(Year(Today), "0000")
            txtMesACerrar.Text = "01 - Enero"
        Else
            ObtenerProximoPeriodo(CInt(VB.Left(txtUltimoMes.Text, 2)), CInt(txtUltimoAño.Text), txtMesACerrar, txtAño)
        End If
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub frmBancosProcesoMensualCierreConciliacion_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmBancosProcesoMensualCierreConciliacion_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoMensualCierreConciliacion_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            ModEstandar.AvanzarTab(Me)
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
            If VB6.GetActiveControl().Name <> "txtUltimoMes" Then
                ModEstandar.RetrocederTab(Me)
            Else
                mblnSalir = True
                Me.Close()
            End If
        End If
    End Sub

    Private Sub frmBancosProcesoMensualCierreConciliacion_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me, MDIMenuPrincipalCorpo)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ChecaUltimoCierre()
        ObtenerUltimoCierre()
    End Sub

    Private Sub frmBancosProcesoMensualCierreConciliacion_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If mblnSalir Then
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0
        '        Case MsgBoxResult.No
        '            mblnSalir = False
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmBancosProcesoMensualCierreConciliacion_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub
End Class