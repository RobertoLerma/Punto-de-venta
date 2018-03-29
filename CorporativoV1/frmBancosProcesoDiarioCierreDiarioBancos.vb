Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosProcesoDiarioCierreDiarioBancos
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             CIERRE DIARIO DE BANCOS                                                                      *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      MARTES 05 DE AGOSTO DE 2003                                                                  *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents prgCierre As System.Windows.Forms.ProgressBar
    Public WithEvents lblFechaActual As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents lblFechaUltimoCorte As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents btnGuardar As Button
    Public WithEvents lblAvance As System.Windows.Forms.Label

    Function Guardar() As Boolean
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        System.Windows.Forms.Application.DoEvents()
        On Error GoTo MErr
        Dim blnTransaccion As Boolean
        Dim sglTiempo As Single
        If lblFechaUltimoCorte.Text = lblFechaActual.Text Then
            MsgBox("Este Proceso ya Fue Ejecutado, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
            Exit Function
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        System.Windows.Forms.Application.DoEvents()
        blnTransaccion = True
        'Guardar la Fecha del Corte
        ModStoredProcedures.PR_IME_ConfiguracionBancos(VB6.Format(lblFechaActual.Text, C_FORMATFECHAGUARDAR), "01/01/1900", C_MODIFICACION, CStr(0))
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

    Sub ObtenerUltimoCierre()
        On Error GoTo MErr
        gStrSql = "SELECT UltCierreBancos FROM ConfiguracionBancos"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Dim fechaCorte As String = AgregarHoraAFecha(RsGral.Fields("UltCierreBancos").Value)
            lblFechaUltimoCorte.Text = fechaCorte
        End If
        Dim fechaActual As String = AgregarHoraAFecha(Today)
        lblFechaActual.Text = fechaActual
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub frmBancosProcesoDiarioCierreDiarioBancos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        If lblFechaUltimoCorte.Text = lblFechaActual.Text Then
            MsgBox("Este Proceso ya Fue Ejecutado, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            IsNothing(Me)
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioCierreDiarioBancos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoDiarioCierreDiarioBancos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me, MDIMenuPrincipalCorpo)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ChecaUltimoCierre()
        ObtenerUltimoCierre()
    End Sub

    Private Sub frmBancosProcesoDiarioCierreDiarioBancos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.prgCierre = New System.Windows.Forms.ProgressBar()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.lblFechaActual = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblFechaUltimoCorte = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblAvance = New System.Windows.Forms.Label()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'prgCierre
        '
        Me.prgCierre.Location = New System.Drawing.Point(18, 95)
        Me.prgCierre.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.prgCierre.Name = "prgCierre"
        Me.prgCierre.Size = New System.Drawing.Size(307, 20)
        Me.prgCierre.TabIndex = 5
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.lblFechaActual)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.lblFechaUltimoCorte)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(18, 11)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(316, 66)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'lblFechaActual
        '
        Me.lblFechaActual.BackColor = System.Drawing.SystemColors.Window
        Me.lblFechaActual.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFechaActual.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFechaActual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFechaActual.Location = New System.Drawing.Point(117, 37)
        Me.lblFechaActual.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFechaActual.Name = "lblFechaActual"
        Me.lblFechaActual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFechaActual.Size = New System.Drawing.Size(178, 19)
        Me.lblFechaActual.TabIndex = 4
        Me.lblFechaActual.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(6, 37)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(79, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Fecha Actual :"
        '
        'lblFechaUltimoCorte
        '
        Me.lblFechaUltimoCorte.BackColor = System.Drawing.SystemColors.Window
        Me.lblFechaUltimoCorte.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFechaUltimoCorte.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFechaUltimoCorte.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFechaUltimoCorte.Location = New System.Drawing.Point(117, 14)
        Me.lblFechaUltimoCorte.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFechaUltimoCorte.Name = "lblFechaUltimoCorte"
        Me.lblFechaUltimoCorte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFechaUltimoCorte.Size = New System.Drawing.Size(178, 17)
        Me.lblFechaUltimoCorte.TabIndex = 2
        Me.lblFechaUltimoCorte.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(6, 15)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(107, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Fecha Ultimo Corte :"
        '
        'lblAvance
        '
        Me.lblAvance.BackColor = System.Drawing.Color.Transparent
        Me.lblAvance.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAvance.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblAvance.Location = New System.Drawing.Point(92, 117)
        Me.lblAvance.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblAvance.Name = "lblAvance"
        Me.lblAvance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAvance.Size = New System.Drawing.Size(180, 17)
        Me.lblAvance.TabIndex = 6
        Me.lblAvance.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(104, 137)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 74
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoDiarioCierreDiarioBancos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(345, 186)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.prgCierre)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.lblAvance)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(402, 257)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBancosProcesoDiarioCierreDiarioBancos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Cierre Diario de Bancos"
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

End Class