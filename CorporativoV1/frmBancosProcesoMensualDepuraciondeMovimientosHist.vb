Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmBancosProcesoMensualDepuraciondeMovimientosHist
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             DEPURACION DE MOVIMIENTOS HISTORICOS                                                         *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      JUEVES 14 DE AGOSTO DE 2003                                                                  *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmbAño As System.Windows.Forms.ComboBox
    Public WithEvents cmbMes As System.Windows.Forms.ComboBox
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents prgCierre As System.Windows.Forms.ProgressBar
    Public WithEvents Image1 As System.Windows.Forms.PictureBox
    Public WithEvents lblPorc As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents btnNuevo As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents Label1 As System.Windows.Forms.Label

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmbAño = New System.Windows.Forms.ComboBox()
        Me.cmbMes = New System.Windows.Forms.ComboBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.prgCierre = New System.Windows.Forms.ProgressBar()
        Me.Image1 = New System.Windows.Forms.PictureBox()
        Me.lblPorc = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Frame3.SuspendLayout()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbAño
        '
        Me.cmbAño.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAño.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAño.Location = New System.Drawing.Point(321, 27)
        Me.cmbAño.Name = "cmbAño"
        Me.cmbAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAño.Size = New System.Drawing.Size(112, 21)
        Me.cmbAño.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.cmbAño, "Año.")
        '
        'cmbMes
        '
        Me.cmbMes.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMes.Items.AddRange(New Object() {"01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"})
        Me.cmbMes.Location = New System.Drawing.Point(65, 27)
        Me.cmbMes.Name = "cmbMes"
        Me.cmbMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMes.Size = New System.Drawing.Size(168, 21)
        Me.cmbMes.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.cmbMes, "Mes.")
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.cmbAño)
        Me.Frame3.Controls.Add(Me.cmbMes)
        Me.Frame3.Controls.Add(Me.Label6)
        Me.Frame3.Controls.Add(Me.Label5)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(16, 168)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(465, 70)
        Me.Frame3.TabIndex = 5
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Información del Periodo"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(280, 29)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(33, 21)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Año :"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(24, 29)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(33, 21)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Mes :"
        '
        'prgCierre
        '
        Me.prgCierre.Location = New System.Drawing.Point(16, 256)
        Me.prgCierre.Name = "prgCierre"
        Me.prgCierre.Size = New System.Drawing.Size(465, 25)
        Me.prgCierre.TabIndex = 8
        '
        'Image1
        '
        Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image1.Location = New System.Drawing.Point(48, 56)
        Me.Image1.Name = "Image1"
        Me.Image1.Size = New System.Drawing.Size(56, 55)
        Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Image1.TabIndex = 9
        Me.Image1.TabStop = False
        '
        'lblPorc
        '
        Me.lblPorc.BackColor = System.Drawing.SystemColors.Control
        Me.lblPorc.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPorc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPorc.Location = New System.Drawing.Point(16, 288)
        Me.lblPorc.Name = "lblPorc"
        Me.lblPorc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPorc.Size = New System.Drawing.Size(465, 17)
        Me.lblPorc.TabIndex = 9
        Me.lblPorc.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(126, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(305, 57)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Al Ejecutar este Proceso Todos los Movimientos Hasta la Fecha Dada Serán Eliminad" &
    "os de los Archivos Actuales"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(160, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(233, 41)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Advertencia"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(465, 145)
        Me.Label1.TabIndex = 2
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(134, 308)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 103
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(19, 308)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 102
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoMensualDepuraciondeMovimientosHist
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(503, 358)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.prgCierre)
        Me.Controls.Add(Me.Image1)
        Me.Controls.Add(Me.lblPorc)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoMensualDepuraciondeMovimientosHist"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Depuración de Movimientos Históricos"
        Me.Frame3.ResumeLayout(False)
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


    'Variables
    Dim mblnSALIR As Boolean

    Function Guardar() As Boolean
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        System.Windows.Forms.Application.DoEvents()
        On Error GoTo MErr
        Dim blnTransaccion As Boolean
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim nReg As Integer
        Dim nRegBorrados As Integer
        Dim PorcBorrado As Integer
        Dim sglTiempo As Single
        ObtenerLimitedeFechas(CInt(VB.Left(Trim(cmbMes.Text), 2)), CInt(Trim(cmbAño.Text)), FechaInicial, FechaFinal)
        gStrSql = "(SELECT MOA.FOLIOMOVTO FROM MOVIMIENTOSBANCARIOS MB INNER JOIN MOVIMIENTOSORIGENAPLIC MOA ON MB.FOLIOMOVTO = MOA.FOLIOMOVTO " & "WHERE MB.FECHAMOVTO < '" & FechaFinal & "') " & "UNION ALL " & "(SELECT MR.FOLIOMOVTO FROM MOVIMIENTOSBANCARIOS MB INNER JOIN MOVIMIENTOSREFERENCIAS MR ON MB.FOLIOMOVTO = MR.FOLIOMOVTO " & "WHERE MB.FECHAMOVTO < '" & FechaFinal & "') " & "UNION ALL " & "(SELECT FOLIOMOVTO FROM MOVIMIENTOSBANCARIOS WHERE FECHAMOVTO < '" & FechaFinal & "')"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            MsgBox("No Existen Movimientos por Depurar, ni durante este periodo, ni hacia atras...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
            Exit Function
        Else
            nReg = RsGral.RecordCount
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        System.Windows.Forms.Application.DoEvents()
        nRegBorrados = 0
        gStrSql = "SELECT MR.FOLIOMOVTO FROM MOVIMIENTOSBANCARIOS MB INNER JOIN MOVIMIENTOSREFERENCIAS MR ON MB.FOLIOMOVTO = MR.FOLIOMOVTO " & "WHERE MB.FECHAMOVTO < '" & FechaFinal & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Do While Not RsGral.EOF
                ModStoredProcedures.PR_IMEMovimientosReferencias(RsGral.Fields("FolioMovto").Value, "0", "0", "", "0", "", "", C_ELIMINACION, CStr(0))
                Cmd.Execute()
                RsGral.MoveNext()
            Loop
        End If
        nRegBorrados = nRegBorrados + RsGral.RecordCount
        PorcBorrado = Int((nRegBorrados / nReg) * 100)
        prgCierre.Value = prgCierre.Value + PorcBorrado
        lblPorc.Text = PorcBorrado & " % "
        System.Windows.Forms.Application.DoEvents()
        sglTiempo = VB.Timer()
        Do While (VB.Timer() - sglTiempo) < 1.5
        Loop
        gStrSql = "SELECT MOA.FOLIOMOVTO FROM MOVIMIENTOSBANCARIOS MB INNER JOIN MOVIMIENTOSORIGENAPLIC MOA ON MB.FOLIOMOVTO = MOA.FOLIOMOVTO " & "WHERE MB.FECHAMOVTO < '" & FechaFinal & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Do While Not RsGral.EOF
                ModStoredProcedures.PR_IMEMovimientosOrigenAplic(RsGral.Fields("FolioMovto").Value, "0", "0", "0", "0", "", "0", "", "01/01/1900", C_ELIMINACION, CStr(0))
                Cmd.Execute()
                RsGral.MoveNext()
            Loop
        End If
        nRegBorrados = nRegBorrados + RsGral.RecordCount
        PorcBorrado = Int((nRegBorrados / nReg) * 100)
        prgCierre.Value = PorcBorrado
        lblPorc.Text = PorcBorrado & " % "
        System.Windows.Forms.Application.DoEvents()
        sglTiempo = VB.Timer()
        Do While (VB.Timer() - sglTiempo) < 1.5
        Loop
        gStrSql = "SELECT FOLIOMOVTO FROM MOVIMIENTOSBANCARIOS WHERE FECHAMOVTO < '" & FechaFinal & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Do While Not RsGral.EOF
                ModStoredProcedures.PR_IMEMovimientosBancarios(RsGral.Fields("FolioMovto").Value, "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "0", "01/01/1900", "", "", "", C_ELIMINACION, CStr(0))
                Cmd.Execute()
                RsGral.MoveNext()
            Loop
        End If
        nRegBorrados = nRegBorrados + RsGral.RecordCount
        PorcBorrado = Int((nRegBorrados / nReg) * 100)
        prgCierre.Value = PorcBorrado
        lblPorc.Text = PorcBorrado & " % "
        System.Windows.Forms.Application.DoEvents()
        sglTiempo = VB.Timer()
        Do While (VB.Timer() - sglTiempo) < 2
        Loop
        lblPorc.Text = "Proceso Completado con Exito ..."
        System.Windows.Forms.Application.DoEvents()
        sglTiempo = VB.Timer()
        Do While (VB.Timer() - sglTiempo) < 3
        Loop
        lblPorc.Text = ""
        prgCierre.Value = 0
        System.Windows.Forms.Application.DoEvents()
        blnTransaccion = True
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
MErr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Function

    Sub ObtenerAños()
        Dim AñoActual As Integer
        Dim I As Integer
        AñoActual = CInt(VB6.Format(Year(Today), "0000"))
        For I = AñoActual - 2 To 1980 Step -1
            cmbAño.Items.Add(CStr(I))
        Next
    End Sub

    Private Sub cmbAño_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAño.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbMes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMes.Enter
        Pon_Tool()
    End Sub

    Private Sub frmBancosProcesoMensualDepuraciondeMovimientosHist_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoMensualDepuraciondeMovimientosHist_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoMensualDepuraciondeMovimientosHist_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If VB6.GetActiveControl().Name <> "cmbMes" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSALIR = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoMensualDepuraciondeMovimientosHist_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoMensualDepuraciondeMovimientosHist_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ObtenerAños()
        cmbMes.SelectedIndex = 0
        cmbAño.SelectedIndex = 0
    End Sub

    Private Sub frmBancosProcesoMensualDepuraciondeMovimientosHist_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If mblnSALIR Then
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0
        '        Case MsgBoxResult.No
        '            mblnSALIR = False
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmBancosProcesoMensualDepuraciondeMovimientosHist_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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