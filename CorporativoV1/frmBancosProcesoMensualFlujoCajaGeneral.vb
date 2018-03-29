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
'Imports CrystalDecisions.CrystalReports.Engine

Public Class frmBancosProcesoMensualFlujoCajaGeneral
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE FLUJO DE LA CAJA GENERAL                                                          *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      JUEVES 06 DE NOVIEMBRE DE 2003                                                               *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmbAño As System.Windows.Forms.ComboBox
    Public WithEvents cmbMes As System.Windows.Forms.ComboBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmbAño = New System.Windows.Forms.ComboBox()
        Me.cmbMes = New System.Windows.Forms.ComboBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame3.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbAño
        '
        Me.cmbAño.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAño.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAño.Location = New System.Drawing.Point(65, 59)
        Me.cmbAño.Name = "cmbAño"
        Me.cmbAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAño.Size = New System.Drawing.Size(185, 21)
        Me.cmbAño.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.cmbAño, "Año.")
        '
        'cmbMes
        '
        Me.cmbMes.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMes.Items.AddRange(New Object() {"01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"})
        Me.cmbMes.Location = New System.Drawing.Point(65, 32)
        Me.cmbMes.Name = "cmbMes"
        Me.cmbMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMes.Size = New System.Drawing.Size(185, 21)
        Me.cmbMes.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.cmbMes, "Mes.")
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.cmbAño)
        Me.Frame3.Controls.Add(Me.cmbMes)
        Me.Frame3.Controls.Add(Me.Label5)
        Me.Frame3.Controls.Add(Me.Label4)
        Me.Frame3.Location = New System.Drawing.Point(16, 16)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(273, 105)
        Me.Frame3.TabIndex = 0
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Información del Periodo"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(24, 61)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(33, 21)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Año :"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(24, 34)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(33, 21)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Mes :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(134, 137)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 73
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.Location = New System.Drawing.Point(19, 137)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 72
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoMensualFlujoCajaGeneral
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(312, 180)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoMensualFlujoCajaGeneral"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Flujo de la Caja General"
        Me.Frame3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub


    'Variables
    Dim mblnSalir As Boolean
    Dim rsReporte As ADODB.Recordset
    Dim FueraChange As Boolean

    Sub Imprime()

        Dim RptBancosProcesoMensualReportedeFlujodeCajaGeneral As New RptBancosProcesoMensualReportedeFlujodeCajaGeneral
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim cpfds As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
        Dim cpfd As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo ImprimeErr
        Dim Sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim PeriodoReporte As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim Concepto As String
        Dim SaldoPesos As String
        Dim SaldoDolares As String

        ObtenerLimitedeFechas(CInt(VB.Left(Trim(cmbMes.Text), 2)), CInt(Trim(cmbAño.Text)), FechaInicial, FechaFinal)
        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = "Reporte de Flujo de Caja General"
        'Dim fechaInicial1 As String = AgregarHoraAFecha(FechaInicial)
        'Dim fechaFinal2 As String = AgregarHoraAFecha(FechaFinal)
        'PeriodoReporte = "Del " & fechaInicial1 & " al " & fechaFinal2
        'PeriodoReporte = "Del  " & VB6.Format(FechaInicial, "dd/mmm/yyyy") & "  al  " & VB6.Format(FechaFinal, "dd/mmm/yyyy")

        gStrSql = "SELECT (CB.SaldoInicial + (ISNULL(SUM(CASE MB.TipoMovto WHEN 'I' THEN MB.IMPORTE END),0)) - " & "ISNULL(SUM(CASE MB.TipoMovto WHEN 'E' THEN MB.IMPORTE END),0)) AS Saldo,CB.CtaBancaria,CB.Moneda," & "CatBan.DescBanco " & "FROM CatCuentasBancarias CB LEFT OUTER JOIN " & "(SELECT * FROM MovimientosBancarios WHERE FechaMovto < '" & FechaInicial & "' AND " & "CodBanco = (SELECT CodBanco FROM CatBancos WHERE ControlInterno = 1 AND Sucursal = 0)) MB " & "ON CB.CodBanco = MB.CodBanco AND CB.CtaBancaria = MB.CtaBancaria " & "INNER JOIN CatBancos CatBan ON CB.CodBanco = CatBan.CodBanco " & "WHERE CB.CodBanco = (SELECT CodBanco FROM CatBancos WHERE ControlInterno = 1 AND Sucursal = 0) " & "GROUP BY CB.SaldoInicial,CB.CtaBancaria,CB.Moneda,CatBan.DescBanco"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount > 0 Then
            Concepto = "SALDO INICIAL DEL " & Trim(frmReportes.rsReport.Fields("DescBanco").Value)
            Do While Not frmReportes.rsReport.EOF
                If frmReportes.rsReport.Fields("Moneda").Value = "P" Then
                    SaldoPesos = VB6.Format(frmReportes.rsReport.Fields("SALDO").Value, "###,##0.00")
                ElseIf frmReportes.rsReport.Fields("Moneda").Value = "D" Then
                    SaldoDolares = VB6.Format(frmReportes.rsReport.Fields("SALDO").Value, "###,##0.00")
                End If
                frmReportes.rsReport.MoveNext()
            Loop
        End If
        Sql = "SELECT * FROM DBO.SaldoCuentasPesosYDolares('" & FechaInicial & "','" & FechaFinal & "') ORDER BY FechaMovimiento, Agrupador, Rubro "
        BorraCmd()
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = Sql
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existen movimientos en el periodo especificado, Favor de verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        Else
            'frmReportes.Report = RptBancosProcesoMensualReportedeFlujodeCajaGeneral
            RptBancosProcesoMensualReportedeFlujodeCajaGeneral.SetDataSource(frmReportes.rsReport)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.rsReport = rsReporte
        'pvNum = New pvNum = {"NombreEmpresa", "NombreReporte", "PeriodoReporte", "Concepto", "SaldoPesos", "SaldoDolares"}
        'pdvNum.Value = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte, Concepto, SaldoPesos, SaldoDolares}

        'RptBancosProcesoMensualReportedeFlujodeCajaGeneral.SetParameterValue("SaldoPesos", SaldoPesos)


        'pdvNum.Value = SaldoPesos
        'cpfds = RptBancosProcesoMensualReportedeFlujodeCajaGeneral.DataDefinition.ParameterFields
        'cpfd = cpfds.Item("SaldoPesos")
        'pvNum = cpfd.CurrentValues

        'pvNum.Clear()
        'pvNum.Add(pdvNum)
        'cpfd.ApplyCurrentValues(pvNum)


        'If (SaldoPesos <> Nothing) Then
        '    pdvNum.Value = SaldoPesos : pvNum.Add(pdvNum)
        '    RptBancosProcesoMensualReportedeFlujodeCajaGeneral.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        'Else
        '    pdvNum.Value = "" : pvNum.Add(pdvNum)
        '    RptBancosProcesoMensualReportedeFlujodeCajaGeneral.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        'End If


        frmReportes.Text = "Reporte de Origen y Aplicación de los Recursos"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.reporteActual = RptBancosProcesoMensualReportedeFlujodeCajaGeneral
        frmReportes.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error al Imprimir : " & Err.Description, MsgBoxStyle.Exclamation, "Error de Operacion")
    End Sub

    Sub ObtenerEjercicios()
        On Error GoTo Merr
        gStrSql = "SELECT DISTINCT Ejercicio FROM EjercicioPeriodo"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Do While Not RsGral.EOF
                cmbAño.Items.Add(RsGral.Fields("Ejercicio").Value)
                RsGral.MoveNext()
            Loop
        Else
            cmbAño.Items.Add("")
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Limpiar()
        Nuevo()
        cmbMes.Focus()
    End Sub

    Sub Nuevo()
        Dim lMes As Byte
        Select Case Month(Today)
            Case 1 : lMes = 0
            Case 2 : lMes = 1
            Case 3 : lMes = 2
            Case 4 : lMes = 3
            Case 5 : lMes = 4
            Case 6 : lMes = 5
            Case 7 : lMes = 6
            Case 8 : lMes = 7
            Case 9 : lMes = 8
            Case 10 : lMes = 9
            Case 11 : lMes = 10
            Case 12 : lMes = 11
        End Select
        cmbMes.SelectedIndex = lMes
        cmbAño.SelectedIndex = cmbAño.Items.Count - 1
    End Sub

    Private Sub cmbAño_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAño.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbMes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMes.Enter
        Pon_Tool()
    End Sub

    'UPGRADE_WARNING: Form event frmBancosProcesoMensualFlujoCajaGeneral.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmBancosProcesoMensualFlujoCajaGeneral_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        'UPGRADE_WARNING: Form method frmBancosProcesoMensualFlujoCajaGeneral.ZOrder has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        Me.BringToFront()
    End Sub

    'UPGRADE_WARNING: Form event frmBancosProcesoMensualFlujoCajaGeneral.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmBancosProcesoMensualFlujoCajaGeneral_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoMensualFlujoCajaGeneral_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If Me.ActiveControl.Name <> "cmbMes" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoMensualFlujoCajaGeneral_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoMensualFlujoCajaGeneral_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        ObtenerEjercicios()
        Nuevo()
    End Sub

    Private Sub frmBancosProcesoMensualFlujoCajaGeneral_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        '            cmbMes.Focus()
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmBancosProcesoMensualFlujoCajaGeneral_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class