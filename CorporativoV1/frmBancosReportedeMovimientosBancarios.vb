Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmBancosReportedeMovimientosBancarios
    Inherits System.Windows.Forms.Form

    Public rptBancosReporteMovimientosBancarios As New rptBancosReportedeMovimientosBancarios
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkIncluirBancos As System.Windows.Forms.CheckBox
    Public WithEvents dbcBanco As System.Windows.Forms.ComboBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents chkIncluirFoliosCancel As System.Windows.Forms.CheckBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox

    Dim mblnSalir As Boolean
    Dim rsReporte As ADODB.Recordset
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    Dim tecla As Integer
    Public WithEvents btnImprimir As Button
    Public WithEvents btnNuevo As Button
    Dim intCodBanco As Integer

    Sub Imprime()

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        Dim Sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim PeriodoReporte As String
        Dim Nota As String
        Dim strWhere As String
        Dim TipoCambio As String
        On Error GoTo ImprimeErr
        strWhere = ""
        Do While (VB.Timer() - sglTiempoCambio) <= 2.1
        Loop
        System.Windows.Forms.Application.DoEvents()
        If dtpFechaInicial.Value > dtpFechaFinal.Value Then
            MsgBox("La Fecha Inicial no Puede ser Mayor que la Fecha Final.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaInicial.Focus()
            Exit Sub
        End If
        If dtpFechaInicial.Value > Now Then
            MsgBox("la Fecha Inicial no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaInicial.Focus()
            Exit Sub
        End If
        If dtpFechaFinal.Value > Now Then
            MsgBox("la Fecha Final no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaFinal.Focus()
            Exit Sub
        End If
        If chkIncluirBancos.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            If intCodBanco = 0 Then
                MsgBox("No se ha seleccionado ningun banco, Favor de verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Exit Sub
            End If
        End If
        If chkIncluirBancos.CheckState = System.Windows.Forms.CheckState.Checked And chkIncluirFoliosCancel.CheckState = System.Windows.Forms.CheckState.Checked Then
            Sql = "SELECT * FROM ((SELECT MB.CodBanco,CB.DescBanco,MB.CtaBancaria,CASE WHEN MB.Moneda = 'P' THEN 'PESOS' WHEN MB.Moneda = 'D' THEN 'DOLARES' END Moneda,MB.FolioMovto,MB.FechaMovto,CASE WHEN MB.Movimiento = 'PA' THEN 'Pagos' WHEN MB.Movimiento = 'DE' THEN 'Depositos' " & "WHEN MB.Movimiento = 'TB' THEN 'Traspasos Bancarios' WHEN MB.Movimiento = 'CD' THEN 'Cargos Diversos' WHEN MB.Movimiento = 'AP' THEN 'Anticipo a Proveedores' WHEN MB.Movimiento = 'OI' THEN 'Otros Ingresos' WHEN MB.Movimiento = 'CA' THEN 'Cancelacion' END Movimiento,MB.Beneficiario,MB.Referencia,CASE WHEN MB.TipoMovto = 'E' THEN MB.Importe ELSE 0 END Abonos," & "CASE WHEN MB.TipoMovto = 'I' THEN MB.Importe ELSE 0 END Cargos FROM MovimientosBancarios MB INNER JOIN CatBancos CB ON MB.CodBanco = CB.CodBanco INNER JOIN CatCuentasBancarias CTA ON CB.CodBanco = CTA.CodBanco AND MB.CtaBancaria = CTA.CtaBancaria " & "WHERE MB.Movimiento <> 'CA' AND FechaMovto BETWEEN '" & VB6.Format(dtpFechaInicial.Value, "MM/DD/YYYY") & "' AND '" & VB6.Format(dtpFechaFinal.Value, "MM/DD/YYYY") & "') " & "UNION " & "(SELECT MB.CodBanco,CB.DescBanco,MB.CtaBancaria,CASE WHEN MB.Moneda = 'P' THEN 'PESOS' WHEN MB.Moneda = 'D' THEN 'DOLARES' END Moneda,MB.FolioMovto,MB.FechaMovto,CASE WHEN MB.Movimiento = 'PA' THEN 'Pagos' WHEN MB.Movimiento = 'DE' THEN 'Depositos' " & "WHEN MB.Movimiento = 'TB' THEN 'Traspasos Bancarios' WHEN MB.Movimiento = 'CD' THEN 'Cargos Diversos' WHEN MB.Movimiento = 'AP' THEN 'Anticipo a Proveedores' WHEN MB.Movimiento = 'OI' THEN 'Otros Ingresos' WHEN MB.Movimiento = 'CA' THEN 'Cancelacion' END Movimiento,MB.Beneficiario,MB.Referencia,CASE WHEN MB.TipoMovto = 'E' THEN MB.Importe ELSE 0 END Abonos," & "CASE WHEN MB.TipoMovto = 'I' THEN MB.Importe ELSE 0 END Cargos FROM (SELECT CAN.CodBanco,CAN.CtaBancaria,CAN.Moneda,CAN.FolioMovto,CAN.FechaMovto,CAN.Movimiento,CAN.TipoMovto,CAN.Importe,CAN.Beneficiario,CAN.Referencia " & "FROM ((SELECT * FROM MovimientosBancarios WHERE Movimiento <> 'CA' AND FechaMovto BETWEEN '" & VB6.Format(dtpFechaInicial.Value, "MM/DD/YYYY") & "' AND '" & VB6.Format(dtpFechaFinal.Value, "MM/DD/YYYY") & "') VIG INNER JOIN (SELECT * FROM MovimientosBancarios WHERE Movimiento = 'CA') CAN ON VIG.FolioMovto = CAN.Referencia)) MB INNER JOIN CatBancos CB ON MB.CodBanco = CB.CodBanco INNER JOIN CatCuentasBancarias CTA ON CB.CodBanco = CTA.CodBanco AND MB.CtaBancaria = CTA.CtaBancaria)) T " & "ORDER BY T.CodBanco,T.DescBanco,T.CtaBancaria,T.FechaMovto"
        ElseIf chkIncluirBancos.CheckState = System.Windows.Forms.CheckState.Checked And chkIncluirFoliosCancel.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            Sql = "SELECT * FROM (SELECT MB.CodBanco,CB.DescBanco,MB.CtaBancaria," & "CASE WHEN MB.Moneda = 'P' THEN 'PESOS' WHEN MB.Moneda = 'D' THEN 'DOLARES' END Moneda," & "MB.FolioMovto,MB.FechaMovto,CASE WHEN MB.Movimiento = 'PA' THEN 'Pagos' WHEN MB.Movimiento = 'DE' THEN 'Depositos' " & "WHEN MB.Movimiento = 'TB' THEN 'Traspasos Bancarios' WHEN MB.Movimiento = 'CD' THEN 'Cargos Diversos' " & "WHEN MB.Movimiento = 'AP' THEN 'Anticipo a Proveedores' WHEN MB.Movimiento = 'OI' THEN 'Otros Ingresos' " & "WHEN MB.Movimiento = 'CA' THEN 'Cancelacion' END Movimiento,MB.Beneficiario,MB.Referencia," & "CASE WHEN MB.TipoMovto = 'E' THEN MB.Importe ELSE 0 END Abonos," & "CASE WHEN MB.TipoMovto = 'I' THEN MB.Importe ELSE 0 END Cargos " & "FROM MovimientosBancarios MB INNER JOIN CatBancos CB ON MB.CodBanco = CB.CodBanco " & "INNER JOIN CatCuentasBancarias CTA ON CB.CodBanco = CTA.CodBanco AND MB.CtaBancaria = CTA.CtaBancaria " & "WHERE MB.Movimiento <> 'CA' AND FechaMovto BETWEEN '" & VB6.Format(dtpFechaInicial.Value, "MM/DD/YYYY") & "' AND '" & VB6.Format(dtpFechaFinal.Value, "MM/DD/YYYY") & "' AND MB.FolioMovto NOT IN (SELECT Referencia FROM MovimientosBancarios WHERE Movimiento = 'CA')) T " & "ORDER BY T.CodBanco,T.DescBanco,T.CtaBancaria,T.FechaMovto"
        ElseIf chkIncluirBancos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkIncluirFoliosCancel.CheckState = System.Windows.Forms.CheckState.Checked Then
            Sql = "SELECT * FROM ((SELECT MB.CodBanco,CB.DescBanco,MB.CtaBancaria,CASE WHEN MB.Moneda = 'P' THEN 'PESOS' WHEN MB.Moneda = 'D' THEN 'DOLARES' END Moneda,MB.FolioMovto,MB.FechaMovto,CASE WHEN MB.Movimiento = 'PA' THEN 'Pagos' WHEN MB.Movimiento = 'DE' THEN 'Depositos' " & "WHEN MB.Movimiento = 'TB' THEN 'Traspasos Bancarios' WHEN MB.Movimiento = 'CD' THEN 'Cargos Diversos' WHEN MB.Movimiento = 'AP' THEN 'Anticipo a Proveedores' WHEN MB.Movimiento = 'OI' THEN 'Otros Ingresos' WHEN MB.Movimiento = 'CA' THEN 'Cancelacion' END Movimiento,MB.Beneficiario,MB.Referencia," & "CASE WHEN MB.TipoMovto = 'E' THEN MB.Importe ELSE 0 END Abonos,CASE WHEN MB.TipoMovto = 'I' THEN MB.Importe ELSE 0 END Cargos FROM MovimientosBancarios MB INNER JOIN CatBancos CB ON MB.CodBanco = CB.CodBanco INNER JOIN CatCuentasBancarias CTA ON CB.CodBanco = CTA.CodBanco AND MB.CtaBancaria = CTA.CtaBancaria " & "WHERE MB.CodBanco = " & intCodBanco & " AND MB.Movimiento <> 'CA' AND FechaMovto BETWEEN '" & VB6.Format(dtpFechaInicial.Value, "MM/DD/YYYY") & "' AND '" & VB6.Format(dtpFechaFinal.Value, "MM/DD/YYYY") & "') " & "UNION " & "(SELECT MB.CodBanco,CB.DescBanco,MB.CtaBancaria,CASE WHEN MB.Moneda = 'P' THEN 'PESOS' WHEN MB.Moneda = 'D' THEN 'DOLARES' END Moneda,MB.FolioMovto,MB.FechaMovto,CASE WHEN MB.Movimiento = 'PA' THEN 'Pagos' WHEN MB.Movimiento = 'DE' THEN 'Depositos' " & "WHEN MB.Movimiento = 'TB' THEN 'Traspasos Bancarios' WHEN MB.Movimiento = 'CD' THEN 'Cargos Diversos' WHEN MB.Movimiento = 'AP' THEN 'Anticipo a Proveedores' WHEN MB.Movimiento = 'OI' THEN 'Otros Ingresos' WHEN MB.Movimiento = 'CA' THEN 'Cancelacion' END Movimiento,MB.Beneficiario,MB.Referencia," & "CASE WHEN MB.TipoMovto = 'E' THEN MB.Importe ELSE 0 END Abonos,CASE WHEN MB.TipoMovto = 'I' THEN MB.Importe ELSE 0 END Cargos FROM (SELECT CAN.CodBanco,CAN.CtaBancaria,CAN.Moneda,CAN.FolioMovto,CAN.FechaMovto,CAN.Movimiento,CAN.TipoMovto,CAN.Importe,CAN.Beneficiario,CAN.Referencia FROM ((SELECT * FROM MovimientosBancarios WHERE CodBanco = " & intCodBanco & " AND Movimiento <> 'CA' " & "AND FechaMovto BETWEEN '" & VB6.Format(dtpFechaInicial.Value, "MM/DD/YYYY") & "' AND '" & VB6.Format(dtpFechaFinal.Value, "MM/DD/YYYY") & "') VIG INNER JOIN (SELECT * FROM MovimientosBancarios WHERE Movimiento = 'CA') CAN ON VIG.FolioMovto = CAN.Referencia)) MB INNER JOIN CatBancos CB ON MB.CodBanco = CB.CodBanco INNER JOIN CatCuentasBancarias CTA ON CB.CodBanco = CTA.CodBanco AND MB.CtaBancaria = CTA.CtaBancaria WHERE MB.CodBanco = " & intCodBanco & ")) T " & "ORDER BY T.CodBanco,T.DescBanco,T.CtaBancaria,T.FechaMovto"
        ElseIf chkIncluirBancos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkIncluirFoliosCancel.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            Sql = "SELECT * FROM (SELECT MB.CodBanco,CB.DescBanco,MB.CtaBancaria," & "CASE WHEN MB.Moneda = 'P' THEN 'PESOS' WHEN MB.Moneda = 'D' THEN 'DOLARES' END Moneda," & "MB.FolioMovto,MB.FechaMovto,CASE WHEN MB.Movimiento = 'PA' THEN 'Pagos' WHEN MB.Movimiento = 'DE' THEN 'Depositos' " & "WHEN MB.Movimiento = 'TB' THEN 'Traspasos Bancarios' WHEN MB.Movimiento = 'CD' THEN 'Cargos Diversos' " & "WHEN MB.Movimiento = 'AP' THEN 'Anticipo a Proveedores' WHEN MB.Movimiento = 'OI' THEN 'Otros Ingresos' " & "WHEN MB.Movimiento = 'CA' THEN 'Cancelacion' END Movimiento,MB.Beneficiario,MB.Referencia," & "CASE WHEN MB.TipoMovto = 'E' THEN MB.Importe ELSE 0 END Abonos," & "CASE WHEN MB.TipoMovto = 'I' THEN MB.Importe ELSE 0 END Cargos " & "FROM MovimientosBancarios MB INNER JOIN CatBancos CB ON MB.CodBanco = CB.CodBanco " & "INNER JOIN CatCuentasBancarias CTA ON CB.CodBanco = CTA.CodBanco AND MB.CtaBancaria = CTA.CtaBancaria " & "WHERE MB.CodBanco = " & intCodBanco & " AND MB.Movimiento <> 'CA' AND FechaMovto BETWEEN '" & VB6.Format(dtpFechaInicial.Value, "MM/DD/YYYY") & "' AND '" & VB6.Format(dtpFechaFinal.Value, "MM/DD/YYYY") & "' AND MB.FolioMovto NOT IN (SELECT Referencia FROM MovimientosBancarios WHERE Movimiento = 'CA')) T " & "ORDER BY T.CodBanco,T.DescBanco,T.CtaBancaria,T.FechaMovto"
        End If
        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = UCase("Reporte de Movimientos Bancarios")
        PeriodoReporte = "Del  " & Format(dtpFechaInicial.Value, "DD/MMM/YYYY") & "  Hasta el  " & VB6.Format(dtpFechaFinal.Value, "DD/MMM/YYYY")
        TipoCambio = Format(gcurCorpoTIPOCAMBIODOLAR, "###,##0.00")

        BorraCmd()
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = Sql
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existen movimientos en el periodo especificado, Favor de verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        Else
            rptBancosReporteMovimientosBancarios.SetDataSource(frmReportes.rsReport)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'pdvNum.Value = New Object() {"NombreEmpresa", "NombreReporte", "Periodo", "TipoCambio"}
        'pdvNum.Value = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte, TipoCambio}

        'If (NombreEmpresa <> Nothing) Then
        '    pdvNum.Value = NombreEmpresa : pvNum.Add(pdvNum)
        '    rptBancosReporteMovimientosBancarios.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
        'End If
        'If (NombreReporte <> Nothing) Then
        '    pdvNum.Value = NombreReporte : pvNum.Add(pdvNum)
        '    rptBancosReporteMovimientosBancarios.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
        'End If
        'If (PeriodoReporte <> Nothing) Then
        '    pdvNum.Value = PeriodoReporte : pvNum.Add(pdvNum)
        '    rptBancosReporteMovimientosBancarios.DataDefinition.ParameterFields("Periodo").ApplyCurrentValues(pvNum)
        'End If
        'If (TipoCambio <> Nothing) Then
        '    pdvNum.Value = TipoCambio : pvNum.Add(pdvNum)
        '    rptBancosReporteMovimientosBancarios.DataDefinition.ParameterFields("TipoCambio").ApplyCurrentValues(pvNum)
        'End If


        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.Text = "Reporte de Movimientos Bancarios"
        frmReportes.reporteActual = rptBancosReporteMovimientosBancarios
        frmReportes.Show()

        Me.Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error al Imprimir : " & Err.Description, MsgBoxStyle.Exclamation, "Error de Operacion")
    End Sub

    Sub Limpiar()
        Nuevo()
        dtpFechaInicial.Focus()
    End Sub

    Sub Nuevo()
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
        chkIncluirFoliosCancel.CheckState = System.Windows.Forms.CheckState.Checked
        chkIncluirBancos.CheckState = System.Windows.Forms.CheckState.Checked
        dbcBanco.Text = ""
    End Sub

    Private Sub chkIncluirBancos_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIncluirBancos.CheckStateChanged
        If chkIncluirBancos.CheckState = System.Windows.Forms.CheckState.Checked Then
            dbcBanco.Enabled = False
            dbcBanco.Text = ""
        ElseIf chkIncluirBancos.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            dbcBanco.Enabled = True
        End If
    End Sub

    Private Sub chkIncluirFoliosCancel_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIncluirFoliosCancel.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcBanco_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.CursorChanged
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcBanco.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,RTRIM(LTRIM(DescBanco)) AS DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' ORDER BY DescBanco"
        DCChange(gStrSql, tecla)
        intCodBanco = 0
    End Sub

    Private Sub dbcBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcBanco.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,RTRIM(LTRIM(DescBanco)) AS DescBanco FROM CatBancos ORDER BY DescBanco"
        DCGotFocus(gStrSql, dbcBanco)
        Pon_Tool()
    End Sub

    Private Sub dbcBanco_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcBanco.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            If chkIncluirBancos.Enabled Then
                chkIncluirBancos.Focus()
            Else
                dtpFechaFinal.Focus()
            End If
        End If
    End Sub

    Private Sub dbcBanco_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dbcBanco.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcBanco_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcBanco.KeyUp
        Dim Aux As String
        Aux = dbcBanco.Text
        If dbcBanco.SelectedItem <> 0 Then
            dbcBanco_Leave(dbcBanco, New System.EventArgs())
        End If
        dbcBanco.Text = Aux
    End Sub

    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        gStrSql = "SELECT CodBanco,RTRIM(LTRIM(DescBanco)) AS DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' ORDER BY DescBanco"
        DCLostFocus(dbcBanco, gStrSql, intCodBanco)
    End Sub

    Private Sub dbcBanco_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcBanco.MouseUp
        Dim Aux As String
        Aux = dbcBanco.Text
        'If dbcBanco.SelectedItem <> 0 Then
        'dbcBanco_Leave(dbcBanco, New System.EventArgs())
        'End If
        dbcBanco.Text = Aux
    End Sub

    Private Sub dtpFechaFinal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.CursorChanged
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.Click
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaFinal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dtpFechaFinal.KeyPress
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.CursorChanged
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Click
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaInicial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dtpFechaInicial.KeyPress
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub frmBancosReportedeMovimientosBancarios_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosReportedeMovimientosBancarios_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosReportedeMovimientosBancarios_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dtpFechaInicial" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosReportedeMovimientosBancarios_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosReportedeMovimientosBancarios_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        Nuevo()
    End Sub

    Private Sub frmBancosReportedeMovimientosBancarios_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
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

    Private Sub frmBancosReportedeMovimientosBancarios_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkIncluirBancos = New System.Windows.Forms.CheckBox()
        Me.chkIncluirFoliosCancel = New System.Windows.Forms.CheckBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkIncluirBancos
        '
        Me.chkIncluirBancos.BackColor = System.Drawing.SystemColors.Control
        Me.chkIncluirBancos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncluirBancos.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkIncluirBancos.Location = New System.Drawing.Point(10, 18)
        Me.chkIncluirBancos.Name = "chkIncluirBancos"
        Me.chkIncluirBancos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncluirBancos.Size = New System.Drawing.Size(123, 18)
        Me.chkIncluirBancos.TabIndex = 2
        Me.chkIncluirBancos.Text = "Todos los Bancos"
        Me.ToolTip1.SetToolTip(Me.chkIncluirBancos, "Incluir Movimientos Cancelados")
        Me.chkIncluirBancos.UseVisualStyleBackColor = False
        '
        'chkIncluirFoliosCancel
        '
        Me.chkIncluirFoliosCancel.BackColor = System.Drawing.SystemColors.Control
        Me.chkIncluirFoliosCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncluirFoliosCancel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkIncluirFoliosCancel.Location = New System.Drawing.Point(202, 167)
        Me.chkIncluirFoliosCancel.Name = "chkIncluirFoliosCancel"
        Me.chkIncluirFoliosCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncluirFoliosCancel.Size = New System.Drawing.Size(154, 18)
        Me.chkIncluirFoliosCancel.TabIndex = 4
        Me.chkIncluirFoliosCancel.Text = "Incluír Folios Cancelados"
        Me.ToolTip1.SetToolTip(Me.chkIncluirFoliosCancel, "Incluir Movimientos Cancelados")
        Me.chkIncluirFoliosCancel.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.chkIncluirBancos)
        Me.Frame2.Controls.Add(Me.dbcBanco)
        Me.Frame2.Controls.Add(Me.Label3)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(13, 80)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(354, 77)
        Me.Frame2.TabIndex = 8
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Información del Banco"
        '
        'dbcBanco
        '
        Me.dbcBanco.Location = New System.Drawing.Point(67, 39)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(276, 21)
        Me.dbcBanco.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(20, 42)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(49, 21)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Banco :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFechaInicial)
        Me.Frame1.Controls.Add(Me.dtpFechaFinal)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(13, 10)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(354, 65)
        Me.Frame1.TabIndex = 5
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(72, 24)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaInicial.TabIndex = 0
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(238, 24)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(105, 20)
        Me.dtpFechaFinal.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(62, 21)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Desde el :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(186, 27)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(62, 21)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Hasta el :"
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(13, 217)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 9
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(128, 217)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 10
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'frmBancosReportedeMovimientosBancarios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(379, 265)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.chkIncluirFoliosCancel)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(278, 156)
        Me.MaximizeBox = False
        Me.Name = "frmBancosReportedeMovimientosBancarios"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Movimientos Bancarios"
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class