Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosReporteMovBancariosXTipo
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE MOVIMIENTOS BANCARIOS POR TIPO                                                    *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      MARTES 17 DE FEBRERO DE 2004                                                                 *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkIncluirFoliosCancel As System.Windows.Forms.CheckBox
    Public WithEvents chkDolares As System.Windows.Forms.CheckBox
    Public WithEvents chkPesos As System.Windows.Forms.CheckBox
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents chkOtrosIngresos As System.Windows.Forms.CheckBox
    Public WithEvents chkAnticipoProveedores As System.Windows.Forms.CheckBox
    Public WithEvents chkTraspasosBancarios As System.Windows.Forms.CheckBox
    Public WithEvents chkCargosDiversos As System.Windows.Forms.CheckBox
    Public WithEvents chkDepositos As System.Windows.Forms.CheckBox
    Public WithEvents chkPagos As System.Windows.Forms.CheckBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox

    Dim mblnSalir As Boolean
    Dim rsReporte As ADODB.Recordset
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo

    Sub Imprime()

        Dim rptBancosReportedeMovimientosporTipo1 As New rptBancosReportedeMovimientosporTipo
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        Dim sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim PeriodoReporte As String
        Dim Nota As String
        On Error GoTo ImprimeErr

        '''Actualizan el dato registrado
        dtpFechaInicial.Refresh()
        dtpFechaFinal.Refresh()

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
        If chkPagos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkDepositos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkCargosDiversos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkTraspasosBancarios.CheckState = System.Windows.Forms.CheckState.Unchecked And chkAnticipoProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked And chkOtrosIngresos.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            MsgBox("No es posible generar el reporte, Favor de seleccionar por lo menos un tipo de movimiento.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            chkPagos.Focus()
            Exit Sub
        End If
        If chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            MsgBox("No es posible generar el reporte, Favor de Seleccionar por lo menos una moneda.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            chkPesos.Focus()
            Exit Sub
        End If

        'Armamos el Query para el reporte
        If optPesos.Checked Then
            sql = "SELECT MB.CodBanco,CB.DescBanco,MB.FolioMovto,CASE WHEN MB.Moneda = 'P' THEN 'PES' WHEN MB.Moneda = 'D' THEN 'DOL' END AS Moneda," & "MB.Beneficiario, MB.Concepto, " & "ROUND(CASE WHEN MB.Moneda = 'P' THEN Importe WHEN MB.Moneda = 'D' THEN MB.Importe * MB.TipoCambio END,2) Importe,"
            Nota = "**Importes Expresados en Pesos"
        Else
            sql = "SELECT MB.CodBanco,CB.DescBanco,MB.FolioMovto,CASE WHEN MB.Moneda = 'P' THEN 'PES' WHEN MB.Moneda = 'D' THEN 'DOL' END AS Moneda," & "MB.Beneficiario, MB.Concepto, " & "ROUND(CASE WHEN MB.Moneda = 'D' THEN Importe WHEN MB.Moneda = 'P' THEN MB.Importe/MB.TipoCambio END,2) Importe,"
            Nota = "**Importes Expresados en Dólares"
        End If
        sql = sql & "CASE WHEN MB.Movimiento = 'PA' THEN 'Pagos' WHEN MB.Movimiento = 'DE' THEN 'Depositos' " & "WHEN MB.Movimiento = 'TB' THEN 'Traspasos Bancarios' WHEN MB.Movimiento = 'CD' THEN 'Cargos Diversos' " & "WHEN MB.Movimiento = 'AP' THEN 'Anticipo a Proveedores' WHEN MB.Movimiento = 'OI' THEN 'Otros Ingresos' END Movimiento,"
        If chkIncluirFoliosCancel.CheckState Then
            sql = sql & "CASE WHEN MB.FolioMovto = Can.FolioMovto THEN 'Cancelado' ELSE '' END Estatus " & "FROM MovimientosBancarios MB LEFT OUTER JOIN " & "(SELECT Vig.FolioMovto,Can.Movimiento FROM (SELECT * FROM MOVIMIENTOSBANCARIOS WHERE Movimiento <> 'CA') Vig " & "INNER JOIN (SELECT * FROM MOVIMIENTOSBANCARIOS WHERE Movimiento = 'CA') Can ON Vig.FolioMovto = Can.Referencia) Can " & "ON MB.FolioMovto = Can.FolioMovto INNER JOIN CatBancos CB ON MB.CodBanco = CB.CodBanco WHERE "
        Else
            sql = sql & "'' Estatus FROM MovimientosBancarios MB " & "INNER JOIN (SELECT FolioMovto FROM MovimientosBancarios WHERE FolioMovto NOT IN " & "(SELECT Referencia FROM MovimientosBancarios WHERE Movimiento = 'CA')AND Movimiento <> 'CA') Vig " & "ON MB.FolioMovto = Vig.FolioMovto INNER JOIN CatBancos CB ON MB.CodBanco = CB.CodBanco WHERE "
        End If
        If chkPagos.CheckState = System.Windows.Forms.CheckState.Checked Then
            sql = sql & "(MB.Movimiento = '" & C_MOVPAGO & "' OR "
        Else
            sql = sql & "(MB.Movimiento = '' OR "
        End If
        If chkIncluirFoliosCancel.CheckState = System.Windows.Forms.CheckState.Checked Then
            sql = sql & "MB.Movimiento = '" & C_MOVCANCELACION & "' OR "
        Else
            sql = sql & "MB.Movimiento = '' OR "
        End If
        If chkDepositos.CheckState = System.Windows.Forms.CheckState.Checked Then
            sql = sql & "MB.Movimiento = '" & C_MOVDEPOSITO & "' OR "
        Else
            sql = sql & "MB.Movimiento = '' OR "
        End If
        If chkCargosDiversos.CheckState = System.Windows.Forms.CheckState.Checked Then
            sql = sql & "MB.Movimiento = '" & C_MOVCARGOS & "' OR "
        Else
            sql = sql & "MB.Movimiento = '' OR "
        End If
        If chkTraspasosBancarios.CheckState = System.Windows.Forms.CheckState.Checked Then
            sql = sql & "MB.Movimiento = '" & C_MOVTRASPASO & "' OR "
        Else
            sql = sql & "MB.Movimiento = '' OR "
        End If
        If chkAnticipoProveedores.CheckState = System.Windows.Forms.CheckState.Checked Then
            sql = sql & "MB.Movimiento = '" & C_MOVANTICIPOS & "' OR "
        Else
            sql = sql & "MB.Movimiento = '' OR "
        End If
        If chkOtrosIngresos.CheckState = System.Windows.Forms.CheckState.Checked Then
            sql = sql & "MB.Movimiento = '" & C_OTROSINGRESOS & "') AND "
        Else
            sql = sql & "MB.Movimiento = '') AND "
        End If
        If chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            sql = sql & "MB.Moneda = 'P' AND "
        End If
        If chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            sql = sql & "MB.Moneda = 'D' AND "
        End If
        sql = sql & "FechaMovto BETWEEN '" & VB6.Format(dtpFechaInicial.Value, "MM/DD/YYYY") & "' AND '" & VB6.Format(dtpFechaFinal.Value, "MM/DD/YYYY") & "' "
        sql = sql & "ORDER BY MB.FolioMovto DESC"
        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = UCase("Reporte de Movimientos Bancarios por Tipo")
        PeriodoReporte = "Del  " & VB6.Format(dtpFechaInicial.Value, "DD/MMM/YYYY") & "  Hasta el  " & VB6.Format(dtpFechaFinal.Value, "DD/MMM/YYYY")
        BorraCmd()
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = sql
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existen movimientos en el periodo especificado, Favor de verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        Else
            rptBancosReportedeMovimientosporTipo1.SetDataSource(frmReportes.rsReport)
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "Periodo", "Nota"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte, Nota} 


        'If (NombreEmpresa <> Nothing) Then
        '    pdvNum.Value = NombreEmpresa.ToString() : pvNum.Add(pdvNum)
        '    rptBancosReportedeMovimientosporTipo1.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pdvNum)
        'End If
        'If (NombreReporte <> Nothing) Then
        '    pdvNum.Value = NombreReporte.ToString() : pvNum.Add(pdvNum)
        '    rptBancosReportedeMovimientosporTipo1.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
        'End If
        'If (PeriodoReporte <> Nothing) Then
        '    pdvNum.Value = PeriodoReporte.ToString() : pvNum.Add(pdvNum)
        '    rptBancosReportedeMovimientosporTipo1.DataDefinition.ParameterFields("Periodo").ApplyCurrentValues(pvNum)
        'End If
        'If (Nota <> Nothing) Then
        '    pdvNum.Value = Nota.ToString() : pvNum.Add(pdvNum)
        '    rptBancosReportedeMovimientosporTipo1.DataDefinition.ParameterFields("Nota").ApplyCurrentValues(pvNum)
        'End If

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.Text = "Reporte de Movimientos Bancarios por Tipo"
        frmReportes.reporteActual = rptBancosReportedeMovimientosporTipo1
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
        optPesos.Checked = True
        optDolares.Checked = False
        chkPagos.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkDepositos.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkCargosDiversos.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkTraspasosBancarios.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkAnticipoProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkOtrosIngresos.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkIncluirFoliosCancel.CheckState = System.Windows.Forms.CheckState.Checked
    End Sub

    Private Sub chkAnticipoProveedores_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAnticipoProveedores.Enter
        Pon_Tool()
    End Sub

    Private Sub chkCargosDiversos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCargosDiversos.Enter
        Pon_Tool()
    End Sub

    Private Sub chkDepositos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDepositos.Enter
        Pon_Tool()
    End Sub

    Private Sub chkDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub chkIncluirFoliosCancel_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIncluirFoliosCancel.Enter
        Pon_Tool()
    End Sub

    Private Sub chkOtrosIngresos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkOtrosIngresos.Enter
        Pon_Tool()
    End Sub

    Private Sub chkPagos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPagos.Enter
        Pon_Tool()
    End Sub

    Private Sub chkPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPesos.Enter
        Pon_Tool()
    End Sub

    Private Sub chkTraspasosBancarios_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTraspasosBancarios.Enter
        Pon_Tool()
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

    Private Sub dtpFechaFinal_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dtpFechaFinal.KeyPress
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

    Private Sub dtpFechaInicial_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dtpFechaInicial.KeyPress
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub frmBancosReporteMovBancariosXTipo_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosReporteMovBancariosXTipo_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosReporteMovBancariosXTipo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
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

    Private Sub frmBancosReporteMovBancariosXTipo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosReporteMovBancariosXTipo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        Nuevo()
    End Sub

    Private Sub frmBancosReporteMovBancariosXTipo_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        If mblnSalir Then
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSalir = False
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmBancosReporteMovBancariosXTipo_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
    End Sub

    Private Sub optDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub optPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPesos.Enter
        Pon_Tool()
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkIncluirFoliosCancel = New System.Windows.Forms.CheckBox()
        Me.chkDolares = New System.Windows.Forms.CheckBox()
        Me.chkPesos = New System.Windows.Forms.CheckBox()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.chkOtrosIngresos = New System.Windows.Forms.CheckBox()
        Me.chkAnticipoProveedores = New System.Windows.Forms.CheckBox()
        Me.chkTraspasosBancarios = New System.Windows.Forms.CheckBox()
        Me.chkCargosDiversos = New System.Windows.Forms.CheckBox()
        Me.chkDepositos = New System.Windows.Forms.CheckBox()
        Me.chkPagos = New System.Windows.Forms.CheckBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkIncluirFoliosCancel
        '
        Me.chkIncluirFoliosCancel.BackColor = System.Drawing.SystemColors.Control
        Me.chkIncluirFoliosCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncluirFoliosCancel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkIncluirFoliosCancel.Location = New System.Drawing.Point(288, 260)
        Me.chkIncluirFoliosCancel.Name = "chkIncluirFoliosCancel"
        Me.chkIncluirFoliosCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncluirFoliosCancel.Size = New System.Drawing.Size(147, 18)
        Me.chkIncluirFoliosCancel.TabIndex = 12
        Me.chkIncluirFoliosCancel.Text = "Incluír Folios Cancelados"
        Me.ToolTip1.SetToolTip(Me.chkIncluirFoliosCancel, "Incluir Movimientos Cancelados")
        Me.chkIncluirFoliosCancel.UseVisualStyleBackColor = False
        '
        'chkDolares
        '
        Me.chkDolares.BackColor = System.Drawing.SystemColors.Control
        Me.chkDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDolares.Location = New System.Drawing.Point(22, 38)
        Me.chkDolares.Name = "chkDolares"
        Me.chkDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDolares.Size = New System.Drawing.Size(66, 22)
        Me.chkDolares.TabIndex = 11
        Me.chkDolares.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me.chkDolares, "Muestra los Movimientos Cuya Moneda de Operación es Dolares")
        Me.chkDolares.UseVisualStyleBackColor = False
        '
        'chkPesos
        '
        Me.chkPesos.BackColor = System.Drawing.SystemColors.Control
        Me.chkPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPesos.Location = New System.Drawing.Point(22, 16)
        Me.chkPesos.Name = "chkPesos"
        Me.chkPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPesos.Size = New System.Drawing.Size(66, 22)
        Me.chkPesos.TabIndex = 10
        Me.chkPesos.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me.chkPesos, "Muestra los Movimientos Cuya Moneda de Operación es Pesos")
        Me.chkPesos.UseVisualStyleBackColor = False
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(22, 38)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(90, 21)
        Me.optDolares.TabIndex = 9
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me.optDolares, "Muestra los Importes en Dolares")
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(22, 16)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(90, 21)
        Me.optPesos.TabIndex = 8
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me.optPesos, "Muestra los Importes en Pesos")
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'chkOtrosIngresos
        '
        Me.chkOtrosIngresos.BackColor = System.Drawing.SystemColors.Control
        Me.chkOtrosIngresos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkOtrosIngresos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkOtrosIngresos.Location = New System.Drawing.Point(231, 74)
        Me.chkOtrosIngresos.Name = "chkOtrosIngresos"
        Me.chkOtrosIngresos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkOtrosIngresos.Size = New System.Drawing.Size(106, 21)
        Me.chkOtrosIngresos.TabIndex = 7
        Me.chkOtrosIngresos.Text = "Otros Ingresos"
        Me.ToolTip1.SetToolTip(Me.chkOtrosIngresos, "Muestra los Movimientos de Otros Ingresos")
        Me.chkOtrosIngresos.UseVisualStyleBackColor = False
        '
        'chkAnticipoProveedores
        '
        Me.chkAnticipoProveedores.BackColor = System.Drawing.SystemColors.Control
        Me.chkAnticipoProveedores.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAnticipoProveedores.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAnticipoProveedores.Location = New System.Drawing.Point(231, 47)
        Me.chkAnticipoProveedores.Name = "chkAnticipoProveedores"
        Me.chkAnticipoProveedores.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAnticipoProveedores.Size = New System.Drawing.Size(141, 21)
        Me.chkAnticipoProveedores.TabIndex = 6
        Me.chkAnticipoProveedores.Text = "Anticipos a Proveedores"
        Me.ToolTip1.SetToolTip(Me.chkAnticipoProveedores, "Muestra los Movimientos de Anticipo a Proveedores")
        Me.chkAnticipoProveedores.UseVisualStyleBackColor = False
        '
        'chkTraspasosBancarios
        '
        Me.chkTraspasosBancarios.BackColor = System.Drawing.SystemColors.Control
        Me.chkTraspasosBancarios.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTraspasosBancarios.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTraspasosBancarios.Location = New System.Drawing.Point(231, 20)
        Me.chkTraspasosBancarios.Name = "chkTraspasosBancarios"
        Me.chkTraspasosBancarios.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTraspasosBancarios.Size = New System.Drawing.Size(141, 21)
        Me.chkTraspasosBancarios.TabIndex = 5
        Me.chkTraspasosBancarios.Text = "Traspasos Bancarios"
        Me.ToolTip1.SetToolTip(Me.chkTraspasosBancarios, "Muestra los Movimientos de Traspasos Bancarios")
        Me.chkTraspasosBancarios.UseVisualStyleBackColor = False
        '
        'chkCargosDiversos
        '
        Me.chkCargosDiversos.BackColor = System.Drawing.SystemColors.Control
        Me.chkCargosDiversos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCargosDiversos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCargosDiversos.Location = New System.Drawing.Point(60, 74)
        Me.chkCargosDiversos.Name = "chkCargosDiversos"
        Me.chkCargosDiversos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCargosDiversos.Size = New System.Drawing.Size(106, 21)
        Me.chkCargosDiversos.TabIndex = 4
        Me.chkCargosDiversos.Text = "Cargos Diversos"
        Me.ToolTip1.SetToolTip(Me.chkCargosDiversos, "Muestra los Movimientos de Cargos Diversos")
        Me.chkCargosDiversos.UseVisualStyleBackColor = False
        '
        'chkDepositos
        '
        Me.chkDepositos.BackColor = System.Drawing.SystemColors.Control
        Me.chkDepositos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDepositos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDepositos.Location = New System.Drawing.Point(60, 47)
        Me.chkDepositos.Name = "chkDepositos"
        Me.chkDepositos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDepositos.Size = New System.Drawing.Size(81, 21)
        Me.chkDepositos.TabIndex = 3
        Me.chkDepositos.Text = "Depósitos"
        Me.ToolTip1.SetToolTip(Me.chkDepositos, "Muestra los Movimientos de Depósitos")
        Me.chkDepositos.UseVisualStyleBackColor = False
        '
        'chkPagos
        '
        Me.chkPagos.BackColor = System.Drawing.SystemColors.Control
        Me.chkPagos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPagos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPagos.Location = New System.Drawing.Point(60, 20)
        Me.chkPagos.Name = "chkPagos"
        Me.chkPagos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPagos.Size = New System.Drawing.Size(61, 21)
        Me.chkPagos.TabIndex = 2
        Me.chkPagos.Text = "Pagos"
        Me.ToolTip1.SetToolTip(Me.chkPagos, "Muestra los Movimientos de Pagos")
        Me.chkPagos.UseVisualStyleBackColor = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.chkDolares)
        Me.Frame4.Controls.Add(Me.chkPesos)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(152, 209)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(121, 67)
        Me.Frame4.TabIndex = 18
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Moneda"
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.optDolares)
        Me.Frame3.Controls.Add(Me.optPesos)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(16, 209)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(121, 67)
        Me.Frame3.TabIndex = 17
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Presentar en"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.chkOtrosIngresos)
        Me.Frame2.Controls.Add(Me.chkAnticipoProveedores)
        Me.Frame2.Controls.Add(Me.chkTraspasosBancarios)
        Me.Frame2.Controls.Add(Me.chkCargosDiversos)
        Me.Frame2.Controls.Add(Me.chkDepositos)
        Me.Frame2.Controls.Add(Me.chkPagos)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(16, 89)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(437, 106)
        Me.Frame2.TabIndex = 16
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Movimientos"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFechaInicial)
        Me.Frame1.Controls.Add(Me.dtpFechaFinal)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(16, 15)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(437, 65)
        Me.Frame1.TabIndex = 13
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(108, 24)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(95, 20)
        Me.dtpFechaInicial.TabIndex = 0
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(275, 24)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaFinal.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(223, 27)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(62, 21)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Hasta el :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(53, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(62, 21)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Desde el :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(131, 301)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 20
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(16, 301)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 19
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmBancosReporteMovBancariosXTipo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(469, 343)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.chkIncluirFoliosCancel)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(244, 96)
        Me.MaximizeBox = False
        Me.Name = "frmBancosReporteMovBancariosXTipo"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Movimientos Bancarios por Tipo"
        Me.Frame4.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
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