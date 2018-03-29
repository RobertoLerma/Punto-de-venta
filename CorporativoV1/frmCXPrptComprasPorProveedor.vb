Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCXPrptComprasPorProveedor
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents chkTodosProv As System.Windows.Forms.CheckBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_2 As System.Windows.Forms.RadioButton
    Public WithEvents _fraRpt_2 As System.Windows.Forms.GroupBox
    Public WithEvents _chkMoneda_2 As System.Windows.Forms.CheckBox
    Public WithEvents _chkMoneda_1 As System.Windows.Forms.CheckBox
    Public WithEvents _chkMoneda_0 As System.Windows.Forms.CheckBox
    Public WithEvents _fraRpt_1 As System.Windows.Forms.GroupBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblRpt_1 As System.Windows.Forms.Label
    Public WithEvents _lblRpt_0 As System.Windows.Forms.Label
    Public WithEvents _fraRpt_0 As System.Windows.Forms.GroupBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents chkMoneda As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
    Public WithEvents fraRpt As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray

    Const C_DOLARES As Integer = 0
    Const C_PESOS As Integer = 1
    Const C_EUROS As Integer = 2

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean
    Dim mintCodProveedor As Integer
    Dim FueraChange As Boolean
    Dim Tecla As Integer
    Dim cMonedaDeCantidades As String 'Moneda en la que estarán expresadas las cantidades en el reporte
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim mblnSalir As Boolean

    Public Sub Limpiar()
        On Error Resume Next
        Call Me.Nuevo()
        Me.chkTodosProv.Focus()
    End Sub

    Public Sub Nuevo()
        Me.chkTodosProv.CheckState = System.Windows.Forms.CheckState.Checked
        Me.dbcProveedor.Text = ""
        Me.dtpDesde.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.dtpHasta.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.chkMoneda(C_DOLARES).CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMoneda(C_PESOS).CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMoneda(C_EUROS).CheckState = System.Windows.Forms.CheckState.Checked
        Me.optMoneda(C_DOLARES).Checked = True
        Me.optMoneda(C_PESOS).Checked = False
        Me.optMoneda(C_EUROS).Checked = False
        FueraChange = False
        Me.txtMensaje.Text = ""
        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
        mintCodProveedor = 0
    End Sub

    Public Function DevuelveQuery() As String
        On Error Resume Next
        Dim I As Integer
        Dim cMoneda As String

        Dim cSELECT As String
        Dim cFROM As String
        Dim cWHERE As String
        Dim cGROUPBY As String
        Dim cORDERBY As String

        'Obtener el WHERE
        cWHERE = " where ( a.Estatus = '" & C_STGENERADA & "' or a.Estatus = '" & C_STREGISTRADA & "' ) "
        'Por tipo de moneda
        Select Case True
            Case Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Checked
                ' Dólares - Pesos - Euros
                'No importa qué tipo de moneda quiera, las quiere todas
            Case Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                ' Dólares - Pesos
                cWHERE = cWHERE & " and a.Moneda <> 'E' "
            Case Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Checked
                ' Dólares - Euros
                cWHERE = cWHERE & " and a.Moneda <> 'P' "
            Case Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                ' Dólares
                cWHERE = cWHERE & " and a.Moneda = 'D' "
            Case Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                ' ERROR - Debe seleccionar por lo menos un tipo de moneda
                DevuelveQuery = ""
                Exit Function
            Case Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                ' Pesos
                cWHERE = cWHERE & " and a.Moneda = 'P' "
            Case Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Checked
                ' Pesos - Euros
                cWHERE = cWHERE & " and a.Moneda <> 'D' "
            Case Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Checked
                ' Euros
                cWHERE = cWHERE & " and a.Moneda = 'E' "
        End Select

        'Por Intervalo de Fechas
        cWHERE = cWHERE & " and ( a.FechaCompraEI >= '" & VB6.Format(Me.dtpDesde.Value, "mm/dd/yyyy") & "' and a.FechaCompraEI <= '" & VB6.Format(Me.dtpHasta.Value, "mm/dd/yyyy") & "')" & IIf(mintCodProveedor <> 0, " AND A.CodProvAcreed = " & mintCodProveedor & " ", "")

        'Convertir los totales a la moneda indicada
        If Me.optMoneda(C_DOLARES).Checked Then
            cMoneda = C_DOLAR
            cMonedaDeCantidades = "** Los importes están expresados en Dólares (USD)"
        ElseIf Me.optMoneda(C_PESOS).Checked Then
            cMoneda = C_PESO
            cMonedaDeCantidades = "** Los importes están expresados en Pesos"
        Else
            cMoneda = C_EURO
            cMonedaDeCantidades = "** Los importes están expresados en Euros"
        End If

        cSELECT = " select a.codProvAcreed, a.FolioOrdenCompra, dbo.RegresaCHAR(c.FolioFactura) as FolioFactura, a.FechaCompraEI as FechaCompraEI, LTrim(RTrim( b.DescProvAcreed )) as NomProv, " & "           round(dbo.ConvertirCantidad( a.Moneda, '" & cMoneda & "', a.SubTotal, a.TipoCambioC, a.TipoCambioEuroC ),2) as SubTotal, " & "           round(dbo.ConvertirCantidad( a.Moneda, '" & cMoneda & "', a.Descuento, a.TipoCambioC, a.TipoCambioEuroC ),2) as Descuento, " & "           round(dbo.ConvertirCantidad( a.Moneda, '" & cMoneda & "', a.Iva, a.TipoCambioC, a.TipoCambioEuroC ),2) as Iva, " & "           SUM(ISNULL(CASE WHEN N.Moneda = '" & cMoneda & "' THEN N.Total ELSE DBO.ConvertirCantidad(N.Moneda,'" & cMoneda & "',N.Total,CASE WHEN N.Estatus = 'A' THEN N.TipoCambioAplic WHEN N.Estatus = 'V' THEN N.TipoCambio END,CASE WHEN N.Estatus = 'A' THEN N.TipoCambioEuroAplic WHEN N.Estatus = 'V' THEN N.TipoCambioEuro END)END,0)) AS NotasCred," & "           (round(dbo.ConvertirCantidad( a.Moneda, '" & cMoneda & "', a.Total, a.TipoCambioC, a.TipoCambioEuroC ),2) - " & "           SUM(ISNULL(CASE WHEN N.Moneda = '" & cMoneda & "' THEN N.Total ELSE DBO.ConvertirCantidad(N.Moneda,'" & cMoneda & "',N.Total,CASE WHEN N.Estatus = 'A' THEN N.TipoCambioAplic WHEN N.Estatus = 'V' THEN N.TipoCambio END,CASE WHEN N.Estatus = 'A' THEN N.TipoCambioEuroAplic WHEN N.Estatus = 'V' THEN N.TipoCambioEuro END)END,0))) AS Total, " & "           a.Estatus as EstausOC "

        cFROM = " from    OrdenesCompra a INNER JOIN CatProvAcreed b " & "                               ON a.CodProvAcreed = b.CodProvAcreed " & "                           LEFT OUTER JOIN CxPFacturas c " & "                               ON a.FolioOrdenCompra = c.FolioOrdenCompra and c.Estatus <> '" & C_STCANCELADA & "' " & "LEFT OUTER JOIN (SELECT * FROM NotasCreditoCab WHERE TipoNotaCredito = 'D' AND Estatus <> 'C') N ON A.CodProvAcreed = N.CodProvAcreed AND c.CodProvAcreed = N.CodProvAcreed AND c.FolioFactura = N.FolioFactura "

        cGROUPBY = ""

        cORDERBY = " group by a.codProvAcreed,a.FolioOrdenCompra,c.FolioFactura,a.FechaCompraEI,b.DescProvAcreed,a.SubTotal," & "a.Descuento,a.Iva,a.Total,a.Estatus,a.TipoCambioEuroC,a.TipoCambioC,a.Moneda Order by a.codProvAcreed, a.FolioOrdenCompra "

        DevuelveQuery = cSELECT & cFROM & cWHERE & cGROUPBY & cORDERBY

    End Function

    Public Sub Imprime()

        Dim rptMejoresProvDetalle As New rptMejoresProvDetalle
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        'On Error GoTo MErr

        Dim lStrSql As String
        'Declarar vectores para almacenar los parámetros que se le enviarán al reporte
        Dim aParam(5) As Object
        Dim aValues(5) As Object

        If Not ValidaDatos() Then
            Exit Sub
        End If

        lStrSql = DevuelveQuery()
        gStrSql = lStrSql
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existen datos para el rango de fechas indicado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            rptMejoresProvDetalle.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "Mensaje"
        'aValues(1) = Trim(Me.txtMensaje.Text)
        'aParam(2) = "dDesde"
        'aValues(2) = Me.dtpDesde.Value
        'aParam(3) = "dHasta"
        'aValues(3) = Me.dtpHasta.Value
        'aParam(4) = "MonedaDeCantidades"
        'aValues(4) = Trim(cMonedaDeCantidades)
        'aParam(5) = "Empresa"
        'aValues(5) = Trim(gstrNombCortoEmpresa)

        If (txtMensaje.Text <> Nothing Or txtMensaje.Text <> "") Then
            pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
            rptMejoresProvDetalle.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptMejoresProvDetalle.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        End If

        If (dtpDesde.Value <> Nothing) Then
            pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
            rptMejoresProvDetalle.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
        End If

        If (dtpHasta.Value <> Nothing) Then
            pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
            rptMejoresProvDetalle.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
        End If

        If (cMonedaDeCantidades <> Nothing Or cMonedaDeCantidades <> "") Then
            pdvNum.Value = cMonedaDeCantidades : pvNum.Add(pdvNum)
            rptMejoresProvDetalle.DataDefinition.ParameterFields("MonedaDeCantidades").ApplyCurrentValues(pvNum)
        End If

        If (gstrNombCortoEmpresa <> Nothing Or gstrNombCortoEmpresa <> "") Then
            pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
            rptMejoresProvDetalle.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
        End If

        'frmReportes.Report = rptMejoresProvDetalle 'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Imprime(Trim(Me.Text), aParam, aValues)
        frmReportes.reporteActual = rptMejoresProvDetalle
        frmReportes.Show()

MErr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function ValidaDatos() As Boolean
        If mblnTecleoFechaI Then
            Do While (VB.Timer() - msglTiempoCambioI) <= 2.1
            Loop
            mblnTecleoFechaI = False
        End If
        If mblnTecleoFechaF Then
            Do While (VB.Timer() - msglTiempoCambioF) <= 2.1
            Loop
            mblnTecleoFechaF = False
        End If
        System.Windows.Forms.Application.DoEvents()
        Select Case True
            Case Me.dtpDesde.Value > Me.dtpHasta.Value
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Límite", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dtpDesde.Focus()
                Exit Function
            Case Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                MsgBox("Debe seleccionar por lo menos un tipo de moneda", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.chkMoneda(0).Focus()
                Exit Function
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Private Sub chkMoneda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkMoneda.Enter
        Dim Index As Integer = chkMoneda.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub chkTodosProv_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodosProv.CheckStateChanged
        If chkTodosProv.CheckState = System.Windows.Forms.CheckState.Checked Then
            mintCodProveedor = 0
            FueraChange = True
            dbcProveedor.Text = ""
            dbcProveedor.Enabled = False
            FueraChange = False
        Else
            dbcProveedor.Enabled = True
        End If
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo MErr
        Dim lStrSql As String
        If FueraChange Then Exit Sub

        lStrSql = "SELECT CodProvAcreed, RTRIM(LTrim(RTrim(descProvAcreed))) as descProvAcreed FROM CatProvAcreed Where DescProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%' AND Tipo = 'P'"
        ModDCombo.DCChange(lStrSql, Tecla, (Me.dbcProveedor))

        If Trim(Me.dbcProveedor.Text) = "" Then
            mintCodProveedor = 0
        End If
MErr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        Pon_Tool()
        gStrSql = "SELECT CodProvAcreed, RTRIM(LTrim(RTrim(DescProvAcreed))) as DescProvAcreed FROM CatProvAcreed WHERE Tipo = 'P'"
        ModDCombo.DCGotFocus(gStrSql, Me.dbcProveedor)
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcProveedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodosProv.Focus()
        End If
        Tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcProveedor_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcProveedor.KeyUp
        Dim Aux As String
        Aux = Trim(Me.dbcProveedor.Text)
        'If Me.dbcProveedor.SelectedItem <> 0 Then
        '    dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        'End If
        Me.dbcProveedor.Text = Aux
    End Sub

    Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        gStrSql = "SELECT CodProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM CatProvAcreed Where DescProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%' AND Tipo = 'P'"
        ModDCombo.DCLostFocus((Me.dbcProveedor), gStrSql, mintCodProveedor)
    End Sub

    Private Sub dbcProveedor_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcProveedor.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcProveedor.Text)
        'If Me.dbcProveedor.SelectedItem <> 0 Then
        '    dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        'End If
        Me.dbcProveedor.Text = Aux
    End Sub

    Private Sub dtpDesde_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpDesde.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpDesde_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dtpDesde.KeyPress
        mblnTecleoFechaI = True
        msglTiempoCambioI = VB.Timer()
    End Sub

    Private Sub dtpHasta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpHasta.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpHasta_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dtpHasta.KeyPress
        mblnTecleoFechaF = True
        msglTiempoCambioF = VB.Timer()
    End Sub

    Private Sub frmCXPrptComprasPorProveedor_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCXPrptComprasPorProveedor_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPrptComprasPorProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "CHKTODOSPROV" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCXPrptComprasPorProveedor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPrptComprasPorProveedor_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Me.dtpDesde.MinDate = C_FECHAINICIAL
        Me.dtpDesde.MaxDate = C_FECHAFINAL
        Me.dtpHasta.MinDate = C_FECHAINICIAL
        Me.dtpHasta.MaxDate = C_FECHAFINAL
        Call Me.Nuevo()
    End Sub

    Private Sub frmCXPrptComprasPorProveedor_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If mblnSalir Then
        '    mblnSalir = False
        '    Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Me.chkTodosProv.Focus()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCXPrptComprasPorProveedor_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optMoneda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoneda.Enter
        Dim Index As Integer = optMoneda.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_2 = New System.Windows.Forms.RadioButton()
        Me._chkMoneda_2 = New System.Windows.Forms.CheckBox()
        Me._chkMoneda_1 = New System.Windows.Forms.CheckBox()
        Me._chkMoneda_0 = New System.Windows.Forms.CheckBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me.chkTodosProv = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me._fraRpt_2 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_1 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_0 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblRpt_1 = New System.Windows.Forms.Label()
        Me._lblRpt_0 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.chkMoneda = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.fraRpt = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me._fraRpt_2.SuspendLayout()
        Me._fraRpt_1.SuspendLayout()
        Me._fraRpt_0.SuspendLayout()
        CType(Me.chkMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(8, 263)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(345, 73)
        Me.txtMensaje.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        '_optMoneda_0
        '
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_0, CType(0, Short))
        Me._optMoneda_0.Location = New System.Drawing.Point(24, 24)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(66, 17)
        Me._optMoneda_0.TabIndex = 12
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me._optMoneda_0, "Los importes del reporte aparecerán en dólares")
        Me._optMoneda_0.UseVisualStyleBackColor = False
        '
        '_optMoneda_1
        '
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_1, CType(1, Short))
        Me._optMoneda_1.Location = New System.Drawing.Point(24, 48)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(66, 17)
        Me._optMoneda_1.TabIndex = 13
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me._optMoneda_1, "Los importes del reporte aparecerán en Pesos")
        Me._optMoneda_1.UseVisualStyleBackColor = False
        '
        '_optMoneda_2
        '
        Me._optMoneda_2.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_2, CType(2, Short))
        Me._optMoneda_2.Location = New System.Drawing.Point(24, 72)
        Me._optMoneda_2.Name = "_optMoneda_2"
        Me._optMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_2.Size = New System.Drawing.Size(66, 17)
        Me._optMoneda_2.TabIndex = 14
        Me._optMoneda_2.TabStop = True
        Me._optMoneda_2.Text = "Euros"
        Me.ToolTip1.SetToolTip(Me._optMoneda_2, "Los importes del reporte aparecerán en Euros")
        Me._optMoneda_2.UseVisualStyleBackColor = False
        '
        '_chkMoneda_2
        '
        Me._chkMoneda_2.BackColor = System.Drawing.SystemColors.Control
        Me._chkMoneda_2.Checked = True
        Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMoneda.SetIndex(Me._chkMoneda_2, CType(2, Short))
        Me._chkMoneda_2.Location = New System.Drawing.Point(24, 72)
        Me._chkMoneda_2.Name = "_chkMoneda_2"
        Me._chkMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_2.Size = New System.Drawing.Size(81, 17)
        Me._chkMoneda_2.TabIndex = 10
        Me._chkMoneda_2.Text = "Euros"
        Me.ToolTip1.SetToolTip(Me._chkMoneda_2, "Selecciona todas las compras en Euros")
        Me._chkMoneda_2.UseVisualStyleBackColor = False
        '
        '_chkMoneda_1
        '
        Me._chkMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._chkMoneda_1.Checked = True
        Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMoneda.SetIndex(Me._chkMoneda_1, CType(1, Short))
        Me._chkMoneda_1.Location = New System.Drawing.Point(24, 48)
        Me._chkMoneda_1.Name = "_chkMoneda_1"
        Me._chkMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_1.Size = New System.Drawing.Size(81, 17)
        Me._chkMoneda_1.TabIndex = 9
        Me._chkMoneda_1.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me._chkMoneda_1, "Selecciona todas las compras en Pesos")
        Me._chkMoneda_1.UseVisualStyleBackColor = False
        '
        '_chkMoneda_0
        '
        Me._chkMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._chkMoneda_0.Checked = True
        Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMoneda.SetIndex(Me._chkMoneda_0, CType(0, Short))
        Me._chkMoneda_0.Location = New System.Drawing.Point(24, 24)
        Me._chkMoneda_0.Name = "_chkMoneda_0"
        Me._chkMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_0.Size = New System.Drawing.Size(81, 17)
        Me._chkMoneda_0.TabIndex = 8
        Me._chkMoneda_0.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me._chkMoneda_0, "Selecciona todas las compras en Dólares")
        Me._chkMoneda_0.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dbcProveedor)
        Me.Frame1.Controls.Add(Me.chkTodosProv)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(9, 5)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(345, 60)
        Me.Frame1.TabIndex = 17
        Me.Frame1.TabStop = False
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(85, 28)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(254, 21)
        Me.dbcProveedor.TabIndex = 1
        '
        'chkTodosProv
        '
        Me.chkTodosProv.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodosProv.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodosProv.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTodosProv.Location = New System.Drawing.Point(16, 11)
        Me.chkTodosProv.Name = "chkTodosProv"
        Me.chkTodosProv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodosProv.Size = New System.Drawing.Size(269, 21)
        Me.chkTodosProv.TabIndex = 0
        Me.chkTodosProv.Text = "Todos los Proveedores"
        Me.chkTodosProv.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(26, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(63, 21)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Proveedor"
        '
        '_fraRpt_2
        '
        Me._fraRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_2.Controls.Add(Me._optMoneda_0)
        Me._fraRpt_2.Controls.Add(Me._optMoneda_1)
        Me._fraRpt_2.Controls.Add(Me._optMoneda_2)
        Me._fraRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_2, CType(2, Short))
        Me._fraRpt_2.Location = New System.Drawing.Point(184, 143)
        Me._fraRpt_2.Name = "_fraRpt_2"
        Me._fraRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_2.Size = New System.Drawing.Size(169, 97)
        Me._fraRpt_2.TabIndex = 11
        Me._fraRpt_2.TabStop = False
        Me._fraRpt_2.Text = "Presentar en ..."
        '
        '_fraRpt_1
        '
        Me._fraRpt_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_1.Controls.Add(Me._chkMoneda_2)
        Me._fraRpt_1.Controls.Add(Me._chkMoneda_1)
        Me._fraRpt_1.Controls.Add(Me._chkMoneda_0)
        Me._fraRpt_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_1, CType(1, Short))
        Me._fraRpt_1.Location = New System.Drawing.Point(8, 143)
        Me._fraRpt_1.Name = "_fraRpt_1"
        Me._fraRpt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_1.Size = New System.Drawing.Size(169, 97)
        Me._fraRpt_1.TabIndex = 7
        Me._fraRpt_1.TabStop = False
        Me._fraRpt_1.Text = "Moneda"
        '
        '_fraRpt_0
        '
        Me._fraRpt_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_0.Controls.Add(Me.dtpDesde)
        Me._fraRpt_0.Controls.Add(Me.dtpHasta)
        Me._fraRpt_0.Controls.Add(Me._lblRpt_1)
        Me._fraRpt_0.Controls.Add(Me._lblRpt_0)
        Me._fraRpt_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_0, CType(0, Short))
        Me._fraRpt_0.Location = New System.Drawing.Point(9, 71)
        Me._fraRpt_0.Name = "_fraRpt_0"
        Me._fraRpt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_0.Size = New System.Drawing.Size(345, 57)
        Me._fraRpt_0.TabIndex = 2
        Me._fraRpt_0.TabStop = False
        Me._fraRpt_0.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(72, 21)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(97, 20)
        Me.dtpDesde.TabIndex = 4
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(232, 21)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(97, 20)
        Me.dtpHasta.TabIndex = 6
        '
        '_lblRpt_1
        '
        Me._lblRpt_1.AutoSize = True
        Me._lblRpt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRpt.SetIndex(Me._lblRpt_1, CType(1, Short))
        Me._lblRpt_1.Location = New System.Drawing.Point(184, 25)
        Me._lblRpt_1.Name = "_lblRpt_1"
        Me._lblRpt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_1.Size = New System.Drawing.Size(44, 13)
        Me._lblRpt_1.TabIndex = 5
        Me._lblRpt_1.Text = "hasta el"
        '
        '_lblRpt_0
        '
        Me._lblRpt_0.AutoSize = True
        Me._lblRpt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRpt.SetIndex(Me._lblRpt_0, CType(0, Short))
        Me._lblRpt_0.Location = New System.Drawing.Point(16, 25)
        Me._lblRpt_0.Name = "_lblRpt_0"
        Me._lblRpt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_0.Size = New System.Drawing.Size(49, 13)
        Me._lblRpt_0.TabIndex = 3
        Me._lblRpt_0.Text = "Desde el"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblRpt.SetIndex(Me._lblRpt_2, CType(2, Short))
        Me._lblRpt_2.Location = New System.Drawing.Point(8, 249)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 15
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'chkMoneda
        '
        '
        'optMoneda
        '
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(128, 352)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 121
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(13, 352)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 120
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(243, 353)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 119
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmCXPrptComprasPorProveedor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(361, 410)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me._fraRpt_2)
        Me.Controls.Add(Me._fraRpt_1)
        Me.Controls.Add(Me._fraRpt_0)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 29)
        Me.MaximizeBox = False
        Me.Name = "frmCXPrptComprasPorProveedor"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Análisis anual de compras"
        Me.Frame1.ResumeLayout(False)
        Me._fraRpt_2.ResumeLayout(False)
        Me._fraRpt_1.ResumeLayout(False)
        Me._fraRpt_0.ResumeLayout(False)
        Me._fraRpt_0.PerformLayout()
        CType(Me.chkMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class