Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCXPrptOC
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents _chkOrdenes_0 As System.Windows.Forms.CheckBox
    Public WithEvents _chkOrdenes_1 As System.Windows.Forms.CheckBox
    Public WithEvents _chkOrdenes_2 As System.Windows.Forms.CheckBox
    Public WithEvents _fraRpt_5 As System.Windows.Forms.GroupBox
    Public WithEvents _optOC_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optOC_0 As System.Windows.Forms.RadioButton
    Public WithEvents _fraRpt_4 As System.Windows.Forms.GroupBox
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
    Public WithEvents chkOrdenes As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
    Public WithEvents fraRpt As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optOC As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray

    Const C_DOLARES As Integer = 0
    Const C_PESOS As Integer = 1
    Const C_EUROS As Integer = 2

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean

    Dim cMonedaDeCantidades As String 'Moneda en la que estarán expresadas las cantidades en el reporte
    Dim cTituloOC As String
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim mblnSALIR As Boolean

    Public Sub Limpiar()
        On Error Resume Next
        Call Me.Nuevo()
        Me.dtpDesde.Focus()
    End Sub

    Public Sub Nuevo()
        Me._optOC_0.Checked = True
        Me._optOC_1.Checked = False
        Call _optOC_0_CheckedChanged(_optOC_0, New System.EventArgs())
        Call _optOC_1_CheckedChanged(_optOC_1, New System.EventArgs())
        Me.dtpDesde.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.dtpHasta.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Checked
        Me._optMoneda_0.Checked = True
        Me._optMoneda_1.Checked = False
        Me._optMoneda_2.Checked = False
        Me.txtMensaje.Text = ""
        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
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

        If Me._optOC_0.Checked Then
            'Título del reporte
            cTituloOC = "EMITIDAS"
        ElseIf Me._optOC_1.Checked Then
            cTituloOC = "RECIBIDAS"
        End If

        'Obtener el WHERE
        'Por Estatus de la orden de compra
        cWHERE = " WHERE "
        Select Case True
            Case Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Checked
                'Pendientes - Recibidas - Canceladas
                'No Importa, las quiere todas
                cWHERE = cWHERE & " a.Estatus <> '' "
            Case Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                'Pendientes - Recibidas
                cWHERE = cWHERE & " a.Estatus <> '" & C_STCANCELADA & "' "
            Case Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Checked
                'Pendientes - Canceladas
                cWHERE = cWHERE & " a.Estatus <> '" & C_STGENERADA & "' and a.Estatus <> '" & C_STREGISTRADA & "' "
            Case Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                'Pendientes
                cWHERE = cWHERE & " a.Estatus = '" & C_STVIGENTE & "' "
            Case Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Checked
                'Recibidas - Canceladas
                cWHERE = cWHERE & " a.Estatus <> '" & C_STVIGENTE & "' "
            Case Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                'Recibidas
                cWHERE = cWHERE & " (a.Estatus = '" & C_STGENERADA & "' or a.Estatus = '" & C_STREGISTRADA & "') "
            Case Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Checked
                'Canceladas
                cWHERE = cWHERE & " a.Estatus = '" & C_STCANCELADA & "' "
            Case Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                'Ninguna
                'Error
                DevuelveQuery = ""
                Exit Function
        End Select
        'Por tipo de moneda
        Select Case True
            Case Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Checked
                ' Dólares - Pesos - Euros
                'No importa qué tipo de moneda quiera, las quiere todas
            Case Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                ' Dólares - Pesos
                cWHERE = cWHERE & " and a.Moneda <> 'E' "
            Case Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Checked
                ' Dólares - Euros
                cWHERE = cWHERE & " and a.Moneda <> 'P' "
            Case Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                ' Dólares
                cWHERE = cWHERE & " and a.Moneda = 'D' "
            Case Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                ' ERROR - Debe seleccionar por lo menos un tipo de moneda
                DevuelveQuery = ""
                Exit Function
            Case Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                ' Pesos
                cWHERE = cWHERE & " and a.Moneda = 'P' "
            Case Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Checked And Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Checked
                ' Pesos - Euros
                cWHERE = cWHERE & " and a.Moneda <> 'D' "
            Case Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Checked
                ' Euros
                cWHERE = cWHERE & " and a.Moneda = 'E' "
        End Select

        'Por Intervalo de Fechas
        If Me._optOC_1.Checked Then
            cWHERE = cWHERE & " and ( a.FechaCompraEI >= '" & VB6.Format(Me.dtpDesde.Value, "mm/dd/yyyy") & "' and a.FechaCompraEI <= '" & VB6.Format(Me.dtpHasta.Value, "mm/dd/yyyy") & "')"
        Else
            cWHERE = cWHERE & " and a.FechaOrdenCompra between '" & VB6.Format(Me.dtpDesde.Value, "mm/dd/yyyy") & "' and '" & VB6.Format(Me.dtpHasta.Value, "mm/dd/yyyy") & "'"
        End If

        'Convertir los totales a la moneda indicada
        If Me._optMoneda_0.Checked Then
            cMoneda = C_DOLAR
            cMonedaDeCantidades = "** Los importes están expresados en Dólares (USD)"
        End If

        If Me._optMoneda_1.Checked Then
            cMoneda = C_PESO
            cMonedaDeCantidades = "** Los importes están expresados en Pesos"
        End If

        If Me._optMoneda_2.Checked Then
            cMoneda = C_EURO
            cMonedaDeCantidades = "** Los importes están expresados en Euros"
        End If

        cSELECT = " select a.codProvAcreed, LTrim(RTrim( b.DescProvAcreed )) as NomProv, a.FechaOrdenCompra, a.FechaCompraEI, a.FolioOrdenCompra, " & "       Round(dbo.ConvertirCantidad( a.Moneda, '" & cMoneda & "', a.SubTotal, a.TipoCambioC, a.TipoCambioEuroC ),2) as SubTotal, " & "       Round(dbo.ConvertirCantidad( a.Moneda, '" & cMoneda & "', a.Descuento, a.TipoCambioC, a.TipoCambioEuroC ),2) as Descuento, " & "       Round(dbo.ConvertirCantidad( a.Moneda, '" & cMoneda & "', a.Iva, a.TipoCambioC, a.TipoCambioEuroC ),2) as Iva, " & "       Round(dbo.ConvertirCantidad( a.Moneda, '" & cMoneda & "', a.Total,  a.TipoCambioC, a.TipoCambioEuroC ),2) as Total, " & "       dbo.EstatusRpt(a.Estatus) as EstausOC "
        cFROM = " from    OrdenesCompra a INNER JOIN CatProvAcreed b " & "                               ON a.CodProvAcreed = b.CodProvAcreed "

        cGROUPBY = ""

        cORDERBY = " Order by a.codProvAcreed, a.FolioOrdenCompra"

        DevuelveQuery = cSELECT & cFROM & cWHERE & cGROUPBY & cORDERBY

    End Function

    Public Sub Imprime()
        Dim rptCXPrptOC As New rptCXPrptOC
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        'On Error GoTo Merr

        Dim lStrSql As String
        'Declarar vectores para almacenar los parámetros que se le enviarán al reporte
        Dim aParam(6) As Object
        Dim aValues(6) As Object

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
            rptCXPrptOC.SetDataSource(frmReportes.rsReport)
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
        'aParam(6) = "TituloOC"
        'aValues(6) = Trim(cTituloOC)

        If (txtMensaje.Text <> Nothing Or txtMensaje.Text <> "") Then
            pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
            rptCXPrptOC.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptCXPrptOC.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        End If

        If (dtpDesde.Value <> Nothing) Then
            pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
            rptCXPrptOC.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
        End If

        If (dtpHasta.Value <> Nothing) Then
            pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
            rptCXPrptOC.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
        End If

        If (cMonedaDeCantidades <> Nothing Or cMonedaDeCantidades <> "") Then
            pdvNum.Value = cMonedaDeCantidades : pvNum.Add(pdvNum)
            rptCXPrptOC.DataDefinition.ParameterFields("MonedaDeCantidades").ApplyCurrentValues(pvNum)
        End If

        If (gstrNombCortoEmpresa <> Nothing Or gstrNombCortoEmpresa <> "") Then
            pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
            rptCXPrptOC.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
        End If

        If (cTituloOC <> Nothing Or cTituloOC <> "") Then
            pdvNum.Value = cTituloOC : pvNum.Add(pdvNum)
            rptCXPrptOC.DataDefinition.ParameterFields("TituloOC").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptCXPrptOC.DataDefinition.ParameterFields("TituloOC").ApplyCurrentValues(pvNum)
        End If

        'frmReportes.Report = rptCXPrptOC 'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Imprime(Trim(Me.Text), aParam, aValues)
        frmReportes.reporteActual = rptCXPrptOC
        frmReportes.Show()

Merr:
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
            Case Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                MsgBox("Debe seleccionar por lo menos un tipo de moneda", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me._chkMoneda_0.Focus()
                Exit Function
            Case Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Unchecked And Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                MsgBox("Debe seleccionar por lo menos un grupo de órdenes de compra", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me._chkOrdenes_0.Focus()
                Exit Function
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Private Sub _chkMoneda_0_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _chkMoneda_0.Enter
        Dim Index As Integer
        '= _chkMoneda_0.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub _chkMoneda_1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _chkMoneda_1.Enter
        Dim Index As Integer
        '= _chkMoneda_1.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub _chkMoneda_2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _chkMoneda_2.Enter
        Dim Index As Integer
        '= _chkMoneda_2.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub chkOrdenes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkOrdenes.Enter
        Dim Index As Integer
        '= chkOrdenes.GetIndex(eventSender)
        Pon_Tool()
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

    Private Sub frmCXPrptOC_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCXPrptOC_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPrptOC_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "OPTOC" Then
                    mblnSALIR = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCXPrptOC_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPrptOC_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmCXPrptOC_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If mblnSALIR Then
        '    mblnSALIR = False
        '    Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Me.dtpDesde.Focus()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCXPrptOC_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub _optMoneda_0_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optMoneda_0.Enter
        Dim Index As Integer
        '= _optMoneda_0.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub _optMoneda_1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optMoneda_1.Enter
        Dim Index As Integer
        '= _optMoneda_1.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub _optMoneda_2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optMoneda_2.Enter
        Dim Index As Integer
        '= _optMoneda_2.GetIndex(eventSender)
        Pon_Tool()
    End Sub


    Private Sub _optOC_0_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optOC_0.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer
            '= _optOC_0.GetIndex(eventSender)
            If Index = 1 Then
                Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Checked
                Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me._chkOrdenes_0.Enabled = False
                Me._chkOrdenes_1.Enabled = False
                Me._chkOrdenes_2.Enabled = False
            Else
                Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Checked
                Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Checked
                Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Checked
                Me._chkOrdenes_0.Enabled = True
                Me._chkOrdenes_1.Enabled = True
                Me._chkOrdenes_2.Enabled = True
            End If
        End If
    End Sub


    Private Sub _optOC_1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optOC_1.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer
            '= _optOC_1.GetIndex(eventSender)
            If Index = 1 Then
                Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Checked
                Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me._chkOrdenes_0.Enabled = False
                Me._chkOrdenes_1.Enabled = False
                Me._chkOrdenes_2.Enabled = False
            Else
                Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Checked
                Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Checked
                Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Checked
                Me._chkOrdenes_0.Enabled = True
                Me._chkOrdenes_1.Enabled = True
                Me._chkOrdenes_2.Enabled = True
            End If
        End If
    End Sub

    Private Sub _optOC_0_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optOC_0.Enter
        Dim Index As Integer
        '= _optOC_0.GetIndex(eventSender)
        Pon_Tool()
    End Sub


    Private Sub _optOC_1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _optOC_1.Enter
        Dim Index As Integer
        '= _optOC_1.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._chkOrdenes_0 = New System.Windows.Forms.CheckBox()
        Me._chkOrdenes_1 = New System.Windows.Forms.CheckBox()
        Me._chkOrdenes_2 = New System.Windows.Forms.CheckBox()
        Me._optOC_1 = New System.Windows.Forms.RadioButton()
        Me._optOC_0 = New System.Windows.Forms.RadioButton()
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_2 = New System.Windows.Forms.RadioButton()
        Me._chkMoneda_2 = New System.Windows.Forms.CheckBox()
        Me._chkMoneda_1 = New System.Windows.Forms.CheckBox()
        Me._chkMoneda_0 = New System.Windows.Forms.CheckBox()
        Me._fraRpt_5 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_4 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_2 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_1 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_0 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblRpt_1 = New System.Windows.Forms.Label()
        Me._lblRpt_0 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.chkMoneda = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.chkOrdenes = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.fraRpt = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optOC = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me._fraRpt_5.SuspendLayout()
        Me._fraRpt_4.SuspendLayout()
        Me._fraRpt_2.SuspendLayout()
        Me._fraRpt_1.SuspendLayout()
        Me._fraRpt_0.SuspendLayout()
        CType(Me.chkMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkOrdenes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optOC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_chkOrdenes_0
        '
        Me._chkOrdenes_0.BackColor = System.Drawing.SystemColors.Control
        Me._chkOrdenes_0.Checked = True
        Me._chkOrdenes_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkOrdenes_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkOrdenes_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkOrdenes.SetIndex(Me._chkOrdenes_0, CType(0, Short))
        Me._chkOrdenes_0.Location = New System.Drawing.Point(16, 24)
        Me._chkOrdenes_0.Name = "_chkOrdenes_0"
        Me._chkOrdenes_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkOrdenes_0.Size = New System.Drawing.Size(81, 17)
        Me._chkOrdenes_0.TabIndex = 16
        Me._chkOrdenes_0.Text = "Por recibir"
        Me.ToolTip1.SetToolTip(Me._chkOrdenes_0, "Selecciona todas las órdenes de compra que no se han conciliado")
        Me._chkOrdenes_0.UseVisualStyleBackColor = False
        '
        '_chkOrdenes_1
        '
        Me._chkOrdenes_1.BackColor = System.Drawing.SystemColors.Control
        Me._chkOrdenes_1.Checked = True
        Me._chkOrdenes_1.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkOrdenes_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkOrdenes_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkOrdenes.SetIndex(Me._chkOrdenes_1, CType(1, Short))
        Me._chkOrdenes_1.Location = New System.Drawing.Point(16, 48)
        Me._chkOrdenes_1.Name = "_chkOrdenes_1"
        Me._chkOrdenes_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkOrdenes_1.Size = New System.Drawing.Size(81, 17)
        Me._chkOrdenes_1.TabIndex = 17
        Me._chkOrdenes_1.Text = "Recibidas"
        Me.ToolTip1.SetToolTip(Me._chkOrdenes_1, "Selecciona las órdenes de compra que ya se recibieron")
        Me._chkOrdenes_1.UseVisualStyleBackColor = False
        '
        '_chkOrdenes_2
        '
        Me._chkOrdenes_2.BackColor = System.Drawing.SystemColors.Control
        Me._chkOrdenes_2.Checked = True
        Me._chkOrdenes_2.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkOrdenes_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkOrdenes_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkOrdenes.SetIndex(Me._chkOrdenes_2, CType(2, Short))
        Me._chkOrdenes_2.Location = New System.Drawing.Point(16, 72)
        Me._chkOrdenes_2.Name = "_chkOrdenes_2"
        Me._chkOrdenes_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkOrdenes_2.Size = New System.Drawing.Size(81, 17)
        Me._chkOrdenes_2.TabIndex = 18
        Me._chkOrdenes_2.Text = "Canceladas"
        Me.ToolTip1.SetToolTip(Me._chkOrdenes_2, "Selecciona las órdenes de compra que se han cancelado")
        Me._chkOrdenes_2.UseVisualStyleBackColor = False
        '
        '_optOC_1
        '
        Me._optOC_1.BackColor = System.Drawing.SystemColors.Control
        Me._optOC_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optOC_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optOC.SetIndex(Me._optOC_1, CType(1, Short))
        Me._optOC_1.Location = New System.Drawing.Point(200, 17)
        Me._optOC_1.Name = "_optOC_1"
        Me._optOC_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optOC_1.Size = New System.Drawing.Size(93, 19)
        Me._optOC_1.TabIndex = 2
        Me._optOC_1.TabStop = True
        Me._optOC_1.Text = "Recibidas"
        Me.ToolTip1.SetToolTip(Me._optOC_1, "Órdenes Recibidas")
        Me._optOC_1.UseVisualStyleBackColor = False
        '
        '_optOC_0
        '
        Me._optOC_0.BackColor = System.Drawing.SystemColors.Control
        Me._optOC_0.Checked = True
        Me._optOC_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optOC_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optOC.SetIndex(Me._optOC_0, CType(0, Short))
        Me._optOC_0.Location = New System.Drawing.Point(72, 17)
        Me._optOC_0.Name = "_optOC_0"
        Me._optOC_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optOC_0.Size = New System.Drawing.Size(85, 19)
        Me._optOC_0.TabIndex = 1
        Me._optOC_0.TabStop = True
        Me._optOC_0.Text = "Emitidas"
        Me.ToolTip1.SetToolTip(Me._optOC_0, "Órdenes Generadas")
        Me._optOC_0.UseVisualStyleBackColor = False
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(8, 240)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(345, 73)
        Me.txtMensaje.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        '_optMoneda_0
        '
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_0, CType(0, Short))
        Me._optMoneda_0.Location = New System.Drawing.Point(16, 24)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(64, 17)
        Me._optMoneda_0.TabIndex = 13
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
        Me._optMoneda_1.Location = New System.Drawing.Point(16, 48)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(64, 17)
        Me._optMoneda_1.TabIndex = 14
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
        Me._optMoneda_2.Location = New System.Drawing.Point(16, 72)
        Me._optMoneda_2.Name = "_optMoneda_2"
        Me._optMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_2.Size = New System.Drawing.Size(64, 17)
        Me._optMoneda_2.TabIndex = 15
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
        Me._chkMoneda_2.Location = New System.Drawing.Point(16, 72)
        Me._chkMoneda_2.Name = "_chkMoneda_2"
        Me._chkMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_2.Size = New System.Drawing.Size(65, 17)
        Me._chkMoneda_2.TabIndex = 11
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
        Me._chkMoneda_1.Location = New System.Drawing.Point(16, 48)
        Me._chkMoneda_1.Name = "_chkMoneda_1"
        Me._chkMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_1.Size = New System.Drawing.Size(65, 17)
        Me._chkMoneda_1.TabIndex = 10
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
        Me._chkMoneda_0.Location = New System.Drawing.Point(16, 24)
        Me._chkMoneda_0.Name = "_chkMoneda_0"
        Me._chkMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_0.Size = New System.Drawing.Size(65, 17)
        Me._chkMoneda_0.TabIndex = 9
        Me._chkMoneda_0.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me._chkMoneda_0, "Selecciona todas las compras en Dólares")
        Me._chkMoneda_0.UseVisualStyleBackColor = False
        '
        '_fraRpt_5
        '
        Me._fraRpt_5.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_5.Controls.Add(Me._chkOrdenes_0)
        Me._fraRpt_5.Controls.Add(Me._chkOrdenes_1)
        Me._fraRpt_5.Controls.Add(Me._chkOrdenes_2)
        Me._fraRpt_5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_5, CType(5, Short))
        Me._fraRpt_5.Location = New System.Drawing.Point(248, 120)
        Me._fraRpt_5.Name = "_fraRpt_5"
        Me._fraRpt_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_5.Size = New System.Drawing.Size(105, 97)
        Me._fraRpt_5.TabIndex = 19
        Me._fraRpt_5.TabStop = False
        Me._fraRpt_5.Text = "Incluir órdenes"
        '
        '_fraRpt_4
        '
        Me._fraRpt_4.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_4.Controls.Add(Me._optOC_1)
        Me._fraRpt_4.Controls.Add(Me._optOC_0)
        Me._fraRpt_4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_4, CType(4, Short))
        Me._fraRpt_4.Location = New System.Drawing.Point(8, 8)
        Me._fraRpt_4.Name = "_fraRpt_4"
        Me._fraRpt_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_4.Size = New System.Drawing.Size(345, 42)
        Me._fraRpt_4.TabIndex = 0
        Me._fraRpt_4.TabStop = False
        Me._fraRpt_4.Text = "Imprimir órdenes ..."
        '
        '_fraRpt_2
        '
        Me._fraRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_2.Controls.Add(Me._optMoneda_0)
        Me._fraRpt_2.Controls.Add(Me._optMoneda_1)
        Me._fraRpt_2.Controls.Add(Me._optMoneda_2)
        Me._fraRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_2, CType(2, Short))
        Me._fraRpt_2.Location = New System.Drawing.Point(128, 120)
        Me._fraRpt_2.Name = "_fraRpt_2"
        Me._fraRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_2.Size = New System.Drawing.Size(113, 97)
        Me._fraRpt_2.TabIndex = 12
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
        Me._fraRpt_1.Location = New System.Drawing.Point(8, 120)
        Me._fraRpt_1.Name = "_fraRpt_1"
        Me._fraRpt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_1.Size = New System.Drawing.Size(113, 97)
        Me._fraRpt_1.TabIndex = 8
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
        Me._fraRpt_0.Location = New System.Drawing.Point(8, 56)
        Me._fraRpt_0.Name = "_fraRpt_0"
        Me._fraRpt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_0.Size = New System.Drawing.Size(345, 57)
        Me._fraRpt_0.TabIndex = 3
        Me._fraRpt_0.TabStop = False
        Me._fraRpt_0.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(72, 21)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(97, 20)
        Me.dtpDesde.TabIndex = 5
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(232, 21)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(97, 20)
        Me.dtpHasta.TabIndex = 7
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
        Me._lblRpt_1.TabIndex = 6
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
        Me._lblRpt_0.TabIndex = 4
        Me._lblRpt_0.Text = "Desde el"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblRpt.SetIndex(Me._lblRpt_2, CType(2, Short))
        Me._lblRpt_2.Location = New System.Drawing.Point(8, 226)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 20
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'chkMoneda
        '
        '
        'chkOrdenes
        '
        '
        'optMoneda
        '
        '
        'optOC
        '
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(123, 329)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 124
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(8, 329)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 123
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(238, 330)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 122
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmCXPrptOC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(361, 379)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me._fraRpt_5)
        Me.Controls.Add(Me._fraRpt_4)
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
        Me.Name = "frmCXPrptOC"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Órdenes de compra"
        Me._fraRpt_5.ResumeLayout(False)
        Me._fraRpt_4.ResumeLayout(False)
        Me._fraRpt_2.ResumeLayout(False)
        Me._fraRpt_1.ResumeLayout(False)
        Me._fraRpt_0.ResumeLayout(False)
        Me._fraRpt_0.PerformLayout()
        CType(Me.chkMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkOrdenes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optOC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub
End Class