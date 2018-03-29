Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCXPrptFacturas
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents _chkTipoFac_0 As System.Windows.Forms.CheckBox
    Public WithEvents _chkTipoFac_1 As System.Windows.Forms.CheckBox
    Public WithEvents _fraRpt_1 As System.Windows.Forms.GroupBox
    Public WithEvents chkTodos As System.Windows.Forms.CheckBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents _fraRpt_0 As System.Windows.Forms.GroupBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblRpt_1 As System.Windows.Forms.Label
    Public WithEvents _lblRpt_0 As System.Windows.Forms.Label
    Public WithEvents _fraRpt_5 As System.Windows.Forms.GroupBox
    Public WithEvents chkSaldadas As System.Windows.Forms.CheckBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents chkTipoFac As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
    Public WithEvents fraRpt As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Const C_DOLARES As Integer = 0
    Const C_PESOS As Integer = 1
    Const C_EUROS As Integer = 2

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean

    Dim cMonedaDeCantidades As String 'Moneda en la que estarán expresadas las cantidades en el reporte

    Dim tecla As Integer
    Dim mblnFueraChange As Boolean
    Dim mintCodProveedor As Integer

    Dim mTitulo As String
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim mblnSALIR As Boolean

    Public Sub Limpiar()
        On Error Resume Next
        Call Me.Nuevo()
        Me.chkTipoFac(0).Focus()
    End Sub

    Public Sub Nuevo()
        Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Checked
        Call chkTipoFac_CheckStateChanged(chkTipoFac.Item(0), New System.EventArgs())
        Me.dtpDesde.Value = Format(Today, "dd/MMM/yyyy")
        Me.dtpHasta.Value = Format(Today, "dd/MMM/yyyy")
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

        'Obtener el WHERE
        cWHERE = " where a.Estatus <> '" & C_STCANCELADA & "' "
        'Por Tipos: Proveedores y/o Acreedores
        mTitulo = ""
        If mintCodProveedor <> 0 Then
            cWHERE = cWHERE & " and a.codProvAcreed = " & mintCodProveedor
        Else
            'Validar si son proveedores o acreedores, o ambos
            Select Case True
                Case Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Checked
                    'Proveedores - Acreedores
                    'Todos, no le importa
                    mTitulo = " POR PROVEEDORES Y ACREEDORES "
                Case Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Unchecked
                    'Proveedores
                    cWHERE = cWHERE & " and a.TipoFacturaCxP = '" & C_TPROVEEDOR & "'"
                    mTitulo = " POR PROVEEDORES "
                Case Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Checked
                    'Acreedores
                    cWHERE = cWHERE & " and a.TipoFacturaCxP = '" & C_TACREEDOR & "'"
                    mTitulo = " POR ACREEDORES "
                Case Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Unchecked
                    'Error, debe seleccionar alguno, por lo menos
                    DevuelveQuery = ""
                    Exit Function
            End Select
        End If

        'Por Intervalo de Fechas
        cWHERE = cWHERE & " and ( a.FechaFactura between '" & Format(Me.dtpDesde.Value, C_FORMATFECHAGUARDAR) & "' and '" & Format(Me.dtpHasta.Value, C_FORMATFECHAGUARDAR) & "')"

        'Por facturas saldadas
        If Me.chkSaldadas.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = cWHERE & " and (Round(a.Total,2) - dbo.RegresaCERO((select sum(c.TotalPago) from pagos c where c.CodProvAcreed = a.CodProvAcreed and c.FolioFactura = a.FolioFactura and c.Estatus <> 'C'))) > 0 "
        End If

        cSELECT = " select a.FolioFactura, a.Moneda, " & gcurCorpoTIPOCAMBIODOLAR & " as TipoCambio, " & gcurCorpoTIPOCAMBIOEURO & " as TipoCambioEuro, " & " Case a.Moneda " & " when 'D' then 'DOL' " & " when 'P' then 'PES' " & " when 'E' then 'EUR' end as DescMoneda, " & " a.FechaFactura, a.FechaVencto, a.CodProvAcreed, b.DescProvAcreed, " & " a.SubTotal, a.Descuento, a.Iva, Round(a.Total,2) as Total, " & " dbo.RegresaCERO((select sum(c.TotalPago) from pagos c where c.CodProvAcreed = a.CodProvAcreed and c.FolioFactura = a.FolioFactura and c.Estatus <> 'C')) as Abonos, " & " Round(a.Total,2) - dbo.RegresaCERO((select sum(c.TotalPago) from pagos c where c.CodProvAcreed = a.CodProvAcreed and c.FolioFactura = a.FolioFactura and c.Estatus <> 'C')) as Saldo "
        cFROM = " from cxpFacturas a " & " inner join catprovacreed b on b.codprovacreed = a.codprovacreed "
        cGROUPBY = ""
        cORDERBY = " order by a.CodProvAcreed, a.Moneda, a.FechaFactura, a.FolioFactura "

        DevuelveQuery = cSELECT & cFROM & cWHERE & cGROUPBY & cORDERBY

    End Function

    Public Sub Imprime()

        Dim rptCxPrptFacturas As New rptCxPrptFacturas
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        'On Error GoTo Merr

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
            rptCxPrptFacturas.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "Mensaje"
        'aValues(1) = Trim(Me.txtMensaje.Text)
        'aParam(2) = "dDesde"
        'aValues(2) = Me.dtpDesde.Value
        'aParam(3) = "dHasta"
        'aValues(3) = Me.dtpHasta.Value
        'aParam(4) = "Empresa"
        'aValues(4) = Trim(gstrNombCortoEmpresa)
        'aParam(5) = "Titulo"
        'aValues(5) = Trim(mTitulo)

        If (txtMensaje.Text <> Nothing Or txtMensaje.Text <> "") Then
            pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
            rptCxPrptFacturas.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptCxPrptFacturas.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        End If

        If (dtpDesde.Value <> Nothing) Then
            pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
            rptCxPrptFacturas.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
        End If

        If (dtpHasta.Value <> Nothing) Then
            pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
            rptCxPrptFacturas.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
        End If

        If (gstrNombCortoEmpresa <> Nothing Or gstrNombCortoEmpresa <> "") Then
            pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
            rptCxPrptFacturas.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
        End If

        If (mTitulo <> Nothing Or mTitulo <> "") Then
            pdvNum.Value = mTitulo : pvNum.Add(pdvNum)
            rptCxPrptFacturas.DataDefinition.ParameterFields("Titulo").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptCxPrptFacturas.DataDefinition.ParameterFields("Titulo").ApplyCurrentValues(pvNum)
        End If

        'frmReportes.Report = rptCxPrptFacturas 'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Imprime(Trim(Me.Text), aParam, aValues)
        frmReportes.reporteActual = rptCxPrptFacturas
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
            Case Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Unchecked
                MsgBox("Debe seleccionar, por lo menos, un tipo de factura", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.chkTipoFac(0).Focus()
                Exit Function
            Case Me.dbcProveedor.Enabled And mintCodProveedor = 0
                MsgBox("Debe elegir un proveedor, o habilitar la casilla de verificación para seleccionarlos todos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.chkTodos.Focus()
                Exit Function
            Case Me.dtpDesde.Value > Me.dtpHasta.Value
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Límite", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dtpDesde.Focus()
                Exit Function
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Private Sub chkTipoFac_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTipoFac.CheckStateChanged
        Dim Index As Integer = chkTipoFac.GetIndex(eventSender)
        Select Case True
            Case Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.Enabled = False

            Case Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Unchecked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.Enabled = True

            Case Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.Enabled = True

            Case Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Unchecked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me.chkTodos.Enabled = False

        End Select
        mblnFueraChange = True
        mintCodProveedor = 0
        Me.dbcProveedor.Text = "[ Todos ... ]"
        Me.dbcProveedor.Tag = ""
        Me.dbcProveedor.Enabled = False
        mblnFueraChange = False
    End Sub

    Private Sub chkTipoFac_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTipoFac.Enter
        Dim Index As Integer = chkTipoFac.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub chkTodos_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodos.CheckStateChanged
        If Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
            mblnFueraChange = True
            mintCodProveedor = 0
            Me.dbcProveedor.Text = "[ Todos ... ]"
            Me.dbcProveedor.Tag = ""
            mblnFueraChange = False
            Me.dbcProveedor.Enabled = False
        Else
            mblnFueraChange = True
            mintCodProveedor = 0
            Me.dbcProveedor.Text = ""
            Me.dbcProveedor.Tag = ""
            mblnFueraChange = False
            Me.dbcProveedor.Enabled = True
        End If
    End Sub

    Private Sub chkTodos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodos.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        If Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Checked Then
            lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ElseIf Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Checked Then
            lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TACREEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        End If

        ModDCombo.DCChange(lStrSql, tecla, dbcProveedor)

        If Trim(Me.dbcProveedor.Text) = "" Then
            dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        Pon_Tool()
        If Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' ORDER BY descProvAcreed"
        ElseIf Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & C_TACREEDOR & "' ORDER BY descProvAcreed"
        End If
        ModDCombo.DCGotFocus(gStrSql, dbcProveedor)
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcProveedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkTodos.Focus()
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        If Me.chkTipoFac(0).CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ElseIf Me.chkTipoFac(1).CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TACREEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        End If
        Aux = mintCodProveedor
        mintCodProveedor = 0
        ModDCombo.DCLostFocus(dbcProveedor, gStrSql, mintCodProveedor)
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

    Private Sub frmCXPrptFacturas_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCXPrptFacturas_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPrptFacturas_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "CHKTIPOFAC" Then
                    If Me.ActiveControl.TabIndex = 0 Then
                        mblnSALIR = True
                        Me.Close()
                    Else
                        ModEstandar.RetrocederTab(Me)
                    End If
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCXPrptFacturas_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPrptFacturas_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmCXPrptFacturas_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmCXPrptFacturas_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optMoneda_GotFocus(ByRef Index As Integer)
        Pon_Tool()
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub



    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._chkTipoFac_0 = New System.Windows.Forms.CheckBox()
        Me._chkTipoFac_1 = New System.Windows.Forms.CheckBox()
        Me.chkTodos = New System.Windows.Forms.CheckBox()
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me._fraRpt_1 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_0 = New System.Windows.Forms.GroupBox()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me._fraRpt_5 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblRpt_1 = New System.Windows.Forms.Label()
        Me._lblRpt_0 = New System.Windows.Forms.Label()
        Me.chkSaldadas = New System.Windows.Forms.CheckBox()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.chkTipoFac = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.fraRpt = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me._fraRpt_1.SuspendLayout()
        Me._fraRpt_0.SuspendLayout()
        Me._fraRpt_5.SuspendLayout()
        CType(Me.chkTipoFac, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_chkTipoFac_0
        '
        Me._chkTipoFac_0.BackColor = System.Drawing.SystemColors.Control
        Me._chkTipoFac_0.Checked = True
        Me._chkTipoFac_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkTipoFac_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkTipoFac_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTipoFac.SetIndex(Me._chkTipoFac_0, CType(0, Short))
        Me._chkTipoFac_0.Location = New System.Drawing.Point(16, 24)
        Me._chkTipoFac_0.Name = "_chkTipoFac_0"
        Me._chkTipoFac_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkTipoFac_0.Size = New System.Drawing.Size(163, 27)
        Me._chkTipoFac_0.TabIndex = 1
        Me._chkTipoFac_0.Text = "de Proveedores (compras)"
        Me.ToolTip1.SetToolTip(Me._chkTipoFac_0, "Facturas de Compras")
        Me._chkTipoFac_0.UseVisualStyleBackColor = False
        '
        '_chkTipoFac_1
        '
        Me._chkTipoFac_1.BackColor = System.Drawing.SystemColors.Control
        Me._chkTipoFac_1.Checked = True
        Me._chkTipoFac_1.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkTipoFac_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkTipoFac_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTipoFac.SetIndex(Me._chkTipoFac_1, CType(1, Short))
        Me._chkTipoFac_1.Location = New System.Drawing.Point(200, 24)
        Me._chkTipoFac_1.Name = "_chkTipoFac_1"
        Me._chkTipoFac_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkTipoFac_1.Size = New System.Drawing.Size(147, 27)
        Me._chkTipoFac_1.TabIndex = 2
        Me._chkTipoFac_1.Text = "de Acreedores (gastos)"
        Me.ToolTip1.SetToolTip(Me._chkTipoFac_1, "Facturas de Gastos")
        Me._chkTipoFac_1.UseVisualStyleBackColor = False
        '
        'chkTodos
        '
        Me.chkTodos.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodos.Checked = True
        Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodos.Location = New System.Drawing.Point(16, 28)
        Me.chkTodos.Name = "chkTodos"
        Me.chkTodos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodos.Size = New System.Drawing.Size(58, 17)
        Me.chkTodos.TabIndex = 4
        Me.chkTodos.Text = "Todos"
        Me.ToolTip1.SetToolTip(Me.chkTodos, "Selecciona todos los proveedores")
        Me.chkTodos.UseVisualStyleBackColor = False
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(8, 253)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(353, 66)
        Me.txtMensaje.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        '_fraRpt_1
        '
        Me._fraRpt_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_1.Controls.Add(Me._chkTipoFac_0)
        Me._fraRpt_1.Controls.Add(Me._chkTipoFac_1)
        Me._fraRpt_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_1, CType(1, Short))
        Me._fraRpt_1.Location = New System.Drawing.Point(8, 8)
        Me._fraRpt_1.Name = "_fraRpt_1"
        Me._fraRpt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_1.Size = New System.Drawing.Size(353, 57)
        Me._fraRpt_1.TabIndex = 0
        Me._fraRpt_1.TabStop = False
        Me._fraRpt_1.Text = "Facturas ..."
        '
        '_fraRpt_0
        '
        Me._fraRpt_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_0.Controls.Add(Me.chkTodos)
        Me._fraRpt_0.Controls.Add(Me.dbcProveedor)
        Me._fraRpt_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_0, CType(0, Short))
        Me._fraRpt_0.Location = New System.Drawing.Point(8, 72)
        Me._fraRpt_0.Name = "_fraRpt_0"
        Me._fraRpt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_0.Size = New System.Drawing.Size(353, 65)
        Me._fraRpt_0.TabIndex = 3
        Me._fraRpt_0.TabStop = False
        Me._fraRpt_0.Text = "Proveedor"
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(80, 24)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(257, 21)
        Me.dbcProveedor.TabIndex = 5
        '
        '_fraRpt_5
        '
        Me._fraRpt_5.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_5.Controls.Add(Me.dtpDesde)
        Me._fraRpt_5.Controls.Add(Me.dtpHasta)
        Me._fraRpt_5.Controls.Add(Me._lblRpt_1)
        Me._fraRpt_5.Controls.Add(Me._lblRpt_0)
        Me._fraRpt_5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_5, CType(5, Short))
        Me._fraRpt_5.Location = New System.Drawing.Point(8, 144)
        Me._fraRpt_5.Name = "_fraRpt_5"
        Me._fraRpt_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_5.Size = New System.Drawing.Size(353, 57)
        Me._fraRpt_5.TabIndex = 6
        Me._fraRpt_5.TabStop = False
        Me._fraRpt_5.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(64, 21)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(96, 20)
        Me.dtpDesde.TabIndex = 8
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(232, 21)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(105, 20)
        Me.dtpHasta.TabIndex = 10
        '
        '_lblRpt_1
        '
        Me._lblRpt_1.AutoSize = True
        Me._lblRpt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRpt.SetIndex(Me._lblRpt_1, CType(1, Short))
        Me._lblRpt_1.Location = New System.Drawing.Point(192, 25)
        Me._lblRpt_1.Name = "_lblRpt_1"
        Me._lblRpt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_1.Size = New System.Drawing.Size(33, 13)
        Me._lblRpt_1.TabIndex = 9
        Me._lblRpt_1.Text = "hasta"
        '
        '_lblRpt_0
        '
        Me._lblRpt_0.AutoSize = True
        Me._lblRpt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRpt.SetIndex(Me._lblRpt_0, CType(0, Short))
        Me._lblRpt_0.Location = New System.Drawing.Point(24, 25)
        Me._lblRpt_0.Name = "_lblRpt_0"
        Me._lblRpt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_0.Size = New System.Drawing.Size(38, 13)
        Me._lblRpt_0.TabIndex = 7
        Me._lblRpt_0.Text = "Desde"
        '
        'chkSaldadas
        '
        Me.chkSaldadas.BackColor = System.Drawing.SystemColors.Control
        Me.chkSaldadas.Checked = True
        Me.chkSaldadas.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSaldadas.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSaldadas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSaldadas.Location = New System.Drawing.Point(24, 208)
        Me.chkSaldadas.Name = "chkSaldadas"
        Me.chkSaldadas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSaldadas.Size = New System.Drawing.Size(144, 26)
        Me.chkSaldadas.TabIndex = 11
        Me.chkSaldadas.Text = "Incluir facturas saldadas"
        Me.chkSaldadas.UseVisualStyleBackColor = False
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblRpt.SetIndex(Me._lblRpt_2, CType(2, Short))
        Me._lblRpt_2.Location = New System.Drawing.Point(12, 237)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 12
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'chkTipoFac
        '
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(124, 332)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 112
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(9, 332)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 111
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(239, 332)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 110
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmCXPrptFacturas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(369, 380)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me._fraRpt_1)
        Me.Controls.Add(Me._fraRpt_0)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me._fraRpt_5)
        Me.Controls.Add(Me.chkSaldadas)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 29)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCXPrptFacturas"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Facturas por Proveedor / Acreedor"
        Me._fraRpt_1.ResumeLayout(False)
        Me._fraRpt_0.ResumeLayout(False)
        Me._fraRpt_5.ResumeLayout(False)
        Me._fraRpt_5.PerformLayout()
        CType(Me.chkTipoFac, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class