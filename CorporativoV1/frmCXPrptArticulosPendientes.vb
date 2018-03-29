Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCXPrptArticulosPendientes
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents optRecibidas As System.Windows.Forms.RadioButton
    Public WithEvents optPendientes As System.Windows.Forms.RadioButton
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblRpt_0 As System.Windows.Forms.Label
    Public WithEvents _lblRpt_1 As System.Windows.Forms.Label
    Public WithEvents _fraRpt_5 As System.Windows.Forms.GroupBox
    Public WithEvents _chkMoneda_0 As System.Windows.Forms.CheckBox
    Public WithEvents _chkMoneda_1 As System.Windows.Forms.CheckBox
    Public WithEvents _chkMoneda_2 As System.Windows.Forms.CheckBox
    Public WithEvents _fraRpt_4 As System.Windows.Forms.GroupBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents chkTodos As System.Windows.Forms.CheckBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents _fraRpt_0 As System.Windows.Forms.GroupBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents chkMoneda As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
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
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim mblnSalir As Boolean

    Public Sub Limpiar()
        On Error Resume Next
        Call Me.Nuevo()
        Me.chkTodos.Focus()
    End Sub

    Public Sub Nuevo()
        Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
        mblnFueraChange = True
        mintCodProveedor = 0
        Me.dbcProveedor.Text = "[ Todos ... ]"
        Me.dbcProveedor.Tag = ""
        mblnFueraChange = False
        Me.dbcProveedor.Enabled = False
        Me.dtpDesde.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.dtpHasta.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.chkMoneda(C_DOLARES).CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMoneda(C_PESOS).CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMoneda(C_EUROS).CheckState = System.Windows.Forms.CheckState.Checked
        Me.txtMensaje.Text = ""
        Me.optPendientes.Checked = True
        Me.optRecibidas.Checked = False
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
        If optPendientes.Checked = True Then
            cWHERE = " where a.Estatus = '" & C_STVIGENTE & "' and a.fechacompraei = '01/01/1900' "
        ElseIf optRecibidas.Checked = True Then
            cWHERE = " where (a.estatus = '" & C_STGENERADA & "' or a.estatus = '" & C_STREGISTRADA & "') and a.fechacompraei <> '01/01/1900' "
        End If
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
        cWHERE = cWHERE & " and ( a.FechaOrdenCompra >= '" & VB6.Format(Me.dtpDesde.Value, "mm/dd/yyyy") & "' and a.FechaOrdenCompra <= '" & VB6.Format(Me.dtpHasta.Value, "mm/dd/yyyy") & "')"

        'Por Proveedor, si acaso seleccionó uno en especial
        If mintCodProveedor <> 0 Then
            cWHERE = cWHERE & " and a.CodProvAcreed = " & mintCodProveedor
        End If

        '    cSELECT = " SELECT a.* FROM(SELECT a.FolioOrdenCompra, a.CodProvAcreed, b.DescProvAcreed, " & _
        ''                "       c.CodArticulo , c.DescArticulo, d.DescUnidad, c.CantidadRegistro "
        '    cFROM = "FROM OrdenesCompra a " & _
        ''                    " INNER JOIN OrdenesCompraPreCat c ON a.FolioOrdenCompra = c.FolioOrdenCompra " & _
        ''                    " INNER JOIN CatProvAcreed b ON a.CodProvAcreed = b.CodProvAcreed " & _
        ''                    " INNER JOIN CatUnidades d ON c.CodUnidad = d.CodUnidad "
        '    cGROUPBY = ""
        '
        '    cORDERBY = " ORDER BY a.CodProvAcreed, a.FolioOrdenCompra, a.CodArticulo "
        '
        '    DevuelveQuery = cSELECT & cFROM & cWHERE & _
        ''    ") a Inner Join " & _
        ''    "(SELECT FolioOrdenCompra,SUM(CantidadRegistro) as Cantidad " & _
        ''    "From OrdenesCompraPrecat " & _
        ''    "GROUP BY FolioOrdenCompra " & _
        ''    "HAVING SUM(CantidadRegistro) " & IIf(optPendientes.Value = True, "<>", "=") & " 0 ) b " & _
        ''    "on a.folioordencompra = b.folioordencompra " & cGROUPBY & cORDERBY

        cSELECT = "select a.folioordencompra,a.codprovacreed,b.descprovacreed,c.codarticulo,c.codigoarticuloprov,c.descarticulo," & "D.descunidad , c.CantidadRecepcion, V.Medidas "
        cFROM = "from ordenescompra a inner join ordenescompraprecat c on a.folioordencompra = c.folioordencompra left outer join MovimientosVentasDet V on a.folioApartado = V.FolioVenta " & "inner join catprovacreed b on a.codprovacreed = b.codprovacreed " & "left outer join catunidades d on c.codunidad = d.codunidad"

        DevuelveQuery = cSELECT & cFROM & cWHERE
    End Function

    Public Sub Imprime()

        Dim rptCXPrptArticulosPendientes As New rptCXPrptArticulosPendientes
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        'On Error GoTo Merr
        Try
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
                rptCXPrptArticulosPendientes.SetDataSource(frmReportes.rsReport)
            End If

            'aParam(1) = "Mensaje"
            'aValues(1) = Trim(Me.txtMensaje.Text)
            'aParam(2) = "dDesde"
            'aValues(2) = Me.dtpDesde.Value
            'aParam(3) = "dHasta"
            'aValues(3) = Me.dtpHasta.Value
            'aParam(4) = "Empresa"
            'aValues(4) = Trim(gstrNombCortoEmpresa)
            'aParam(5) = "Mensajes"
            'aValues(5) = IIf(optPendientes.Checked = True, "PENDIENTES", "RECIBIDAS")


            If (txtMensaje.Text <> Nothing Or txtMensaje.Text <> "") Then
                pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
                rptCXPrptArticulosPendientes.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptCXPrptArticulosPendientes.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
            End If

            If (dtpDesde.Value <> Nothing) Then
                pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
                rptCXPrptArticulosPendientes.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
            End If

            If (dtpHasta.Value <> Nothing) Then
                pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
                rptCXPrptArticulosPendientes.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
            End If

            If (gstrNombCortoEmpresa <> Nothing Or gstrNombCortoEmpresa <> "") Then
                pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
                rptCXPrptArticulosPendientes.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
            End If

            If (optPendientes.Checked <> True) Then
                pdvNum.Value = IIf(optPendientes.Checked = True, "PENDIENTES", "RECIBIDAS") : pvNum.Add(pdvNum)
                rptCXPrptArticulosPendientes.DataDefinition.ParameterFields("Mensajes").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptCXPrptArticulosPendientes.DataDefinition.ParameterFields("Mensajes").ApplyCurrentValues(pvNum)
            End If

            'frmReportes.Report = rptCXPrptArticulosPendientes 'Es el nombre del archivo que se incluyó en el proyecto
            'frmReportes.Imprime(Trim(Me.Text), aParam, aValues)
            frmReportes.reporteActual = rptCXPrptArticulosPendientes
            frmReportes.Show()

            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then
                ModEstandar.MostrarError()
            End If
        End Try
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
            Case Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Unchecked And mintCodProveedor = 0
                MsgBox("Debe seleccionar el proveedor, o bien, activar la casilla de verificación" & vbNewLine & "para imprimir todos los proveedores", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dbcProveedor.Focus()
                ModEstandar.SelTxt()
                Exit Function
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

        lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcProveedor)

        'If Trim(Me.dbcProveedor.Text) = "" Then
        '    dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        'End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        Pon_Tool()
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' ORDER BY descProvAcreed"
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
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
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

    Private Sub chkMoneda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkMoneda.Enter
        Dim Index As Integer = chkMoneda.GetIndex(eventSender)
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

    Private Sub frmCXPrptArticulosPendientes_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCXPrptArticulosPendientes_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPrptArticulosPendientes_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "CHKTODOS" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCXPrptArticulosPendientes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPrptArticulosPendientes_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmCXPrptArticulosPendientes_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If mblnSalir Then
        '    mblnSalir = False
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

    Private Sub frmCXPrptArticulosPendientes_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._chkMoneda_0 = New System.Windows.Forms.CheckBox()
        Me._chkMoneda_1 = New System.Windows.Forms.CheckBox()
        Me._chkMoneda_2 = New System.Windows.Forms.CheckBox()
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.chkTodos = New System.Windows.Forms.CheckBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.optRecibidas = New System.Windows.Forms.RadioButton()
        Me.optPendientes = New System.Windows.Forms.RadioButton()
        Me._fraRpt_5 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblRpt_0 = New System.Windows.Forms.Label()
        Me._lblRpt_1 = New System.Windows.Forms.Label()
        Me._fraRpt_4 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_0 = New System.Windows.Forms.GroupBox()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.chkMoneda = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.fraRpt = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me._fraRpt_5.SuspendLayout()
        Me._fraRpt_4.SuspendLayout()
        Me._fraRpt_0.SuspendLayout()
        CType(Me.chkMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_chkMoneda_0
        '
        Me._chkMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._chkMoneda_0.Checked = True
        Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMoneda.SetIndex(Me._chkMoneda_0, CType(0, Short))
        Me._chkMoneda_0.Location = New System.Drawing.Point(40, 24)
        Me._chkMoneda_0.Name = "_chkMoneda_0"
        Me._chkMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_0.Size = New System.Drawing.Size(81, 17)
        Me._chkMoneda_0.TabIndex = 9
        Me._chkMoneda_0.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me._chkMoneda_0, "Selecciona todas las compras en Dólares")
        Me._chkMoneda_0.UseVisualStyleBackColor = False
        '
        '_chkMoneda_1
        '
        Me._chkMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._chkMoneda_1.Checked = True
        Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMoneda.SetIndex(Me._chkMoneda_1, CType(1, Short))
        Me._chkMoneda_1.Location = New System.Drawing.Point(40, 48)
        Me._chkMoneda_1.Name = "_chkMoneda_1"
        Me._chkMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_1.Size = New System.Drawing.Size(81, 17)
        Me._chkMoneda_1.TabIndex = 10
        Me._chkMoneda_1.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me._chkMoneda_1, "Selecciona todas las compras en Pesos")
        Me._chkMoneda_1.UseVisualStyleBackColor = False
        '
        '_chkMoneda_2
        '
        Me._chkMoneda_2.BackColor = System.Drawing.SystemColors.Control
        Me._chkMoneda_2.Checked = True
        Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMoneda.SetIndex(Me._chkMoneda_2, CType(2, Short))
        Me._chkMoneda_2.Location = New System.Drawing.Point(40, 72)
        Me._chkMoneda_2.Name = "_chkMoneda_2"
        Me._chkMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_2.Size = New System.Drawing.Size(81, 17)
        Me._chkMoneda_2.TabIndex = 11
        Me._chkMoneda_2.Text = "Euros"
        Me.ToolTip1.SetToolTip(Me._chkMoneda_2, "Selecciona todas las compras en Euros")
        Me._chkMoneda_2.UseVisualStyleBackColor = False
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(8, 264)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(353, 71)
        Me.txtMensaje.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'chkTodos
        '
        Me.chkTodos.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodos.Checked = True
        Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodos.Location = New System.Drawing.Point(16, 24)
        Me.chkTodos.Name = "chkTodos"
        Me.chkTodos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodos.Size = New System.Drawing.Size(145, 18)
        Me.chkTodos.TabIndex = 1
        Me.chkTodos.Text = "Todos los proveedores"
        Me.ToolTip1.SetToolTip(Me.chkTodos, "Selecciona todos los proveedores")
        Me.chkTodos.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optRecibidas)
        Me.Frame1.Controls.Add(Me.optPendientes)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 200)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(353, 49)
        Me.Frame1.TabIndex = 16
        Me.Frame1.TabStop = False
        '
        'optRecibidas
        '
        Me.optRecibidas.BackColor = System.Drawing.SystemColors.Control
        Me.optRecibidas.Cursor = System.Windows.Forms.Cursors.Default
        Me.optRecibidas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optRecibidas.Location = New System.Drawing.Point(208, 18)
        Me.optRecibidas.Name = "optRecibidas"
        Me.optRecibidas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optRecibidas.Size = New System.Drawing.Size(105, 21)
        Me.optRecibidas.TabIndex = 13
        Me.optRecibidas.TabStop = True
        Me.optRecibidas.Text = "Ya Recibidas"
        Me.optRecibidas.UseVisualStyleBackColor = False
        '
        'optPendientes
        '
        Me.optPendientes.BackColor = System.Drawing.SystemColors.Control
        Me.optPendientes.Checked = True
        Me.optPendientes.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPendientes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPendientes.Location = New System.Drawing.Point(56, 18)
        Me.optPendientes.Name = "optPendientes"
        Me.optPendientes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPendientes.Size = New System.Drawing.Size(105, 21)
        Me.optPendientes.TabIndex = 12
        Me.optPendientes.TabStop = True
        Me.optPendientes.Text = "Solo Pendientes"
        Me.optPendientes.UseVisualStyleBackColor = False
        '
        '_fraRpt_5
        '
        Me._fraRpt_5.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_5.Controls.Add(Me.dtpDesde)
        Me._fraRpt_5.Controls.Add(Me.dtpHasta)
        Me._fraRpt_5.Controls.Add(Me._lblRpt_0)
        Me._fraRpt_5.Controls.Add(Me._lblRpt_1)
        Me._fraRpt_5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_5, CType(5, Short))
        Me._fraRpt_5.Location = New System.Drawing.Point(8, 96)
        Me._fraRpt_5.Name = "_fraRpt_5"
        Me._fraRpt_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_5.Size = New System.Drawing.Size(177, 97)
        Me._fraRpt_5.TabIndex = 3
        Me._fraRpt_5.TabStop = False
        Me._fraRpt_5.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(64, 29)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(105, 20)
        Me.dtpDesde.TabIndex = 5
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(64, 61)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(105, 20)
        Me.dtpHasta.TabIndex = 7
        '
        '_lblRpt_0
        '
        Me._lblRpt_0.AutoSize = True
        Me._lblRpt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRpt.SetIndex(Me._lblRpt_0, CType(0, Short))
        Me._lblRpt_0.Location = New System.Drawing.Point(16, 33)
        Me._lblRpt_0.Name = "_lblRpt_0"
        Me._lblRpt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_0.Size = New System.Drawing.Size(49, 13)
        Me._lblRpt_0.TabIndex = 4
        Me._lblRpt_0.Text = "Desde el"
        '
        '_lblRpt_1
        '
        Me._lblRpt_1.AutoSize = True
        Me._lblRpt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRpt.SetIndex(Me._lblRpt_1, CType(1, Short))
        Me._lblRpt_1.Location = New System.Drawing.Point(16, 65)
        Me._lblRpt_1.Name = "_lblRpt_1"
        Me._lblRpt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_1.Size = New System.Drawing.Size(44, 13)
        Me._lblRpt_1.TabIndex = 6
        Me._lblRpt_1.Text = "hasta el"
        '
        '_fraRpt_4
        '
        Me._fraRpt_4.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_4.Controls.Add(Me._chkMoneda_0)
        Me._fraRpt_4.Controls.Add(Me._chkMoneda_1)
        Me._fraRpt_4.Controls.Add(Me._chkMoneda_2)
        Me._fraRpt_4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_4, CType(4, Short))
        Me._fraRpt_4.Location = New System.Drawing.Point(192, 96)
        Me._fraRpt_4.Name = "_fraRpt_4"
        Me._fraRpt_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_4.Size = New System.Drawing.Size(169, 97)
        Me._fraRpt_4.TabIndex = 8
        Me._fraRpt_4.TabStop = False
        Me._fraRpt_4.Text = "Moneda"
        '
        '_fraRpt_0
        '
        Me._fraRpt_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_0.Controls.Add(Me.chkTodos)
        Me._fraRpt_0.Controls.Add(Me.dbcProveedor)
        Me._fraRpt_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_0, CType(0, Short))
        Me._fraRpt_0.Location = New System.Drawing.Point(8, 8)
        Me._fraRpt_0.Name = "_fraRpt_0"
        Me._fraRpt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_0.Size = New System.Drawing.Size(353, 81)
        Me._fraRpt_0.TabIndex = 0
        Me._fraRpt_0.TabStop = False
        Me._fraRpt_0.Text = "Proveedor"
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(16, 48)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(321, 21)
        Me.dbcProveedor.TabIndex = 2
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblRpt.SetIndex(Me._lblRpt_2, CType(2, Short))
        Me._lblRpt_2.Location = New System.Drawing.Point(11, 250)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 15
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'chkMoneda
        '
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(122, 359)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 127
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(7, 359)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 126
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(237, 360)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 125
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmCXPrptArticulosPendientes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(369, 408)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me._fraRpt_5)
        Me.Controls.Add(Me._fraRpt_4)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me._fraRpt_0)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 29)
        Me.MaximizeBox = False
        Me.Name = "frmCXPrptArticulosPendientes"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Artículos pendientes por recibir"
        Me.Frame1.ResumeLayout(False)
        Me._fraRpt_5.ResumeLayout(False)
        Me._fraRpt_5.PerformLayout()
        Me._fraRpt_4.ResumeLayout(False)
        Me._fraRpt_0.ResumeLayout(False)
        CType(Me.chkMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub
End Class