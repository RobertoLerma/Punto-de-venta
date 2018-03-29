Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCxpReporteSaldoXProveedor
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE SALDOS POR PROVEEDOR                                                              *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :                                                                                                   *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents chkEuros As System.Windows.Forms.CheckBox
    Public WithEvents chkDolares As System.Windows.Forms.CheckBox
    Public WithEvents chkPesos As System.Windows.Forms.CheckBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents chkProveedores As System.Windows.Forms.CheckBox
    Public WithEvents chkAcreedores As System.Windows.Forms.CheckBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents _fraRpt_3 As System.Windows.Forms.GroupBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents fraRpt As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Dim mblnSalir As Boolean
    Dim mblnFueraChange As Boolean
    Dim tecla As Integer
    Dim mintCodProveedor As Integer
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    Dim rsReporte As ADODB.Recordset
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Const C_TODOS As String = "[ Todos ... ]"

    Sub Limpiar()
        Nuevo()
        chkProveedores.Focus()
    End Sub

    Sub Nuevo()
        chkProveedores.CheckState = System.Windows.Forms.CheckState.Checked
        chkAcreedores.CheckState = System.Windows.Forms.CheckState.Checked
        mblnFueraChange = True
        dbcProveedor.Text = C_TODOS
        Me.dbcProveedor.Tag = Me.dbcProveedor.Text
        mblnFueraChange = False
        dtpFechaFinal.Value = Today
        dtpFechaInicial.Value = Today
        chkDolares.CheckState = System.Windows.Forms.CheckState.Checked
        chkEuros.CheckState = System.Windows.Forms.CheckState.Checked
        chkPesos.CheckState = System.Windows.Forms.CheckState.Checked
        '    chkIncluir.Value = vbUnchecked
        txtMensaje.Text = ""
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
        mblnFueraChange = False
        tecla = 0
        mintCodProveedor = 0
    End Sub

    Function Query() As String
        Dim sSELECT As String
        Dim sFROM As String
        Dim sWHERE As String
        Dim sORDER As String
        sSELECT = "SELECT a.* "
        sFROM = "FROM DBO.ObtenerSaldoXProveedor('" & VB6.Format(dtpFechaInicial.Value, "MM/DD/YYYY") & "','" & VB6.Format(dtpFechaFinal.Value, "MM/DD/YYYY") & "',0) a "
        sWHERE = ""
        If mintCodProveedor <> 0 Then
            sWHERE = sWHERE & "WHERE CodProveedor = " & mintCodProveedor & " "
        ElseIf chkProveedores.CheckState = System.Windows.Forms.CheckState.Checked And chkAcreedores.CheckState = System.Windows.Forms.CheckState.Checked Then
            sWHERE = sWHERE & ""
        ElseIf chkProveedores.CheckState = System.Windows.Forms.CheckState.Checked And chkAcreedores.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            sWHERE = sWHERE & "WHERE Tipo = 'P' "
        ElseIf chkProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked And chkAcreedores.CheckState = System.Windows.Forms.CheckState.Checked Then
            sWHERE = sWHERE & "WHERE Tipo = 'A' "
        End If
        If chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked Then 'pesos y dolares
            If Trim(sWHERE) = "" Then
                sWHERE = "WHERE "
                sWHERE = sWHERE & "(Moneda = 'D' OR Moneda = 'P') "
            Else
                sWHERE = sWHERE & "AND (Moneda = 'D' OR Moneda = 'P') "
            End If
        ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then  'dolares y euros
            If Trim(sWHERE) = "" Then
                sWHERE = "WHERE "
                sWHERE = sWHERE & "(Moneda = 'D' OR Moneda = 'E') "
            Else
                sWHERE = sWHERE & "AND (Moneda = 'D' OR Moneda = 'E') "
            End If
        ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked Then  'dolares
            If Trim(sWHERE) = "" Then
                sWHERE = "WHERE "
                sWHERE = sWHERE & "Moneda = 'D' "
            Else
                sWHERE = sWHERE & "AND Moneda = 'D' "
            End If
        ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then  'PESOS Y EUROS
            If Trim(sWHERE) = "" Then
                sWHERE = "WHERE "
                sWHERE = sWHERE & "(Moneda = 'P' OR Moneda = 'E') "
            Else
                sWHERE = sWHERE & "AND (Moneda = 'P' OR Moneda = 'E') "
            End If
        ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked Then  'PESOS
            If Trim(sWHERE) = "" Then
                sWHERE = "WHERE "
                sWHERE = sWHERE & "Moneda = 'P' "
            Else
                sWHERE = sWHERE & "AND Moneda = 'P' "
            End If
        ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then  'EUROS
            If Trim(sWHERE) = "" Then
                sWHERE = "WHERE "
                sWHERE = sWHERE & "Moneda = 'E' "
            Else
                sWHERE = sWHERE & "AND Moneda = 'E' "
            End If
        End If
        sORDER = "ORDER BY CodProveedor,Moneda,Fecha"
        Query = sSELECT & sFROM & sWHERE & sORDER
    End Function

    Sub Imprime()
        Dim rptCXPSaldoXProveedor As New rptCXPSaldoXProveedor
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Err_Renamed
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim Periodo As String
        Dim TextoAdicional As String
        Dim Sql As String

        dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        If Not ValidaDatos() Then Exit Sub
        Sql = Query()
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existen datos para el rango de fechas indicado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            If frmReportes.rsReport.RecordCount = 1 Then
                If IsDBNull(frmReportes.rsReport.Fields("CodProveedor").Value.ToString()) Then
                    MsgBox("No existen datos para el rango de fechas indicado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Sub
                End If
            End If
        End If



        'NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        'NombreReporte = UCase("Auxiliar de Proveedores")
        'Periodo = "Del  " & VB6.Format(dtpFechaInicial.Value, "dd/MMM/yyyy") & "  al  " & VB6.Format(dtpFechaFinal.Value, "dd/MMM/yyyy")
        'TextoAdicional = txtMensaje.Text
        'frmReportes.Report = rptCXPSaldoXProveedor
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        rptCXPSaldoXProveedor.SetDataSource(frmReportes.rsReport)
        'frmReportes.rsReport = rsReporte
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "Periodo", "TextoAdicional"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, Periodo, TextoAdicional}
        frmReportes.Text = "Auxiliar de Proveedores"
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            frmReportes.reporteActual = rptCXPSaldoXProveedor
            frmReportes.Show()
            Me.Cursor = System.Windows.Forms.Cursors.Default

Err_Renamed:
        If Err.Number <> 0 Then
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Sub

    Function ValidaDatos() As Boolean
        On Error GoTo Err_Renamed
        ValidaDatos = False
        Do While (VB.Timer() - sglTiempoCambio) <= 2.1
        Loop
        System.Windows.Forms.Application.DoEvents()
        If dtpFechaInicial.Value > dtpFechaFinal.Value Then
            MsgBox("La Fecha Inicial no Puede ser Mayor que la Fecha Final.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Function
        End If
        If dtpFechaInicial.Value > Now Then
            MsgBox("la Fecha Inicial no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Function
        End If
        If dtpFechaFinal.Value > Now Then
            MsgBox("la Fecha Final no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Function
        End If
        If chkProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked And chkAcreedores.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            MsgBox("Debe especificar si desea ver saldos de proveedores, acreedores o ambos, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            chkProveedores.Focus()
            Exit Function
        End If
        If chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            MsgBox("Debe seleccionar la moneda de los movimientos que desea ver, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Function
        End If
        ValidaDatos = True
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Private Sub chkAcreedores_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAcreedores.CheckStateChanged
        dbcProveedor.Text = C_TODOS
    End Sub

    Private Sub chkProveedores_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkProveedores.CheckStateChanged
        dbcProveedor.Text = C_TODOS
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String
        Dim cWHERE As String

        If mblnFueraChange Then Exit Sub

        cWHERE = ""
        If Me.chkProveedores.CheckState = System.Windows.Forms.CheckState.Checked And Me.chkAcreedores.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = cWHERE & " Tipo = '" & C_TPROVEEDOR & "' and "
        ElseIf Me.chkProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkAcreedores.CheckState = System.Windows.Forms.CheckState.Checked Then
            cWHERE = cWHERE & " Tipo = '" & C_TACREEDOR & "' and "
        End If
        lStrSql = "SELECT CodProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM CatProvAcreed Where " & cWHERE & " DescProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%' Order by DescProvAcreed "
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcProveedor))

        If Trim(Me.dbcProveedor.Text) = "" Then
            mintCodProveedor = 0
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        Dim cWHERE As String
        cWHERE = ""
        If Me.chkProveedores.CheckState = System.Windows.Forms.CheckState.Checked And Me.chkAcreedores.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = cWHERE & " WHERE Tipo = '" & C_TPROVEEDOR & "'"
        ElseIf Me.chkProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkAcreedores.CheckState = System.Windows.Forms.CheckState.Checked Then
            cWHERE = cWHERE & " WHERE Tipo = '" & C_TACREEDOR & "'"
        End If
        Pon_Tool()
        gStrSql = "SELECT CodProvAcreed, LTrim(RTrim(DescProvAcreed)) as DescProvAcreed FROM CatProvAcreed " & cWHERE & " Order by DescProvAcreed "
        ModDCombo.DCGotFocus(gStrSql, dbcProveedor)
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcProveedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkAcreedores.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
        Dim Aux As Integer
        Dim cWHERE As String
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        cWHERE = ""
        If Me.chkProveedores.CheckState = System.Windows.Forms.CheckState.Checked And Me.chkAcreedores.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = cWHERE & " Tipo = '" & C_TPROVEEDOR & "' and "
        ElseIf Me.chkProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkAcreedores.CheckState = System.Windows.Forms.CheckState.Checked Then
            cWHERE = cWHERE & " Tipo = '" & C_TACREEDOR & "' and "
        End If
        gStrSql = "SELECT CodProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM CatProvAcreed Where " & cWHERE & " DescProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%' Order by DescProvAcreed "
        Aux = mintCodProveedor
        mintCodProveedor = 0
        If Trim(Me.dbcProveedor.Text) <> Trim(C_TODOS) Or Trim(Me.dbcProveedor.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcProveedor), gStrSql, mintCodProveedor)
        End If

        If Aux <> mintCodProveedor Then
            If mintCodProveedor = 0 Then
                mblnFueraChange = True
                Me.dbcProveedor.Text = C_TODOS
                mblnFueraChange = False
            End If
        End If
        If Trim(Me.dbcProveedor.Text) = "" Then Me.dbcProveedor.Text = C_TODOS
    End Sub

    Private Sub dbcProveedor_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcProveedor.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcProveedor.Text)
        If Me.dbcProveedor.SelectedItem <> 0 Then
            dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        End If
        Me.dbcProveedor.Text = Aux
    End Sub

    Private Sub dtpFechaFinal_CursorChangede(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.CursorChanged
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

    Private Sub frmCxpReporteSaldoXProveedor_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCxpReporteSaldoXProveedor_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCxpReporteSaldoXProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "chkProveedores" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmCxpReporteSaldoXProveedor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCxpReporteSaldoXProveedor_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        dtpFechaInicial.MinDate = C_FECHAINICIAL
        dtpFechaInicial.MaxDate = C_FECHAFINAL
        dtpFechaFinal.MinDate = C_FECHAINICIAL
        dtpFechaFinal.MaxDate = C_FECHAFINAL
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
        Nuevo()
    End Sub

    Private Sub frmCxpReporteSaldoXProveedor_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
        'Else
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

    Private Sub frmCxpReporteSaldoXProveedor_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.chkEuros = New System.Windows.Forms.CheckBox()
        Me.chkDolares = New System.Windows.Forms.CheckBox()
        Me.chkPesos = New System.Windows.Forms.CheckBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me._Label2_1 = New System.Windows.Forms.Label()
        Me._Label2_0 = New System.Windows.Forms.Label()
        Me._fraRpt_3 = New System.Windows.Forms.GroupBox()
        Me.chkProveedores = New System.Windows.Forms.CheckBox()
        Me.chkAcreedores = New System.Windows.Forms.CheckBox()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.fraRpt = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me._fraRpt_3.SuspendLayout()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(15, 245)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(377, 77)
        Me.txtMensaje.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.chkEuros)
        Me.Frame2.Controls.Add(Me.chkDolares)
        Me.Frame2.Controls.Add(Me.chkPesos)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(15, 160)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(377, 41)
        Me.Frame2.TabIndex = 14
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Movimientos en ...."
        '
        'chkEuros
        '
        Me.chkEuros.BackColor = System.Drawing.SystemColors.Control
        Me.chkEuros.Checked = True
        Me.chkEuros.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkEuros.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEuros.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEuros.Location = New System.Drawing.Point(260, 14)
        Me.chkEuros.Name = "chkEuros"
        Me.chkEuros.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEuros.Size = New System.Drawing.Size(93, 21)
        Me.chkEuros.TabIndex = 7
        Me.chkEuros.Text = "Euros"
        Me.chkEuros.UseVisualStyleBackColor = False
        '
        'chkDolares
        '
        Me.chkDolares.BackColor = System.Drawing.SystemColors.Control
        Me.chkDolares.Checked = True
        Me.chkDolares.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDolares.Location = New System.Drawing.Point(56, 14)
        Me.chkDolares.Name = "chkDolares"
        Me.chkDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDolares.Size = New System.Drawing.Size(93, 21)
        Me.chkDolares.TabIndex = 5
        Me.chkDolares.Text = "Dolares"
        Me.chkDolares.UseVisualStyleBackColor = False
        '
        'chkPesos
        '
        Me.chkPesos.BackColor = System.Drawing.SystemColors.Control
        Me.chkPesos.Checked = True
        Me.chkPesos.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPesos.Location = New System.Drawing.Point(158, 14)
        Me.chkPesos.Name = "chkPesos"
        Me.chkPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPesos.Size = New System.Drawing.Size(93, 21)
        Me.chkPesos.TabIndex = 6
        Me.chkPesos.Text = "Pesos"
        Me.chkPesos.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFechaInicial)
        Me.Frame1.Controls.Add(Me.dtpFechaFinal)
        Me.Frame1.Controls.Add(Me._Label2_1)
        Me.Frame1.Controls.Add(Me._Label2_0)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(15, 97)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(377, 57)
        Me.Frame1.TabIndex = 11
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(87, 21)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaInicial.TabIndex = 3
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(245, 21)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaFinal.TabIndex = 4
        '
        '_Label2_1
        '
        Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.SetIndex(Me._Label2_1, CType(1, Short))
        Me._Label2_1.Location = New System.Drawing.Point(192, 24)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_1.Size = New System.Drawing.Size(49, 21)
        Me._Label2_1.TabIndex = 13
        Me._Label2_1.Text = "Hasta el :"
        '
        '_Label2_0
        '
        Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.SetIndex(Me._Label2_0, CType(0, Short))
        Me._Label2_0.Location = New System.Drawing.Point(32, 24)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_0.Size = New System.Drawing.Size(57, 21)
        Me._Label2_0.TabIndex = 12
        Me._Label2_0.Text = "Desde el :"
        '
        '_fraRpt_3
        '
        Me._fraRpt_3.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_3.Controls.Add(Me.chkProveedores)
        Me._fraRpt_3.Controls.Add(Me.chkAcreedores)
        Me._fraRpt_3.Controls.Add(Me.dbcProveedor)
        Me._fraRpt_3.Controls.Add(Me._lblVentas_5)
        Me._fraRpt_3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_3, CType(3, Short))
        Me._fraRpt_3.Location = New System.Drawing.Point(15, 8)
        Me._fraRpt_3.Name = "_fraRpt_3"
        Me._fraRpt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_3.Size = New System.Drawing.Size(377, 82)
        Me._fraRpt_3.TabIndex = 9
        Me._fraRpt_3.TabStop = False
        Me._fraRpt_3.Text = "Proveedores/Acreedores"
        '
        'chkProveedores
        '
        Me.chkProveedores.BackColor = System.Drawing.SystemColors.Control
        Me.chkProveedores.Checked = True
        Me.chkProveedores.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkProveedores.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkProveedores.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkProveedores.Location = New System.Drawing.Point(77, 20)
        Me.chkProveedores.Name = "chkProveedores"
        Me.chkProveedores.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkProveedores.Size = New System.Drawing.Size(98, 18)
        Me.chkProveedores.TabIndex = 0
        Me.chkProveedores.Text = "Proveedores"
        Me.chkProveedores.UseVisualStyleBackColor = False
        '
        'chkAcreedores
        '
        Me.chkAcreedores.BackColor = System.Drawing.SystemColors.Control
        Me.chkAcreedores.Checked = True
        Me.chkAcreedores.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAcreedores.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAcreedores.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAcreedores.Location = New System.Drawing.Point(224, 20)
        Me.chkAcreedores.Name = "chkAcreedores"
        Me.chkAcreedores.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAcreedores.Size = New System.Drawing.Size(89, 18)
        Me.chkAcreedores.TabIndex = 1
        Me.chkAcreedores.Text = "Acreedores"
        Me.chkAcreedores.UseVisualStyleBackColor = False
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(83, 44)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(281, 21)
        Me.dbcProveedor.TabIndex = 2
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.AutoSize = True
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_5, CType(5, Short))
        Me._lblVentas_5.Location = New System.Drawing.Point(11, 48)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(68, 13)
        Me._lblVentas_5.TabIndex = 10
        Me._lblVentas_5.Text = "Prov/Acreed"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblRpt.SetIndex(Me._lblRpt_2, CType(2, Short))
        Me._lblRpt_2.Location = New System.Drawing.Point(15, 231)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 15
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(133, 341)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 133
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(18, 341)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 132
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(248, 342)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 131
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmCxpReporteSaldoXProveedor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(407, 389)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me._fraRpt_3)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmCxpReporteSaldoXProveedor"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte Auxiliar de Proveedores"
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me._fraRpt_3.ResumeLayout(False)
        Me._fraRpt_3.PerformLayout()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
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