Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmCXPReporteCuentasporPagar
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkPesos As System.Windows.Forms.CheckBox
    Public WithEvents chkDolares As System.Windows.Forms.CheckBox
    Public WithEvents chkEuros As System.Windows.Forms.CheckBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents chkDetalle As System.Windows.Forms.CheckBox
    Public WithEvents _chkCxP_0 As System.Windows.Forms.CheckBox
    Public WithEvents _chkCxP_1 As System.Windows.Forms.CheckBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents _fraRpt_3 As System.Windows.Forms.GroupBox
    Public WithEvents _chkOrigen_1 As System.Windows.Forms.CheckBox
    Public WithEvents _chkOrigen_0 As System.Windows.Forms.CheckBox
    Public WithEvents _fraCXP_0 As System.Windows.Forms.GroupBox
    Public WithEvents _optTipoRPT_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optTipoRPT_1 As System.Windows.Forms.RadioButton
    Public WithEvents _fraCXP_4 As System.Windows.Forms.GroupBox
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents lblFecha As System.Windows.Forms.Label
    Public WithEvents fraFechaCorte As System.Windows.Forms.Panel
    'Public WithEvents txtAnio As System.Windows.Forms.TextBox
    Public WithEvents cboMes As System.Windows.Forms.ComboBox
    Public WithEvents spnAnio As System.Windows.Forms.NumericUpDown
    Public WithEvents lblAnio As System.Windows.Forms.Label
    Public WithEvents lblMes As System.Windows.Forms.Label
    Public WithEvents fraMensual As System.Windows.Forms.Panel
    Public WithEvents _fraCXP_2 As System.Windows.Forms.GroupBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents chkCxP As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
    Public WithEvents chkOrigen As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
    Public WithEvents fraCXP As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents fraRpt As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optTipoRPT As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray

    Dim mblnSalir As Boolean
    Dim mblnFueraChange As Boolean

    Dim tecla As Integer
    Dim mintCodProveedor As Integer

    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaF As Boolean

    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Const C_NINGUNA As String = "[ Vacío ... ]"

    Function DiaFinal(ByRef Mes As Integer, ByRef Anio As Integer) As Integer
        Select Case Mes
            Case 1
                DiaFinal = 31
            Case 2
                If BICIESTO(Anio) Then
                    DiaFinal = 29
                Else
                    DiaFinal = 28
                End If
            Case 3
                DiaFinal = 31
            Case 4
                DiaFinal = 30
            Case 5
                DiaFinal = 31
            Case 6
                DiaFinal = 30
            Case 7
                DiaFinal = 31
            Case 8
                DiaFinal = 31
            Case 9
                DiaFinal = 30
            Case 10
                DiaFinal = 31
            Case 11
                DiaFinal = 30
            Case 12
                DiaFinal = 31
        End Select
    End Function

    Function DevuelveQuery() As String
        Dim Sql As String
        Dim Where As String
        Dim TipoGasto As String
        Dim FechaInicial As String
        Dim FechaFinal As String

        Sql = ""
        Where = ""
        FechaInicial = ""
        FechaFinal = ""

        If optTipoRPT(0).Checked = True Then
            'Determinamos la fecha inicial y final
            FechaInicial = VB.Right("00" & (cboMes.SelectedIndex + 1), 2) & "/01/" & spnAnio.Text
            FechaFinal = VB.Right("00" & (cboMes.SelectedIndex + 1), 2) & "/" & DiaFinal((cboMes.SelectedIndex) + 1, CInt(spnAnio.Text)) & "/" & spnAnio.Text

            If mintCodProveedor = 0 Then
                If chkCxP(0).CheckState = System.Windows.Forms.CheckState.Checked And chkCxP(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                    Where = Where & ""
                ElseIf chkCxP(0).CheckState = System.Windows.Forms.CheckState.Checked And chkCxP(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
                    Where = Where & "Tipo = '" & C_TPROVEEDOR & "' "
                ElseIf chkCxP(0).CheckState = System.Windows.Forms.CheckState.Unchecked And chkCxP(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                    Where = Where & "Tipo = '" & C_TACREEDOR & "' "
                End If
            End If

            If chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then
                Where = Where & ""
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                If Trim(Where) = "" Then
                    Where = Where & "(Moneda = '" & C_DOLAR & "' OR Moneda = '" & C_PESO & "') "
                Else
                    Where = Where & "AND (Moneda = '" & C_DOLAR & "' OR Moneda = '" & C_PESO & "') "
                End If
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                If Trim(Where) = "" Then
                    Where = Where & "Moneda = '" & C_DOLAR & "' "
                Else
                    Where = Where & "AND Moneda = '" & C_DOLAR & "' "
                End If
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then
                If Trim(Where) = "" Then
                    Where = Where & "(Moneda = '" & C_DOLAR & "' OR Moneda = '" & C_EURO & "') "
                Else
                    Where = Where & "AND (Moneda = '" & C_DOLAR & "' OR Moneda = '" & C_EURO & "') "
                End If
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then
                If Trim(Where) = "" Then
                    Where = Where & "(Moneda = '" & C_PESO & "' OR Moneda = '" & C_EURO & "') "
                Else
                    Where = Where & "AND (Moneda = '" & C_PESO & "' OR Moneda = '" & C_EURO & "') "
                End If
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                If Trim(Where) = "" Then
                    Where = Where & "Moneda = '" & C_PESO & "' "
                Else
                    Where = Where & "AND Moneda = '" & C_PESO & "' "
                End If
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then
                If Trim(Where) = "" Then
                    Where = Where & "Moneda = '" & C_EURO & "' "
                Else
                    Where = Where & "AND Moneda = '" & C_EURO & "' "
                End If
            End If

            If chkOrigen(0).CheckState = System.Windows.Forms.CheckState.Checked And chkOrigen(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                TipoGasto = ""
            ElseIf chkOrigen(0).CheckState = System.Windows.Forms.CheckState.Checked And chkOrigen(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
                TipoGasto = C_GASTOPERSONAL
            ElseIf chkOrigen(0).CheckState = System.Windows.Forms.CheckState.Unchecked And chkOrigen(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                TipoGasto = C_GASTOJOYERIA
            End If

            Sql = "SELECT * FROM DBO.CXPConcentrado('" & FechaInicial & "','" & FechaFinal & "'," & IIf(mintCodProveedor = 0, 0, mintCodProveedor) & ",'" & Trim(TipoGasto) & "') " & IIf(Trim(Where) <> "", "Where ", "") & Where

        ElseIf optTipoRPT(1).Checked = True Then
            If mintCodProveedor = 0 Then
                If chkCxP(0).CheckState = System.Windows.Forms.CheckState.Checked And chkCxP(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                    Where = Where & ""
                ElseIf chkCxP(0).CheckState = System.Windows.Forms.CheckState.Checked And chkCxP(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
                    Where = Where & "Tipo = '" & C_TPROVEEDOR & "' "
                ElseIf chkCxP(0).CheckState = System.Windows.Forms.CheckState.Unchecked And chkCxP(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                    Where = Where & "Tipo = '" & C_TACREEDOR & "' "
                End If
            End If

            If chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then
                Where = Where & ""
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                If Trim(Where) = "" Then
                    Where = Where & "(Moneda = '" & C_DOLAR & "' OR Moneda = '" & C_PESO & "') "
                Else
                    Where = Where & "AND (Moneda = '" & C_DOLAR & "' OR Moneda = '" & C_PESO & "') "
                End If
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                If Trim(Where) = "" Then
                    Where = Where & "Moneda = '" & C_DOLAR & "' "
                Else
                    Where = Where & "AND Moneda = '" & C_DOLAR & "' "
                End If
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Checked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then
                If Trim(Where) = "" Then
                    Where = Where & "(Moneda = '" & C_DOLAR & "' OR Moneda = '" & C_EURO & "') "
                Else
                    Where = Where & "AND (Moneda = '" & C_DOLAR & "' OR Moneda = '" & C_EURO & "') "
                End If
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then
                If Trim(Where) = "" Then
                    Where = Where & "(Moneda = '" & C_PESO & "' OR Moneda = '" & C_EURO & "') "
                Else
                    Where = Where & "AND (Moneda = '" & C_PESO & "' OR Moneda = '" & C_EURO & "') "
                End If
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Checked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                If Trim(Where) = "" Then
                    Where = Where & "Moneda = '" & C_PESO & "' "
                Else
                    Where = Where & "AND Moneda = '" & C_PESO & "' "
                End If
            ElseIf chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Checked Then
                If Trim(Where) = "" Then
                    Where = Where & "Moneda = '" & C_EURO & "' "
                Else
                    Where = Where & "AND Moneda = '" & C_EURO & "' "
                End If
            End If

            If chkOrigen(0).CheckState = System.Windows.Forms.CheckState.Checked And chkOrigen(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                TipoGasto = ""
            ElseIf chkOrigen(0).CheckState = System.Windows.Forms.CheckState.Checked And chkOrigen(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
                TipoGasto = C_GASTOPERSONAL
            ElseIf chkOrigen(0).CheckState = System.Windows.Forms.CheckState.Unchecked And chkOrigen(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                TipoGasto = C_GASTOJOYERIA
            End If
            Dim fechaHasta As String = AgregarHoraAFecha(dtpHasta.Value)
            Sql = "SELECT * FROM DBO.CXPFechaCorte('" & fechaHasta & "'," & IIf(mintCodProveedor = 0, 0, mintCodProveedor) & ",'" & Trim(TipoGasto) & "') " & IIf(Trim(Where) <> "", "Where ", "") & Where
        End If

        DevuelveQuery = Sql
    End Function

    Sub Imprime()
        Dim rptCXPRepCuentasXPagar As New rptCXPRepCuentasXPagar
        Dim rptCXPRepCuentasXPagarMensual As New rptCXPRepCuentasXPagarMensual
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr
        Dim lStrSql As String
        'Declarar vectores para almacenar los parámetros que se le enviarán al reporte
        Dim aParam(5) As Object
        Dim aValues(5) As Object
        Dim lValor As Boolean

        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim Periodo As String
        Dim TextoAdicional As String

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
            'Aqui se vincula con el reporte
            If optTipoRPT(0).Checked = True Then
                'NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
                'NombreReporte = UCase("REPORTE DE CUENTAS POR PAGAR CONCENTRADO MENSUAL")
                'Periodo = "CUENTAS POR PAGAR DEL MES DE " & UCase(cboMes.Text)
                'TextoAdicional = txtMensaje.Text
                With rptCXPRepCuentasXPagarMensual
                    'If chkDetalle.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                    '    'Ocultar
                    '    .Text1.Suppress = True
                    '    .Text6.Suppress = True
                    '    .Text7.Suppress = True
                    '    .Text8.Suppress = True
                    '    .Text9.Suppress = True
                    '    .Text10.Suppress = True
                    '    .Text5.Suppress = True
                    '    .Section5.Suppress = True
                    '    .Field13.Suppress = True
                    '    'Mostrar
                    '    .Text12.Suppress = False
                    '    .Field6.Suppress = False
                    '    .Section8.Height = 500
                    '    .Section9.Suppress = True
                    'ElseIf chkDetalle.CheckState = System.Windows.Forms.CheckState.Checked Then
                    '    'Mostrar
                    '    .Text1.Suppress = False
                    '    .Text6.Suppress = False
                    '    .Text7.Suppress = False
                    '    .Text8.Suppress = False
                    '    .Text9.Suppress = False
                    '    .Text10.Suppress = False
                    '    .Text5.Suppress = False
                    '    .Section5.Suppress = False
                    '    .Field13.Suppress = False
                    '    'Ocultar
                    '    .Text12.Suppress = True
                    '    .Field6.Suppress = True
                    '    .Section8.Height = 270
                    '    .Section9.Suppress = False
                    'End If
                End With
                rptCXPRepCuentasXPagarMensual.SetDataSource(frmReportes.rsReport)
                'frmReportes.Report = rptCXPRepCuentasXPagarMensual
                'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "Periodo", "TextoAdicional"}
                'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, Periodo, ModEstandar.QuitaEnter(TextoAdicional)}
                frmReportes.Text = "Reporte de cuentas por pagar a fecha de corte"
                frmReportes.reporteActual = rptCXPRepCuentasXPagarMensual
                frmReportes.Show()
            ElseIf optTipoRPT(1).Checked = True Then
                'NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
                'NombreReporte = UCase("REPORTE DE CUENTAS POR PAGAR A FECHA DE CORTE")
                'Periodo = "CUENTAS POR PAGAR HASTA EL " & UCase(Format(dtpHasta.Value, "dd/MMM/yyyy"))
                'TextoAdicional = txtMensaje.Text
                With rptCXPRepCuentasXPagar
                    'If chkDetalle.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                    '    'oculta
                    '    .Text1.Suppress = True
                    '    .Text2.Suppress = True
                    '    .Text4.Suppress = True
                    '    .Text5.Suppress = True
                    '    .Text6.Suppress = True
                    '    .Text7.Suppress = True
                    '    .Text8.Suppress = True
                    '    .Text11.Suppress = True
                    '    .Section8.Suppress = True
                    '    .Section5.Suppress = True
                    '    'muestra
                    '    .Text13.Suppress = False
                    '    .Text14.Suppress = False
                    '    .Field18.Suppress = False
                    '    .Section9.Height = 350
                    'ElseIf chkDetalle.CheckState = System.Windows.Forms.CheckState.Checked Then
                    '    'muestra
                    '    .Text1.Suppress = False
                    '    .Text2.Suppress = False
                    '    .Text4.Suppress = False
                    '    .Text5.Suppress = False
                    '    .Text6.Suppress = False
                    '    .Text7.Suppress = False
                    '    .Text8.Suppress = False
                    '    .Text11.Suppress = False
                    '    .Section8.Suppress = False
                    '    .Section5.Suppress = False
                    '    'oculta
                    '    .Text13.Suppress = True
                    '    .Text14.Suppress = True
                    '    .Field18.Suppress = True
                    '    .Section9.Height = 481
                    'End If
                End With
                rptCXPRepCuentasXPagar.SetDataSource(frmReportes.rsReport)
                'frmReportes.Report = rptCXPRepCuentasXPagar
                'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "Periodo", "TextoAdicional"}
                'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, Periodo, ModEstandar.QuitaEnter(TextoAdicional)}
                frmReportes.Text = "Reporte de cuentas por pagar a fecha de corte"
                frmReportes.reporteActual = rptCXPRepCuentasXPagar
                frmReportes.Show()
            End If
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Sub Limpiar()
        On Error Resume Next
        Call Me.Nuevo()
        Me.optTipoRPT(0).Focus()
    End Sub

    Sub Nuevo()
        Me.optTipoRPT(0).Checked = True
        Me.optTipoRPT(1).Checked = False
        optTipoRPT_CheckedChanged(optTipoRPT.Item(0), New System.EventArgs())
        'Marco de reporte mensual
        Me.cboMes.SelectedIndex = Month(Today) - 1
        Me.spnAnio.Text = CStr(Year(Today))

        'Marco de reporte a fecha de corte
        Me.dtpHasta.Value = Format(Today, "dd/MMM/yyyy")

        Me.chkOrigen(0).CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOrigen(1).CheckState = System.Windows.Forms.CheckState.Checked

        Me.chkCxP(0).CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCxP(1).CheckState = System.Windows.Forms.CheckState.Checked
        chkDolares.CheckState = System.Windows.Forms.CheckState.Checked
        chkPesos.CheckState = System.Windows.Forms.CheckState.Checked
        chkEuros.CheckState = System.Windows.Forms.CheckState.Checked
        mblnFueraChange = True
        Me.dbcProveedor.Text = C_TODOS
        Me.dbcProveedor.Tag = Me.dbcProveedor.Text
        mblnFueraChange = False
        Me.txtMensaje.Text = ""
        mblnTecleoFechaF = False
        mintCodProveedor = 0
        chkDetalle.CheckState = System.Windows.Forms.CheckState.Checked
    End Sub

    Function ValidaDatos() As Boolean
        If mblnTecleoFechaF Then
            Do While (VB.Timer() - msglTiempoCambioF) <= 2.1
            Loop
            mblnTecleoFechaF = False
        End If
        System.Windows.Forms.Application.DoEvents()


        Select Case True
            Case Me.optTipoRPT(0).Checked
                If CInt(Numerico((Me.spnAnio.Text))) < 1900 Or CInt(Numerico((Me.spnAnio.Text))) > 2075 Then
                    MsgBox("El año que especificó está fuera del rango permitido", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    ValidaDatos = False
                    Me.spnAnio.Focus()
                    ModEstandar.SelTxt()
                    Exit Function
                End If
            Case Else
                If Year(Me.dtpHasta.Value) < 1900 Or Year(Me.dtpHasta.Value) > 2075 Then
                    MsgBox("La fecha que especificó está fuera del rango permitido", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    ValidaDatos = False
                    Me.dtpHasta.Focus()
                    Exit Function
                End If
        End Select

        If chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked And chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked And chkEuros.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            MsgBox("Debe seleccionar minimo una moneda, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            ValidaDatos = False
            chkDolares.Focus()
            Exit Function
        End If

        Select Case True
            Case Me.chkOrigen(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkOrigen(1).CheckState = System.Windows.Forms.CheckState.Unchecked
                MsgBox("Debe escoger el origen del que provendrán los gastos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.chkOrigen(0).Focus()
                Exit Function
            Case Me.chkCxP(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkCxP(1).CheckState = System.Windows.Forms.CheckState.Unchecked
                MsgBox("Debe elegir algún tipo de cuentas por pagar: gastos o compras; o bien, los dos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.chkCxP(0).Focus()
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Private Sub chkCxP_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCxP.CheckStateChanged
        Dim Index As Integer = chkCxP.GetIndex(eventSender)
        If mblnFueraChange Then Exit Sub
        If chkCxP(0).CheckState = System.Windows.Forms.CheckState.Unchecked And chkCxP(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
            mblnFueraChange = True
            dbcProveedor.Text = ""
            dbcProveedor.Tag = dbcProveedor.Text
            mintCodProveedor = 0
            mblnFueraChange = False
        Else
            mblnFueraChange = True
            dbcProveedor.Text = C_TODOS
            dbcProveedor.Tag = dbcProveedor.Text
            mintCodProveedor = 0
            mblnFueraChange = False
        End If
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String
        Dim cWHERE As String

        If mblnFueraChange Then Exit Sub

        cWHERE = ""
        If Me.chkCxP(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkCxP(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = cWHERE & " Tipo = '" & C_TPROVEEDOR & "' and "
        ElseIf Me.chkCxP(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkCxP(1).CheckState = System.Windows.Forms.CheckState.Checked Then
            cWHERE = cWHERE & " Tipo = '" & C_TACREEDOR & "' and "
        ElseIf chkCxP(0).CheckState = System.Windows.Forms.CheckState.Unchecked And chkCxP(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = cWHERE & " Tipo = '' and "
        End If

        lStrSql = "SELECT CodProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM CatProvAcreed Where " & cWHERE & " DescProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
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
        If Me.chkCxP(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkCxP(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = cWHERE & " WHERE Tipo = '" & C_TPROVEEDOR & "'"
        ElseIf Me.chkCxP(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkCxP(1).CheckState = System.Windows.Forms.CheckState.Checked Then
            cWHERE = cWHERE & " WHERE Tipo = '" & C_TACREEDOR & "'"
        ElseIf chkCxP(0).CheckState = System.Windows.Forms.CheckState.Unchecked And chkCxP(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = cWHERE & " WHERE Tipo = '' "
        End If
        Pon_Tool()
        gStrSql = "SELECT CodProvAcreed, LTrim(RTrim(DescProvAcreed)) as DescProvAcreed FROM CatProvAcreed " & cWHERE
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcProveedor))
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcProveedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkCxP(1).Focus()
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
        If Me.chkCxP(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkCxP(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = cWHERE & " Tipo = '" & C_TPROVEEDOR & "' and "
        ElseIf Me.chkCxP(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkCxP(1).CheckState = System.Windows.Forms.CheckState.Checked Then
            cWHERE = cWHERE & " Tipo = '" & C_TACREEDOR & "' and "
        ElseIf chkCxP(0).CheckState = System.Windows.Forms.CheckState.Unchecked And chkCxP(1).CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = cWHERE & " Tipo = '' and "
        End If
        gStrSql = "SELECT CodProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM CatProvAcreed Where " & cWHERE & " DescProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        Aux = mintCodProveedor
        mintCodProveedor = 0
        If Trim(Me.dbcProveedor.Text) <> Trim(C_TODOS) Or Trim(Me.dbcProveedor.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcProveedor), gStrSql, mintCodProveedor)
        End If

        If Aux <> mintCodProveedor Then
            If mintCodProveedor = 0 Then
                mblnFueraChange = True
                If chkCxP(0).CheckState = System.Windows.Forms.CheckState.Checked Or chkCxP(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                    Me.dbcProveedor.Text = C_TODOS
                Else
                    Me.dbcProveedor.Text = ""
                End If
                mblnFueraChange = False
            End If
        End If
        If Trim(Me.dbcProveedor.Text) = "" Then
            mblnFueraChange = True
            If chkCxP(0).CheckState = System.Windows.Forms.CheckState.Checked Or chkCxP(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                Me.dbcProveedor.Text = C_TODOS
            Else
                Me.dbcProveedor.Text = ""
            End If
            mblnFueraChange = False
        End If
    End Sub

    Private Sub dbcProveedor_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcProveedor.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcProveedor.Text)
        'If Me.dbcProveedor.SelectedItem <> 0 Then
        '    dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        'End If
        Me.dbcProveedor.Text = Aux
    End Sub

    Private Sub frmCXPReporteCuentasporPagar_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCXPReporteCuentasporPagar_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPReporteCuentasporPagar_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "OPTTIPORPT" Then
                    If Me.optTipoRPT(0).Checked Or Me.optTipoRPT(1).Checked Then
                        mblnSalir = True
                        Me.Close()
                    Else
                        ModEstandar.RetrocederTab(Me)
                    End If
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCXPReporteCuentasporPagar_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPReporteCuentasporPagar_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
    End Sub

    Private Sub frmCXPReporteCuentasporPagar_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If mblnSalir Then
        '    mblnSalir = False
        '    Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "optTipoRPT" Then
        '                If optTipoRPT(0).Checked Then
        '                    Me.optTipoRPT(0).Focus()
        '                ElseIf optTipoRPT(1).Checked Then
        '                    Me.optTipoRPT(1).Focus()
        '                End If
        '            End If
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCXPReporteCuentasporPagar_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optTipoRPT_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTipoRPT.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = optTipoRPT.GetIndex(eventSender)
            Select Case True
                Case Me.optTipoRPT(0).Checked
                    Me.fraMensual.Visible = True
                    Me.fraMensual.Enabled = True
                    Me.fraFechaCorte.Visible = False
                    Me.fraFechaCorte.Enabled = False
                Case Me.optTipoRPT(1).Checked
                    Me.fraFechaCorte.Visible = True
                    Me.fraFechaCorte.Enabled = True
                    Me.fraMensual.Visible = False
                    Me.fraMensual.Enabled = False
            End Select
        End If
    End Sub

    Private Sub spnAnio_DownClick()
        Dim nAnio As Integer
        nAnio = CInt(Numerico((Me.spnAnio.Text)))
        nAnio = nAnio - 1
        If nAnio >= 2075 Then
            Me.spnAnio.Text = CStr(2075)
        ElseIf nAnio <= 1900 Then
            Me.spnAnio.Text = CStr(1900)
        ElseIf nAnio > 1900 And nAnio < 2075 Then
            Me.spnAnio.Text = CStr(nAnio)
        End If
    End Sub

    Private Sub spnAnio_UpClick()
        Dim nAnio As Integer
        nAnio = CInt(Numerico((Me.spnAnio.Text)))
        nAnio = nAnio + 1
        If nAnio >= 2075 Then
            Me.spnAnio.Text = CStr(2075)
        ElseIf nAnio <= 1900 Then
            Me.spnAnio.Text = CStr(1900)
        ElseIf nAnio > 1900 And nAnio < 2075 Then
            Me.spnAnio.Text = CStr(nAnio)
        End If
    End Sub

    Private Sub spnAnio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub spnAnio_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs)
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Up
                'spnAnio_UpClick(spnAnio, New System.EventArgs())
                spnAnio_UpClick()
            Case System.Windows.Forms.Keys.Down
                'spnAnio_DownClick(spnAnio, New System.EventArgs())
                spnAnio_UpClick()
        End Select
    End Sub

    Private Sub spnAnio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.spnAnio.Text = Format(Numerico((Me.spnAnio.Text)), "###0")
        End If
        KeyAscii = ModEstandar.MskCantidad((Me.spnAnio.Text), KeyAscii, 4, 0, (Convert.ToInt32(Me.spnAnio)))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub spnAnio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim nAnio As Integer
        nAnio = CInt(Numerico((Me.spnAnio.Text)))
        If nAnio < 1900 Then
            Me.spnAnio.Text = CStr(1900)
        ElseIf nAnio > 2075 Then
            Me.spnAnio.Text = CStr(2075)
        End If
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me._chkOrigen_1 = New System.Windows.Forms.CheckBox()
        Me._chkOrigen_0 = New System.Windows.Forms.CheckBox()
        Me._optTipoRPT_0 = New System.Windows.Forms.RadioButton()
        Me._optTipoRPT_1 = New System.Windows.Forms.RadioButton()
        Me.cboMes = New System.Windows.Forms.ComboBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.chkPesos = New System.Windows.Forms.CheckBox()
        Me.chkDolares = New System.Windows.Forms.CheckBox()
        Me.chkEuros = New System.Windows.Forms.CheckBox()
        Me.chkDetalle = New System.Windows.Forms.CheckBox()
        Me._fraRpt_3 = New System.Windows.Forms.GroupBox()
        Me._chkCxP_0 = New System.Windows.Forms.CheckBox()
        Me._chkCxP_1 = New System.Windows.Forms.CheckBox()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me._fraCXP_0 = New System.Windows.Forms.GroupBox()
        Me._fraCXP_2 = New System.Windows.Forms.GroupBox()
        Me._fraCXP_4 = New System.Windows.Forms.GroupBox()
        Me.fraFechaCorte = New System.Windows.Forms.Panel()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me.lblFecha = New System.Windows.Forms.Label()
        Me.fraMensual = New System.Windows.Forms.Panel()
        Me.spnAnio = New System.Windows.Forms.NumericUpDown()
        Me.lblAnio = New System.Windows.Forms.Label()
        Me.lblMes = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.chkCxP = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.chkOrigen = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.fraCXP = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.fraRpt = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optTipoRPT = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me._fraRpt_3.SuspendLayout()
        Me._fraCXP_0.SuspendLayout()
        Me._fraCXP_2.SuspendLayout()
        Me.fraFechaCorte.SuspendLayout()
        Me.fraMensual.SuspendLayout()
        CType(Me.spnAnio, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkCxP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkOrigen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraCXP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optTipoRPT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(16, 326)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(393, 70)
        Me.txtMensaje.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        '_chkOrigen_1
        '
        Me._chkOrigen_1.BackColor = System.Drawing.SystemColors.Control
        Me._chkOrigen_1.Checked = True
        Me._chkOrigen_1.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkOrigen_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkOrigen_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkOrigen.SetIndex(Me._chkOrigen_1, CType(1, Short))
        Me._chkOrigen_1.Location = New System.Drawing.Point(96, 24)
        Me._chkOrigen_1.Name = "_chkOrigen_1"
        Me._chkOrigen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkOrigen_1.Size = New System.Drawing.Size(61, 19)
        Me._chkOrigen_1.TabIndex = 13
        Me._chkOrigen_1.Text = "Joyería"
        Me.ToolTip1.SetToolTip(Me._chkOrigen_1, "Selecciona todos los gastos de la joyería")
        Me._chkOrigen_1.UseVisualStyleBackColor = False
        '
        '_chkOrigen_0
        '
        Me._chkOrigen_0.BackColor = System.Drawing.SystemColors.Control
        Me._chkOrigen_0.Checked = True
        Me._chkOrigen_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkOrigen_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkOrigen_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkOrigen.SetIndex(Me._chkOrigen_0, CType(0, Short))
        Me._chkOrigen_0.Location = New System.Drawing.Point(16, 24)
        Me._chkOrigen_0.Name = "_chkOrigen_0"
        Me._chkOrigen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkOrigen_0.Size = New System.Drawing.Size(74, 19)
        Me._chkOrigen_0.TabIndex = 12
        Me._chkOrigen_0.Text = "Personal"
        Me.ToolTip1.SetToolTip(Me._chkOrigen_0, "Selecciona todos los gastos personales")
        Me._chkOrigen_0.UseVisualStyleBackColor = False
        '
        '_optTipoRPT_0
        '
        Me._optTipoRPT_0.BackColor = System.Drawing.SystemColors.Control
        Me._optTipoRPT_0.Checked = True
        Me._optTipoRPT_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoRPT_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTipoRPT.SetIndex(Me._optTipoRPT_0, CType(0, Short))
        Me._optTipoRPT_0.Location = New System.Drawing.Point(16, 32)
        Me._optTipoRPT_0.Name = "_optTipoRPT_0"
        Me._optTipoRPT_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoRPT_0.Size = New System.Drawing.Size(130, 21)
        Me._optTipoRPT_0.TabIndex = 0
        Me._optTipoRPT_0.TabStop = True
        Me._optTipoRPT_0.Text = "Concentrado mensual"
        Me.ToolTip1.SetToolTip(Me._optTipoRPT_0, "Concentrado mensual")
        Me._optTipoRPT_0.UseVisualStyleBackColor = False
        '
        '_optTipoRPT_1
        '
        Me._optTipoRPT_1.BackColor = System.Drawing.SystemColors.Control
        Me._optTipoRPT_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoRPT_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTipoRPT.SetIndex(Me._optTipoRPT_1, CType(1, Short))
        Me._optTipoRPT_1.Location = New System.Drawing.Point(16, 56)
        Me._optTipoRPT_1.Name = "_optTipoRPT_1"
        Me._optTipoRPT_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoRPT_1.Size = New System.Drawing.Size(114, 25)
        Me._optTipoRPT_1.TabIndex = 1
        Me._optTipoRPT_1.TabStop = True
        Me._optTipoRPT_1.Text = "A fecha de corte"
        Me.ToolTip1.SetToolTip(Me._optTipoRPT_1, "A fecha de corte")
        Me._optTipoRPT_1.UseVisualStyleBackColor = False
        '
        'cboMes
        '
        Me.cboMes.BackColor = System.Drawing.SystemColors.Window
        Me.cboMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboMes.Items.AddRange(New Object() {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"})
        Me.cboMes.Location = New System.Drawing.Point(72, 24)
        Me.cboMes.Name = "cboMes"
        Me.cboMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboMes.Size = New System.Drawing.Size(89, 21)
        Me.cboMes.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.cboMes, "Mes del corte")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.chkPesos)
        Me.Frame2.Controls.Add(Me.chkDolares)
        Me.Frame2.Controls.Add(Me.chkEuros)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(16, 208)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(393, 41)
        Me.Frame2.TabIndex = 27
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Movimientos en ...."
        '
        'chkPesos
        '
        Me.chkPesos.BackColor = System.Drawing.SystemColors.Control
        Me.chkPesos.Checked = True
        Me.chkPesos.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPesos.Location = New System.Drawing.Point(166, 14)
        Me.chkPesos.Name = "chkPesos"
        Me.chkPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPesos.Size = New System.Drawing.Size(93, 21)
        Me.chkPesos.TabIndex = 10
        Me.chkPesos.Text = "Pesos"
        Me.chkPesos.UseVisualStyleBackColor = False
        '
        'chkDolares
        '
        Me.chkDolares.BackColor = System.Drawing.SystemColors.Control
        Me.chkDolares.Checked = True
        Me.chkDolares.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDolares.Location = New System.Drawing.Point(64, 14)
        Me.chkDolares.Name = "chkDolares"
        Me.chkDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDolares.Size = New System.Drawing.Size(93, 21)
        Me.chkDolares.TabIndex = 9
        Me.chkDolares.Text = "Dolares"
        Me.chkDolares.UseVisualStyleBackColor = False
        '
        'chkEuros
        '
        Me.chkEuros.BackColor = System.Drawing.SystemColors.Control
        Me.chkEuros.Checked = True
        Me.chkEuros.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkEuros.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEuros.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEuros.Location = New System.Drawing.Point(268, 14)
        Me.chkEuros.Name = "chkEuros"
        Me.chkEuros.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEuros.Size = New System.Drawing.Size(93, 21)
        Me.chkEuros.TabIndex = 11
        Me.chkEuros.Text = "Euros"
        Me.chkEuros.UseVisualStyleBackColor = False
        '
        'chkDetalle
        '
        Me.chkDetalle.BackColor = System.Drawing.SystemColors.Control
        Me.chkDetalle.Checked = True
        Me.chkDetalle.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDetalle.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDetalle.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDetalle.Location = New System.Drawing.Point(296, 288)
        Me.chkDetalle.Name = "chkDetalle"
        Me.chkDetalle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDetalle.Size = New System.Drawing.Size(105, 17)
        Me.chkDetalle.TabIndex = 14
        Me.chkDetalle.Text = "Detalle por folio"
        Me.chkDetalle.UseVisualStyleBackColor = False
        '
        '_fraRpt_3
        '
        Me._fraRpt_3.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_3.Controls.Add(Me._chkCxP_0)
        Me._fraRpt_3.Controls.Add(Me._chkCxP_1)
        Me._fraRpt_3.Controls.Add(Me.dbcProveedor)
        Me._fraRpt_3.Controls.Add(Me._lblVentas_5)
        Me._fraRpt_3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_3, CType(3, Short))
        Me._fraRpt_3.Location = New System.Drawing.Point(16, 120)
        Me._fraRpt_3.Name = "_fraRpt_3"
        Me._fraRpt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_3.Size = New System.Drawing.Size(393, 82)
        Me._fraRpt_3.TabIndex = 24
        Me._fraRpt_3.TabStop = False
        Me._fraRpt_3.Text = "Cuentas por pagar ..."
        '
        '_chkCxP_0
        '
        Me._chkCxP_0.BackColor = System.Drawing.SystemColors.Control
        Me._chkCxP_0.Checked = True
        Me._chkCxP_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkCxP_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkCxP_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCxP.SetIndex(Me._chkCxP_0, CType(0, Short))
        Me._chkCxP_0.Location = New System.Drawing.Point(24, 24)
        Me._chkCxP_0.Name = "_chkCxP_0"
        Me._chkCxP_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkCxP_0.Size = New System.Drawing.Size(154, 18)
        Me._chkCxP_0.TabIndex = 6
        Me._chkCxP_0.Text = "de Proveedores (compras)"
        Me._chkCxP_0.UseVisualStyleBackColor = False
        '
        '_chkCxP_1
        '
        Me._chkCxP_1.BackColor = System.Drawing.SystemColors.Control
        Me._chkCxP_1.Checked = True
        Me._chkCxP_1.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkCxP_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkCxP_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCxP.SetIndex(Me._chkCxP_1, CType(1, Short))
        Me._chkCxP_1.Location = New System.Drawing.Point(224, 24)
        Me._chkCxP_1.Name = "_chkCxP_1"
        Me._chkCxP_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkCxP_1.Size = New System.Drawing.Size(127, 18)
        Me._chkCxP_1.TabIndex = 7
        Me._chkCxP_1.Text = "de Acreedores (gastos)"
        Me._chkCxP_1.UseVisualStyleBackColor = False
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(96, 48)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(281, 21)
        Me.dbcProveedor.TabIndex = 8
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.AutoSize = True
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_5, CType(5, Short))
        Me._lblVentas_5.Location = New System.Drawing.Point(24, 52)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(68, 13)
        Me._lblVentas_5.TabIndex = 25
        Me._lblVentas_5.Text = "Prov/Acreed"
        '
        '_fraCXP_0
        '
        Me._fraCXP_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraCXP_0.Controls.Add(Me._chkOrigen_1)
        Me._fraCXP_0.Controls.Add(Me._chkOrigen_0)
        Me._fraCXP_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraCXP.SetIndex(Me._fraCXP_0, CType(0, Short))
        Me._fraCXP_0.Location = New System.Drawing.Point(16, 256)
        Me._fraCXP_0.Name = "_fraCXP_0"
        Me._fraCXP_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraCXP_0.Size = New System.Drawing.Size(161, 49)
        Me._fraCXP_0.TabIndex = 23
        Me._fraCXP_0.TabStop = False
        Me._fraCXP_0.Text = "Origen ..."
        '
        '_fraCXP_2
        '
        Me._fraCXP_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraCXP_2.Controls.Add(Me._optTipoRPT_0)
        Me._fraCXP_2.Controls.Add(Me._optTipoRPT_1)
        Me._fraCXP_2.Controls.Add(Me._fraCXP_4)
        Me._fraCXP_2.Controls.Add(Me.fraFechaCorte)
        Me._fraCXP_2.Controls.Add(Me.fraMensual)
        Me._fraCXP_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraCXP.SetIndex(Me._fraCXP_2, CType(2, Short))
        Me._fraCXP_2.Location = New System.Drawing.Point(16, 8)
        Me._fraCXP_2.Name = "_fraCXP_2"
        Me._fraCXP_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraCXP_2.Size = New System.Drawing.Size(393, 105)
        Me._fraCXP_2.TabIndex = 16
        Me._fraCXP_2.TabStop = False
        Me._fraCXP_2.Text = "Tipo de Reporte ..."
        '
        '_fraCXP_4
        '
        Me._fraCXP_4.BackColor = System.Drawing.SystemColors.Control
        Me._fraCXP_4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraCXP.SetIndex(Me._fraCXP_4, CType(4, Short))
        Me._fraCXP_4.Location = New System.Drawing.Point(168, 8)
        Me._fraCXP_4.Name = "_fraCXP_4"
        Me._fraCXP_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraCXP_4.Size = New System.Drawing.Size(2, 89)
        Me._fraCXP_4.TabIndex = 17
        Me._fraCXP_4.TabStop = False
        '
        'fraFechaCorte
        '
        Me.fraFechaCorte.BackColor = System.Drawing.SystemColors.Control
        Me.fraFechaCorte.Controls.Add(Me.dtpHasta)
        Me.fraFechaCorte.Controls.Add(Me.lblFecha)
        Me.fraFechaCorte.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraFechaCorte.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraFechaCorte.Location = New System.Drawing.Point(184, 16)
        Me.fraFechaCorte.Name = "fraFechaCorte"
        Me.fraFechaCorte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraFechaCorte.Size = New System.Drawing.Size(202, 81)
        Me.fraFechaCorte.TabIndex = 21
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(67, 34)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(100, 20)
        Me.dtpHasta.TabIndex = 5
        '
        'lblFecha
        '
        Me.lblFecha.AutoSize = True
        Me.lblFecha.BackColor = System.Drawing.SystemColors.Control
        Me.lblFecha.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFecha.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFecha.Location = New System.Drawing.Point(27, 38)
        Me.lblFecha.Name = "lblFecha"
        Me.lblFecha.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFecha.Size = New System.Drawing.Size(37, 13)
        Me.lblFecha.TabIndex = 22
        Me.lblFecha.Text = "Fecha"
        '
        'fraMensual
        '
        Me.fraMensual.BackColor = System.Drawing.SystemColors.Control
        Me.fraMensual.Controls.Add(Me.cboMes)
        Me.fraMensual.Controls.Add(Me.spnAnio)
        Me.fraMensual.Controls.Add(Me.lblAnio)
        Me.fraMensual.Controls.Add(Me.lblMes)
        Me.fraMensual.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraMensual.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraMensual.Location = New System.Drawing.Point(176, 8)
        Me.fraMensual.Name = "fraMensual"
        Me.fraMensual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMensual.Size = New System.Drawing.Size(202, 89)
        Me.fraMensual.TabIndex = 18
        '
        'spnAnio
        '
        Me.spnAnio.BackColor = System.Drawing.SystemColors.Control
        Me.spnAnio.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.spnAnio.Location = New System.Drawing.Point(95, 60)
        Me.spnAnio.Maximum = New Decimal(New Integer() {9000, 0, 0, 0})
        Me.spnAnio.Name = "spnAnio"
        Me.spnAnio.Size = New System.Drawing.Size(90, 20)
        Me.spnAnio.TabIndex = 4
        '
        'lblAnio
        '
        Me.lblAnio.AutoSize = True
        Me.lblAnio.BackColor = System.Drawing.SystemColors.Control
        Me.lblAnio.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAnio.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAnio.Location = New System.Drawing.Point(67, 60)
        Me.lblAnio.Name = "lblAnio"
        Me.lblAnio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAnio.Size = New System.Drawing.Size(26, 13)
        Me.lblAnio.TabIndex = 20
        Me.lblAnio.Text = "Año"
        '
        'lblMes
        '
        Me.lblMes.AutoSize = True
        Me.lblMes.BackColor = System.Drawing.SystemColors.Control
        Me.lblMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMes.Location = New System.Drawing.Point(48, 28)
        Me.lblMes.Name = "lblMes"
        Me.lblMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMes.Size = New System.Drawing.Size(27, 13)
        Me.lblMes.TabIndex = 19
        Me.lblMes.Text = "Mes"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblRpt.SetIndex(Me._lblRpt_2, CType(2, Short))
        Me._lblRpt_2.Location = New System.Drawing.Point(19, 312)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 26
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'chkCxP
        '
        '
        'optTipoRPT
        '
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(131, 415)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 130
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(16, 415)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 129
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(246, 416)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 128
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmCXPReporteCuentasporPagar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(420, 462)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me.chkDetalle)
        Me.Controls.Add(Me._fraRpt_3)
        Me.Controls.Add(Me._fraCXP_0)
        Me.Controls.Add(Me._fraCXP_2)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmCXPReporteCuentasporPagar"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Cuentas por pagar"
        Me.Frame2.ResumeLayout(False)
        Me._fraRpt_3.ResumeLayout(False)
        Me._fraRpt_3.PerformLayout()
        Me._fraCXP_0.ResumeLayout(False)
        Me._fraCXP_2.ResumeLayout(False)
        Me.fraFechaCorte.ResumeLayout(False)
        Me.fraFechaCorte.PerformLayout()
        Me.fraMensual.ResumeLayout(False)
        Me.fraMensual.PerformLayout()
        CType(Me.spnAnio, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkCxP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkOrigen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraCXP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optTipoRPT, System.ComponentModel.ISupportInitialize).EndInit()
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