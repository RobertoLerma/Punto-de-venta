Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource

Public Class frmCXPPresupuestado
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents cmbAño As System.Windows.Forms.ComboBox
    Public WithEvents cmbMes As System.Windows.Forms.ComboBox
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Dim mblnSalir As Boolean
    Dim I As Integer
    Dim TipoCambioDol As Decimal
    Dim TipoCambioEuro As Decimal
    Dim EfectivoCaja As Decimal
    Dim BancosPesos As Decimal
    Dim Dolares As Decimal
    Dim BancosDolares As Decimal
    Dim TarjetasNoAcreditadas As Decimal
    Dim TotalSinTarjetas As Decimal
    Dim TotalGeneral As Decimal
    Dim TotalCXP As Decimal
    Dim rsReporte As New ADODB.Recordset
    Dim RsAux As New ADODB.Recordset
    Dim SaldoXPagarPesos As Decimal
    Dim SaldoXPagarDolares As Decimal
    Dim VentaDiariaPesos As Decimal
    Dim VentaDiariaDolares As Decimal
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim DiasXTranscurrir As Integer

    Sub CalculaImportes()

        SaldoXPagarPesos = 0
        SaldoXPagarDolares = 0
        TotalCXP = 0

        EfectivoCaja = 0
        Dolares = 0
        TarjetasNoAcreditadas = 0
        Dim fecha As String = AgregarHoraAFecha(Today)
        gStrSql = "SELECT ISNULL((ABS(SUM(CASE WHEN Tipo = 'I' THEN ImporteDolares ELSE 0 END)) - " & "ABS(SUM(CASE WHEN TIPO = 'D' THEN ImporteDolares ELSE 0 END)) - " & "ABS(SUM(CASE WHEN TIPO = 'R' THEN ImporteDolares ELSE 0 END))),0) AS EfectivoDolares," & "ISNULL((ABS(SUM(CASE WHEN Tipo = 'I' THEN ImportePesos ELSE 0 END)) - " & "ABS(SUM(CASE WHEN TIPO = 'D' THEN ImportePesos ELSE 0 END)) - " & "ABS(SUM(CASE WHEN TIPO = 'R' THEN ImportePesos ELSE 0 END))),0) AS EfectivoPesos," & "ISNULL(ABS(SUM(CASE WHEN TIPO = 'T' THEN ImportePesos ELSE 0 END)),0) AS ImporteTarjetas " & "FROM DBO.vw_ObtenerIngresos " & "WHERE FechaIngreso <= '" & fecha & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount > 0 Then
            EfectivoCaja = frmReportes.rsReport.Fields("EfectivoPesos").Value
            Dolares = frmReportes.rsReport.Fields("EfectivoDolares").Value
            TarjetasNoAcreditadas = frmReportes.rsReport.Fields("ImporteTarjetas").Value
        End If
        'Obtener los Importes de bancos
        'obtener los saldos de bancos en pesos
        BancosPesos = 0
        BancosDolares = 0
        gStrSql = "SELECT CB.CODBANCO,CB.CTABANCARIA,CB.MONEDA FROM CATCUENTASBANCARIAS CB INNER JOIN CATBANCOS B ON CB.CODBANCO = B.CODBANCO WHERE B.CONTROLINTERNO = 0 AND CB.MONEDA = 'P'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount > 0 Then
            Do While Not frmReportes.rsReport.EOF
                gStrSql = "SELECT DBO.SALDOXCUENTA('" & fecha & "'," & frmReportes.rsReport.Fields("CodBanco").Value & ",'" & Trim(frmReportes.rsReport.Fields("CtaBancaria").Value) & "') AS Saldo"
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                frmReportes.rsReport = Cmd.Execute
                BancosPesos = BancosPesos + frmReportes.rsReport.Fields("SALDO").Value
                frmReportes.rsReport.MoveNext()
            Loop
        End If
        'obtener los saldos de bancos en dolares
        gStrSql = "SELECT CB.CODBANCO,CB.CTABANCARIA,CB.MONEDA FROM CATCUENTASBANCARIAS CB INNER JOIN CATBANCOS B ON CB.CODBANCO = B.CODBANCO WHERE B.CONTROLINTERNO = 0 AND CB.MONEDA = 'D' "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount > 0 Then
            Do While Not frmReportes.rsReport.EOF
                gStrSql = "SELECT DBO.SALDOXCUENTA('" & fecha & "'," & frmReportes.rsReport.Fields("CodBanco").Value & ",'" & Trim(frmReportes.rsReport.Fields("CtaBancaria").Value) & "') AS Saldo "
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                frmReportes.rsReport = Cmd.Execute
                BancosDolares = BancosDolares + frmReportes.rsReport.Fields("SALDO").Value
                frmReportes.rsReport.MoveNext()
            Loop
        End If
        Dolares = System.Math.Round(Dolares * TipoCambioDol, 2)
        BancosDolares = System.Math.Round(BancosDolares * TipoCambioDol, 2)
        TotalSinTarjetas = (EfectivoCaja + BancosPesos + Dolares + BancosDolares)
        TotalGeneral = (EfectivoCaja + BancosPesos + Dolares + BancosDolares + TarjetasNoAcreditadas)
        'Obtenemos el Total de CXP
        'RsAux = rsReporte
        Do While Not frmReportes.rsReport.EOF
            TotalCXP = TotalCXP + frmReportes.rsReport.Fields("importe").Value
            frmReportes.rsReport.MoveNext()
        Loop
        SaldoXPagarPesos = TotalCXP - TotalGeneral
        SaldoXPagarDolares = System.Math.Round(SaldoXPagarPesos / TipoCambioDol, 1)
        DiasXTranscurrir = DiaFinal() - VB.Day(Today)
        VentaDiariaPesos = System.Math.Round(SaldoXPagarPesos / DiasXTranscurrir, 2)
        VentaDiariaDolares = System.Math.Round(SaldoXPagarDolares / DiasXTranscurrir, 2)
    End Sub

    Function DiaFinal() As Integer
        Select Case Month(Today)
            Case 1
                DiaFinal = 31
            Case 2
                If BICIESTO(Year(Today)) Then
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

    Sub Imprime()
        Dim rptCXPReportePresupuesto As New rptCXPReportePresupuesto
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Err_Renamed

        Dim lStrSql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim Periodo As String
        Dim TextoAdicional As String

        TipoCambioDol = gcurCorpoTIPOCAMBIODOLAR
        TipoCambioEuro = gcurCorpoTIPOCAMBIOEURO
        lStrSql = Query()
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, lStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existe informacion para este periodo...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            rptCXPReportePresupuesto.SetDataSource(frmReportes.rsReport)
        End If

        CalculaImportes()
        'NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        'NombreReporte = UCase("Reporte de Cuentas por Pagar Presupuestado")
        'Periodo = Trim(cmbMes.Text)
        'TextoAdicional = txtMensaje.Text
        'frmReportes.Report = rptCXPReportePresupuesto 
        '        rptCXPReportePresupuesto.Section4.ReportObjects.Item("Text14").Kind.TextObject
        '.SetText(Format(EfectivoCaja, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text15.SetText(Format(BancosPesos, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text16.SetText(Format(Dolares, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text17.SetText(Format(BancosDolares, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text18.SetText(Format(TarjetasNoAcreditadas, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text20.SetText(Format(TotalSinTarjetas, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text19.SetText(Format(TotalGeneral, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text28.SetText(Format(SaldoXPagarPesos, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text30.SetText(Format(SaldoXPagarDolares, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text29.SetText(Format(VentaDiariaPesos, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text31.SetText(Format(VentaDiariaDolares, "###,##0.00"))
        '        rptCXPReportePresupuesto.Text33.SetText(Format(TotalCXP, "###,##0.00"))
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        'frmReportes.rsReport = rsReporte
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "Periodo", "TextoAdicional", "TipoCambio"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, Periodo, TextoAdicional, Format(TipoCambioDol, "###,##0.00")}
        frmReportes.Text = "Auxiliar de Proveedores"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.reporteActual = rptCXPReportePresupuesto
        frmReportes.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Default

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Limpiar()
        Nuevo()
        cmbMes.Focus()
    End Sub

    Sub LlenaAños()
        cmbAño.Items.Clear()
        For I = 1900 To 2075
            cmbAño.Items.Add(CStr(I))
        Next
    End Sub

    Sub Nuevo()
        LlenaAños()
        cmbMes.SelectedIndex = Month(Today) - 1
        cmbAño.Text = CStr(Year(Today))
        txtMensaje.Text = ""
    End Sub

    Function Query() As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        ModCorporativo.ObtenerLimitedeFechas(cmbMes.SelectedIndex + 1, CInt(cmbAño.Text), FechaInicial, FechaFinal)
        Query = "/*PROGRAMADO*/ select a.fechapago, a.codprovacreed, b.descprovacreed, a.foliofactura,'P' as tipo,sum(isnull(case when moneda = 'P' then round(a.totalpago,2) when moneda = 'D' then round(a.totalpago * " & TipoCambioDol & ",2) when moneda = 'E' then round(a.totalpago/" & TipoCambioEuro & ",2) end,0)) as importe " & "from     programacionpagos a (Nolock) inner join catprovacreed b on a.codprovacreed = b.codprovacreed where fechapago between '" & FechaInicial & "' and '" & FechaFinal & "' and estatus <> 'C' group by a.fechapago,a.codprovacreed,b.descprovacreed,a.foliofactura " & "Union " & "/*PENDIENTE*/ select '01/01/1900' as fechapago,a.codprovacreed,b.descprovacreed,'' as foliofactura,'A' as tipo,(isnull(a.totalpago,0) - (isnull(p.totalpago,0) + isnull(n.total,0) + isnull(an.total,0))) as total " & "from     (select codprovacreed,(sum(isnull(case when moneda = 'P' then round(totalpago,2) when moneda = 'D' then round((totalpago * " & TipoCambioDol & "),2) when moneda = 'E' then round((totalpago/" & TipoCambioEuro & "),2) end,0))) as totalpago " & "from programacionpagos (Nolock) where fechapago < '" & FechaInicial & "' and estatus <> 'C' group by codprovacreed) a left outer join catprovacreed b on a.codprovacreed = b.codprovacreed " & "left outer join (select codprovacreed,sum(isnull(case when moneda = 'P' then round(totalpago,2) when moneda = 'D' then round((totalpago * " & TipoCambioDol & "),2) when moneda = 'E' then round((totalpago/" & TipoCambioEuro & "),2) end,0)) as totalpago " & "from pagos (Nolock) where fechapago < '" & FechaInicial & "' and estatus <> 'C' group by codprovacreed) p on a.codprovacreed = p.codprovacreed " & "left outer join (select codprovacreed,sum(isnull(case when moneda = 'P' then round(total,2) when moneda = 'D' then round((total * " & TipoCambioDol & "),2) when moneda = 'E' then round((total/" & TipoCambioEuro & "),2) end,0)) as total " & "from notascreditocab (Nolock) where estatus = 'V' group by codprovacreed) n on p.codprovacreed = n.codprovacreed and a.codprovacreed = n.codprovacreed " & "left outer join (select codprovacreed,sum(isnull(case when moneda = 'P' then round(total,2) when moneda = 'D' then round((total * " & TipoCambioDol & "),2) when moneda = 'E' then round((total/" & TipoCambioEuro & "),2) end,0)) as total " & "from anticipos (Nolock) where estatus = 'V' group by codprovacreed) an on p.codprovacreed = an.codprovacreed and a.codprovacreed = an.codprovacreed"
        End Function

    Private Sub frmCXPPresupuestado_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCXPPresupuestado_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPPresupuestado_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "cmbMes" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmCXPPresupuestado_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPPresupuestado_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Nuevo()
    End Sub

    Private Sub frmCXPPresupuestado_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmCXPPresupuestado_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.cmbAño = New System.Windows.Forms.ComboBox()
        Me.cmbMes = New System.Windows.Forms.ComboBox()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(16, 86)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(339, 88)
        Me.txtMensaje.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cmbAño)
        Me.Frame1.Controls.Add(Me.cmbMes)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(16, 13)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(339, 52)
        Me.Frame1.TabIndex = 3
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo"
        '
        'cmbAño
        '
        Me.cmbAño.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAño.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAño.Location = New System.Drawing.Point(265, 19)
        Me.cmbAño.Name = "cmbAño"
        Me.cmbAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAño.Size = New System.Drawing.Size(68, 21)
        Me.cmbAño.TabIndex = 1
        '
        'cmbMes
        '
        Me.cmbMes.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMes.Items.AddRange(New Object() {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"})
        Me.cmbMes.Location = New System.Drawing.Point(12, 19)
        Me.cmbMes.Name = "cmbMes"
        Me.cmbMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMes.Size = New System.Drawing.Size(232, 21)
        Me.cmbMes.TabIndex = 0
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblRpt.SetIndex(Me._lblRpt_2, CType(2, Short))
        Me._lblRpt_2.Location = New System.Drawing.Point(16, 72)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 4
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(131, 192)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 136
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(16, 192)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 135
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(246, 193)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 134
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmCXPPresupuestado
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(377, 240)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmCXPPresupuestado"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de CxP Presupuestado"
        Me.Frame1.ResumeLayout(False)
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
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