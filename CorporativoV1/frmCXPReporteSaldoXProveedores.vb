Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCXPReporteSaldoXProveedores
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents cmbMes As System.Windows.Forms.ComboBox
    Public WithEvents cmbAño As System.Windows.Forms.ComboBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents chkTodosProveedores As System.Windows.Forms.CheckBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Dim mblnSalir As Boolean
    Dim mintCodProveedor As Integer
    Dim FueraChange As Boolean
    Dim Tecla As Integer
    Dim I As Integer
    Dim Fechas(2, 2) As String
    Dim OrdenMeses(11) As String
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim rsReporte As ADODB.Recordset

    Sub CalculaFechas()
        Dim Mes As Integer
        Dim Periodo As Integer
        Dim FechaInicial As String
        Dim FechaFinal As String
        Mes = cmbMes.SelectedIndex + 1
        Periodo = CInt(cmbAño.Text)
        For I = 0 To 11
            If I = 0 Or I = 11 Then
                ModCorporativo.ObtenerLimitedeFechas(Mes, Periodo, FechaInicial, FechaFinal)
                If I = 0 Then
                    Fechas(0, 0) = FechaInicial
                    Fechas(0, 1) = FechaFinal
                End If
                If I = 11 Then
                    Fechas(1, 0) = FechaInicial
                    Fechas(1, 1) = FechaFinal
                End If
            End If
            OrdenMeses(I) = CStr(Mes)
            If Mes = 12 Then
                Mes = 1
                Periodo = Periodo + 1
            Else
                Mes = Mes + 1
            End If
        Next
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
        mintCodProveedor = 0
        FueraChange = False
        Tecla = 0
    End Sub

    Sub Imprime()
        Dim rptCXPReporteSaldoXProveedor As New rptCXPReporteSaldoXProveedor
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Err_Renamed
        Dim NombreEmpresa As Object
        Dim NombreReporte As Object
        Dim Periodo As Object
        Dim TextoAdicional As Object
        Dim sql1 As Object
        Dim sql2 As String
        Dim Nota As Object
        Dim Moneda As String
        Dim aParam(12) As Object
        Dim aValues(12) As Object
        Dim TipoCambio As Decimal
        If Not ValidaDatos() Then Exit Sub
        CalculaFechas()
        TipoCambio = gcurCorpoTIPOCAMBIODOLAR
        If optPesos.Checked = True Then
            Moneda = "P"
        ElseIf optDolares.Checked = True Then
            Moneda = "D"
        End If
        sql1 = "select b.codprovacreed,b.descprovacreed,(isnull(a.ppmes1,0) - (isnull(p.pmes1,0) + isnull(n.ncmes1,0) + isnull(an.anmes1,0))) as importemes1,(isnull(a.ppmes2,0) - (isnull(p.pmes2,0) + isnull(n.ncmes2,0) + isnull(an.anmes2,0))) as importemes2,(isnull(a.ppmes3,0) - (isnull(p.pmes3,0) + isnull(n.ncmes3,0) + isnull(an.anmes3,0))) as importemes3,(isnull(a.ppmes4,0) - (isnull(p.pmes4,0) + isnull(n.ncmes4,0) + isnull(an.anmes4,0))) as importemes4,(isnull(a.ppmes5,0) - (isnull(p.pmes5,0) + isnull(n.ncmes5,0) + isnull(an.anmes5,0))) as importemes5,(isnull(a.ppmes6,0) - (isnull(p.pmes6,0) + isnull(n.ncmes6,0) + isnull(an.anmes6,0))) as importemes6," & "(isnull(a.ppmes7,0) - (isnull(p.pmes7,0) + isnull(n.ncmes7,0) + isnull(an.anmes7,0))) as importemes7,(isnull(a.ppmes8,0) - (isnull(p.pmes8,0) + isnull(n.ncmes8,0) + isnull(an.anmes8,0))) as importemes8,(isnull(a.ppmes9,0) - (isnull(p.pmes9,0) + isnull(n.ncmes9,0) + isnull(an.anmes9,0))) as importemes9,(isnull(a.ppmes10,0) - (isnull(p.pmes10,0) + isnull(n.ncmes10,0) + isnull(an.anmes10,0))) as importemes10,(isnull(a.ppmes11,0) - (isnull(p.pmes11,0) + isnull(n.ncmes11,0) + isnull(an.anmes11,0))) as importemes11,(isnull(a.ppmes12,0) - (isnull(p.pmes12,0) + isnull(n.ncmes12,0) + isnull(an.anmes12,0))) as importemes12 " & "from (select codprovacreed,sum(isnull(case when month(fechapago) = 1 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes1,sum(isnull(case when month(fechapago) = 2 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes2,sum(isnull(case when month(fechapago) = 3 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes3,sum(isnull(case when month(fechapago) = 4 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes4,sum(isnull(case when month(fechapago) = 5 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes5,sum(isnull(case when month(fechapago) = 6 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes6," & "sum(isnull(case when month(fechapago) = 7 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes7,sum(isnull(case when month(fechapago) = 8 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes8,sum(isnull(case when month(fechapago) = 9 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes9,sum(isnull(case when month(fechapago) = 10 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes10,sum(isnull(case when month(fechapago) = 11 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes11,sum(isnull(case when month(fechapago) = 12 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as ppmes12 from programacionpagos where fechapago between '" & Fechas(0, 0) & "' and '" & Fechas(1, 1) & "' and estatus <> 'C' group by codprovacreed) a " & "full join catprovacreed b on a.codprovacreed = b.codprovacreed full Join " & "(select codprovacreed,sum(isnull(case when month(fechapago) = 1 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes1,sum(isnull(case when month(fechapago) = 2 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes2,sum(isnull(case when month(fechapago) = 3 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes3,sum(isnull(case when month(fechapago) = 4 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes4,sum(isnull(case when month(fechapago) = 5 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes5,sum(isnull(case when month(fechapago) = 6 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes6," & "sum(isnull(case when month(fechapago) = 7 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes7,sum(isnull(case when month(fechapago) = 8 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes8,sum(isnull(case when month(fechapago) = 9 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes9,sum(isnull(case when month(fechapago) = 10 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes10,sum(isnull(case when month(fechapago) = 11 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes11,sum(isnull(case when month(fechapago) = 12 then dbo.convertircantidad(moneda,'" & Moneda & "',totalpago," & TipoCambio & ",12) end,0)) as pmes12 from pagos where fechapago between '" & Fechas(0, 0) & "' and '" & Fechas(1, 1) & "' and estatus <> 'C' group by codprovacreed) p on p.codprovacreed = b.codprovacreed full Join "
        sql2 = "(select codprovacreed,sum(isnull(case when month(fechanotacredito) = 1 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes1,sum(isnull(case when month(fechanotacredito) = 2 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes2,sum(isnull(case when month(fechanotacredito) = 3 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes3,sum(isnull(case when month(fechanotacredito) = 4 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes4,sum(isnull(case when month(fechanotacredito) = 5 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes5,sum(isnull(case when month(fechanotacredito) = 6 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes6," & "sum(isnull(case when month(fechanotacredito) = 7 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes7,sum(isnull(case when month(fechanotacredito) = 8 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes8,sum(isnull(case when month(fechanotacredito) = 9 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes9,sum(isnull(case when month(fechanotacredito) = 10 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes10,sum(isnull(case when month(fechanotacredito) = 11 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes11,sum(isnull(case when month(fechanotacredito) = 12 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as ncmes12 from notascreditocab where fechanotacredito between '" & Fechas(0, 0) & "' and '" & Fechas(1, 1) & "' and estatus = 'V' group by codprovacreed) n on b.codprovacreed = n.codprovacreed full Join " & "(select codprovacreed,sum(isnull(case when month(fechaanticipo) = 1 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes1,sum(isnull(case when month(fechaanticipo) = 2 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes2,sum(isnull(case when month(fechaanticipo) = 3 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes3,sum(isnull(case when month(fechaanticipo) = 4 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes4,sum(isnull(case when month(fechaanticipo) = 5 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes5,sum(isnull(case when month(fechaanticipo) = 6 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes6," & "sum(isnull(case when month(fechaanticipo) = 7 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes7,sum(isnull(case when month(fechaanticipo) = 8 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes8,sum(isnull(case when month(fechaanticipo) = 9 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes9,sum(isnull(case when month(fechaanticipo) = 10 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes10,sum(isnull(case when month(fechaanticipo) = 11 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes11,sum(isnull(case when month(fechaanticipo) = 12 then dbo.convertircantidad(moneda,'" & Moneda & "',total," & TipoCambio & ",12) end,0)) as anmes12 from anticipos where fechaanticipo between '" & Fechas(0, 0) & "' and '" & Fechas(1, 1) & "' and estatus = 'V' group by codprovacreed) an on b.codprovacreed = an.codprovacreed " & "where b.tipo = 'P' " & IIf(mintCodProveedor <> 0, "and b.codprovacreed = " & mintCodProveedor & " ", "") & "and ((isnull(a.ppmes1,0) - (isnull(p.pmes1,0) + isnull(n.ncmes1,0) + isnull(an.anmes1,0))) + (isnull(a.ppmes2,0) - (isnull(p.pmes2,0) + isnull(n.ncmes2,0) + isnull(an.anmes2,0))) + (isnull(a.ppmes3,0) - (isnull(p.pmes3,0) + isnull(n.ncmes3,0) + isnull(an.anmes3,0))) + (isnull(a.ppmes4,0) - (isnull(p.pmes4,0) + isnull(n.ncmes4,0) + isnull(an.anmes4,0))) + (isnull(a.ppmes5,0) - (isnull(p.pmes5,0) + isnull(n.ncmes5,0) + isnull(an.anmes5,0))) + (isnull(a.ppmes6,0) - (isnull(p.pmes6,0) + isnull(n.ncmes6,0) + isnull(an.anmes6,0))) + " & "(isnull(a.ppmes7,0) - (isnull(p.pmes7,0) + isnull(n.ncmes7,0) + isnull(an.anmes7,0))) + (isnull(a.ppmes8,0) - (isnull(p.pmes8,0) + isnull(n.ncmes8,0) + isnull(an.anmes8,0))) + (isnull(a.ppmes9,0) - (isnull(p.pmes9,0) + isnull(n.ncmes9,0) + isnull(an.anmes9,0))) + (isnull(a.ppmes10,0) - (isnull(p.pmes10,0) + isnull(n.ncmes10,0) + isnull(an.anmes10,0))) + (isnull(a.ppmes11,0) - (isnull(p.pmes11,0) + isnull(n.ncmes11,0) + isnull(an.anmes11,0))) + (isnull(a.ppmes12,0) - (isnull(p.pmes12,0) + isnull(n.ncmes12,0) + isnull(an.anmes12,0)))) <> 0"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_DatosSql"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, sql1))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, sql2))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existen datos para el periodo indicado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            rptCXPReporteSaldoXProveedor.SetDataSource(frmReportes.rsReport)
        End If


        If optPesos.Checked Then
            Nota = "Importes expresados en pesos"
        ElseIf optDolares.Checked Then
            Nota = "Importes expresados en dólares"
        End If
        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = UCase("Reporte de Saldos por Proveedores")
        Periodo = "Del  " & VB6.Format(Fechas(0, 0), "DD/MM/YYYY") & "  al  " & VB6.Format(Fechas(1, 0), "DD/MM/YYYY")
        TextoAdicional = txtMensaje.Text
        'frmReportes.Report = rptCXPReporteSaldoXProveedor
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.rsReport = rsReporte
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "Periodo", "TextoAdicional", "Nota", "Importe1", "Importe2", "Importe3", "Importe4", "Importe5", "Importe6", "Importe7", "Importe8", "Importe9", "Importe10", "Importe11", "Importe12"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, Periodo, TextoAdicional, Nota, OrdenMeses(0), OrdenMeses(1), OrdenMeses(2), OrdenMeses(3), OrdenMeses(4), OrdenMeses(5), OrdenMeses(6), OrdenMeses(7), OrdenMeses(8), OrdenMeses(9), OrdenMeses(10), OrdenMeses(11)}

        'If (NombreEmpresa <> Nothing) Then
        '    pdvNum.Value = NombreEmpresa : pvNum.Add(pdvNum)
        '    rptCXPReporteSaldoXProveedor.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
        'Else
        '    pdvNum.Value = "" : pvNum.Add(pdvNum)
        '    rptCXPReporteSaldoXProveedor.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
        'End If

        'If (NombreReporte <> Nothing) Then
        '    pdvNum.Value = NombreReporte : pvNum.Add(pdvNum)
        '    rptCXPReporteSaldoXProveedor.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
        'Else
        '    pdvNum.Value = "" : pvNum.Add(pdvNum)
        '    rptCXPReporteSaldoXProveedor.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
        'End If

        'If (Periodo <> Nothing) Then
        '    pdvNum.Value = Periodo : pvNum.Add(pdvNum)
        '    rptCXPReporteSaldoXProveedor.DataDefinition.ParameterFields("Periodo").ApplyCurrentValues(pvNum)
        'Else
        '    pdvNum.Value = "" : pvNum.Add(pdvNum)
        '    rptCXPReporteSaldoXProveedor.DataDefinition.ParameterFields("Periodo").ApplyCurrentValues(pvNum)
        'End If

        'If (TextoAdicional <> Nothing) Then
        '    pdvNum.Value = TextoAdicional : pvNum.Add(pdvNum)
        '    rptCXPReporteSaldoXProveedor.DataDefinition.ParameterFields("TextoAdicional").ApplyCurrentValues(pvNum)
        'Else
        '    pdvNum.Value = "" : pvNum.Add(pdvNum)
        '    rptCXPReporteSaldoXProveedor.DataDefinition.ParameterFields("TextoAdicional").ApplyCurrentValues(pvNum)
        'End If

        'If (Nota <> Nothing) Then
        '    pdvNum.Value = Nota : pvNum.Add(pdvNum)
        '    rptCXPReporteSaldoXProveedor.DataDefinition.ParameterFields("Nota").ApplyCurrentValues(pvNum)
        'Else
        '    pdvNum.Value = "" : pvNum.Add(pdvNum)
        '    rptCXPReporteSaldoXProveedor.DataDefinition.ParameterFields("Nota").ApplyCurrentValues(pvNum)
        'End If



        frmReportes.Text = "Reporte de Saldos por Proveedores"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.reporteActual = rptCXPReporteSaldoXProveedor
        frmReportes.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Default

Err_Renamed:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Limpiar()
        Nuevo()
        chkTodosProveedores.Focus()
    End Sub

    Sub LlenaAños()
        cmbAño.Items.Clear()
        For I = 1900 To 2075
            cmbAño.Items.Add(CStr(I))
        Next
    End Sub

    Sub Nuevo()
        LlenaAños()
        chkTodosProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked
        FueraChange = True
        dbcProveedor.Text = ""
        FueraChange = False
        cmbMes.SelectedIndex = Month(Today) - 1
        cmbAño.Text = CStr(Year(Today))
        optPesos.Checked = True
        txtMensaje.Text = ""
        InicializaVariables()
    End Sub

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If chkTodosProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked And Trim(dbcProveedor.Text) = "" Then
            MsgBox("Proporcione el proveedor que desea consultar, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub chkTodosProveedores_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodosProveedores.CheckStateChanged
        If chkTodosProveedores.CheckState = System.Windows.Forms.CheckState.Checked Then
            FueraChange = True
            dbcProveedor.Text = ""
            FueraChange = False
            mintCodProveedor = 0
            dbcProveedor.Enabled = False
        ElseIf chkTodosProveedores.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            dbcProveedor.Enabled = True
        End If
    End Sub

    Private Sub chkTodosProveedores_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodosProveedores.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbAño_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAño.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbMes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMes.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo Merr
        If FueraChange Then Exit Sub
        gStrSql = "SELECT CodProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM CatProvAcreed Where DescProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%' AND Tipo = 'P'"
        ModDCombo.DCChange(gStrSql, Tecla, (Me.dbcProveedor))
        If Trim(Me.dbcProveedor.Text) = "" Then
            mintCodProveedor = 0
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        Pon_Tool()
        gStrSql = "SELECT CodProvAcreed, LTrim(RTrim(DescProvAcreed)) as DescProvAcreed FROM CatProvAcreed WHERE Tipo = 'P'"
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcProveedor))
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcProveedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodosProveedores.Focus()
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
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodProvAcreed,LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM CatProvAcreed Where DescProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%' AND Tipo = 'P'"
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

    Private Sub frmCXPReporteSaldoXProveedores_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCXPReporteSaldoXProveedores_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPReporteSaldoXProveedores_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "chkTodosProveedores" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmCXPReporteSaldoXProveedores_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPReporteSaldoXProveedores_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Nuevo()
    End Sub

    Private Sub frmCXPReporteSaldoXProveedores_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmCXPReporteSaldoXProveedores_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub optPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPesos.Enter
        Pon_Tool()
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.cmbMes = New System.Windows.Forms.ComboBox()
        Me.cmbAño = New System.Windows.Forms.ComboBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.chkTodosProveedores = New System.Windows.Forms.CheckBox()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
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
        Me.txtMensaje.Location = New System.Drawing.Point(16, 208)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(339, 80)
        Me.txtMensaje.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.optDolares)
        Me.Frame3.Controls.Add(Me.optPesos)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(16, 139)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(339, 50)
        Me.Frame3.TabIndex = 11
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Moneda del Reporte"
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(166, 19)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(138, 21)
        Me.optDolares.TabIndex = 5
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Presentar en Dolares"
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Checked = True
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(32, 19)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(118, 21)
        Me.optPesos.TabIndex = 4
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Presentar en Pesos"
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.cmbMes)
        Me.Frame2.Controls.Add(Me.cmbAño)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(16, 82)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(339, 52)
        Me.Frame2.TabIndex = 9
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Periodo"
        '
        'cmbMes
        '
        Me.cmbMes.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMes.Items.AddRange(New Object() {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"})
        Me.cmbMes.Location = New System.Drawing.Point(14, 19)
        Me.cmbMes.Name = "cmbMes"
        Me.cmbMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMes.Size = New System.Drawing.Size(206, 21)
        Me.cmbMes.TabIndex = 2
        '
        'cmbAño
        '
        Me.cmbAño.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAño.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAño.Location = New System.Drawing.Point(235, 19)
        Me.cmbAño.Name = "cmbAño"
        Me.cmbAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAño.Size = New System.Drawing.Size(68, 21)
        Me.cmbAño.TabIndex = 3
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.chkTodosProveedores)
        Me.Frame1.Controls.Add(Me.dbcProveedor)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(16, 6)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(339, 71)
        Me.Frame1.TabIndex = 7
        Me.Frame1.TabStop = False
        '
        'chkTodosProveedores
        '
        Me.chkTodosProveedores.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodosProveedores.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodosProveedores.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodosProveedores.Location = New System.Drawing.Point(14, 13)
        Me.chkTodosProveedores.Name = "chkTodosProveedores"
        Me.chkTodosProveedores.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodosProveedores.Size = New System.Drawing.Size(172, 21)
        Me.chkTodosProveedores.TabIndex = 0
        Me.chkTodosProveedores.Text = "Todos los Proveedores"
        Me.chkTodosProveedores.UseVisualStyleBackColor = False
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(104, 36)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(206, 21)
        Me.dbcProveedor.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(32, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(66, 18)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Proveedor :"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblRpt.SetIndex(Me._lblRpt_2, CType(2, Short))
        Me._lblRpt_2.Location = New System.Drawing.Point(16, 194)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 10
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(131, 304)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 139
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(16, 304)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 138
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(246, 305)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 137
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmCXPReporteSaldoXProveedores
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(374, 351)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmCXPReporteSaldoXProveedores"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Saldos por Proveedores"
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
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