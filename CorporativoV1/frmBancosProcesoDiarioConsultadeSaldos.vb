Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosProcesoDiarioConsultadeSaldos
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             CONSULTA DE SALDOS                                                                           *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      VIERNES 01 DE AGOSTO DE 2003                                                                 *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdConsultaPtoVta As System.Windows.Forms.Button
    Public WithEvents Line1 As System.Windows.Forms.Label
    Public WithEvents lblTotal As System.Windows.Forms.Label
    Public WithEvents lblDolares As System.Windows.Forms.Label
    Public WithEvents lblPesos As System.Windows.Forms.Label
    Public WithEvents lblTarjetas As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Line2 As System.Windows.Forms.Label
    'Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents lblTotalDolaresSaldoActual As System.Windows.Forms.Label
    Public WithEvents lblTotalDolaresRetiros As System.Windows.Forms.Label
    Public WithEvents lblTotalDolaresDepositos As System.Windows.Forms.Label
    Public WithEvents lblTotalDolaresalDiaAnterior As System.Windows.Forms.Label
    Public WithEvents lblTotalPesosSaldoActual As System.Windows.Forms.Label
    Public WithEvents lblTotalPesosRetiros As System.Windows.Forms.Label
    Public WithEvents lblTotalPesosDepositos As System.Windows.Forms.Label
    Public WithEvents lblTotalPesosalDiaAnterior As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    'Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    'Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaCorte As System.Windows.Forms.DateTimePicker
    Public WithEvents chkDolares As System.Windows.Forms.CheckBox
    Public WithEvents chkPesos As System.Windows.Forms.CheckBox
    'Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dbcBanco As System.Windows.Forms.ComboBox
    Public WithEvents chkTodoslosBancos As System.Windows.Forms.CheckBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    'Public WithEvents Frame1 As System.Windows.Forms.GroupBox

    Dim mblnSalir As Boolean 'Para Salir Con el Esc
    Dim tecla As Integer
    Dim intCodBanco As Integer
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label9 As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Label10 As Label
    Friend WithEvents Panel5 As Panel
    Friend WithEvents Label11 As Label
    Friend WithEvents Label12 As Label
    Dim FueraChange As Boolean

    Sub CalculaTotales()
        Dim I As Integer
        lblTotalDolaresalDiaAnterior.Text = "0.00"
        lblTotalDolaresDepositos.Text = "0.00"
        lblTotalDolaresRetiros.Text = "0.00"
        lblTotalDolaresSaldoActual.Text = "0.00"
        lblTotalPesosalDiaAnterior.Text = "0.00"
        lblTotalPesosDepositos.Text = "0.00"
        lblTotalPesosRetiros.Text = "0.00"
        lblTotalPesosSaldoActual.Text = "0.00"
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 6)) = C_PESO Then
                    lblTotalPesosalDiaAnterior.Text = CStr(CDbl(Numerico(lblTotalPesosalDiaAnterior.Text)) + CDbl(Numerico(.get_TextMatrix(I, 2))))
                    lblTotalPesosDepositos.Text = CStr(CDbl(Numerico(lblTotalPesosDepositos.Text)) + CDbl(Numerico(.get_TextMatrix(I, 3))))
                    lblTotalPesosRetiros.Text = CStr(CDbl(Numerico(lblTotalPesosRetiros.Text)) + CDbl(Numerico(.get_TextMatrix(I, 4))))
                    lblTotalPesosSaldoActual.Text = CStr(CDbl(Numerico(lblTotalPesosSaldoActual.Text)) + CDbl(Numerico(.get_TextMatrix(I, 5))))
                ElseIf Trim(.get_TextMatrix(I, 6)) = C_DOLAR Then
                    lblTotalDolaresalDiaAnterior.Text = CStr(CDbl(Numerico(lblTotalDolaresalDiaAnterior.Text)) + CDbl(Numerico(.get_TextMatrix(I, 2))))
                    lblTotalDolaresDepositos.Text = CStr(CDbl(Numerico(lblTotalDolaresDepositos.Text)) + CDbl(Numerico(.get_TextMatrix(I, 3))))
                    lblTotalDolaresRetiros.Text = CStr(CDbl(Numerico(lblTotalDolaresRetiros.Text)) + CDbl(Numerico(.get_TextMatrix(I, 4))))
                    lblTotalDolaresSaldoActual.Text = CStr(CDbl(Numerico(lblTotalDolaresSaldoActual.Text)) + CDbl(Numerico(.get_TextMatrix(I, 5))))
                End If
            Next
            lblTotalDolaresalDiaAnterior.Text = VB6.Format(lblTotalDolaresalDiaAnterior.Text, "###,##0.00")
            lblTotalDolaresDepositos.Text = VB6.Format(lblTotalDolaresDepositos.Text, "###,##0.00")
            lblTotalDolaresRetiros.Text = VB6.Format(lblTotalDolaresRetiros.Text, "###,##0.00")
            lblTotalDolaresSaldoActual.Text = VB6.Format(lblTotalDolaresSaldoActual.Text, "###,##0.00")
            lblTotalPesosalDiaAnterior.Text = VB6.Format(lblTotalPesosalDiaAnterior.Text, "###,##0.00")
            lblTotalPesosDepositos.Text = VB6.Format(lblTotalPesosDepositos.Text, "###,##0.00")
            lblTotalPesosRetiros.Text = VB6.Format(lblTotalPesosRetiros.Text, "###,##0.00")
            lblTotalPesosSaldoActual.Text = VB6.Format(lblTotalPesosSaldoActual.Text, "###,##0.00")
        End With
    End Sub

    Sub ConfiguraGrid()
        Dim I As Integer
        With flexDetalle
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 1800)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Cuenta"
            .Col = 1
            .set_ColWidth(1, 0, 2200)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Banco"
            .Col = 2
            .set_ColWidth(2, 0, 1530)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Al día Anterior"
            .Col = 3
            .set_ColWidth(3, 0, 1530)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Depósitos"
            .Col = 4
            .set_ColWidth(4, 0, 1530)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Retiros"
            .Col = 5
            .set_ColWidth(5, 0, 1530)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Saldo Actual"
            .Col = 6
            .set_ColWidth(6, 0, 0)
            For I = 1 To .Rows - 1
                .set_TextMatrix(I, 0, VB6.Format(.get_TextMatrix(I, 0), "0000-0000-0000-0000"))
                .set_TextMatrix(I, 2, VB6.Format(.get_TextMatrix(I, 2), "###,##0.00"))
                .set_TextMatrix(I, 3, VB6.Format(.get_TextMatrix(I, 3), "###,##0.00"))
                .set_TextMatrix(I, 4, VB6.Format(.get_TextMatrix(I, 4), "###,##0.00"))
                .set_TextMatrix(I, 5, VB6.Format(.get_TextMatrix(I, 5), "###,##0.00"))
            Next
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub Encabezado()
        Dim I As Integer
        With flexDetalle
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 1800)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Cuenta"
            .Col = 1
            .set_ColWidth(1, 0, 2200)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Banco"
            .Col = 2
            .set_ColWidth(2, 0, 1530)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Al día Anterior"
            .Col = 3
            .set_ColWidth(3, 0, 1530)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Depósitos"
            .Col = 4
            .set_ColWidth(4, 0, 1530)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Retiros"
            .Col = 5
            .set_ColWidth(5, 0, 1530)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Saldo Actual"
            .Rows = 10
            .Row = 1
            .Col = 0
        End With
    End Sub

    Sub Limpiar()
        Nuevo()
        chkTodoslosBancos.Focus()
    End Sub

    Sub Nuevo()
        chkTodoslosBancos.CheckState = System.Windows.Forms.CheckState.Unchecked
        dbcBanco.Enabled = True
        dbcBanco.Text = ""
        dbcBanco.Text = Nothing
        chkPesos.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkDolares.CheckState = System.Windows.Forms.CheckState.Unchecked
        dtpFechaCorte.Value = VB6.Format(Now, "dd/mmm/yyyy")
        flexDetalle.Clear()
        Encabezado()
        lblTotalDolaresalDiaAnterior.Text = "0.00"
        lblTotalDolaresDepositos.Text = "0.00"
        lblTotalDolaresRetiros.Text = "0.00"
        lblTotalDolaresSaldoActual.Text = "0.00"
        lblTotalPesosalDiaAnterior.Text = "0.00"
        lblTotalPesosDepositos.Text = "0.00"
        lblTotalPesosRetiros.Text = "0.00"
        lblTotalPesosSaldoActual.Text = "0.00"
    End Sub

    Sub ObtenerIngresos()
        On Error GoTo Err_Renamed
        gStrSql = "Select sum(IsNull(Efe.EfectivoDolares,0)) as EfectivoDolares,sum(IsNull(Efe.EfectivoPesos,0)) as EfectivoPesos,sum(IsNull(TC.ImporteTarjetas,0)) as ImporteTarjetas " & "From (SELECT FechaIngreso,CodSucursal,Sum(IsNull(CASE WHEN Tipo = 'I' THEN ImporteDolares END,0) - " & "(ABS(IsNull(CASE WHEN TIPO = 'D' THEN ImporteDolares END,0)) + IsNull(CASE WHEN TIPO = 'R' THEN ImporteDolares END,0))) AS EfectivoDolares, " & "Sum(IsNull(CASE WHEN Tipo = 'I' THEN ImportePesos END,0) - (ABS(IsNull(CASE WHEN TIPO = 'D' THEN ImportePesos END,0)) + IsNull(CASE WHEN TIPO = 'R' THEN ImportePesos END,0))) AS EfectivoPesos " & "From DBO.vw_Ingresos (Nolock) WHERE FechaIngreso <= '" & VB6.Format(dtpFechaCorte.Value, "MM/DD/YYYY") & "' Group by FechaIngreso, CodSucursal ) EFE Left outer Join ( " & "SELECT   FechaIngreso, CodSucursal, Sum(ABS(IsNull(CASE WHEN TIPO = 'T' THEN ImportePesos END,0))) AS ImporteTarjetas " & "From DBO.vw_Ingresos WHERE FechaIngreso <= '" & VB6.Format(dtpFechaCorte.Value, "MM/DD/YYYY") & "' Group by FechaIngreso, CodSucursal " & ") TC On Efe.FEchaIngreso = TC.FEchaIngreso And Efe.CodSucursal = TC.CodSucursal "

        '''gStrSql = "SELECT ISNULL((ABS(SUM(CASE WHEN Tipo = 'I' THEN ImporteDolares END)) - " & _
        '"ABS(SUM(CASE WHEN TIPO = 'D' THEN ImporteDolares END)) - " & _
        '"ABS(SUM(CASE WHEN TIPO = 'R' THEN ImporteDolares END))),0) AS EfectivoDolares," & _
        '"ISNULL((ABS(SUM(CASE WHEN Tipo = 'I' THEN ImportePesos END)) - " & _
        '"ABS(SUM(CASE WHEN TIPO = 'D' THEN ImportePesos END)) - " & _
        '"ABS(SUM(CASE WHEN TIPO = 'R' THEN ImportePesos END))),0) AS EfectivoPesos," & _
        '"ISNULL(ABS(SUM(CASE WHEN TIPO = 'T' THEN ImportePesos END)),0) AS ImporteTarjetas " & _
        '"FROM DBO.vw_ObtenerIngresos " & _
        '"WHERE FechaIngreso <= '" & Format(Date, "MM/DD/YYYY") & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        lblPesos.Text = CStr(System.Math.Round(RsGral.Fields("EfectivoPesos").Value, 2))
        lblDolares.Text = CStr(System.Math.Round(RsGral.Fields("EfectivoDolares").Value * gcurCorpoTIPOCAMBIODOLAR, 2))
        lblTarjetas.Text = CStr(System.Math.Round(RsGral.Fields("ImporteTarjetas").Value, 2))
        lblTotal.Text = VB6.Format(CDbl(Numerico(lblPesos.Text)) + CDbl(Numerico(lblDolares.Text)) + CDbl(Numerico(lblTarjetas.Text)), "###,##0.00")
        lblPesos.Text = VB6.Format(lblPesos.Text, "###,##0.00")
        lblDolares.Text = VB6.Format(lblDolares.Text, "###,##0.00")
        lblTarjetas.Text = VB6.Format(lblTarjetas.Text, "###,##0.00")

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Resultado()
        On Error GoTo Merr
        Dim strFecha As String
        strFecha = VB6.Format(Year(dtpFechaCorte.Value), "0000") & "-" & VB6.Format(Month(dtpFechaCorte.Value), "00") & "-" & VB6.Format((dtpFechaCorte.Value).Day, "00")
        If chkTodoslosBancos.CheckState = 1 Then
            If chkPesos.CheckState = 1 Or chkDolares.CheckState = 1 Then
                If chkPesos.CheckState = 1 And chkDolares.CheckState = 1 Then
                    gStrSql = "SELECT T1.CtaBancaria AS Cuenta ,T2.DescBanco AS Banco,T1.SALDO AS 'Saldo al Dia Anterior'," & "T2.Ingresos AS 'Depósitos',T2.Egresos AS 'Retiros',(T1.Saldo + (T2.Ingresos - T2.Egresos)) AS 'Saldo Actual' ,T1.Moneda " & "FROM (SELECT CB.CtaBancaria,CB.Moneda,(CB.SaldoInicial + (ISNULL(SUM(CASE MB.TipoMovto WHEN 'I' THEN MB.Importe END),0))" & " - (ISNULL(SUM(CASE MB.TipoMovto WHEN 'E' THEN MB.Importe END),0))) AS SALDO,CB.Codbanco " & "FROM CatCuentasBancarias CB LEFT OUTER JOIN (SELECT * FROM MovimientosBancarios WHERE FechaMovto < '" & Trim(strFecha) & "') MB ON " & "CB.CtaBancaria = MB.CtaBancaria and CB.CodBanco = MB.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias) " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,CB.CodBanco,CB.Moneda) T1 " & "INNER JOIN " & "(SELECT CB.CtaBancaria,SUM(ISNULL(CASE MB.TipoMovto WHEN 'I' THEN (MB.Importe) END,0)) AS Ingresos," & "SUM(ISNULL(CASE MB.TipoMovto WHEN 'E' THEN (MB.Importe) END,0)) AS Egresos," & "B.DescBanco,B.CodBanco " & "FROM CatCuentasBancarias CB LEFT OUTER JOIN MovimientosBancarios MB ON CB.CodBanco = MB.Codbanco AND CB.CtaBancaria = MB.CtaBancaria AND MB.FechaMovto = '" & Trim(strFecha) & "' " & "INNER JOIN CatBancos B ON B.CodBanco = CB.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias) " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,B.DescBanco,B.CodBanco) " & "T2 ON T1.CtaBancaria = T2.CtaBancaria AND T1.CodBanco = T2.CodBanco " & "GROUP BY T1.CtaBancaria,T2.DescBanco,T1.Saldo,T2.Ingresos,T2.Egresos,T1.Moneda " & "ORDER BY T2.DescBanco"
                ElseIf chkPesos.CheckState = 1 And chkDolares.CheckState = 0 Then
                    gStrSql = "SELECT T1.CtaBancaria AS Cuenta ,T2.DescBanco AS Banco,T1.SALDO AS 'Saldo al Dia Anterior'," & "T2.Ingresos AS 'Depósitos',T2.Egresos AS 'Retiros',(T1.Saldo + (T2.Ingresos - T2.Egresos)) AS 'Saldo Actual' ,T1.Moneda " & "FROM (SELECT CB.CtaBancaria,CB.Moneda,(CB.SaldoInicial + (ISNULL(SUM(CASE MB.TipoMovto WHEN 'I' THEN MB.Importe END),0))" & " - (ISNULL(SUM(CASE MB.TipoMovto WHEN 'E' THEN MB.Importe END),0))) AS SALDO,CB.Codbanco " & "FROM CatCuentasBancarias CB LEFT OUTER JOIN (SELECT * FROM MovimientosBancarios WHERE FechaMovto < '" & Trim(strFecha) & "') MB ON " & "CB.CtaBancaria = MB.CtaBancaria and CB.CodBanco = MB.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias) " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,CB.CodBanco,CB.Moneda) T1 " & "INNER JOIN " & "(SELECT CB.CtaBancaria,SUM(ISNULL(CASE MB.TipoMovto WHEN 'I' THEN (MB.Importe) END,0)) AS Ingresos," & "SUM(ISNULL(CASE MB.TipoMovto WHEN 'E' THEN (MB.Importe) END,0)) AS Egresos," & "B.DescBanco,B.CodBanco " & "FROM CatCuentasBancarias CB LEFT OUTER JOIN MovimientosBancarios MB ON CB.CodBanco = MB.Codbanco AND CB.CtaBancaria = MB.CtaBancaria AND MB.FechaMovto = '" & Trim(strFecha) & "' " & "INNER JOIN CatBancos B ON B.CodBanco = CB.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias) " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,B.DescBanco,B.CodBanco) " & "T2 ON T1.CtaBancaria = T2.CtaBancaria AND T1.CodBanco = T2.CodBanco " & "WHERE T1.Moneda = '" & C_PESO & "' " & "GROUP BY T1.CtaBancaria,T2.DescBanco,T1.Saldo,T2.Ingresos,T2.Egresos,T1.Moneda " & "ORDER BY T2.DescBanco"
                ElseIf chkPesos.CheckState = 0 And chkDolares.CheckState = 1 Then
                    gStrSql = "SELECT T1.CtaBancaria AS Cuenta ,T2.DescBanco AS Banco,T1.SALDO AS 'Saldo al Dia Anterior'," & "T2.Ingresos AS 'Depósitos',T2.Egresos AS 'Retiros',(T1.Saldo + (T2.Ingresos - T2.Egresos)) AS 'Saldo Actual' ,T1.Moneda " & "FROM (SELECT CB.CtaBancaria,CB.Moneda,(CB.SaldoInicial + (ISNULL(SUM(CASE MB.TipoMovto WHEN 'I' THEN MB.Importe END),0))" & " - (ISNULL(SUM(CASE MB.TipoMovto WHEN 'E' THEN MB.Importe END),0))) AS SALDO,CB.Codbanco " & "FROM CatCuentasBancarias CB LEFT OUTER JOIN (SELECT * FROM MovimientosBancarios WHERE FechaMovto < '" & Trim(strFecha) & "') MB ON " & "CB.CtaBancaria = MB.CtaBancaria and CB.CodBanco = MB.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias) " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,CB.CodBanco,CB.Moneda) T1 " & "INNER JOIN " & "(SELECT CB.CtaBancaria,SUM(ISNULL(CASE MB.TipoMovto WHEN 'I' THEN (MB.Importe) END,0)) AS Ingresos," & "SUM(ISNULL(CASE MB.TipoMovto WHEN 'E' THEN (MB.Importe) END,0)) AS Egresos," & "B.DescBanco,B.CodBanco " & "FROM CatCuentasBancarias CB LEFT OUTER JOIN MovimientosBancarios MB ON CB.CodBanco = MB.Codbanco AND CB.CtaBancaria = MB.CtaBancaria AND MB.FechaMovto = '" & Trim(strFecha) & "' " & "INNER JOIN CatBancos B ON B.CodBanco = CB.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias) " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,B.DescBanco,B.CodBanco) " & "T2 ON T1.CtaBancaria = T2.CtaBancaria AND T1.CodBanco = T2.CodBanco " & "WHERE T1.Moneda = '" & C_DOLAR & "' " & "GROUP BY T1.CtaBancaria,T2.DescBanco,T1.Saldo,T2.Ingresos,T2.Egresos,T1.Moneda " & "ORDER BY T2.DescBanco"
                End If
            Else
                flexDetalle.Clear()
                Encabezado()
                lblTotalDolaresalDiaAnterior.Text = "0.00"
                lblTotalDolaresDepositos.Text = "0.00"
                lblTotalDolaresRetiros.Text = "0.00"
                lblTotalDolaresSaldoActual.Text = "0.00"
                lblTotalPesosalDiaAnterior.Text = "0.00"
                lblTotalPesosDepositos.Text = "0.00"
                lblTotalPesosRetiros.Text = "0.00"
                lblTotalPesosSaldoActual.Text = "0.00"
                Exit Sub
            End If
        ElseIf chkTodoslosBancos.CheckState = 0 Then
            If Trim(dbcBanco.Text) = "" Then
                flexDetalle.Clear()
                Encabezado()
                lblTotalDolaresalDiaAnterior.Text = "0.00"
                lblTotalDolaresDepositos.Text = "0.00"
                lblTotalDolaresRetiros.Text = "0.00"
                lblTotalDolaresSaldoActual.Text = "0.00"
                lblTotalPesosalDiaAnterior.Text = "0.00"
                lblTotalPesosDepositos.Text = "0.00"
                lblTotalPesosRetiros.Text = "0.00"
                lblTotalPesosSaldoActual.Text = "0.00"
                Exit Sub
            Else
                If chkPesos.CheckState = 1 Or chkDolares.CheckState = 1 Then
                    If chkPesos.CheckState = 1 And chkDolares.CheckState = 1 Then
                        gStrSql = "SELECT T1.CtaBancaria AS Cuenta ,T2.DescBanco AS Banco,T1.SALDO AS 'Saldo al Dia Anterior'," & "T2.Ingresos AS 'Depósitos',T2.Egresos AS 'Retiros',(T1.Saldo + (T2.Ingresos - T2.Egresos)) AS 'Saldo Actual' ,T1.Moneda " & "FROM (SELECT CB.CtaBancaria,CB.Moneda,(CB.SaldoInicial + (ISNULL(SUM(CASE MB.TipoMovto WHEN 'I' THEN MB.Importe END),0))" & " - (ISNULL(SUM(CASE MB.TipoMovto WHEN 'E' THEN MB.Importe END),0))) AS SALDO,CB.Codbanco " & "FROM (SELECT * FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & ") CB LEFT OUTER JOIN (SELECT * FROM MovimientosBancarios WHERE FechaMovto < '" & Trim(strFecha) & "') MB ON " & "CB.CtaBancaria = MB.CtaBancaria and CB.CodBanco = MB.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & " ) " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,CB.CodBanco,CB.Moneda) T1 " & "INNER JOIN " & "(SELECT CB.CtaBancaria,SUM(ISNULL(CASE MB.TipoMovto WHEN 'I' THEN (MB.Importe) END,0)) AS Ingresos," & "SUM(ISNULL(CASE MB.TipoMovto WHEN 'E' THEN (MB.Importe) END,0)) AS Egresos," & "B.DescBanco,B.CodBanco " & "FROM (SELECT * FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & ") CB LEFT OUTER JOIN MovimientosBancarios MB ON CB.CodBanco = MB.Codbanco AND CB.CtaBancaria = MB.CtaBancaria AND MB.FechaMovto = '" & Trim(strFecha) & "' " & "INNER JOIN (SELECT * FROM CatBancos WHERE CodBanco = " & intCodBanco & ") B ON CB.CodBanco = B.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & ") " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,B.DescBanco,B.CodBanco) " & "T2 ON T1.CtaBancaria = T2.CtaBancaria AND T1.CodBanco = T2.CodBanco " & "GROUP BY T1.CtaBancaria,T2.DescBanco,T1.Saldo,T2.Ingresos,T2.Egresos,T1.Moneda " & "ORDER BY T2.DescBanco"
                    ElseIf chkPesos.CheckState = 1 And chkDolares.CheckState = 0 Then
                        gStrSql = "SELECT T1.CtaBancaria AS Cuenta ,T2.DescBanco AS Banco,T1.SALDO AS 'Saldo al Dia Anterior'," & "T2.Ingresos AS 'Depósitos',T2.Egresos AS 'Retiros',(T1.Saldo + (T2.Ingresos - T2.Egresos)) AS 'Saldo Actual' ,T1.Moneda " & "FROM (SELECT CB.CtaBancaria,CB.Moneda,(CB.SaldoInicial + (ISNULL(SUM(CASE MB.TipoMovto WHEN 'I' THEN MB.Importe END),0))" & " - (ISNULL(SUM(CASE MB.TipoMovto WHEN 'E' THEN MB.Importe END),0))) AS SALDO,CB.Codbanco " & "FROM (SELECT * FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & ") CB LEFT OUTER JOIN (SELECT * FROM MovimientosBancarios WHERE FechaMovto < '" & Trim(strFecha) & "') MB ON " & "CB.CtaBancaria = MB.CtaBancaria and CB.CodBanco = MB.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & " ) " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,CB.CodBanco,CB.Moneda) T1 " & "INNER JOIN " & "(SELECT CB.CtaBancaria,SUM(ISNULL(CASE MB.TipoMovto WHEN 'I' THEN (MB.Importe) END,0)) AS Ingresos," & "SUM(ISNULL(CASE MB.TipoMovto WHEN 'E' THEN (MB.Importe) END,0)) AS Egresos," & "B.DescBanco,B.CodBanco " & "FROM (SELECT * FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & ") CB LEFT OUTER JOIN MovimientosBancarios MB ON CB.CodBanco = MB.Codbanco AND CB.CtaBancaria = MB.CtaBancaria AND MB.FechaMovto = '" & Trim(strFecha) & "' " & "INNER JOIN (SELECT * FROM CatBancos WHERE CodBanco = " & intCodBanco & ") B ON CB.CodBanco = B.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & ") " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,B.DescBanco,B.CodBanco) " & "T2 ON T1.CtaBancaria = T2.CtaBancaria AND T1.CodBanco = T2.CodBanco " & "WHERE T1.Moneda = '" & C_PESO & "' " & "GROUP BY T1.CtaBancaria,T2.DescBanco,T1.Saldo,T2.Ingresos,T2.Egresos,T1.Moneda " & "ORDER BY T2.DescBanco"
                    ElseIf chkPesos.CheckState = 0 And chkDolares.CheckState = 1 Then
                        gStrSql = "SELECT T1.CtaBancaria AS Cuenta ,T2.DescBanco AS Banco,T1.SALDO AS 'Saldo al Dia Anterior'," & "T2.Ingresos AS 'Depósitos',T2.Egresos AS 'Retiros',(T1.Saldo + (T2.Ingresos - T2.Egresos)) AS 'Saldo Actual' ,T1.Moneda " & "FROM (SELECT CB.CtaBancaria,CB.Moneda,(CB.SaldoInicial + (ISNULL(SUM(CASE MB.TipoMovto WHEN 'I' THEN MB.Importe END),0))" & " - (ISNULL(SUM(CASE MB.TipoMovto WHEN 'E' THEN MB.Importe END),0))) AS SALDO,CB.Codbanco " & "FROM (SELECT * FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & ") CB LEFT OUTER JOIN (SELECT * FROM MovimientosBancarios WHERE FechaMovto < '" & Trim(strFecha) & "') MB ON " & "CB.CtaBancaria = MB.CtaBancaria and CB.CodBanco = MB.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & " ) " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,CB.CodBanco,CB.Moneda) T1 " & "INNER JOIN " & "(SELECT CB.CtaBancaria,SUM(ISNULL(CASE MB.TipoMovto WHEN 'I' THEN (MB.Importe) END,0)) AS Ingresos," & "SUM(ISNULL(CASE MB.TipoMovto WHEN 'E' THEN (MB.Importe) END,0)) AS Egresos," & "B.DescBanco,B.CodBanco " & "FROM (SELECT * FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & ") CB LEFT OUTER JOIN MovimientosBancarios MB ON CB.CodBanco = MB.Codbanco AND CB.CtaBancaria = MB.CtaBancaria AND MB.FechaMovto = '" & Trim(strFecha) & "' " & "INNER JOIN (SELECT * FROM CatBancos WHERE CodBanco = " & intCodBanco & ") B ON CB.CodBanco = B.CodBanco " & "WHERE CB.CtaBancaria IN (SELECT CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & ") " & "GROUP BY CB.CtaBancaria,CB.SaldoInicial,B.DescBanco,B.CodBanco) " & "T2 ON T1.CtaBancaria = T2.CtaBancaria AND T1.CodBanco = T2.CodBanco " & "WHERE T1.Moneda = '" & C_DOLAR & "' " & "GROUP BY T1.CtaBancaria,T2.DescBanco,T1.Saldo,T2.Ingresos,T2.Egresos,T1.Moneda " & "ORDER BY T2.DescBanco"
                    End If
                Else
                    flexDetalle.Clear()
                    Encabezado()
                    lblTotalDolaresalDiaAnterior.Text = "0.00"
                    lblTotalDolaresDepositos.Text = "0.00"
                    lblTotalDolaresRetiros.Text = "0.00"
                    lblTotalDolaresSaldoActual.Text = "0.00"
                    lblTotalPesosalDiaAnterior.Text = "0.00"
                    lblTotalPesosDepositos.Text = "0.00"
                    lblTotalPesosRetiros.Text = "0.00"
                    lblTotalPesosSaldoActual.Text = "0.00"
                    Exit Sub
                End If
            End If
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            flexDetalle.Recordset = RsGral
            ConfiguraGrid()
            CalculaTotales()
        Else
            FueraChange = True
            flexDetalle.Clear()
            Encabezado()
            MsgBox("No Existen Cuentas en este Banco, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub chkDolares_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDolares.CheckStateChanged
        Resultado()
    End Sub

    Private Sub chkDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub chkPesos_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPesos.CheckStateChanged
        Resultado()
    End Sub

    Private Sub chkPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPesos.Enter
        Pon_Tool()
    End Sub

    Private Sub chkTodoslosBancos_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodoslosBancos.CheckStateChanged
        If chkTodoslosBancos.CheckState = 1 Then
            dbcBanco.Text = ""
            dbcBanco.Enabled = False
        ElseIf chkTodoslosBancos.CheckState = 0 Then
            dbcBanco.Enabled = True
        End If
        Resultado()
    End Sub

    Private Sub chkTodoslosBancos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodoslosBancos.Enter
        Pon_Tool()
    End Sub

    Private Sub cmdConsultaPtoVta_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConsultaPtoVta.Click
        Dim frmBancosProcesoDiarioImportacionVouchers As New frmBancosProcesoDiarioImportacionVouchers()
        frmBancosProcesoDiarioImportacionVouchers.Tag = "CONSULTASALDOS"
        frmBancosProcesoDiarioImportacionVouchers.Text = "Consulta Efectivo Disponible por Sucursal"
        frmBancosProcesoDiarioImportacionVouchers.ShowDialog()
    End Sub

    Private Sub dbcBanco_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcBanco.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' ORDER BY DescBanco"
        DCChange(gStrSql, tecla)
        'intCodBanco = 0
    End Sub

    Private Sub dbcBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcBanco.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos ORDER BY DescBanco"
        DCGotFocus(gStrSql, dbcBanco)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcBanco_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcBanco.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodoslosBancos.Focus()
        End If
    End Sub

    Private Sub dbcBanco_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dbcBanco.KeyPress
        'eventArgs.keyAscii = ModEstandar.gp_CampoMayusculas(eventArgs.keyAscii)
    End Sub

    Private Sub dbcBanco_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcBanco.KeyUp
        Dim Aux As String
        Aux = dbcBanco.Text
        'If dbcBanco.SelectedItem <> 0 Then
        dbcBanco_Leave(dbcBanco, New System.EventArgs())
        'End If
        dbcBanco.Text = Aux
    End Sub

    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        If Trim(dbcBanco.Text) = "" Then Exit Sub
        If FueraChange = True Then
            FueraChange = False
            Exit Sub
        End If
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' ORDER BY DescBanco"
        DCLostFocus(dbcBanco, gStrSql, intCodBanco)
        Resultado()
    End Sub

    Private Sub dbcBanco_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcBanco.MouseUp
        Dim Aux As String
        Aux = dbcBanco.Text
        'If dbcBanco.SelectedItem <> 0 Then
        'dbcBanco_Leave(dbcBanco, New System.EventArgs())
        'End If
        dbcBanco.Text = Aux
    End Sub

    Private Sub dtpFechaCorte_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaCorte.CursorChanged
        Resultado()
        ObtenerIngresos()
    End Sub

    Private Sub dtpFechaCorte_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaCorte.Enter
        Pon_Tool()
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        Pon_Tool()
        flexDetalle.Col = 0
        flexDetalle.Row = 1
    End Sub

    Private Sub frmBancosProcesoDiarioConsultadeSaldos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoDiarioConsultadeSaldos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoDiarioConsultadeSaldos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "chkTodoslosBancos" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoDiarioConsultadeSaldos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioConsultadeSaldos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        dtpFechaCorte.MinDate = C_FECHAINICIAL
        dtpFechaCorte.MaxDate = C_FECHAFINAL
        dtpFechaCorte.Value = CDate(Today)
        'Nuevo
        chkTodoslosBancos.CheckState = System.Windows.Forms.CheckState.Checked
        chkDolares.CheckState = System.Windows.Forms.CheckState.Checked
        chkPesos.CheckState = System.Windows.Forms.CheckState.Checked
        ObtenerIngresos()
    End Sub

    Private Sub frmBancosProcesoDiarioConsultadeSaldos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
        '    'If Cambios = True And mblnNuevo = False Then
        '    'Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
        '    'Case vbYes: 'Guardar el registro
        '    'If Guardar = False Then
        '    'Cancel = 1
        '    'End If
        '    'Case vbNo: 'No hace nada y permite el cierre del formulario
        '    'Case vbCancel: 'Cancela el cierre del formulario sin guardar
        '    'Cancel = 1
        '    'End Select
        '    'End If
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

    Private Sub frmBancosProcesoDiarioConsultadeSaldos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBancosProcesoDiarioConsultadeSaldos))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkDolares = New System.Windows.Forms.CheckBox()
        Me.chkPesos = New System.Windows.Forms.CheckBox()
        Me.chkTodoslosBancos = New System.Windows.Forms.CheckBox()
        Me.cmdConsultaPtoVta = New System.Windows.Forms.Button()
        Me.Line1 = New System.Windows.Forms.Label()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.lblDolares = New System.Windows.Forms.Label()
        Me.lblPesos = New System.Windows.Forms.Label()
        Me.lblTarjetas = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Line2 = New System.Windows.Forms.Label()
        Me.lblTotalDolaresSaldoActual = New System.Windows.Forms.Label()
        Me.lblTotalDolaresRetiros = New System.Windows.Forms.Label()
        Me.lblTotalDolaresDepositos = New System.Windows.Forms.Label()
        Me.lblTotalDolaresalDiaAnterior = New System.Windows.Forms.Label()
        Me.lblTotalPesosSaldoActual = New System.Windows.Forms.Label()
        Me.lblTotalPesosRetiros = New System.Windows.Forms.Label()
        Me.lblTotalPesosDepositos = New System.Windows.Forms.Label()
        Me.lblTotalPesosalDiaAnterior = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.dtpFechaCorte = New System.Windows.Forms.DateTimePicker()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkDolares
        '
        Me.chkDolares.BackColor = System.Drawing.SystemColors.Control
        Me.chkDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDolares.Location = New System.Drawing.Point(13, 36)
        Me.chkDolares.Name = "chkDolares"
        Me.chkDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDolares.Size = New System.Drawing.Size(67, 17)
        Me.chkDolares.TabIndex = 3
        Me.chkDolares.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me.chkDolares, "Muestra solo las cuentas en Dolares.")
        Me.chkDolares.UseVisualStyleBackColor = False
        '
        'chkPesos
        '
        Me.chkPesos.BackColor = System.Drawing.SystemColors.Control
        Me.chkPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPesos.Location = New System.Drawing.Point(13, 17)
        Me.chkPesos.Name = "chkPesos"
        Me.chkPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPesos.Size = New System.Drawing.Size(67, 17)
        Me.chkPesos.TabIndex = 2
        Me.chkPesos.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me.chkPesos, "Muestra solo Las cuentas en Pesos.")
        Me.chkPesos.UseVisualStyleBackColor = False
        '
        'chkTodoslosBancos
        '
        Me.chkTodoslosBancos.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodoslosBancos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodoslosBancos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodoslosBancos.Location = New System.Drawing.Point(12, 15)
        Me.chkTodoslosBancos.Name = "chkTodoslosBancos"
        Me.chkTodoslosBancos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodoslosBancos.Size = New System.Drawing.Size(121, 17)
        Me.chkTodoslosBancos.TabIndex = 0
        Me.chkTodoslosBancos.Text = "Todos los Bancos"
        Me.ToolTip1.SetToolTip(Me.chkTodoslosBancos, "Selecciona Todos los Bancos.")
        Me.chkTodoslosBancos.UseVisualStyleBackColor = False
        '
        'cmdConsultaPtoVta
        '
        Me.cmdConsultaPtoVta.BackColor = System.Drawing.SystemColors.Control
        Me.cmdConsultaPtoVta.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdConsultaPtoVta.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdConsultaPtoVta.Location = New System.Drawing.Point(593, 495)
        Me.cmdConsultaPtoVta.Name = "cmdConsultaPtoVta"
        Me.cmdConsultaPtoVta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdConsultaPtoVta.Size = New System.Drawing.Size(138, 30)
        Me.cmdConsultaPtoVta.TabIndex = 31
        Me.cmdConsultaPtoVta.Text = "Consulta por &Sucursal"
        Me.cmdConsultaPtoVta.UseVisualStyleBackColor = False
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(9, 87)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(179, 1)
        Me.Line1.TabIndex = 0
        '
        'lblTotal
        '
        Me.lblTotal.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotal.Location = New System.Drawing.Point(85, 98)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotal.Size = New System.Drawing.Size(103, 21)
        Me.lblTotal.TabIndex = 30
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblDolares
        '
        Me.lblDolares.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDolares.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDolares.Location = New System.Drawing.Point(85, 35)
        Me.lblDolares.Name = "lblDolares"
        Me.lblDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDolares.Size = New System.Drawing.Size(103, 21)
        Me.lblDolares.TabIndex = 29
        Me.lblDolares.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblPesos
        '
        Me.lblPesos.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblPesos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPesos.Location = New System.Drawing.Point(85, 8)
        Me.lblPesos.Name = "lblPesos"
        Me.lblPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPesos.Size = New System.Drawing.Size(103, 21)
        Me.lblPesos.TabIndex = 28
        Me.lblPesos.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTarjetas
        '
        Me.lblTarjetas.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTarjetas.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTarjetas.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTarjetas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTarjetas.Location = New System.Drawing.Point(85, 62)
        Me.lblTarjetas.Name = "lblTarjetas"
        Me.lblTarjetas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTarjetas.Size = New System.Drawing.Size(104, 21)
        Me.lblTarjetas.TabIndex = 27
        Me.lblTarjetas.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(8, 99)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(41, 21)
        Me.Label8.TabIndex = 26
        Me.Label8.Text = "Total"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(5, 62)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(79, 13)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "T.C. No Acred."
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(9, 36)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(68, 21)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "Dólares"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(9, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(68, 21)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Pesos"
        '
        'Line2
        '
        Me.Line2.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line2.Location = New System.Drawing.Point(4, 104)
        Me.Line2.Name = "Line2"
        Me.Line2.Size = New System.Drawing.Size(179, 1)
        Me.Line2.TabIndex = 31
        '
        'lblTotalDolaresSaldoActual
        '
        Me.lblTotalDolaresSaldoActual.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalDolaresSaldoActual.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalDolaresSaldoActual.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalDolaresSaldoActual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalDolaresSaldoActual.Location = New System.Drawing.Point(364, 35)
        Me.lblTotalDolaresSaldoActual.Name = "lblTotalDolaresSaldoActual"
        Me.lblTotalDolaresSaldoActual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalDolaresSaldoActual.Size = New System.Drawing.Size(101, 21)
        Me.lblTotalDolaresSaldoActual.TabIndex = 20
        Me.lblTotalDolaresSaldoActual.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalDolaresRetiros
        '
        Me.lblTotalDolaresRetiros.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalDolaresRetiros.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalDolaresRetiros.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalDolaresRetiros.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalDolaresRetiros.Location = New System.Drawing.Point(262, 35)
        Me.lblTotalDolaresRetiros.Name = "lblTotalDolaresRetiros"
        Me.lblTotalDolaresRetiros.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalDolaresRetiros.Size = New System.Drawing.Size(101, 21)
        Me.lblTotalDolaresRetiros.TabIndex = 19
        Me.lblTotalDolaresRetiros.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalDolaresDepositos
        '
        Me.lblTotalDolaresDepositos.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalDolaresDepositos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalDolaresDepositos.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalDolaresDepositos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalDolaresDepositos.Location = New System.Drawing.Point(160, 35)
        Me.lblTotalDolaresDepositos.Name = "lblTotalDolaresDepositos"
        Me.lblTotalDolaresDepositos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalDolaresDepositos.Size = New System.Drawing.Size(101, 21)
        Me.lblTotalDolaresDepositos.TabIndex = 18
        Me.lblTotalDolaresDepositos.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalDolaresalDiaAnterior
        '
        Me.lblTotalDolaresalDiaAnterior.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalDolaresalDiaAnterior.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalDolaresalDiaAnterior.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalDolaresalDiaAnterior.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalDolaresalDiaAnterior.Location = New System.Drawing.Point(58, 35)
        Me.lblTotalDolaresalDiaAnterior.Name = "lblTotalDolaresalDiaAnterior"
        Me.lblTotalDolaresalDiaAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalDolaresalDiaAnterior.Size = New System.Drawing.Size(102, 21)
        Me.lblTotalDolaresalDiaAnterior.TabIndex = 17
        Me.lblTotalDolaresalDiaAnterior.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalPesosSaldoActual
        '
        Me.lblTotalPesosSaldoActual.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalPesosSaldoActual.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalPesosSaldoActual.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalPesosSaldoActual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalPesosSaldoActual.Location = New System.Drawing.Point(364, 8)
        Me.lblTotalPesosSaldoActual.Name = "lblTotalPesosSaldoActual"
        Me.lblTotalPesosSaldoActual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalPesosSaldoActual.Size = New System.Drawing.Size(101, 21)
        Me.lblTotalPesosSaldoActual.TabIndex = 16
        Me.lblTotalPesosSaldoActual.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalPesosRetiros
        '
        Me.lblTotalPesosRetiros.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalPesosRetiros.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalPesosRetiros.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalPesosRetiros.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalPesosRetiros.Location = New System.Drawing.Point(262, 8)
        Me.lblTotalPesosRetiros.Name = "lblTotalPesosRetiros"
        Me.lblTotalPesosRetiros.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalPesosRetiros.Size = New System.Drawing.Size(101, 21)
        Me.lblTotalPesosRetiros.TabIndex = 15
        Me.lblTotalPesosRetiros.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalPesosDepositos
        '
        Me.lblTotalPesosDepositos.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalPesosDepositos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalPesosDepositos.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalPesosDepositos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalPesosDepositos.Location = New System.Drawing.Point(160, 8)
        Me.lblTotalPesosDepositos.Name = "lblTotalPesosDepositos"
        Me.lblTotalPesosDepositos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalPesosDepositos.Size = New System.Drawing.Size(101, 21)
        Me.lblTotalPesosDepositos.TabIndex = 14
        Me.lblTotalPesosDepositos.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalPesosalDiaAnterior
        '
        Me.lblTotalPesosalDiaAnterior.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalPesosalDiaAnterior.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalPesosalDiaAnterior.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalPesosalDiaAnterior.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalPesosalDiaAnterior.Location = New System.Drawing.Point(58, 8)
        Me.lblTotalPesosalDiaAnterior.Name = "lblTotalPesosalDiaAnterior"
        Me.lblTotalPesosalDiaAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalPesosalDiaAnterior.Size = New System.Drawing.Size(102, 21)
        Me.lblTotalPesosalDiaAnterior.TabIndex = 13
        Me.lblTotalPesosalDiaAnterior.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(9, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(49, 16)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Dólares :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(9, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(44, 21)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Pesos :"
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(12, 14)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(696, 196)
        Me.flexDetalle.TabIndex = 5
        '
        'dtpFechaCorte
        '
        Me.dtpFechaCorte.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaCorte.Location = New System.Drawing.Point(603, 18)
        Me.dtpFechaCorte.Name = "dtpFechaCorte"
        Me.dtpFechaCorte.Size = New System.Drawing.Size(83, 20)
        Me.dtpFechaCorte.TabIndex = 4
        '
        'dbcBanco
        '
        Me.dbcBanco.Location = New System.Drawing.Point(76, 43)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(217, 21)
        Me.dbcBanco.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(21, 46)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(49, 21)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Banco :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(510, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(92, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Fecha de Corte :"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.chkTodoslosBancos)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.dbcBanco)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.dtpFechaCorte)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(719, 111)
        Me.Panel1.TabIndex = 0
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(351, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(46, 13)
        Me.Label9.TabIndex = 22
        Me.Label9.Text = "Moneda"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.chkPesos)
        Me.Panel2.Controls.Add(Me.chkDolares)
        Me.Panel2.Location = New System.Drawing.Point(354, 32)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(99, 63)
        Me.Panel2.TabIndex = 4
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.flexDetalle)
        Me.Panel3.Location = New System.Drawing.Point(12, 153)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(719, 224)
        Me.Panel3.TabIndex = 1
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Label5)
        Me.Panel4.Controls.Add(Me.Label8)
        Me.Panel4.Controls.Add(Me.Label7)
        Me.Panel4.Controls.Add(Me.Label6)
        Me.Panel4.Controls.Add(Me.lblPesos)
        Me.Panel4.Controls.Add(Me.Line1)
        Me.Panel4.Controls.Add(Me.lblDolares)
        Me.Panel4.Controls.Add(Me.lblTotal)
        Me.Panel4.Controls.Add(Me.lblTarjetas)
        Me.Panel4.Location = New System.Drawing.Point(12, 401)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(209, 134)
        Me.Panel4.TabIndex = 2
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(12, 384)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(92, 13)
        Me.Label10.TabIndex = 3
        Me.Label10.Text = "Saldos sucursales"
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.Label3)
        Me.Panel5.Controls.Add(Me.Label2)
        Me.Panel5.Controls.Add(Me.lblTotalPesosalDiaAnterior)
        Me.Panel5.Controls.Add(Me.lblTotalPesosSaldoActual)
        Me.Panel5.Controls.Add(Me.lblTotalPesosRetiros)
        Me.Panel5.Controls.Add(Me.lblTotalPesosDepositos)
        Me.Panel5.Controls.Add(Me.lblTotalDolaresalDiaAnterior)
        Me.Panel5.Controls.Add(Me.lblTotalDolaresDepositos)
        Me.Panel5.Controls.Add(Me.lblTotalDolaresSaldoActual)
        Me.Panel5.Controls.Add(Me.lblTotalDolaresRetiros)
        Me.Panel5.Location = New System.Drawing.Point(246, 403)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(485, 75)
        Me.Panel5.TabIndex = 4
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(248, 386)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(77, 13)
        Me.Label11.TabIndex = 5
        Me.Label11.Text = "Saldos bancos"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(16, 135)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(191, 13)
        Me.Label12.TabIndex = 32
        Me.Label12.Text = "DETALLE DE CUENTAS BANCARIAS"
        '
        'frmBancosProcesoDiarioConsultadeSaldos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(743, 550)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdConsultaPtoVta)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(168, 131)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoDiarioConsultadeSaldos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Consulta de Saldos de Cuentas Bancarias"
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

End Class