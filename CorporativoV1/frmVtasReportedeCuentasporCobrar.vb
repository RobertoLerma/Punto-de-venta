Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmVtasReportedeCuentasporCobrar
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtNombre As System.Windows.Forms.TextBox
    Public WithEvents chkSoloVigentes As System.Windows.Forms.CheckBox
    Public WithEvents chkCuentasSaldadas As System.Windows.Forms.CheckBox
    Public WithEvents optResumenGeneral As System.Windows.Forms.RadioButton
    Public WithEvents chkTodaslasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents optDetalladoporCliente As System.Windows.Forms.RadioButton
    Public WithEvents txtCodCliente As System.Windows.Forms.TextBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents Line1 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Line2 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Line3 As System.Windows.Forms.Label
    Public WithEvents Line4 As System.Windows.Forms.Label


    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim rsReporte As ADODB.Recordset
    Dim sglTiempoCambio As Single '''Para Esperar un Tiempo

    Public gFueraChange As Boolean '''03MAR2008

    Const ColorGris As Integer = &H8000000F '''03MAR2008 - MAVF
    Const ColorAmarillo As Integer = &HC0FFFF '''03MAR2008 - MAVF
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Public WithEvents btnBuscar As Button
    Const ColorBlanco As Integer = &HFFFFFF '''03MAR2008 - MAVF
    Public strControlActual As String 'Nombre del control actual

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendrá el string del tag que se le mandara al fromulario de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        If strControlActual = "TXTNOMBRE" Then
            strCaptionForm = "Consulta de Clientes"
            gStrSql = "Select Right('00000' + ltrim(rtrim(CodCliente)),5) as Codigo, DescCliente as Nombre From CatClientes (Nolock) Where DescCliente Like '" & Trim(txtNombre.Text) & "%' Order by DescCliente "
        End If

        strSQL = gStrSql 'Se hace uso de una variable temporal para el query
        gStrSql = strSQL 'Se regresa el valor de la variable temporal a la variable original

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsgBox(C_msgSINDATOS & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta 
        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 7000, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTNOMBRE"
                    .set_ColWidth(0, 0, 1000) 'Columna del Código
                    .set_ColWidth(1, 0, 6000) 'Columna de la Descripción
                    .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            End Select
        End With
        FrmConsultas.ShowDialog()

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Imprime()

        Dim RptVtasReportedeCXCDetalladoporCliente As New RptVtasReportedeCXCDetalladoporCliente
        Dim rptVtasReportedeCuentasXCobrar As New RptVtasReporteGeneralCuentasXCobrar

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        'On Error GoTo ImprimeErr

        Dim Sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim PeriodoReporte As String
        Dim CodigoCliente As String
        Dim NombreCliente As String
        Dim strWhere As String
        Dim strHaving As String
        Dim strSucursal As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim RsAux As ADODB.Recordset

        'Do While (sglTiempoCambio) <= 2.1
        'Loop
        'System.Windows.Forms.Application.DoEvents()
        If dtpFechaInicial.Value > dtpFechaFinal.Value Then
            MsgBox("La Fecha Inicial no Puede ser Mayor que la Fecha Final.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If dtpFechaInicial.Value > Now Then
            MsgBox("la Fecha Inicial no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If dtpFechaFinal.Value > Now Then
            MsgBox("la Fecha Final no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If optResumenGeneral.Checked = True Then
            If CDbl(Numerico(txtCodSucursal.Text)) = 0 And chkTodaslasSucursales.CheckState = 0 Then
                MsgBox("Proporcione un Codigo de Sucursal, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodSucursal.Focus()
                Exit Sub
            End If
            If Trim(dbcSucursal.Text) = "" And chkTodaslasSucursales.CheckState = 0 Then
                MsgBox("Proporcione la Descripción de la Sucursal, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                dbcSucursal.Focus()
                Exit Sub
            End If
        ElseIf optDetalladoporCliente.Checked = True Then
            If CDbl(Numerico(txtCodCliente.Text)) = 0 Then
                MsgBox("Proporcione el Código de un Cliente, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodCliente.Focus()
                Exit Sub
            End If
            If Trim(txtNombre.Text) = "" Then
                MsgBox("Proporcione el Nombre de un Cliente, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtNombre.Focus()
                Exit Sub
            End If
        End If
        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        If optResumenGeneral.Checked = True Then
            NombreReporte = UCase("Cuentas por Cobrar")
        Else
            NombreReporte = UCase("Cuentas por Cobrar")
            CodigoCliente = txtCodCliente.Text
            NombreCliente = txtNombre.Text
        End If

        'FechaInicial = Format(Month(dtpFechaInicial.Value), "00") & "/" & Format((dtpFechaInicial.Value), "00") & "/" & VB6.Format(Year(dtpFechaInicial.Value), "0000")
        'FechaFinal = Format(Month(dtpFechaFinal.Value), "00") & "/" & Format((dtpFechaFinal.Value), "00") & "/" & VB6.Format(Year(dtpFechaFinal.Value), "0000")
        'PeriodoReporte = "Del " & Format(dtpFechaInicial.Value, "dd/mmm/yyyy") & " al " & Format(dtpFechaFinal.Value, "dd/mmm/yyyy")
        FechaInicial = AgregarHoraAFecha(dtpFechaInicial.Value)
        FechaFinal = AgregarHoraAFecha(dtpFechaFinal.Value)
        PeriodoReporte = "Del " & FechaInicial & " al " & FechaFinal

        strWhere = ""
        Cmd.CommandTimeout = 300

        If optResumenGeneral.Checked = True Then
            If chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Checked Then
                strHaving = "Having (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1)) > 0 /*AND (Round(Sum(Case When VtaCab.Moneda = 'D' Then ISNULL(Ing.Total,0) Else ISNULL(Ing.Total * Ing.TipoCambio,0) End),1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN ISNULL(DC.TOTALDOCTO,0) ELSE ISNULL(DC.TOTALDOCTO * VTACAB.TIPOCAMBIO,0) END,1)) <> 0 */" & " AND (Round(Sum(Case When VtaCab.Moneda = 'D' Then ISNULL(Ing.Total,0) Else ISNULL(Ing.Total * Ing.TipoCambio,0) End),1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN ISNULL(DC.TOTALDOCTO,0) ELSE ISNULL(DC.TOTALDOCTO * VTACAB.TIPOCAMBIO,0) END,1)) < " & "(ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1)) "
                strWhere = ""
            ElseIf chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                strWhere = " AND VtaCab.FechaVenta BETWEEN '" & Format(dtpFechaInicial.Value, C_FORMATFECHAGUARDAR) & "' AND '" & VB6.Format(dtpFechaFinal.Value, C_FORMATFECHAGUARDAR) & "' "
                strHaving = "Having (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1)) > 0 "
            End If
            If chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                strSucursal = " AND VTACAB.CODSUCURSAL = " & CInt(Numerico((txtCodSucursal.Text))) & " "
            Else
                strSucursal = ""
            End If
            Sql = "SELECT SUC.DESCALMACEN,VTACAB.FOLIOVENTA,VTACAB.FECHAVENTA,CATCLI.DESCCLIENTE," & "(ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1)) AS APARTADO,(Round(Sum(Case When VtaCab.Moneda = 'D' Then ISNULL(Ing.Total,0) Else ISNULL(Ing.Total * Ing.TipoCambio,0) End),1)) - (ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN ISNULL(DC.TOTALDOCTO,0) ELSE ISNULL(DC.TOTALDOCTO * VTACAB.TIPOCAMBIO,0) END,1)) AS ABONOS," & "(ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN (VTACAB.TOTAL + VTACAB.REDONDEO) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO END,1) - ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN isnull(DC.TotalDevol + VtaCab.Redondeo,0) ELSE ISNULL((DC.TOTALDEVOL + VtaCab.Redondeo) * VTACAB.TIPOCAMBIO,0) END,1)) - (Round(Sum(Case When VtaCab.Moneda = 'D' Then ISNULL(Ing.Total,0) Else ISNULL(Ing.Total * Ing.TipoCambio,0) End),1) - ROUND(CASE WHEN VTACAB.MONEDA = 'D' THEN ISNULL(DC.TOTALDOCTO,0) ELSE ISNULL(DC.TOTALDOCTO * VTACAB.TIPOCAMBIO,0) END,1)) as SALDO," & "VtaCab.Moneda, VtaCab.TipoCambio, VtaCab.Estatus,Round(Case When VtaCab.Moneda = 'D' Then (VtaCab.Total+VtaCab.Redondeo) Else ((VtaCab.Total+VtaCab.Redondeo)*VtaCab.TipoCambio) End,1) as VtaReal, Round(Sum(Case When VtaCab.Moneda = 'D' Then ISNULL(Ing.Total,0) Else ISNULL(Ing.Total*Ing.TipoCambio,0) End),1) As IngresoReal, Round(IsNull(Case When DF.Moneda = 'D' Then DF.Importe Else (DF.Importe*DF.TipoCambio) End,0),1) As DifCamb, Max(ISNULL(Ing.FechaIngreso,'')) as FechaUltIng  " & "FROM MOVIMIENTOSVENTASCAB VTACAB (Nolock) LEFT OUTER JOIN (SELECT * FROM INGRESOS (Nolock) WHERE ESTATUS <> 'C') ING ON VTACAB.FOLIOVENTA = ING.FOLIOMOVTO LEFT OUTER JOIN (SELECT FolioVenta,SUM(TotalDevol) TotalDevol,SUM(TotalDocto) TotalDocto FROM DevolucionesCab WHERE ESTATUS <> 'C' GROUP BY FolioVenta) DC ON VtaCab.FolioVenta = DC.FolioVenta INNER JOIN CATALMACEN SUC (Nolock) ON VTACAB.CODSUCURSAL = SUC.CODALMACEN INNER JOIN CATCLIENTES CATCLI (Nolock) ON VTACAB.CODCLIENTE = CATCLI.CODCLIENTE LEFT OUTER JOIN DiferenciaCambiaria DF (Nolock) ON VtaCab.FolioVenta = DF.FolioVenta " & "WHERE VTACAB.TIPOMOVTO = 'V' " & strSucursal & " AND VTACAB.ESTATUS <> 'C' AND VTACAB.CONDICION = 'CR' " & strWhere & "GROUP  BY SUC.DESCALMACEN,VTACAB.FOLIOVENTA,VTACAB.FECHAVENTA,CATCLI.DESCCLIENTE,VTACAB.TOTAL,VTACAB.REDONDEO,VtaCab.Moneda,VtaCab.TipoCambio,VtaCab.Estatus,DF.Moneda,DF.TipoCambio,DF.Importe,DC.TotalDocto,DC.TotalDevol " & strHaving & "ORDER  BY SUC.DESCALMACEN,CATCLI.DESCCLIENTE,VTACAB.FECHAVENTA,VTACAB.FOLIOVENTA"
            BorraCmd()
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd.CommandText = Sql
            frmReportes.rsReport = Cmd.Execute

            If frmReportes.rsReport.RecordCount = 0 Then
                MsgBox("No existen movimientos en este periodo de fechas" & vbNewLine & "Favor de verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                Exit Sub
            Else
                'frmReportes.Report = rptVtasReportedeCuentasXCobrar
                rptVtasReportedeCuentasXCobrar.SetDataSource(frmReportes.rsReport)
            End If

            NombreReporte = UCase("Reporte de Cuentas por Cobrar")
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'frmReportes.rsReport = rsReporte
            'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "PeriodoReporte"}
            'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte}
            frmReportes.Text = "Reporte General de Cuentas por Cobrar"
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            frmReportes.reporteActual = rptVtasReportedeCuentasXCobrar
            frmReportes.Show()
            Cursor = System.Windows.Forms.Cursors.Default
            FueraChange = False

        ElseIf optDetalladoporCliente.Checked = True Then

            If chkCuentasSaldadas.CheckState = 0 Then
                strWhere = "WHERE A.ESTATUS <> 'S'"
            ElseIf chkCuentasSaldadas.CheckState = 1 Then
                strWhere = ""
            End If
            '''Sql = "SELECT A.FOLIOVENTA,CASE WHEN A.TIPOMOVTO = 1 THEN A.FECHAVENTA ELSE '' END AS FECHA," & _
            '"A.TIPOMOVTO,A.CODIGOARTICULO,A.DESCARTICULO,A.CANTIDAD," & _
            '"ROUND(CASE WHEN B.MONEDA = 'D' THEN A.PRECIOREAL ELSE A.PRECIOREAL * B.TIPOCAMBIO END,1) AS PRECIOREAL," & _
            '"ROUND(CASE WHEN B.MONEDA = 'D' THEN A.IMPORTE ELSE A.IMPORTE * B.TIPOCAMBIO END,1) AS IMPORTE," & _
            '"A.FOLIOINGRESO,A.FECHAINGRESO,ROUND(CASE WHEN B.MONEDA = 'D' THEN A.TOTAL ELSE A.TOTAL * B.TIPOCAMBIO END,1) AS TOTAL," & _
            '"ROUND(CASE WHEN B.MONEDA = 'D' THEN A.ABONO ELSE A.ABONO * B.TIPOCAMBIO END,1) AS ABONO," & _
            '"ROUND(CASE WHEN B.MONEDA = 'D' THEN A.SALDO ELSE A.SALDO * B.TIPOCAMBIO END,1) AS SALDO,A.NUMPARTIDA,A.ESTATUS,A.FIN,B.MONEDA,C.DESCALMACEN " & _
            '"FROM SALDOCUENTASXCOBRAR(" & txtCodCliente & ",'" & FechaInicial & "','" & FechaFinal & "') A " & _
            '"LEFT OUTER JOIN MOVIMIENTOSVENTASCAB B ON CAST(A.FOLIOVENTA AS NVARCHAR) COLLATE Traditional_Spanish_CI_AI = CAST(B.FOLIOVENTA AS NVARCHAR) COLLATE Traditional_Spanish_CI_AI " & _
            '"INNER JOIN CATALMACEN C ON B.CODSUCURSAL = C.CODALMACEN " & _
            'strWhere & _
            '"GROUP BY A.FOLIOVENTA,A.FECHAVENTA,A.TIPOMOVTO,A.CODIGOARTICULO,A.DESCARTICULO,A.CANTIDAD,A.PRECIOREAL,A.IMPORTE," & _
            '"A.FOLIOINGRESO,A.FECHAINGRESO,A.TOTAL,A.ABONO,A.SALDO,A.NUMPARTIDA,A.ESTATUS,A.FIN,B.MONEDA,B.TIPOCAMBIO,C.DESCALMACEN " & _
            '"ORDER BY A.FOLIOVENTA,A.FECHAVENTA,A.TIPOMOVTO,A.NUMPARTIDA"

            Sql = "SELECT   A.FOLIOVENTA, CASE WHEN A.TIPOMOVTO = 1 THEN A.FECHAVENTA ELSE '' END AS FECHA, A.TIPOMOVTO, A.CODIGOARTICULO, " & "A.DESCARTICULO, A.CANTIDAD, ROUND(A.PRECIOREAL,1) AS PRECIOREAL, ROUND(A.IMPORTE,1) AS IMPORTE, A.FOLIOINGRESO, A.FECHAINGRESO, " & "ROUND(A.TOTAL,1) AS TOTAL, ROUND(A.ABONO,1) AS ABONO, ROUND(A.SALDO,1) AS SALDO, A.NumPartida , A.Estatus, A.FIN, B.Moneda, c.DescAlmacen " & "FROM  SALDOCUENTASXCOBRAR (" & txtCodCliente.Text & ",'" & FechaInicial & "','" & FechaFinal & "') A " & "LEFT OUTER JOIN MOVIMIENTOSVENTASCAB B ON CAST(A.FOLIOVENTA AS NVARCHAR) COLLATE Traditional_Spanish_CI_AI = CAST(B.FOLIOVENTA AS NVARCHAR) COLLATE Traditional_Spanish_CI_AI " & "INNER JOIN CATALMACEN C ON B.CODSUCURSAL = C.CODALMACEN " & strWhere & "GROUP BY A.FOLIOVENTA, A.FECHAVENTA, A.TIPOMOVTO, A.CODIGOARTICULO, A.DESCARTICULO, A.CANTIDAD, " & "A.PRECIOREAL, A.IMPORTE, A.FOLIOINGRESO, A.FECHAINGRESO, A.TOTAL, A.ABONO, A.SALDO, A.NUMPARTIDA, " & "A.ESTATUS, A.FIN, B.MONEDA, B.TIPOCAMBIO, C.DESCALMACEN " & "ORDER BY A.FOLIOVENTA, A.FECHAVENTA, A.TIPOMOVTO, A.NUMPARTIDA "

            BorraCmd()
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd.CommandText = Sql
            frmReportes.rsReport = Cmd.Execute

            If frmReportes.rsReport.RecordCount = 0 Then
                MsgBox("No Existen Movimientos En Este Periodo de Fechas, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Sub
            Else
                'frmReportes.Report = RptVtasReportedeCXCDetalladoporCliente
                RptVtasReportedeCXCDetalladoporCliente.SetDataSource(frmReportes.rsReport)
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'frmReportes.rsReport = rsReporte
            'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "PeriodoReporte", "CodigoCliente", "NombreCliente"}
            'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte, CodigoCliente, NombreCliente}
            frmReportes.Text = "Cuentas por Cobrar - Detallado de Movimientos por Cliente"
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            frmReportes.reporteActual = RptVtasReportedeCXCDetalladoporCliente
            frmReportes.Show()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            FueraChange = False
        End If
        Cmd.CommandTimeout = 90
        Exit Sub

ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error al Imprimir : " & Err.Description, MsgBoxStyle.Exclamation, "Error de Operacion")
        FueraChange = False
    End Sub

    Sub BuscaCliente()
        On Error GoTo Merr

        gStrSql = "SELECT CodCliente,DescCliente,AlmacenVExt FROM CatClientes WHERE CodCliente = " & txtCodCliente.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("CodCliente").Value = 1 Then
                MsgBox("El Cliente Publico en General No Tiene Registradas Cuentas X Cobrar, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodCliente.Text = ""
                txtCodCliente.Focus()
                Exit Sub
            End If
            If Not IsDBNull(RsGral.Fields("AlmacenVExt").Value) Then
                MsgBox("Los Clientes Registrados como Vendedores Externo No Tienen Registradas Cuentas X Cobrar, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodCliente.Text = ""
                txtCodCliente.Focus()
                Exit Sub
            End If
            'txtCodCliente.Text = Format(txtCodCliente.Text, "00000")
            'txtCodCliente.Text = Format(String.Concat(txtCodCliente.Text, "00000"))
            For i = 0 To 5 - txtCodCliente.TextLength
                txtCodCliente.Text = String.Concat("0" + txtCodCliente.Text)
            Next i
            txtNombre.Text = Trim(RsGral.Fields("DescCliente").Value)
        Else
            MsgBox("Código de Cliente No Existe, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodCliente.Text = ""
            txtCodCliente.Focus()
            Exit Sub
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub BuscaSucursal()
        On Error GoTo Merr
        gStrSql = "SELECT DescAlmacen,TipoAlmacen FROM CatAlmacen WHERE CodAlmacen = " & txtCodSucursal.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("TipoAlmacen").Value = "V" Then
                MsgBox("Este Almacen No Es Un Almacen Propio, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodSucursal.Text = ""
                txtCodSucursal.Focus()
                Exit Sub
            Else
                For i = 0 To 3 - txtCodSucursal.TextLength
                    txtCodSucursal.Text = String.Concat("0" + txtCodSucursal.Text)
                Next i
                dbcSucursal.Text = RsGral.Fields("DescAlmacen").Value

            End If
        Else
            MsgBox("Codigo de Almacen no Existe, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodSucursal.Text = ""
            txtCodSucursal.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
    End Sub

    Sub Limpiar()
        InicializaVariables()
        Nuevo()
        optResumenGeneral.Focus()
    End Sub

    Sub Nuevo()
        FueraChange = True
        optResumenGeneral.Checked = True
        chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        txtCodSucursal.Text = ""
        dbcSucursal.Text = ""
        txtCodCliente.Text = ""
        txtNombre.Text = ""
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
        chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Checked
        FueraChange = False
        gFueraChange = False
    End Sub

    Private Sub chkCuentasSaldadas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCuentasSaldadas.Enter
        Pon_Tool()
    End Sub

    Private Sub chkSoloVigentes_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSoloVigentes.CheckStateChanged
        If chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Checked Then
            dtpFechaFinal.Enabled = False
            dtpFechaInicial.Enabled = False
        ElseIf chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            dtpFechaFinal.Enabled = True
            dtpFechaInicial.Enabled = True
        End If
    End Sub

    Private Sub chkTodaslasSucursales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.CheckStateChanged
        If chkTodaslasSucursales.CheckState = 1 Then
            txtCodSucursal.Text = ""
            dbcSucursal.Text = ""
            txtCodSucursal.Enabled = False
            dbcSucursal.Enabled = False
        ElseIf chkTodaslasSucursales.CheckState = 0 Then
            txtCodSucursal.Enabled = True
            dbcSucursal.Enabled = True
        End If
    End Sub

    Private Sub chkTodaslasSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.Enter
        Pon_Tool()
    End Sub

    Private Sub chkTodaslasSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkTodaslasSucursales.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            optResumenGeneral.Focus()
        End If
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        'If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        'If Trim(dbcSucursal.Text) = "" Then
        '    txtCodSucursal.Text = ""
        'End If
        'gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        'DCChange(gStrSql, tecla)
        ''intCodSucursal = 0
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        'If dbcSucursal.Text <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursal)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        tecla = eventArgs.KeyCode
        'If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
        txtCodSucursal.Focus()
        'End If
    End Sub

    Private Sub dbcSucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcSucursal.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcSucursal_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        'If dbcSucursal.SelectedItem <> 0 Then
        'dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        'End If
        FueraChange = True
        dbcSucursal.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        FueraChange = True
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        'DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
        If intCodSucursal <> 0 Then
            txtCodSucursal.Text = Format(String.Concat(intCodSucursal, "000"))
        End If
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        'If dbcSucursal.SelectedItem <> 0 Then
        'dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        'End If
        FueraChange = True
        dbcSucursal.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_SelectedValueChanged(sender As Object, e As EventArgs) Handles dbcSucursal.SelectedValueChanged
        gStrSql = "SELECT CodAlmacen,rtrim(ltrim(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
        txtCodSucursal.Text = intCodSucursal
    End Sub

    Private Sub dtpFechaFinal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.Click
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaFinal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaFinal.KeyPress
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInicial.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInicial.Click
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaInicial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaInicial.KeyPress
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub frmVtasReportedeCuentasporCobrar_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasReportedeCuentasporCobrar_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasReportedeCuentasporCobrar_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optResumenGeneral" And System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optDetalladoporCliente" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmVtasReportedeCuentasporCobrar_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasReportedeCuentasporCobrar_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        InicializaVariables()
        dtpFechaInicial.MinDate = C_FECHAINICIAL
        dtpFechaInicial.MaxDate = C_FECHAFINAL
        dtpFechaFinal.MinDate = C_FECHAINICIAL
        dtpFechaFinal.MaxDate = C_FECHAFINAL
        Nuevo()
    End Sub

    Private Sub frmVtasReportedeCuentasporCobrar_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        'Si se cierra el formulario y existio algun cambio en el registro se
        'informa al usuario del cabio y si desea guardar el registro, ya sea
        'que sea nuevo o un registro modificado
        If Not mblnSalir Then
            'If Cambios = True And mblnNuevo = False Then
            'Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
            'Case vbYes: 'Guardar el registro
            'If Guardar = False Then
            'Cancel = 1
            'End If
            'Case vbNo: 'No hace nada y permite el cierre del formulario
            'Case vbCancel: 'Cancela el cierre del formulario sin guardar
            'Cancel = 1
            'End Select
            'End If
        Else
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

    Private Sub frmVtasReportedeCuentasporCobrar_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optDetalladoporCliente_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDetalladoporCliente.CheckedChanged
        If eventSender.Checked Then
            chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkTodaslasSucursales.Enabled = False
            txtCodSucursal.Enabled = False
            txtCodSucursal.Text = ""
            txtCodSucursal.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            dbcSucursal.Enabled = False
            dbcSucursal.Text = ""
            dbcSucursal.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            Label1.Enabled = False
            txtCodCliente.Enabled = True
            txtCodCliente.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            txtNombre.Enabled = True
            txtNombre.BackColor = System.Drawing.ColorTranslator.FromOle(ColorAmarillo)
            Label2.Enabled = True
            chkSoloVigentes.CheckState = System.Windows.Forms.CheckState.Checked
            chkCuentasSaldadas.Enabled = True
            chkSoloVigentes.Enabled = False
            chkSoloVigentes.Enabled = False
            dtpFechaInicial.Enabled = True
            dtpFechaFinal.Enabled = True
        End If
    End Sub

    Private Sub optDetalladoporCliente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDetalladoporCliente.Enter
        Pon_Tool()
    End Sub

    Private Sub optResumenGeneral_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optResumenGeneral.CheckedChanged
        If eventSender.Checked Then
            chkCuentasSaldadas.CheckState = System.Windows.Forms.CheckState.Unchecked
            txtCodCliente.Enabled = False
            txtCodCliente.Text = ""
            txtCodCliente.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            txtNombre.Enabled = False
            txtNombre.Text = ""
            txtNombre.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            Label2.Enabled = False
            chkCuentasSaldadas.Enabled = False
            chkTodaslasSucursales.Enabled = True
            chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
            txtCodSucursal.Enabled = False
            txtCodSucursal.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            dbcSucursal.Enabled = False
            dbcSucursal.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            Label1.Enabled = True
            chkCuentasSaldadas.Enabled = False
            chkSoloVigentes.Enabled = True
            dtpFechaInicial.Enabled = False
            dtpFechaFinal.Enabled = False
        End If
    End Sub

    Private Sub optResumenGeneral_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optResumenGeneral.Enter
        Pon_Tool()
    End Sub

    Private Sub txtCodCliente_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodCliente.TextChanged
        If FueraChange Then Exit Sub
        If gFueraChange Then Exit Sub

        If Trim(txtCodCliente.Text) = "" Then
            FueraChange = True
            txtNombre.Text = ""
            FueraChange = False
        End If
    End Sub

    Private Sub txtCodCliente_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodCliente.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtCodCliente)
    End Sub

    Private Sub txtCodCliente_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodCliente.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodCliente_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodCliente.Leave
        If Trim(txtCodCliente.Text) <> "" Then
            BuscaCliente()
        End If
    End Sub

    Private Sub txtCodSucursal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.TextChanged
        If FueraChange Then Exit Sub
        dbcSucursal.Text = ""
    End Sub

    Private Sub txtCodSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Enter
        Pon_Tool()
    End Sub

    Private Sub txtCodsucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodSucursal.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Leave
        If Trim(txtCodSucursal.Text) <> "" Then
            BuscaSucursal()
        End If
    End Sub

    '''COMBO CLIENTES
    Private Sub txtNombre_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.TextChanged
        If FueraChange Then Exit Sub
        If gFueraChange Then Exit Sub

        If Trim(txtNombre.Text) = "" Then
            FueraChange = True
            txtCodCliente.Text = ""
            FueraChange = False
        End If
    End Sub

    Private Sub txtNombre_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.Enter
        strControlActual = UCase("txtNombre")
        ModEstandar.SelTextoTxt(txtNombre)
    End Sub

    Private Sub TxtNombre_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNombre.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, ".,:;{}[]+#$%&()/*\-_<>")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub LimpiarCTE()
        FueraChange = True
        txtCodCliente.Text = ""
        txtNombre.Text = ""
        FueraChange = False
    End Sub

    Sub LlenaDatos()
        On Error GoTo Merr

        'txtCodCliente.Text = Format(txtCodCliente.Text, "00000")

        For I = 0 To 5 - txtCodCliente.TextLength
            txtCodCliente.Text = String.Concat("0" + txtCodCliente.Text)
        Next I

        gStrSql = "Select Right('00000' + ltrim(rtrim(CodCliente)),5) as Codigo, DescCliente as Nombre From CatClientes (Nolock) WHERE CodCliente = " & CInt(Numerico(txtCodCliente.Text)) & " "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            txtCodCliente.Text = RsGral.Fields("Codigo").Value
            txtNombre.Text = Trim(RsGral.Fields("Nombre").Value)
        Else
            MsjNoExiste("El cliente no existe." & vbNewLine & "Favor de verificar...", gstrNombCortoEmpresa)
            LimpiarCTE()
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me.chkSoloVigentes = New System.Windows.Forms.CheckBox()
        Me.chkCuentasSaldadas = New System.Windows.Forms.CheckBox()
        Me.optResumenGeneral = New System.Windows.Forms.RadioButton()
        Me.chkTodaslasSucursales = New System.Windows.Forms.CheckBox()
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.optDetalladoporCliente = New System.Windows.Forms.RadioButton()
        Me.txtCodCliente = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.Line1 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Line2 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Line3 = New System.Windows.Forms.Label()
        Me.Line4 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtNombre
        '
        Me.txtNombre.AcceptsReturn = True
        Me.txtNombre.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtNombre.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombre.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombre.Location = New System.Drawing.Point(144, 149)
        Me.txtNombre.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNombre.MaxLength = 0
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombre.Size = New System.Drawing.Size(199, 20)
        Me.txtNombre.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtNombre, "Codigo del Cliente")
        '
        'chkSoloVigentes
        '
        Me.chkSoloVigentes.BackColor = System.Drawing.SystemColors.Control
        Me.chkSoloVigentes.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSoloVigentes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSoloVigentes.Location = New System.Drawing.Point(200, 88)
        Me.chkSoloVigentes.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSoloVigentes.Name = "chkSoloVigentes"
        Me.chkSoloVigentes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSoloVigentes.Size = New System.Drawing.Size(130, 20)
        Me.chkSoloVigentes.TabIndex = 4
        Me.chkSoloVigentes.Text = "Sólo Ventas Vigentes"
        Me.ToolTip1.SetToolTip(Me.chkSoloVigentes, "Muestra Apartados Saldados")
        Me.chkSoloVigentes.UseVisualStyleBackColor = False
        '
        'chkCuentasSaldadas
        '
        Me.chkCuentasSaldadas.BackColor = System.Drawing.SystemColors.Control
        Me.chkCuentasSaldadas.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCuentasSaldadas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCuentasSaldadas.Location = New System.Drawing.Point(45, 173)
        Me.chkCuentasSaldadas.Margin = New System.Windows.Forms.Padding(2)
        Me.chkCuentasSaldadas.Name = "chkCuentasSaldadas"
        Me.chkCuentasSaldadas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCuentasSaldadas.Size = New System.Drawing.Size(230, 28)
        Me.chkCuentasSaldadas.TabIndex = 8
        Me.chkCuentasSaldadas.Text = "Incluir Cuentas por Cobrar Saldadas"
        Me.ToolTip1.SetToolTip(Me.chkCuentasSaldadas, "Muestra Las Cuentas Saldadas")
        Me.chkCuentasSaldadas.UseVisualStyleBackColor = False
        '
        'optResumenGeneral
        '
        Me.optResumenGeneral.BackColor = System.Drawing.SystemColors.Control
        Me.optResumenGeneral.Cursor = System.Windows.Forms.Cursors.Default
        Me.optResumenGeneral.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.optResumenGeneral.Location = New System.Drawing.Point(12, 13)
        Me.optResumenGeneral.Margin = New System.Windows.Forms.Padding(2)
        Me.optResumenGeneral.Name = "optResumenGeneral"
        Me.optResumenGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optResumenGeneral.Size = New System.Drawing.Size(113, 17)
        Me.optResumenGeneral.TabIndex = 0
        Me.optResumenGeneral.TabStop = True
        Me.optResumenGeneral.Text = "&Resumen General"
        Me.ToolTip1.SetToolTip(Me.optResumenGeneral, "Muestra el Resumen General de Cuentas por Cobrar")
        Me.optResumenGeneral.UseVisualStyleBackColor = False
        '
        'chkTodaslasSucursales
        '
        Me.chkTodaslasSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodaslasSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodaslasSucursales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodaslasSucursales.Location = New System.Drawing.Point(28, 34)
        Me.chkTodaslasSucursales.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodaslasSucursales.Name = "chkTodaslasSucursales"
        Me.chkTodaslasSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodaslasSucursales.Size = New System.Drawing.Size(129, 17)
        Me.chkTodaslasSucursales.TabIndex = 1
        Me.chkTodaslasSucursales.Text = "Todas las Sucursales"
        Me.ToolTip1.SetToolTip(Me.chkTodaslasSucursales, "Muestra Todas las Sucursales")
        Me.chkTodaslasSucursales.UseVisualStyleBackColor = False
        '
        'txtCodSucursal
        '
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.Location = New System.Drawing.Point(94, 57)
        Me.txtCodSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucursal.MaxLength = 3
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(41, 20)
        Me.txtCodSucursal.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtCodSucursal, "Codigo de la Sucursal")
        '
        'optDetalladoporCliente
        '
        Me.optDetalladoporCliente.BackColor = System.Drawing.SystemColors.Control
        Me.optDetalladoporCliente.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDetalladoporCliente.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.optDetalladoporCliente.Location = New System.Drawing.Point(12, 124)
        Me.optDetalladoporCliente.Margin = New System.Windows.Forms.Padding(2)
        Me.optDetalladoporCliente.Name = "optDetalladoporCliente"
        Me.optDetalladoporCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDetalladoporCliente.Size = New System.Drawing.Size(123, 17)
        Me.optDetalladoporCliente.TabIndex = 5
        Me.optDetalladoporCliente.TabStop = True
        Me.optDetalladoporCliente.Text = "Detallad&o por Cliente"
        Me.ToolTip1.SetToolTip(Me.optDetalladoporCliente, "Muestra el Reporte Detallado por Cliente")
        Me.optDetalladoporCliente.UseVisualStyleBackColor = False
        '
        'txtCodCliente
        '
        Me.txtCodCliente.AcceptsReturn = True
        Me.txtCodCliente.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodCliente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodCliente.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodCliente.Location = New System.Drawing.Point(94, 149)
        Me.txtCodCliente.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodCliente.MaxLength = 5
        Me.txtCodCliente.Name = "txtCodCliente"
        Me.txtCodCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodCliente.Size = New System.Drawing.Size(41, 20)
        Me.txtCodCliente.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtCodCliente, "Codigo del Cliente")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFechaInicial)
        Me.Frame1.Controls.Add(Me.dtpFechaFinal)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(12, 223)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(397, 46)
        Me.Frame1.TabIndex = 11
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo ..."
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(98, 17)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(96, 20)
        Me.dtpFechaInicial.TabIndex = 9
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(290, 17)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaFinal.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(40, 20)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(64, 17)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Desde el :"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(238, 20)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(58, 15)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Hasta el :"
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(144, 57)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(200, 21)
        Me.dbcSucursal.TabIndex = 3
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(12, 110)
        Me.Line1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(302, 1)
        Me.Line1.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(42, 59)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(61, 17)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Sucursal :"
        '
        'Line2
        '
        Me.Line2.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line2.Location = New System.Drawing.Point(12, 110)
        Me.Line2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Line2.Name = "Line2"
        Me.Line2.Size = New System.Drawing.Size(302, 1)
        Me.Line2.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(42, 151)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(47, 17)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Cliente :"
        '
        'Line3
        '
        Me.Line3.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line3.Location = New System.Drawing.Point(12, 210)
        Me.Line3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Line3.Name = "Line3"
        Me.Line3.Size = New System.Drawing.Size(302, 1)
        Me.Line3.TabIndex = 17
        '
        'Line4
        '
        Me.Line4.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line4.Location = New System.Drawing.Point(12, 210)
        Me.Line4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Line4.Name = "Line4"
        Me.Line4.Size = New System.Drawing.Size(302, 1)
        Me.Line4.TabIndex = 18
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(130, 290)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 106
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(15, 290)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 105
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(245, 291)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 104
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmVtasReportedeCuentasporCobrar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(419, 341)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.txtNombre)
        Me.Controls.Add(Me.chkSoloVigentes)
        Me.Controls.Add(Me.chkCuentasSaldadas)
        Me.Controls.Add(Me.optResumenGeneral)
        Me.Controls.Add(Me.chkTodaslasSucursales)
        Me.Controls.Add(Me.txtCodSucursal)
        Me.Controls.Add(Me.optDetalladoporCliente)
        Me.Controls.Add(Me.txtCodCliente)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.Line1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Line2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Line3)
        Me.Controls.Add(Me.Line4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(334, 175)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasReportedeCuentasporCobrar"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Cuentas por Cobrar"
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

End Class