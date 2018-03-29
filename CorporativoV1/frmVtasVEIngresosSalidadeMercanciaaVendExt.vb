Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasVEIngresosSalidadeMercanciaaVendExt
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE INGRESOS POR SALIDA DE MERCANCIA A VENDEDORES EXTERNOS                            *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :                                                                                                   *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents txtCodVendExterno As System.Windows.Forms.TextBox
    Public WithEvents chkTodosLosVendedores As System.Windows.Forms.CheckBox
    Public WithEvents Label1 As System.Windows.Forms.Label

    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim rsReporte As ADODB.Recordset
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCodVendExterno = New System.Windows.Forms.TextBox()
        Me.chkTodosLosVendedores = New System.Windows.Forms.CheckBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtCodVendExterno
        '
        Me.txtCodVendExterno.AcceptsReturn = True
        Me.txtCodVendExterno.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodVendExterno.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodVendExterno.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodVendExterno.Location = New System.Drawing.Point(89, 33)
        Me.txtCodVendExterno.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodVendExterno.MaxLength = 3
        Me.txtCodVendExterno.Name = "txtCodVendExterno"
        Me.txtCodVendExterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodVendExterno.Size = New System.Drawing.Size(26, 20)
        Me.txtCodVendExterno.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCodVendExterno, "Codigo del Vendedor Externo")
        '
        'chkTodosLosVendedores
        '
        Me.chkTodosLosVendedores.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodosLosVendedores.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodosLosVendedores.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodosLosVendedores.Location = New System.Drawing.Point(12, 6)
        Me.chkTodosLosVendedores.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodosLosVendedores.Name = "chkTodosLosVendedores"
        Me.chkTodosLosVendedores.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodosLosVendedores.Size = New System.Drawing.Size(138, 17)
        Me.chkTodosLosVendedores.TabIndex = 0
        Me.chkTodosLosVendedores.Text = "Todos los Vendedores"
        Me.ToolTip1.SetToolTip(Me.chkTodosLosVendedores, "Muestra Todos los Vendedores Externos")
        Me.chkTodosLosVendedores.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFechaInicial)
        Me.Frame1.Controls.Add(Me.dtpFechaFinal)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(12, 65)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(351, 53)
        Me.Frame1.TabIndex = 6
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(67, 18)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(96, 20)
        Me.dtpFechaInicial.TabIndex = 3
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(238, 17)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(95, 20)
        Me.dtpFechaFinal.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(186, 19)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(64, 17)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Hasta el :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(12, 20)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(59, 17)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Desde el :"
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(118, 33)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(218, 21)
        Me.dbcSucursal.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(24, 33)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(65, 17)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Vendedor :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(128, 133)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 94
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(13, 133)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 93
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasVEIngresosSalidadeMercanciaaVendExt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(375, 177)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.txtCodVendExterno)
        Me.Controls.Add(Me.chkTodosLosVendedores)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasVEIngresosSalidadeMercanciaaVendExt"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Ingresos por Entrega de Mercancia a Vendedores Externos"
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Sub Imprime()
        Dim RptVtasVEReportedeIngresosSalidadeMercancia As New RptVtasVEReportedeIngresosSalidadeMercancia

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        Dim Sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim PeriodoReporte As String
        Dim strWhere As String
        Dim strWhere2 As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim RsAux As ADODB.Recordset
        Dim SubTotal As String
        Dim Descuento As String
        Dim Iva As String
        Dim Total As String
        Dim SubTotalPesos As String
        Dim DescuentoPesos As String
        Dim IvaPesos As String
        Dim TotalPesos As String
        On Error GoTo ImprimeErr
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
        If CDbl(Numerico(txtCodVendExterno.Text)) = 0 And chkTodosLosVendedores.CheckState = 0 Then
            MsgBox("Proporcione un Codigo de Vendedor Externo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodVendExterno.Focus()
            Exit Sub
        End If
        If Trim(dbcSucursal.Text) = "" And chkTodosLosVendedores.CheckState = 0 Then
            MsgBox("Proprcione la Descripción del Vendedor Externo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcSucursal.Focus()
            Exit Sub
        End If

        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = UCase("Ingresos por Entrega de Mercancía a Vendedores Externos")
        FechaInicial = AgregarHoraAFecha(dtpFechaInicial.Value)
        FechaFinal = AgregarHoraAFecha(dtpFechaFinal.Value)
        'FechaInicial = Format(Month(dtpFechaInicial.Value), "00") & "/" & Format((dtpFechaInicial.Value), "00") & "/" & Format(Year(dtpFechaInicial.Value), "0000")
        'FechaFinal = Format(Month(dtpFechaFinal.Value), "00") & "/" & Format((dtpFechaFinal.Value), "00") & "/" & Format(Year(dtpFechaFinal.Value), "0000")
        'PeriodoReporte = "Del " & Format(dtpFechaInicial.Value, "dd/mmm/yyyy") & " al " & Format(dtpFechaFinal.Value, "dd/mmm/yyyy")
        PeriodoReporte = "Del " & FechaInicial & " al " & FechaFinal

        strWhere = ""
        If chkTodosLosVendedores.CheckState = 1 Then
            strWhere = "WHERE MOVCAB.CODMOVTOALM = " & C_SalidaPorVentadeVendedoresExternos & " AND VTACAB.VTAVEXT = 1 " & "AND VTACAB.ESTATUS <> 'C' AND MOVCAB.ESTATUS <> 'C' AND ISNULL(DEV.ESTATUS,'') <> 'C' AND " & "MOVCAB.FECHAALMACEN BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' "
        ElseIf chkTodosLosVendedores.CheckState = 0 Then
            strWhere = "WHERE MOVCAB.CODALMACEN = " & txtCodVendExterno.Text & " AND MOVCAB.CODMOVTOALM = " & C_SalidaPorVentadeVendedoresExternos & " AND " & "VTACAB.VTAVEXT = 1 AND VTACAB.ESTATUS <> 'C' AND MOVCAB.ESTATUS <> 'C' AND ISNULL(DEV.ESTATUS,'') <> 'C' AND " & "MOVCAB.FECHAALMACEN BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' "
        End If

        Sql = "SELECT SUC.CODALMACEN,SUC.DESCALMACEN,VTACAB.FOLIOVENTA,VTACAB.FECHAVENTA,VTACAB.FOLIOFACTURA," & "ISNULL(CASE VTACAB.CONDICION WHEN 'CO' THEN CASE WHEN ISNULL(DEV.TOTALDEVOL,0) <> 0 THEN " & "((VTACAB.TOTAL + VTACAB.REDONDEO) - CASE WHEN DEV.CANTVENDIDA - DEV.CANTIDADDEVOL = 0 THEN DEV.TOTALDEVOL + DEV.TOTALRED ELSE DEV.TOTALDEVOL END) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) END END,0) AS VENTACONTADODOLARES," & "ISNULL(CASE VTACAB.CONDICION WHEN 'CR' THEN CASE WHEN ISNULL(DEV.TOTALDEVOL,0) <> 0 THEN " & "((VTACAB.TOTAL + VTACAB.REDONDEO) - CASE WHEN DEV.CANTVENDIDA - DEV.CANTIDADDEVOL = 0 THEN DEV.TOTALDEVOL + DEV.TOTALRED ELSE DEV.TOTALDEVOL END) ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) END END,0) AS VENTACREDITODOLARES," & "ISNULL(CASE VTACAB.CONDICION WHEN 'CO' THEN CASE WHEN ISNULL(DEV.TOTALDEVOL,0) <> 0 THEN " & "ROUND((((VTACAB.TOTAL + VTACAB.REDONDEO) - CASE WHEN DEV.CANTVENDIDA - DEV.CANTIDADDEVOL = 0 THEN DEV.TOTALDEVOL + DEV.TOTALRED ELSE DEV.TOTALDEVOL END) * VTACAB.TIPOCAMBIO),1) ELSE ROUND(((VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO),1) END END,0) AS TOTALCONTADOPESOS," & "ISNULL(CASE VTACAB.CONDICION WHEN 'CR' THEN CASE WHEN ISNULL(DEV.TOTALDEVOL,0) <> 0 THEN " & "ROUND((((VTACAB.TOTAL + VTACAB.REDONDEO) - CASE WHEN DEV.CANTVENDIDA - DEV.CANTIDADDEVOL = 0 THEN DEV.TOTALDEVOL + DEV.TOTALRED ELSE DEV.TOTALDEVOL END) * VTACAB.TIPOCAMBIO),1) ELSE ROUND(((VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO),1) END END,0) AS TOTALCREDITOPESOS," & "ISNULL(CASE WHEN ISNULL(DEV.TOTALDEVOL,0) <> 0 THEN ((VTACAB.TOTAL + VTACAB.REDONDEO) - CASE WHEN DEV.CANTVENDIDA - DEV.CANTIDADDEVOL = 0 THEN DEV.TOTALDEVOL + DEV.TOTALRED ELSE DEV.TOTALDEVOL END) " & "ELSE (VTACAB.TOTAL + VTACAB.REDONDEO) END,0) AS TOTALDOLARES," & "ISNULL(CASE WHEN ISNULL(DEV.TOTALDEVOL,0) <> 0 THEN ROUND((((VTACAB.TOTAL + VTACAB.REDONDEO) - CASE WHEN DEV.CANTVENDIDA - DEV.CANTIDADDEVOL = 0 THEN DEV.TOTALDEVOL + DEV.TOTALRED ELSE DEV.TOTALDEVOL END) * VTACAB.TIPOCAMBIO),1) " & "ELSE ROUND(((VTACAB.TOTAL + VTACAB.REDONDEO) * VTACAB.TIPOCAMBIO),1) END,0) AS TOTALPESOS," & "ISNULL(DEV.ESTATUS,'') AS ESTATUS " & "FROM MOVIMIENTOSVENTASCAB VTACAB INNER JOIN MOVTOSALMACENCAB MOVCAB ON VTACAB.FOLIOVENTA = MOVCAB.FOLIOVENTA " & "LEFT OUTER JOIN (SELECT MC.FOLIOVENTA,SUM(CASE WHEN D.NUMPARTIDA = 1 THEN ISNULL(C.TOTALDEVOL,0) ELSE 0 END) AS TOTALDEVOL,SUM(CASE WHEN D.NUMPARTIDA = 1 THEN ISNULL(C.REDONDEODEV,0) ELSE 0 END) AS TOTALRED,SUM(MD.CANTIDAD) AS CANTVENDIDA," & "SUM(ISNULL(D.CANTIDADDEVOL,0)) AS CANTIDADDEVOL,ISNULL(C.ESTATUS,'') AS ESTATUS FROM MOVIMIENTOSVENTASCAB MC INNER JOIN MOVIMIENTOSVENTASDET MD ON MC.FOLIOVENTA = MD.FOLIOVENTA " & "LEFT OUTER JOIN DEVOLUCIONESCAB C ON MC.FOLIOVENTA = C.FOLIOVENTA AND MD.FOLIOVENTA = C.FOLIOVENTA LEFT OUTER JOIN DEVOLUCIONESDET D ON C.FOLIODEVOLUCION = D.FOLIODEVOLUCION AND MD.CODARTICULO = D.CODARTICULO " & "WHERE C.ESTATUS <> 'C' AND MC.VTAVEXT = 1 GROUP BY MC.FOLIOVENTA,C.ESTATUS) DEV ON VTACAB.FOLIOVENTA = DEV.FOLIOVENTA " & "INNER JOIN CATCLIENTES CLI ON VTACAB.CODCLIENTE = CLI.CODCLIENTE " & "INNER JOIN CATALMACEN SUC ON CLI.ALMACENVEXT = SUC.CODALMACEN " & strWhere & "GROUP BY SUC.CODALMACEN,SUC.DESCALMACEN,VTACAB.FOLIOVENTA,VTACAB.FECHAVENTA,VTACAB.FOLIOFACTURA,VTACAB.TOTAL," & "VTACAB.Condicion , DEV.TOTALDEVOL,DEV.TOTALRED,DEV.CANTVENDIDA,DEV.CANTIDADDEVOL, DEV.Estatus, VTACAB.Redondeo,VtaCab.tipocambio " & "ORDER BY VTACAB.FECHAVENTA,VTACAB.FOLIOVENTA"
        BorraCmd()
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = Sql
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No Existen Movimientos En Este Periodo de Fechas, Favor de Verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        Else
            RptVtasVEReportedeIngresosSalidadeMercancia.SetDataSource(frmReportes.rsReport)
        End If

        If chkTodosLosVendedores.CheckState = 1 Then
            strWhere = "WHERE CAB.CODMOVTOALM = " & C_SalidaPorVentadeVendedoresExternos & " AND VTACAB.VTAVEXT = 1 AND VTACAB.ESTATUS <> 'C' "
            strWhere2 = "WHERE MOVCAB.CODMOVTOALM = " & C_SalidaPorVentadeVendedoresExternos & " AND " & "VTACAB.VTAVEXT = 1 AND CAB.ESTATUS <> 'C' "
        ElseIf chkTodosLosVendedores.CheckState = 0 Then
            strWhere = "WHERE CAB.CODALMACEN = " & txtCodVendExterno.Text & " AND CAB.CODMOVTOALM = " & C_SalidaPorVentadeVendedoresExternos & " AND " & "VTACAB.VTAVEXT = 1 AND VTACAB.ESTATUS <> 'C' "
            strWhere2 = "WHERE MOVCAB.CODALMACEN = " & txtCodVendExterno.Text & " AND MOVCAB.CODMOVTOALM = " & C_SalidaPorVentadeVendedoresExternos & " AND " & "VTACAB.VTAVEXT = 1 AND CAB.ESTATUS <> 'C' "
        End If
        Sql = "SELECT SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN SUBTOTALDOLARES - CASE WHEN (CANTIDAD - CANTIDADDEVOL) = 0 THEN SUBTOTAL + REDONDEODEV ELSE SUBTOTAL END ELSE SUBTOTALDOLARES END) AS SUBTOTAL," & "SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN DESCUENTODOLARES - DESCUENTOS ELSE DESCUENTODOLARES END) AS DESCUENTO," & "SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN IVADOLARES - IVA ELSE IVADOLARES END) AS IVA," & "((SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN SUBTOTALDOLARES - CASE WHEN (CANTIDAD - CANTIDADDEVOL) = 0 THEN SUBTOTAL + REDONDEODEV ELSE SUBTOTAL END ELSE SUBTOTALDOLARES END) - SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN DESCUENTODOLARES - DESCUENTOS ELSE DESCUENTODOLARES END)) + SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN IVADOLARES - IVA ELSE IVADOLARES END)) AS TOTALDOLARES," & "ROUND(SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN (SUBTOTALDOLARES - CASE WHEN (CANTIDAD - CANTIDADDEVOL) = 0 THEN SUBTOTAL + REDONDEODEV ELSE SUBTOTAL END) * TIPOCAMBIO ELSE (SUBTOTALDOLARES * TIPOCAMBIO) END),1) AS SUBTOTALPESOS," & "ROUND(SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN (DESCUENTODOLARES - DESCUENTOS) * TIPOCAMBIO ELSE DESCUENTODOLARES * TIPOCAMBIO END),1) AS DESCUENTOPESOS," & "ROUND(SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN (IVADOLARES - IVA) * TIPOCAMBIO ELSE (IVADOLARES * TIPOCAMBIO) END),1) AS IVAPESOS," & "ROUND((SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN (SUBTOTALDOLARES - CASE WHEN (CANTIDAD - CANTIDADDEVOL) = 0 THEN SUBTOTAL + REDONDEODEV ELSE SUBTOTAL END) * TIPOCAMBIO ELSE (SUBTOTALDOLARES * TIPOCAMBIO) END) - SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN (DESCUENTODOLARES - DESCUENTOS) * TIPOCAMBIO ELSE (DESCUENTODOLARES * TIPOCAMBIO) END)) + SUM(CASE WHEN VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA THEN (IVADOLARES - IVA) * TIPOCAMBIO ELSE (IVADOLARES * TIPOCAMBIO) END),1) AS TOTALPESOS " & "FROM " & "(SELECT VTACAB.FOLIOVENTA,VTACAB.FECHAVENTA,(VTACAB.SUBTOTAL + VTACAB.REDONDEO) AS SUBTOTALDOLARES," & "(VTACAB.DESCUENTO) AS DESCUENTODOLARES," & "(VTACAB.IVA) AS IVADOLARES,VTACAB.TIPOCAMBIO " & "FROM MOVIMIENTOSVENTASCAB VTACAB INNER JOIN MOVTOSALMACENCAB CAB ON VTACAB.FOLIOVENTA = CAB.FOLIOVENTA " & strWhere & ") VENTAS " & "LEFT OUTER JOIN " & "(SELECT VTACAB.FOLIOVENTA,ROUND(SUM((ISNULL(DET.PRECIOLISTASINIVA,0) * ISNULL(DET.CANTIDADDEVOL,0))),3) AS SUBTOTAL,ROUND(SUM((VTADET.IMPTEPROMOCIONES + VTADET.IMPTEDESCUENTOS) * ISNULL(DET.CANTIDADDEVOL,0)),3) AS DESCUENTOS," & "ROUND(SUM(ISNULL(DET.IVAREAL,0) * ISNULL(DET.CANTIDADDEVOL,0)),3) AS IVA,SUM(VTADET.CANTIDAD) AS CANTIDAD,SUM(ISNULL(DET.CANTIDADDEVOL,0)) AS CANTIDADDEVOL,SUM(CASE WHEN DET.NUMPARTIDA = 1 THEN ISNULL(CAB.REDONDEODEV,0) ELSE 0 END) AS REDONDEODEV " & "FROM MOVIMIENTOSVENTASCAB VTACAB INNER JOIN MOVIMIENTOSVENTASDET VTADET ON VTACAB.FOLIOVENTA = VTADET.FOLIOVENTA " & "INNER JOIN MOVTOSALMACENCAB MOVCAB ON VTACAB.FOLIOVENTA = MOVCAB.FOLIOVENTA AND VTADET.FOLIOVENTA = MOVCAB.FOLIOVENTA " & "LEFT OUTER JOIN DEVOLUCIONESCAB CAB ON VTACAB.FOLIOVENTA = CAB.FOLIOVENTA AND VTADET.FOLIOVENTA = CAB.FOLIOVENTA " & "LEFT OUTER JOIN DEVOLUCIONESDET DET ON CAB.FOLIODEVOLUCION = DET.FOLIODEVOLUCION AND VTADET.CODARTICULO = DET.CODARTICULO " & strWhere2 & "GROUP BY VTACAB.FOLIOVENTA) DEVOLUCIONES " & "ON VENTAS.FOLIOVENTA = DEVOLUCIONES.FOLIOVENTA " & "WHERE FECHAVENTA BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
        frmReportes.rsReport = Cmd.Execute
        If frmReportes.rsReport.RecordCount > 0 Then
            SubTotal = Format(frmReportes.rsReport.Fields("SubTotal").Value, "###,##0.00")
            SubTotalPesos = Format(frmReportes.rsReport.Fields("SubTotalPesos").Value, "###,##0.00")
            Descuento = Format(frmReportes.rsReport.Fields("Descuento").Value, "###,##0.00")
            DescuentoPesos = Format(frmReportes.rsReport.Fields("DescuentoPesos").Value, "###,##0.00")
            Iva = Format(frmReportes.rsReport.Fields("Iva").Value, "###,##0.00")
            IvaPesos = Format(frmReportes.rsReport.Fields("IvaPesos").Value, "###,##0.00")
            Total = Format(frmReportes.rsReport.Fields("TotalDolares").Value, "###,##0.00")
            TotalPesos = Format(frmReportes.rsReport.Fields("TotalPesos").Value, "###,##0.00")
        End If
        'frmReportes.Report = RptVtasVEReportedeIngresosSalidadeMercancia 

        If (NombreEmpresa <> Nothing) Then
            pdvNum.Value = NombreEmpresa : pvNum.Add(pdvNum)
            RptVtasVEReportedeIngresosSalidadeMercancia.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
        End If

        If (NombreReporte <> Nothing) Then
            pdvNum.Value = NombreReporte : pvNum.Add(pdvNum)
            RptVtasVEReportedeIngresosSalidadeMercancia.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
        End If

        If (PeriodoReporte <> Nothing) Then
            pdvNum.Value = PeriodoReporte : pvNum.Add(pdvNum)
            RptVtasVEReportedeIngresosSalidadeMercancia.DataDefinition.ParameterFields("Periodo").ApplyCurrentValues(pvNum)
        End If

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.rsReport = rsReporte
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "PeriodoReporte", "SubTotal", "Descuento", "Iva", "Total", "SubTotalPesos", "DescuentoPesos", "IvaPesos", "TotalPesos"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte, SubTotal, Descuento, Iva, Total, SubTotalPesos, DescuentoPesos, IvaPesos, TotalPesos}
        frmReportes.Text = "Ingresos por Entrega de Mercancía a Vendedores Externos"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.reporteActual = RptVtasVEReportedeIngresosSalidadeMercancia
        frmReportes.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        FueraChange = False
        Exit Sub
ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error al Imprimir : " & Err.Description, MsgBoxStyle.Exclamation, "Error de Operacion")
        FueraChange = False
    End Sub

    Sub BuscaVendedorExterno()
        On Error GoTo Merr
        gStrSql = "SELECT DescAlmacen,TipoAlmacen FROM CatAlmacen WHERE CodAlmacen = " & txtCodVendExterno.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("TipoAlmacen").Value = "P" Then
                MsgBox("Este Almacen no es un Vendedor Externo, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodVendExterno.Text = ""
                txtCodVendExterno.Focus()
                Exit Sub
            Else
                txtCodVendExterno.Text = txtCodVendExterno.Text
                dbcSucursal.Text = RsGral.Fields("DescAlmacen").Value
            End If
        Else
            MsgBox("Codigo de Almacen no Existe, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodVendExterno.Text = ""
            txtCodVendExterno.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Limpiar()
        Nuevo()
        InicializaVariables()
        chkTodosLosVendedores.Focus()
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
    End Sub

    Sub Nuevo()
        chkTodosLosVendedores.CheckState = System.Windows.Forms.CheckState.Checked
        txtCodVendExterno.Text = ""
        dbcSucursal.Text = ""
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
    End Sub

    Private Sub chkTodosLosVendedores_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodosLosVendedores.CheckStateChanged
        If chkTodosLosVendedores.CheckState = 1 Then
            txtCodVendExterno.Text = ""
            txtCodVendExterno.Enabled = False
            dbcSucursal.Text = ""
            dbcSucursal.Enabled = False
        ElseIf chkTodosLosVendedores.CheckState = 0 Then
            txtCodVendExterno.Enabled = True
            dbcSucursal.Enabled = True
        End If
    End Sub

    Private Sub chkTodosLosVendedores_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodosLosVendedores.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        If Trim(dbcSucursal.Text) = "" Then
            txtCodVendExterno.Text = ""
            Exit Sub
        End If
        gStrSql = "SELECT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'V' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursal)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            txtCodVendExterno.Focus()
        End If
    End Sub

    Private Sub dbcSucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcSucursal.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcSucursal_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        'If dbcSucursal.SelectedItem <> 0 Then
        dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        'End If
        FueraChange = True
        dbcSucursal.Text = Aux
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        FueraChange = True
        gStrSql = "SELECT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
        txtCodVendExterno.Text = IIf(intCodSucursal <> 0, intCodSucursal, "")
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

    Private Sub dtpFechaFinal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaFinal.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaFinal.Click
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaFinal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFinal.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaFinal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaFinal.KeyPress
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Click
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaInicial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaInicial.KeyPress
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub frmVtasVEIngresosSalidadeMercanciaaVendExt_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasVEIngresosSalidadeMercanciaaVendExt_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasVEIngresosSalidadeMercanciaaVendExt_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "chkTodosLosVendedores" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmVtasVEIngresosSalidadeMercanciaaVendExt_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasVEIngresosSalidadeMercanciaaVendExt_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        InicializaVariables()
        dtpFechaInicial.MinDate = C_FECHAINICIAL
        dtpFechaInicial.MaxDate = C_FECHAFINAL
        dtpFechaFinal.MinDate = C_FECHAINICIAL
        dtpFechaFinal.MaxDate = C_FECHAFINAL
        Nuevo()
    End Sub

    Private Sub frmVtasVEIngresosSalidadeMercanciaaVendExt_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmVtasVEIngresosSalidadeMercanciaaVendExt_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtCodVendExterno_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodVendExterno.TextChanged
        If Trim(txtCodVendExterno.Text) = "" Then
            txtCodVendExterno.Text = ""
            dbcSucursal.Text = ""
        End If
    End Sub

    Private Sub txtCodVendExterno_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodVendExterno.Enter
        Pon_Tool()
        SelTextoTxt(txtCodVendExterno)
    End Sub

    Private Sub txtCodVendExterno_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodVendExterno.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodVendExterno_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodVendExterno.Leave
        If CDbl(Numerico(txtCodVendExterno.Text)) = 0 Then
            txtCodVendExterno.Text = ""
            dbcSucursal.Text = ""
        Else
            BuscaVendedorExterno()
        End If
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class