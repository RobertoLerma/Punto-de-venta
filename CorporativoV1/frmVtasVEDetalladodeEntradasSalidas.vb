Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasVEDetalladodeEntradasSalidas
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
    Public WithEvents optEntradas As System.Windows.Forms.RadioButton
    Public WithEvents optSalidas As System.Windows.Forms.RadioButton
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents txtCodVendExterno As System.Windows.Forms.TextBox
    Public WithEvents chkTodosLosVendedores As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
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
        Me.optEntradas = New System.Windows.Forms.RadioButton()
        Me.optSalidas = New System.Windows.Forms.RadioButton()
        Me.txtCodVendExterno = New System.Windows.Forms.TextBox()
        Me.chkTodosLosVendedores = New System.Windows.Forms.CheckBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'optEntradas
        '
        Me.optEntradas.BackColor = System.Drawing.SystemColors.Control
        Me.optEntradas.Cursor = System.Windows.Forms.Cursors.Default
        Me.optEntradas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optEntradas.Location = New System.Drawing.Point(180, 13)
        Me.optEntradas.Margin = New System.Windows.Forms.Padding(2)
        Me.optEntradas.Name = "optEntradas"
        Me.optEntradas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optEntradas.Size = New System.Drawing.Size(98, 17)
        Me.optEntradas.TabIndex = 1
        Me.optEntradas.TabStop = True
        Me.optEntradas.Text = "De Recepción"
        Me.ToolTip1.SetToolTip(Me.optEntradas, "Muestra el Reporte Detallado de Entradas")
        Me.optEntradas.UseVisualStyleBackColor = False
        '
        'optSalidas
        '
        Me.optSalidas.BackColor = System.Drawing.SystemColors.Control
        Me.optSalidas.Checked = True
        Me.optSalidas.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSalidas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSalidas.Location = New System.Drawing.Point(60, 13)
        Me.optSalidas.Margin = New System.Windows.Forms.Padding(2)
        Me.optSalidas.Name = "optSalidas"
        Me.optSalidas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSalidas.Size = New System.Drawing.Size(91, 17)
        Me.optSalidas.TabIndex = 0
        Me.optSalidas.TabStop = True
        Me.optSalidas.Text = "De Entrega"
        Me.ToolTip1.SetToolTip(Me.optSalidas, "Muestra el Reporte Detallado de Salidas")
        Me.optSalidas.UseVisualStyleBackColor = False
        '
        'txtCodVendExterno
        '
        Me.txtCodVendExterno.AcceptsReturn = True
        Me.txtCodVendExterno.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodVendExterno.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodVendExterno.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodVendExterno.Location = New System.Drawing.Point(95, 89)
        Me.txtCodVendExterno.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodVendExterno.MaxLength = 3
        Me.txtCodVendExterno.Name = "txtCodVendExterno"
        Me.txtCodVendExterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodVendExterno.Size = New System.Drawing.Size(33, 20)
        Me.txtCodVendExterno.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtCodVendExterno, "Codigo del Vendedor")
        '
        'chkTodosLosVendedores
        '
        Me.chkTodosLosVendedores.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodosLosVendedores.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodosLosVendedores.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodosLosVendedores.Location = New System.Drawing.Point(12, 65)
        Me.chkTodosLosVendedores.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodosLosVendedores.Name = "chkTodosLosVendedores"
        Me.chkTodosLosVendedores.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodosLosVendedores.Size = New System.Drawing.Size(136, 17)
        Me.chkTodosLosVendedores.TabIndex = 2
        Me.chkTodosLosVendedores.Text = "Todos los Vendedores"
        Me.ToolTip1.SetToolTip(Me.chkTodosLosVendedores, "Muestra todos Los Vendedores")
        Me.chkTodosLosVendedores.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optEntradas)
        Me.Frame2.Controls.Add(Me.optSalidas)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(12, 13)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(307, 40)
        Me.Frame2.TabIndex = 11
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Detallado"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFechaInicial)
        Me.Frame1.Controls.Add(Me.dtpFechaFinal)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(12, 121)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(379, 53)
        Me.Frame1.TabIndex = 8
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(68, 17)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(95, 20)
        Me.dtpFechaInicial.TabIndex = 5
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(270, 17)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaFinal.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(209, 22)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(57, 17)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Hasta el :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(7, 22)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(57, 17)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Desde el :"
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(132, 89)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(218, 21)
        Me.dbcSucursal.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(30, 91)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(61, 17)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Vendedor :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(127, 195)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 97
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(12, 195)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 96
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasVEDetalladodeEntradasSalidas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(402, 243)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.txtCodVendExterno)
        Me.Controls.Add(Me.chkTodosLosVendedores)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasVEDetalladodeEntradasSalidas"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte Detallado de Entrega/Recepción"
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Sub Imprime()
        Dim RptVtasVEReportedeEntradadeMercancia As New RptVtasVEReportedeEntradadeMercancias
        Dim RptVtasVEReportedeSalidadeMercancia As New RptVtasVEReportedeSalidadeMercancias

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        Dim Sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim PeriodoReporte As String
        Dim strWhere As String
        Dim FechaInicial As String
        Dim FechaFinal As String
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
        If optSalidas.Checked = True Then
            NombreReporte = UCase("Detallado de Entrega de Articulos a Vendedores Externos")
        Else
            NombreReporte = UCase("Detallado de Recepción de Articulos a Vendedores Externos")
        End If
        'FechaInicial = Format(Month(dtpFechaInicial.Value), "00") & "/" & Format((dtpFechaInicial.Value), "00") & "/" & Format(Year(dtpFechaInicial.Value), "0000")
        'FechaFinal = Format(Month(dtpFechaFinal.Value), "00") & "/" & Format((dtpFechaFinal.Value), "00") & "/" & Format(Year(dtpFechaFinal.Value), "0000")
        'PeriodoReporte = "Del  " & Format(dtpFechaInicial.Value, "dd/mmm/yyyy") & "  al  " & Format(dtpFechaFinal.Value, "dd/mmm/yyyy")
        FechaInicial = AgregarHoraAFecha(dtpFechaInicial.Value)
        FechaFinal = AgregarHoraAFecha(dtpFechaFinal.Value)
        PeriodoReporte = "Del " & FechaInicial & " al " & FechaFinal

        strWhere = ""

        If optSalidas.Checked = True Then
            If chkTodosLosVendedores.CheckState = 1 Then
                strWhere = "WHERE CAB.CODMOVTOALM = " & C_SalidaAVendedoresExternos & " AND CAB.FECHAALMACEN BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND CAB.ESTATUS <> 'C' "
            ElseIf chkTodosLosVendedores.CheckState = 0 Then
                strWhere = "WHERE CAB.CODALMACENREF = " & txtCodVendExterno.Text & " AND CAB.CODMOVTOALM = " & C_SalidaAVendedoresExternos & " AND CAB.FECHAALMACEN BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND CAB.ESTATUS <> 'C' "
            End If

            Sql = "SELECT SUC.CODALMACEN,SUC.DESCALMACEN,CAB.FOLIOALMACEN,CAB.FECHAALMACEN,DET.CODARTICULO,ART.DESCARTICULO,DET.CANTIDAD,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) + '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS ANTERIOR " & "FROM MOVTOSALMACENCAB CAB INNER JOIN MOVTOSALMACENDET DET ON CAB.FOLIOALMACEN = DET.FOLIOALMACEN " & "INNER JOIN CATARTICULOS ART ON DET.CODARTICULO = ART.CODARTICULO " & "INNER JOIN CATALMACEN SUC ON CAB.CODALMACENREF = SUC.CODALMACEN " & strWhere & "ORDER BY CAB.FECHAALMACEN,CAB.FOLIOALMACEN,SUC.DESCALMACEN"
            BorraCmd()
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd.CommandText = Sql
            frmReportes.rsReport = Cmd.Execute

            If frmReportes.rsReport.RecordCount = 0 Then
                MsgBox("No Existen Movimientos En Este Periodo de Fechas, Favor de Verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Exit Sub
            Else
                RptVtasVEReportedeSalidadeMercancia.SetDataSource(frmReportes.rsReport)
            End If

            'If (NombreEmpresa <> Nothing) Then
            '    pdvNum.Value = NombreEmpresa : pvNum.Add(pdvNum)
            '    RptVtasVEReportedeSalidadeMercancia.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
            'End If

            'If (NombreReporte <> Nothing) Then
            '    pdvNum.Value = NombreReporte : pvNum.Add(pdvNum)
            '    RptVtasVEReportedeSalidadeMercancia.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
            'End If

            'If (PeriodoReporte <> Nothing) Then
            '    pdvNum.Value = PeriodoReporte : pvNum.Add(pdvNum)
            '    RptVtasVEReportedeSalidadeMercancia.DataDefinition.ParameterFields("Periodo").ApplyCurrentValues(pvNum)
            'End If


            If optSalidas.Checked = True Then
                frmReportes.Text = "Detallado de Entrega de Mercancia a Vendedores Externos"
            Else
                frmReportes.Text = "Detallado de Recepción de Mercancia a Vendedores Externos"
            End If
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            frmReportes.reporteActual = RptVtasVEReportedeSalidadeMercancia
            frmReportes.Show()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            FueraChange = False
            Exit Sub


        ElseIf optEntradas.Checked = True Then
            If chkTodosLosVendedores.CheckState = 1 Then
                strWhere = "WHERE CAB.CODMOVTOALM = " & C_EntradaPorDevoluciondeVendedoresExternos & " AND CAB.FECHAALMACEN BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND CAB.ESTATUS <> 'C' "
            ElseIf chkTodosLosVendedores.CheckState = 0 Then
                strWhere = "WHERE CAB.CODALMACENREF = " & txtCodVendExterno.Text & " AND CAB.CODMOVTOALM = " & C_EntradaPorDevoluciondeVendedoresExternos & " AND CAB.FECHAALMACEN BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND CAB.ESTATUS <> 'C' "
            End If

            Sql = "SELECT SUC.CODALMACEN,SUC.DESCALMACEN,CAB.FOLIOALMACEN,CAB.FECHAALMACEN,DET.CODARTICULO,ART.DESCARTICULO,DET.CANTIDAD,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) + '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS ANTERIOR " & "FROM MOVTOSALMACENCAB CAB INNER JOIN MOVTOSALMACENDET DET ON CAB.FOLIOALMACEN = DET.FOLIOALMACEN " & "INNER JOIN CATARTICULOS ART ON DET.CODARTICULO = ART.CODARTICULO " & "INNER JOIN CATALMACEN SUC ON CAB.CODALMACENREF = SUC.CODALMACEN " & strWhere & "ORDER BY CAB.FECHAALMACEN,CAB.FOLIOALMACEN,SUC.DESCALMACEN"
            BorraCmd()
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd.CommandText = Sql
            frmReportes.rsReport = Cmd.Execute

            If frmReportes.rsReport.RecordCount = 0 Then
                MsgBox("No Existen Movimientos En Este Periodo de Fechas, Favor de Verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Exit Sub
            Else
                RptVtasVEReportedeEntradadeMercancia.SetDataSource(frmReportes.rsReport)
            End If

            'If (NombreEmpresa <> Nothing) Then
            '    pdvNum.Value = NombreEmpresa : pvNum.Add(pdvNum)
            '    RptVtasVEReportedeEntradadeMercancia.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
            'End If

            'If (NombreReporte <> Nothing) Then
            '    pdvNum.Value = NombreReporte : pvNum.Add(pdvNum)
            '    RptVtasVEReportedeEntradadeMercancia.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
            'End If

            'If (PeriodoReporte <> Nothing) Then
            '    pdvNum.Value = PeriodoReporte : pvNum.Add(pdvNum)
            '    RptVtasVEReportedeEntradadeMercancia.DataDefinition.ParameterFields("Periodo").ApplyCurrentValues(pvNum)
            'End If


            If optSalidas.Checked = True Then
                frmReportes.Text = "Detallado de Entrega de Mercancia a Vendedores Externos"
            Else
                frmReportes.Text = "Detallado de Recepción de Mercancia a Vendedores Externos"
            End If
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            frmReportes.reporteActual = RptVtasVEReportedeEntradadeMercancia
            frmReportes.Show()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            FueraChange = False
            Exit Sub

        End If


        'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.rsReport = rsReporte
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "PeriodoReporte"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte}

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
        optSalidas.Focus()
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
    End Sub

    Sub Nuevo()
        optSalidas.Checked = True
        chkTodosLosVendedores.CheckState = System.Windows.Forms.CheckState.Unchecked
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

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
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

    Private Sub dtpFechaInicial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.CursorChanged       'sglTiempoCambio = VB.Timer()
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

    Private Sub frmVtasVEDetalladodeEntradasSalidas_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasVEDetalladodeEntradasSalidas_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasVEDetalladodeEntradasSalidas_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optSalidas" And System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optEntradas" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmVtasVEDetalladodeEntradasSalidas_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasVEDetalladodeEntradasSalidas_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmVtasVEDetalladodeEntradasSalidas_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmVtasVEDetalladodeEntradasSalidas_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optEntradas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optEntradas.Enter
        Pon_Tool()
    End Sub

    Private Sub optSalidas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSalidas.Enter
        Pon_Tool()
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