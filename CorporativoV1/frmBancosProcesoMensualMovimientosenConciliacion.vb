Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports VB = Microsoft.VisualBasic

Public Class frmBancosProcesoMensualMovimientosenConciliacion
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE CONCILIACION MENSUAL                                                              *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      JUEVES 07 DE AGOSTO DE 2003                                                                  *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmbMes As System.Windows.Forms.ComboBox
    Public WithEvents cmbAño As System.Windows.Forms.ComboBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents dbcBanco As System.Windows.Forms.ComboBox
    Public WithEvents dbcCuentaBancaria As System.Windows.Forms.ComboBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmbMes = New System.Windows.Forms.ComboBox()
        Me.cmbAño = New System.Windows.Forms.ComboBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me.dbcCuentaBancaria = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbMes
        '
        Me.cmbMes.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMes.Items.AddRange(New Object() {"01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"})
        Me.cmbMes.Location = New System.Drawing.Point(105, 16)
        Me.cmbMes.Name = "cmbMes"
        Me.cmbMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMes.Size = New System.Drawing.Size(185, 21)
        Me.cmbMes.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.cmbMes, "Mes.")
        '
        'cmbAño
        '
        Me.cmbAño.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAño.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAño.Location = New System.Drawing.Point(105, 43)
        Me.cmbAño.Name = "cmbAño"
        Me.cmbAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAño.Size = New System.Drawing.Size(185, 21)
        Me.cmbAño.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.cmbAño, "Año.")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Frame3)
        Me.Frame1.Controls.Add(Me.Frame2)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(16, 16)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(345, 209)
        Me.Frame1.TabIndex = 4
        Me.Frame1.TabStop = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.cmbMes)
        Me.Frame3.Controls.Add(Me.cmbAño)
        Me.Frame3.Controls.Add(Me.Label4)
        Me.Frame3.Controls.Add(Me.Label5)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(16, 112)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(313, 81)
        Me.Frame3.TabIndex = 8
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Información del Periodo"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(64, 18)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(33, 21)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Mes :"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(64, 45)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(33, 21)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Año :"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.dbcBanco)
        Me.Frame2.Controls.Add(Me.dbcCuentaBancaria)
        Me.Frame2.Controls.Add(Me.Label1)
        Me.Frame2.Controls.Add(Me.Label2)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(16, 16)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(313, 89)
        Me.Frame2.TabIndex = 5
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Información de la Cuenta Bancaria"
        '
        'dbcBanco
        '
        Me.dbcBanco.Location = New System.Drawing.Point(105, 24)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(185, 21)
        Me.dbcBanco.TabIndex = 0
        '
        'dbcCuentaBancaria
        '
        Me.dbcCuentaBancaria.Location = New System.Drawing.Point(105, 51)
        Me.dbcCuentaBancaria.Name = "dbcCuentaBancaria"
        Me.dbcCuentaBancaria.Size = New System.Drawing.Size(185, 21)
        Me.dbcCuentaBancaria.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(49, 21)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Banco :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(89, 21)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Cuenta Bancaria :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(131, 240)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 71
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.Location = New System.Drawing.Point(16, 240)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 70
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoMensualMovimientosenConciliacion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(377, 284)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoMensualMovimientosenConciliacion"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Movimientos en Conciliación"
        Me.Frame1.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub


    'Variables
    Dim mblnSALIR As Boolean
    Dim intCodBanco As Integer
    Dim tecla As Integer
    'Dim rsReporte As ADODB.Recordset

    Sub Imprime()

        Dim RptBancosProcesoMensualMovimientosenConciliacion As New RptBancosProcesoMensualMovimientosenConciliacion
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        Dim sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim Periodo As String
        Dim Ejercicio As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        'On Error GoTo ImprimeErr

        If Trim(dbcBanco.Text) = "" Then
            MsgBox("Proporcione el Nombre del Banco.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcBanco.Focus()
            Exit Sub
        End If
        If Trim(dbcCuentaBancaria.Text) = "" Then
            MsgBox("Proporcione una Cuenta Bancaria.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcCuentaBancaria.Focus()
            Exit Sub
        End If

        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = UCase("Movimientos en Conciliación")
        ObtenerLimitedeFechas(CInt(VB.Left(Trim(cmbMes.Text), 2)), CInt(Trim(cmbAño.Text)), FechaInicial, FechaFinal)
        Periodo = Mid(cmbMes.Text, 5, 12)
        Ejercicio = cmbAño.Text

        '    sql = "SELECT ISNULL(CASE WHEN MovBancarios.FolioMovto = MovCancelados.Referencia THEN 'C' ELSE CASE WHEN MovBancarios.Conciliado = 1 THEN '*' END END,'') AS Estatus ,MovBancarios.FolioMovto," & _
        ''    "ISNULL(MovCancelados.Referencia,'') AS FolioCancelacion,MovBancarios.FechaMovto,MovBancarios.Referencia,MovBancarios.Concepto," & _
        ''    "(CASE MovBancarios.Movimiento WHEN '" & C_MOVPAGO & "' THEN 'PAGOS' WHEN '" & C_MOVDEPOSITO & "' THEN 'DEPOSITOS' WHEN '" & C_MOVTRASPASO & "' THEN 'TRASP. BANC.' WHEN '" & C_MOVCARGOS & "' THEN 'CARGOS DIV.' WHEN '" & C_MOVANTICIPOS & "' THEN 'ANT. PROV./ACREED.' END) AS Movimiento," & _
        ''    "(CASE WHEN MovBancarios.FolioMovto = MovCancelados.Referencia THEN 0 ELSE MovBancarios.Ingresos END) AS Ingresos,(CASE WHEN MovBancarios.FolioMovto = MovCancelados.Referencia THEN 0 ELSE MovBancarios.Egresos END) AS Egresos " & _
        ''    "FROM ((SELECT FechaMovto,FolioMovto,Referencia,Concepto,Movimiento,ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVINGRESO & "' THEN Importe END,0) AS Ingresos,ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVEGRESO & "' THEN Importe END,0) AS Egresos,Conciliado " & _
        ''    "FROM MovimientosBancarios WHERE Movimiento <> '" & C_MOVCANCELACION & "' AND FechaMovto <= '" & FechaFinal & "' AND FechaMovto >= '" & FechaInicial & "' AND CodBanco = " & intCodBanco & "AND CtaBancaria = '" & dbcCuentaBancaria & "') UNION " & _
        ''    "(SELECT FechaMovto,FolioMovto,Referencia,Concepto,Movimiento,ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVINGRESO & "' THEN Importe END,0) AS Ingresos,ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVEGRESO & "' THEN Importe END,0) AS Egresos,Conciliado " & _
        ''    "FROM MovimientosBancarios WHERE Movimiento <> '" & C_MOVCANCELACION & "' AND ((FechaMovto < '" & FechaInicial & "' AND Conciliado = 0) OR FechaConciliacion = '" & FechaFinal & "') AND CodBanco = " & intCodBanco & " AND CtaBancaria = '" & dbcCuentaBancaria & "')) MovBancarios " & _
        ''    "LEFT OUTER JOIN ((SELECT FechaMovto,FolioMovto,Referencia,Concepto,Movimiento,ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVINGRESO & "' THEN Importe END,0) AS Ingresos," & _
        ''    "ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVEGRESO & "' THEN Importe END,0) AS Egresos,Conciliado FROM MovimientosBancarios WHERE Movimiento = '" & C_MOVCANCELACION & "' AND FechaMovto <= '" & FechaFinal & "' AND FechaMovto >= '" & FechaInicial & "' " & _
        ''    "AND CodBanco = " & intCodBanco & " AND CtaBancaria = '" & dbcCuentaBancaria & "') UNION (SELECT FechaMovto,FolioMovto,Referencia,Concepto,Movimiento,ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVINGRESO & "' THEN Importe END,0) AS Ingresos," & _
        ''    "ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVEGRESO & "' THEN Importe END,0) AS Egresos,Conciliado FROM MovimientosBancarios WHERE Movimiento = '" & C_MOVCANCELACION & "' AND ((FechaMovto < '" & FechaInicial & "' AND Conciliado = 0) OR FechaConciliacion = '" & FechaFinal & "') " & _
        ''    "AND CodBanco = " & intCodBanco & " AND CtaBancaria = '" & dbcCuentaBancaria & "')) MovCancelados ON MovBancarios.FolioMovto = MovCancelados.Referencia"

        sql = "SELECT ISNULL(CASE WHEN MovBancarios.FolioMovto = MovCancelados.Referencia THEN 'C' ELSE CASE WHEN MovBancarios.Conciliado = 1 AND MovBancarios.FechaConciliacion = '" & FechaFinal & "' THEN '*' END END,'') AS Estatus ,MovBancarios.FolioMovto," & "ISNULL(MovCancelados.FolioMovto,'') AS FolioCancelacion,MovBancarios.FechaMovto,MovBancarios.Referencia,MovBancarios.Concepto," & "(CASE MovBancarios.Movimiento WHEN '" & C_MOVPAGO & "' THEN 'PAGOS' WHEN '" & C_MOVDEPOSITO & "' THEN 'DEPOSITOS' WHEN '" & C_MOVTRASPASO & "' THEN 'TRASP. BANC.' WHEN '" & C_MOVCARGOS & "' THEN 'CARGOS DIV.' WHEN '" & C_MOVANTICIPOS & "' THEN 'ANT. PROV./ACREED.' WHEN '" & C_OTROSINGRESOS & "' THEN 'OTROS INGRESOS' END) AS Movimiento," & "(CASE WHEN MovBancarios.FolioMovto = MovCancelados.Referencia THEN 0 ELSE MovBancarios.Ingresos END) AS Ingresos,(CASE WHEN MovBancarios.FolioMovto = MovCancelados.Referencia THEN 0 ELSE MovBancarios.Egresos END) AS Egresos " & "FROM ((SELECT FechaMovto,FolioMovto,Referencia,Concepto,Movimiento,FechaConciliacion,ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVINGRESO & "' THEN Importe END,0) AS Ingresos,ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVEGRESO & "' THEN Importe END,0) AS Egresos,Conciliado " & "FROM MovimientosBancarios WHERE Movimiento <> '" & C_MOVCANCELACION & "' AND ((FechaMovto <= '" & FechaFinal & "' AND FechaMovto >= '" & FechaInicial & "') OR FechaConciliacion = '" & FechaFinal & "' OR (FechaMovto < '" & FechaInicial & "' AND Conciliado = 0) OR (FechaConciliacion > '" & FechaFinal & "' AND FechaMovto <= '" & FechaFinal & "' AND FechaMovto >= '" & FechaInicial & "' AND Conciliado = 1)) AND Movimiento <> '" & C_MOVCANCELACION & "' AND CodBanco = " & intCodBanco & " AND CtaBancaria = '" & dbcCuentaBancaria.Text & "')) MovBancarios " & "LEFT OUTER JOIN ((SELECT FechaMovto,FolioMovto,Referencia,Concepto,Movimiento,ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVINGRESO & "' THEN Importe END,0) AS Ingresos," & "ISNULL(CASE TipoMovto WHEN '" & C_TIPOMOVEGRESO & "' THEN Importe END,0) AS Egresos,Conciliado FROM MovimientosBancarios WHERE Movimiento = '" & C_MOVCANCELACION & "' AND ((FechaMovto <= '" & FechaFinal & "' AND FechaMovto >= '" & FechaInicial & "') OR FechaConciliacion = '" & FechaFinal & "' OR (FechaMovto < '" & FechaInicial & "' AND Conciliado = 0) OR FechaConciliacion > '" & FechaFinal & "') " & "AND CodBanco = " & intCodBanco & " AND CtaBancaria = '" & dbcCuentaBancaria.Text & "')) MovCancelados ON MovBancarios.FolioMovto = MovCancelados.Referencia ORDER BY MovBancarios.FechaMovto"
        BorraCmd()
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = sql
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No Existen Movimientos para esta Cuenta en este Periodo, Favor de Verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        Else
            'frmReportes.Report = RptBancosProcesoMensualMovimientosenConciliacion 
            RptBancosProcesoMensualMovimientosenConciliacion.SetDataSource(frmReportes.rsReport)
        End If

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        'frmReportes.rsReport = rsReporte
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "Periodo", "Ejercicio"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, Periodo, Ejercicio}

        'If (NombreEmpresa <> Nothing) Then
        '    pdvNum.Value = NombreEmpresa : pvNum.Add(pdvNum)
        '    RptBancosProcesoMensualMovimientosenConciliacion.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
        'End If

        'If (NombreReporte <> Nothing) Then
        '    pdvNum.Value = NombreReporte : pvNum.Add(pdvNum)
        '    RptBancosProcesoMensualMovimientosenConciliacion.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
        'End If

        'If (Periodo <> Nothing) Then
        '    pdvNum.Value = Periodo : pvNum.Add(pdvNum)
        '    RptBancosProcesoMensualMovimientosenConciliacion.DataDefinition.ParameterFields("Periodo").ApplyCurrentValues(pvNum)
        'End If

        'If (Ejercicio <> Nothing) Then
        '    pdvNum.Value = Ejercicio : pvNum.Add(pdvNum)
        '    RptBancosProcesoMensualMovimientosenConciliacion.DataDefinition.ParameterFields("Ejercicio").ApplyCurrentValues(pvNum)
        'End If

        frmReportes.Text = "Movimientos en Conciliación"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.reporteActual = RptBancosProcesoMensualMovimientosenConciliacion
        frmReportes.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error al Imprimir : " & Err.Description, MsgBoxStyle.Exclamation, "Error de Operacion")
    End Sub

    Sub ObtenerEjercicios()
        On Error GoTo MErr
        gStrSql = "SELECT DISTINCT Ejercicio FROM EjercicioPeriodo"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Do While Not RsGral.EOF
                cmbAño.Items.Add(RsGral.Fields("Ejercicio").Value)
                RsGral.MoveNext()
            Loop
        Else
            cmbAño.Items.Add("")
        End If
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Limpiar()
        Nuevo()
        dbcBanco.Focus()
    End Sub

    Sub Nuevo()
        On Error GoTo MErr
        dbcBanco.Text = ""
        dbcCuentaBancaria.Text = ""
        'dbcCuentaBancaria.RowSource = Nothing
        cmbMes.SelectedIndex = 0
        cmbAño.SelectedIndex = 0
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub cmbAño_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAño.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbMes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMes.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcBanco_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.CursorChanged
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcBanco.Name Then
        '    Exit Sub
        'End If
        dbcCuentaBancaria.Text = ""
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' ORDER BY DescBanco"
        DCChange(gStrSql, tecla)
        intCodBanco = 0
    End Sub

    Private Sub dbcBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcBanco.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos ORDER BY DescBanco"
        DCGotFocus(gStrSql, dbcBanco)
        Pon_Tool()
    End Sub

    Private Sub dbcBanco_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcBanco.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSALIR = True
            Me.Close()
        End If
    End Sub

    Private Sub dbcBanco_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcBanco.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' ORDER BY DescBanco"
        DCLostFocus(dbcBanco, gStrSql, intCodBanco)
    End Sub

    Private Sub dbcCuentaBancaria_cursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.CursorChanged
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcCuentaBancaria.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCChange(gStrSql, tecla)
        'intCodBanco = 0
    End Sub

    Private Sub dbcCuentaBancaria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcCuentaBancaria.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCGotFocus(gStrSql, dbcCuentaBancaria)
        Pon_Tool()
    End Sub

    Private Sub dbcCuentaBancaria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCuentaBancaria.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            dbcBanco.Focus()
        End If
    End Sub

    Private Sub dbcCuentaBancaria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcCuentaBancaria.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcCuentaBancaria_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCuentaBancaria.KeyUp
        Dim Aux As String
        Aux = dbcCuentaBancaria.Text
        'If dbcCuentaBancaria.SelectedItem <> 0 Then
        '    dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        'End If
        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub dbcCuentaBancaria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Leave
        On Error GoTo Err_Renamed
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCLostFocus(dbcCuentaBancaria, gStrSql, intCodBanco)
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcCuentaBancaria_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcCuentaBancaria.MouseUp
        'Dim Aux As String
        'Aux = dbcCuentaBancaria.Text
        'If dbcCuentaBancaria.SelectedItem <> 0 Then
        '    dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        'End If
        'dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub frmBancosProcesoMensualMovimientosenConciliacion_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoMensualMovimientosenConciliacion_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoMensualMovimientosenConciliacion_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmBancosProcesoMensualMovimientosenConciliacion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.gp_CampoMayusculas(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoMensualMovimientosenConciliacion_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        ObtenerEjercicios()
        Nuevo()
    End Sub

    Private Sub frmBancosProcesoMensualMovimientosenConciliacion_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        'Si se cierra el formulario y existio algun cambio en el registro se
        'informa al usuario del cabio y si desea guardar el registro, ya sea
        'que sea nuevo o un registro modificado
        If mblnSALIR Then
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSALIR = False
                    Cancel = 1
                    dbcBanco.Focus()
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmBancosProcesoMensualMovimientosenConciliacion_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class