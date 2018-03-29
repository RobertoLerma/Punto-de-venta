Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility


Public Class frmVtasVEExistencias
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE EXISTENCIAS A PRECIO PUBLICO                                                      *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      LUNES 25 DE AGOSTO DE 2003                                                                   *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents optCosto As System.Windows.Forms.RadioButton
    Public WithEvents optPrecioPublico As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents txtTipoCambio As System.Windows.Forms.TextBox
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents txtCodVendExterno As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox


    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Public WithEvents btnNuevo As Button
    Friend WithEvents btnImprimir As Button
    Dim rsReporte As ADODB.Recordset

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.optCosto = New System.Windows.Forms.RadioButton()
        Me.optPrecioPublico = New System.Windows.Forms.RadioButton()
        Me.txtTipoCambio = New System.Windows.Forms.TextBox()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.txtCodVendExterno = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.SuspendLayout()
        '
        'optCosto
        '
        Me.optCosto.BackColor = System.Drawing.SystemColors.Control
        Me.optCosto.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCosto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCosto.Location = New System.Drawing.Point(162, 18)
        Me.optCosto.Margin = New System.Windows.Forms.Padding(2)
        Me.optCosto.Name = "optCosto"
        Me.optCosto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCosto.Size = New System.Drawing.Size(66, 17)
        Me.optCosto.TabIndex = 1
        Me.optCosto.TabStop = True
        Me.optCosto.Text = "Al Costo"
        Me.ToolTip1.SetToolTip(Me.optCosto, "Muestra el Reporte de Existencias al Costo")
        Me.optCosto.UseVisualStyleBackColor = False
        '
        'optPrecioPublico
        '
        Me.optPrecioPublico.BackColor = System.Drawing.SystemColors.Control
        Me.optPrecioPublico.Checked = True
        Me.optPrecioPublico.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPrecioPublico.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPrecioPublico.Location = New System.Drawing.Point(48, 18)
        Me.optPrecioPublico.Margin = New System.Windows.Forms.Padding(2)
        Me.optPrecioPublico.Name = "optPrecioPublico"
        Me.optPrecioPublico.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPrecioPublico.Size = New System.Drawing.Size(110, 17)
        Me.optPrecioPublico.TabIndex = 0
        Me.optPrecioPublico.TabStop = True
        Me.optPrecioPublico.Text = "A Precio Público"
        Me.ToolTip1.SetToolTip(Me.optPrecioPublico, "Muestra el Reporte de Existencias a Precio Público")
        Me.optPrecioPublico.UseVisualStyleBackColor = False
        '
        'txtTipoCambio
        '
        Me.txtTipoCambio.AcceptsReturn = True
        Me.txtTipoCambio.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipoCambio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambio.Location = New System.Drawing.Point(188, 17)
        Me.txtTipoCambio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTipoCambio.MaxLength = 5
        Me.txtTipoCambio.Name = "txtTipoCambio"
        Me.txtTipoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambio.Size = New System.Drawing.Size(94, 20)
        Me.txtTipoCambio.TabIndex = 5
        Me.txtTipoCambio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambio, "Tipo de Cambio")
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Checked = True
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(36, 37)
        Me.optDolares.Margin = New System.Windows.Forms.Padding(2)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(61, 18)
        Me.optDolares.TabIndex = 6
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me.optDolares, "Muestra los Precios en Dolares")
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(36, 18)
        Me.optPesos.Margin = New System.Windows.Forms.Padding(2)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(61, 17)
        Me.optPesos.TabIndex = 4
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me.optPesos, "Muestra los Precios en Pesos")
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'txtCodVendExterno
        '
        Me.txtCodVendExterno.AcceptsReturn = True
        Me.txtCodVendExterno.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodVendExterno.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodVendExterno.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodVendExterno.Location = New System.Drawing.Point(72, 76)
        Me.txtCodVendExterno.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodVendExterno.MaxLength = 3
        Me.txtCodVendExterno.Name = "txtCodVendExterno"
        Me.txtCodVendExterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodVendExterno.Size = New System.Drawing.Size(26, 20)
        Me.txtCodVendExterno.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtCodVendExterno, "Código del Vendedor Externo")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Frame3)
        Me.Frame1.Controls.Add(Me.Frame2)
        Me.Frame1.Controls.Add(Me.dbcSucursal)
        Me.Frame1.Controls.Add(Me.txtCodVendExterno)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 12)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(328, 205)
        Me.Frame1.TabIndex = 7
        Me.Frame1.TabStop = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.optCosto)
        Me.Frame3.Controls.Add(Me.optPrecioPublico)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(12, 13)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(283, 46)
        Me.Frame3.TabIndex = 10
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Existencias"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtTipoCambio)
        Me.Frame2.Controls.Add(Me.optDolares)
        Me.Frame2.Controls.Add(Me.optPesos)
        Me.Frame2.Controls.Add(Me.Label2)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(14, 128)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(292, 63)
        Me.Frame2.TabIndex = 9
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Moneda"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(102, 20)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(105, 17)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Tipo de Cambio :"
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(101, 76)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(212, 21)
        Me.dbcSucursal.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 78)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(64, 17)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Vendedor : "
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(120, 234)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 71
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.Location = New System.Drawing.Point(5, 234)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 70
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasVEExistencias
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(345, 278)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(284, 179)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasVEExistencias"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Existencias"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Sub Imprime()
        Dim RptVtasVEReportedeExistenciasalCosto As New RptVtasVEReportedeExistenciasalCosto
        Dim RptVtasVEReportedeExistenciasaPrecioPublico As New RptVtasVEReportedeExistenciasaPrecioPublico


        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue


        Dim Sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim NombreVendedor As String
        Dim strWhere As String
        Dim Moneda As String
        Dim CodigoVendedor As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim PeriodoReporte As String
        'On Error GoTo ImprimeErr

        If CDbl(Numerico(txtCodVendExterno.Text)) = 0 Then
            MsgBox("Proporcione un Codigo de Vendedor Externo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodVendExterno.Focus()
            Exit Sub
        End If
        If Trim(dbcSucursal.Text) = "" Then
            MsgBox("Proprcione la Descripción del Vendedor Externo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcSucursal.Focus()
            Exit Sub
        End If
        If CDbl(Numerico(txtTipoCambio.Text)) = 0 Then
            MsgBox("Proporcione un Tipo de Cambio, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtTipoCambio.Focus()
            Exit Sub
        End If

        'Obtener Fecha Salida
        gStrSql = "select folioalmacen,fechaalmacen from movtosalmacencab where codmovtoalm = " & C_SalidaAVendedoresExternos & " " & "and codalmacenref = " & txtCodVendExterno.Text & " " & "order by folioalmacen desc,fechaalmacen desc"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount > 0 Then
            RsGral.MoveFirst()
            FechaInicial = Format(frmReportes.rsReport.Fields("FechaAlmacen").Value, "dd/mmm/yyyy")
        Else
            'FechaInicial = Format(Today, "dd/mmm/yyyy")
            FechaInicial = AgregarHoraAFecha(Today.ToString())
        End If

        'Obtener Fecha de Entrada
        gStrSql = "select folioalmacen,fechaalmacen from movtosalmacencab where codmovtoalm = " & C_EntradaPorDevoluciondeVendedoresExternos & " " & "and codalmacenref = " & txtCodVendExterno.Text & " " & "order by folioalmacen desc,fechaalmacen desc"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount > 0 Then
            frmReportes.rsReport.MoveFirst()
            FechaFinal = Format(frmReportes.rsReport.Fields("FechaAlmacen").Value, "dd/mmm/yyyy")
        Else
            'FechaFinal = Format(Today, "dd/mmm/yyyy")
            FechaFinal = AgregarHoraAFecha(Today.ToString())
        End If
        If CDate(FechaFinal) < CDate(FechaInicial) Then
            'FechaFinal = Format(Today, "dd/mmm/yyyy")
            FechaFinal = AgregarHoraAFecha(Today.ToString())
        End If

        txtTipoCambio.Text = Format(gcurCorpoTIPOCAMBIODOLAR, "###,##0.00")
        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = UCase("Reporte de Existencias del Vendedor Externo")
        NombreVendedor = UCase(Trim(dbcSucursal.Text))
        CodigoVendedor = txtCodVendExterno.Text
        PeriodoReporte = "Del " & FechaInicial & " al " & FechaFinal

        If optPesos.Checked = True Then
            Moneda = "Pesos"
        ElseIf optDolares.Checked = True Then
            Moneda = "Dólares"
        End If
        If optPrecioPublico.Checked = True Then
            If optDolares.Checked Then
                '''sql = "SELECT CA.CODARTICULO," & _
                '"CA.DESCARTICULO,SUM((I.EXISTENCIAINICIAL + (I.ENTRADAS - I.SALIDAS)) - I.APARTADOS)  AS EXISTENCIA," & _
                '"CA.PRECIOPUBDOLAR, " & _
                '"ISNULL(CASE " & _
                '"WHEN CA.CODGRUPO = " & gCODJOYERIA & " THEN " & _
                '"(SELECT PORCDESCTO FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODJOYERIA & " AND CA.PRECIOPUBDOLAR BETWEEN IMPORTEINI AND IMPORTEFIN) " & _
                '"WHEN CA.CODGRUPO = " & gCODRELOJERIA & " THEN " & _
                '"(SELECT PORCDESCTO FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODRELOJERIA & " AND CA.CODMARCA = CODMARCA) " & _
                '"WHEN CA.CODGRUPO = " & gCODVARIOS & " THEN " & _
                '"(SELECT PORCDESCTO FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODVARIOS & " AND CA.CODFAMILIA = CODFAMILIA) " & _
                '"END,0) AS PORCDESCTO " & _
                '"FROM INVENTARIO I INNER JOIN CATARTICULOS CA ON I.CODARTICULO = CA.CODARTICULO " & _
                '"WHERE I.CODALMACEN = " & txtCodVendExterno & " " & _
                '"GROUP BY CA.CODARTICULO,CA.DESCARTICULO,CA.PRECIOPUBDOLAR,CA.COSTOREAL,CA.CODGRUPO,CA.CODMARCA,CA.CODFAMILIA " & _
                '"HAVING SUM((I.EXISTENCIAINICIAL + (I.ENTRADAS - I.SALIDAS)) - I.APARTADOS) > 0 " & _
                '"ORDER BY CA.CODARTICULO,CA.DESCARTICULO"

                Sql = "SELECT  CA.CODARTICULO,CA.DESCARTICULO, SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS)) AS EXISTENCIA,CASE WHEN CA.PESOSFIJOS = 0 THEN CA.PRECIOPUBDOLAR WHEN CA.PESOSFIJOS = 1 THEN CA.PRECIOPUBDOLAR / " & Numerico(txtTipoCambio.Text) & " END AS PRECIOPUBDOLAR," & "ISNULL(CASE WHEN CA.CODGRUPO = " & gCODJOYERIA & " THEN (SELECT PorcDescto FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODJOYERIA & " And ( (importeini > 0 And ImporteFin > 0) And CA.PrecioPubDolar Between importeini and ImporteFin) or ( (importeini > 0 And ImporteFin = 0) And CA.PrecioPubDolar >= importeini ) ) " & "WHEN CA.CODGRUPO = " & gCODRELOJERIA & " THEN (SELECT PORCDESCTO FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODRELOJERIA & " AND CA.CODMARCA = CODMARCA) WHEN CA.CODGRUPO = " & gCODVARIOS & " THEN (SELECT PORCDESCTO FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODVARIOS & " AND CA.CODFAMILIA = CODFAMILIA) END,0) AS PORCDESCTO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) + '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS ANTERIOR " & "FROM  INVENTARIO I INNER JOIN CATARTICULOS CA ON I.CODARTICULO = CA.CODARTICULO Where I.CodAlmacen = " & txtCodVendExterno.Text & " GROUP BY CA.CODARTICULO,CA.DESCARTICULO,CA.PRECIOPUBDOLAR,CA.COSTOREAL,CA.CODGRUPO,CA.CODMARCA,CA.CODFAMILIA,CA.PESOSFIJOS,CODIGOANT,ORIGENANT Having SUM((I.ExistenciaInicial + I.Entradas) - (I.Salidas + I.Apartados)) > 0 ORDER BY CA.CODARTICULO,CA.DESCARTICULO "
            ElseIf optPesos.Checked Then
                '''sql = "SELECT CA.CODARTICULO," & _
                '"CA.DESCARTICULO,SUM((I.EXISTENCIAINICIAL + (I.ENTRADAS - I.SALIDAS)) - I.APARTADOS)  AS EXISTENCIA," & _
                '"(CA.PRECIOPUBDOLAR * " & txtTipoCambio & ") AS PRECIOPUBDOLAR, " & _
                '"ISNULL(CASE " & _
                '"WHEN CA.CODGRUPO = " & gCODJOYERIA & " THEN " & _
                '"(SELECT PORCDESCTO FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODJOYERIA & " AND CA.PRECIOPUBDOLAR BETWEEN IMPORTEINI AND IMPORTEFIN) " & _
                '"WHEN CA.CODGRUPO = " & gCODRELOJERIA & " THEN " & _
                '"(SELECT PORCDESCTO FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODRELOJERIA & " AND CA.CODMARCA = CODMARCA) " & _
                '"WHEN CA.CODGRUPO = " & gCODVARIOS & " THEN " & _
                '"(SELECT PORCDESCTO FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODVARIOS & " AND CA.CODFAMILIA = CODFAMILIA) " & _
                '"END,0) AS PORCDESCTO " & _
                '"FROM INVENTARIO I INNER JOIN CATARTICULOS CA ON I.CODARTICULO = CA.CODARTICULO " & _
                '"WHERE I.CODALMACEN = " & txtCodVendExterno & " " & _
                '"GROUP BY CA.CODARTICULO,CA.DESCARTICULO,CA.PRECIOPUBDOLAR,CA.COSTOREAL,CA.CODGRUPO,CA.CODMARCA,CA.CODFAMILIA " & _
                '"HAVING SUM((I.EXISTENCIAINICIAL + (I.ENTRADAS - I.SALIDAS)) - I.APARTADOS) > 0 " & _
                '"ORDER BY CA.CODARTICULO,CA.DESCARTICULO"

                Sql = "SELECT CA.CODARTICULO,CA.DESCARTICULO,SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS)) AS EXISTENCIA,CASE WHEN CA.PESOSFIJOS = 0 THEN CA.PRECIOPUBDOLAR * " & Numerico(txtTipoCambio.Text) & " WHEN CA.PESOSFIJOS = 1 THEN CA.PRECIOPUBDOLAR END AS PRECIOPUBDOLAR," & "ISNULL(CASE WHEN CA.CODGRUPO = " & gCODJOYERIA & " THEN (SELECT PorcDescto FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODJOYERIA & " And ( (importeini > 0 And ImporteFin > 0) And CA.PrecioPubDolar Between importeini and ImporteFin) or ( (importeini > 0 And ImporteFin = 0) And CA.PrecioPubDolar >= importeini ) ) " & "WHEN CA.CODGRUPO = " & gCODRELOJERIA & " THEN (SELECT PORCDESCTO FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODRELOJERIA & " AND CA.CODMARCA = CODMARCA) WHEN CA.CODGRUPO = " & gCODVARIOS & " THEN (SELECT PORCDESCTO FROM CATDESCTOSVEXTERNOS WHERE CODGRUPO = " & gCODVARIOS & " AND CA.CODFAMILIA = CODFAMILIA) END,0) AS PORCDESCTO,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) + '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS ANTERIOR " & "FROM  INVENTARIO I INNER JOIN CATARTICULOS CA ON I.CODARTICULO = CA.CODARTICULO Where I.CodAlmacen = " & txtCodVendExterno.Text & " GROUP BY CA.CODARTICULO,CA.DESCARTICULO,CA.PRECIOPUBDOLAR,CA.COSTOREAL,CA.CODGRUPO,CA.CODMARCA,CA.CODFAMILIA,CA.PESOSFIJOS,CODIGOANT,ORIGENANT Having SUM((I.ExistenciaInicial + I.Entradas) - (I.Salidas + I.Apartados)) > 0 ORDER BY CA.CODARTICULO,CA.DESCARTICULO "
            End If
            BorraCmd()
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd.CommandText = Sql
            frmReportes.rsReport = Cmd.Execute

            If frmReportes.rsReport.RecordCount = 0 Then
                MsgBox("No existen movimientos en el almacén del Vendedor Externo seleccionado" & vbNewLine & "Favor de verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Exit Sub
            Else
                'frmReportes.Report = RptVtasVEReportedeExistenciasaPrecioPublico
                RptVtasVEReportedeExistenciasaPrecioPublico.SetDataSource(frmReportes.rsReport)
            End If
        ElseIf optCosto.Checked = True Then
            If optDolares.Checked Then
                Sql = "SELECT CA.CODARTICULO," & "CA.DESCARTICULO,SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS))  AS EXISTENCIA," & "CA.COSTOREAL,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) + '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS ANTERIOR " & "FROM INVENTARIO I INNER JOIN CATARTICULOS CA ON I.CODARTICULO = CA.CODARTICULO INNER JOIN " & "CATALMACEN SUC ON I.CODALMACEN = SUC.CODALMACEN LEFT OUTER JOIN CATCLIENTES CLI  ON SUC.CODALMACEN = CLI.ALMACENVEXT " & "WHERE I.CODALMACEN = " & txtCodVendExterno.Text & " " & "GROUP BY CA.CODARTICULO,CA.DESCARTICULO,CA.COSTOREAL,CODIGOANT,ORIGENANT " & "HAVING SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS)) > 0 " & "ORDER BY CA.CODARTICULO,CA.DESCARTICULO"
            ElseIf optPesos.Checked Then
                Sql = "SELECT CA.CODARTICULO," & "CA.DESCARTICULO,SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS))  AS EXISTENCIA," & "(CA.COSTOREAL * " & txtTipoCambio.Text & ") AS COSTOREAL,CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) + '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS ANTERIOR " & "FROM INVENTARIO I INNER JOIN CATARTICULOS CA ON I.CODARTICULO = CA.CODARTICULO INNER JOIN " & "CATALMACEN SUC ON I.CODALMACEN = SUC.CODALMACEN LEFT OUTER JOIN CATCLIENTES CLI  ON SUC.CODALMACEN = CLI.ALMACENVEXT " & "WHERE I.CODALMACEN = " & txtCodVendExterno.Text & " " & "GROUP BY CA.CODARTICULO,CA.DESCARTICULO,CA.COSTOREAL,CODIGOANT,ORIGENANT " & "HAVING SUM((I.EXISTENCIAINICIAL + I.ENTRADAS) - (I.SALIDAS + I.APARTADOS)) > 0 " & "ORDER BY CA.CODARTICULO,CA.DESCARTICULO"
            End If

            BorraCmd()
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            Cmd.CommandText = Sql
            frmReportes.rsReport = Cmd.Execute

            If frmReportes.rsReport.RecordCount = 0 Then
                MsgBox("No existen movimientos en el almacén del Vendedor Externo seleccionado" & vbNewLine & "Favor de verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Exit Sub
            Else
                'frmReportes.Report = RptVtasVEReportedeExistenciasalCosto
                RptVtasVEReportedeExistenciasalCosto.SetDataSource(frmReportes.rsReport)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.rsReport = rsReporte
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "NombreVendedor", "Moneda", "CodigoVendedor", "PeriodoReporte"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, NombreVendedor, Moneda, CodigoVendedor, PeriodoReporte}
        'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If optPrecioPublico.Checked = True Then
            frmReportes.Text = "Reporte de Existencias a Precio Público"
        Else
            frmReportes.Text = "Reporte de Existencias al Costo"
            frmReportes.reporteActual = RptVtasVEReportedeExistenciasalCosto
            frmReportes.Show()
        End If
        frmReportes.reporteActual = RptVtasVEReportedeExistenciasaPrecioPublico
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
                MsgBox("Este código no pertenece a un Vendedor Externo" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtCodVendExterno.Text = ""
                txtCodVendExterno.Focus()
                Exit Sub
            Else
                txtCodVendExterno.Text = txtCodVendExterno.Text
                dbcSucursal.Text = RsGral.Fields("DescAlmacen").Value
            End If
        Else
            MsgBox("Código de almacen no existe" & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCodVendExterno.Text = ""
            txtCodVendExterno.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Limpiar()
        Nuevo()
        InicializaVariables()
        optPrecioPublico.Focus()
    End Sub

    Sub InicializaVariables()
        mblnSalir = False
    End Sub

    Sub Nuevo()
        optPrecioPublico.Checked = True
        txtCodVendExterno.Text = ""
        dbcSucursal.Text = ""
        txtTipoCambio.Text = Format(gcurCorpoTIPOCAMBIODOLAR, "###,##0.00")
        optDolares.Checked = True
        'txtTipoCambio.Enabled = False
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.CursorChanged
        If FueraChange = True Then Exit Sub
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
            Exit Sub
        End If
        If Trim(dbcSucursal.Text) = "" Then
            txtCodVendExterno.Text = ""
            Exit Sub
        End If
        gStrSql = "SELECT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'V' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
    End Sub
    Private Sub dbcSucursal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dbcSucursal.SelectedIndexChanged
        ''txtCodVendExterno.Text = IIf(intCodSucursal <> 0, intCodSucursal, "")
        'If (dbcSucursal.SelectedItem("DescAlmacen").Value <> "") Then
        '    txtCodVendExterno.Text = dbcSucursal.SelectedItem("CodAlmacen").Value
        'End If
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        ''If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        ''Exit Sub
        ''End If
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
        '    dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
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

    Private Sub frmVtasVEExistencias_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasVEExistencias_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasVEExistencias_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optPrecioPublico" And System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "optCosto" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmVtasVEExistencias_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasVEExistencias_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        InicializaVariables()
        Nuevo()
    End Sub

    Private Sub frmVtasVEExistencias_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmVtasVEExistencias_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optCosto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCosto.Enter
        Pon_Tool()
    End Sub

    Private Sub optDolares_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDolares.CheckedChanged
        If eventSender.Checked Then
            'txtTipoCambio.Enabled = False
        End If
    End Sub

    Private Sub optDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub optPesos_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPesos.CheckedChanged
        If eventSender.Checked Then
            txtTipoCambio.Enabled = True
        End If
    End Sub

    Private Sub optPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPesos.Enter
        Pon_Tool()
    End Sub

    Private Sub optPrecioPublico_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPrecioPublico.Enter
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

    Private Sub txtTipoCambio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.TextChanged
        If Trim(txtTipoCambio.Text) = "" Then
            txtTipoCambio.Text = "0.00"
        End If
    End Sub

    Private Sub txtTipoCambio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Enter
        Pon_Tool()
        SelTextoTxt(txtTipoCambio, 1, 1)
    End Sub

    Private Sub txtTipoCambio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.MskCantidad(txtTipoCambio.Text, KeyAscii, 2, 2, (txtTipoCambio.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Leave
        If CDbl(Numerico(txtTipoCambio.Text)) = 0 Then
            txtTipoCambio.Text = "0.00"
        Else
            txtTipoCambio.Text = Format(txtTipoCambio.Text, "###,##0.00")
        End If
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class