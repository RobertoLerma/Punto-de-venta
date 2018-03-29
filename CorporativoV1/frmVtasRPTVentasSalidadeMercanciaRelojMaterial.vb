Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasRPTVentasSalidadeMercanciaRelojMaterial
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkDetalle As System.Windows.Forms.CheckBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblVentas_2 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_1 As System.Windows.Forms.GroupBox
    Public WithEvents chkTodas As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_0 As System.Windows.Forms.GroupBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents chkImpuesto As System.Windows.Forms.CheckBox
    Public WithEvents dbcMaterial As System.Windows.Forms.ComboBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"
    Const C_NINGUNA As String = "[ Vacío ... ]"

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean

    Dim cTablaTmp As String

    Dim mblnSalir As Boolean

    Dim mblnFueraChange As Boolean
    Dim tecla As Integer
    Dim mintCodSucursal As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim mintCodTipoMaterial As Integer


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.chkDetalle = New System.Windows.Forms.CheckBox()
        Me._fraVtas_1 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me._fraVtas_0 = New System.Windows.Forms.GroupBox()
        Me.chkTodas = New System.Windows.Forms.CheckBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me.chkImpuesto = New System.Windows.Forms.CheckBox()
        Me.dbcMaterial = New System.Windows.Forms.ComboBox()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me._fraVtas_1.SuspendLayout()
        Me._fraVtas_0.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(6, 186)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(359, 60)
        Me.txtMensaje.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'chkDetalle
        '
        Me.chkDetalle.BackColor = System.Drawing.SystemColors.Control
        Me.chkDetalle.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDetalle.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDetalle.Location = New System.Drawing.Point(198, 143)
        Me.chkDetalle.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDetalle.Name = "chkDetalle"
        Me.chkDetalle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDetalle.Size = New System.Drawing.Size(148, 17)
        Me.chkDetalle.TabIndex = 12
        Me.chkDetalle.Text = "Detallar por Sucursal"
        Me.chkDetalle.UseVisualStyleBackColor = False
        '
        '_fraVtas_1
        '
        Me._fraVtas_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_1.Controls.Add(Me.dtpDesde)
        Me._fraVtas_1.Controls.Add(Me.dtpHasta)
        Me._fraVtas_1.Controls.Add(Me._lblVentas_2)
        Me._fraVtas_1.Controls.Add(Me._lblVentas_1)
        Me._fraVtas_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_1.Location = New System.Drawing.Point(6, 91)
        Me._fraVtas_1.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.Name = "_fraVtas_1"
        Me._fraVtas_1.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_1.Size = New System.Drawing.Size(359, 46)
        Me._fraVtas_1.TabIndex = 6
        Me._fraVtas_1.TabStop = False
        Me._fraVtas_1.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(74, 17)
        Me.dtpDesde.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(104, 20)
        Me.dtpDesde.TabIndex = 8
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(240, 17)
        Me.dtpHasta.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(100, 20)
        Me.dtpHasta.TabIndex = 10
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_2.Location = New System.Drawing.Point(191, 22)
        Me._lblVentas_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(46, 13)
        Me._lblVentas_2.TabIndex = 9
        Me._lblVentas_2.Text = "Hasta el"
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_1.Location = New System.Drawing.Point(18, 20)
        Me._lblVentas_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(52, 13)
        Me._lblVentas_1.TabIndex = 7
        Me._lblVentas_1.Text = "Desde el "
        '
        '_fraVtas_0
        '
        Me._fraVtas_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_0.Controls.Add(Me.chkTodas)
        Me._fraVtas_0.Controls.Add(Me.dbcSucursal)
        Me._fraVtas_0.Controls.Add(Me._lblVentas_0)
        Me._fraVtas_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_0.Location = New System.Drawing.Point(6, 6)
        Me._fraVtas_0.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.Name = "_fraVtas_0"
        Me._fraVtas_0.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_0.Size = New System.Drawing.Size(312, 46)
        Me._fraVtas_0.TabIndex = 0
        Me._fraVtas_0.TabStop = False
        '
        'chkTodas
        '
        Me.chkTodas.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodas.Checked = True
        Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodas.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTodas.Location = New System.Drawing.Point(6, 0)
        Me.chkTodas.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodas.Name = "chkTodas"
        Me.chkTodas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodas.Size = New System.Drawing.Size(133, 17)
        Me.chkTodas.TabIndex = 1
        Me.chkTodas.Text = "Todas las sucursales"
        Me.chkTodas.UseVisualStyleBackColor = False
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(80, 17)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(212, 21)
        Me.dbcSucursal.TabIndex = 3
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_0.Location = New System.Drawing.Point(28, 22)
        Me._lblVentas_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(48, 13)
        Me._lblVentas_0.TabIndex = 2
        Me._lblVentas_0.Text = "Sucursal"
        '
        'chkImpuesto
        '
        Me.chkImpuesto.BackColor = System.Drawing.SystemColors.Control
        Me.chkImpuesto.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkImpuesto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkImpuesto.Location = New System.Drawing.Point(12, 143)
        Me.chkImpuesto.Margin = New System.Windows.Forms.Padding(2)
        Me.chkImpuesto.Name = "chkImpuesto"
        Me.chkImpuesto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkImpuesto.Size = New System.Drawing.Size(123, 17)
        Me.chkImpuesto.TabIndex = 11
        Me.chkImpuesto.Text = "Incluir Impuesto"
        Me.chkImpuesto.UseVisualStyleBackColor = False
        '
        'dbcMaterial
        '
        Me.dbcMaterial.Location = New System.Drawing.Point(72, 65)
        Me.dbcMaterial.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcMaterial.Name = "dbcMaterial"
        Me.dbcMaterial.Size = New System.Drawing.Size(226, 21)
        Me.dbcMaterial.TabIndex = 5
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblRpt_2.Location = New System.Drawing.Point(9, 170)
        Me._lblRpt_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 13
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.AutoSize = True
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_5.Location = New System.Drawing.Point(24, 68)
        Me._lblVentas_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(44, 13)
        Me._lblVentas_5.TabIndex = 4
        Me._lblVentas_5.Text = "Material"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(123, 263)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 36
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(8, 263)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 35
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasRPTVentasSalidadeMercanciaRelojMaterial
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(375, 312)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.chkDetalle)
        Me.Controls.Add(Me._fraVtas_1)
        Me.Controls.Add(Me._fraVtas_0)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me.chkImpuesto)
        Me.Controls.Add(Me.dbcMaterial)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Controls.Add(Me._lblVentas_5)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 29)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasRPTVentasSalidadeMercanciaRelojMaterial"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Ventas de Relojería por Material de Fabricación"
        Me._fraVtas_1.ResumeLayout(False)
        Me._fraVtas_1.PerformLayout()
        Me._fraVtas_0.ResumeLayout(False)
        Me._fraVtas_0.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub


    Public Sub Limpiar()
        On Error Resume Next
        Call Me.Nuevo()
        Me.chkTodas.Focus()
    End Sub

    Public Sub Nuevo()
        Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Checked
        chkTodas_CheckStateChanged(chkTodas, New System.EventArgs())

        mblnFueraChange = True
        Me.dbcMaterial.Text = C_TODOS
        Me.dbcMaterial.Tag = ""
        mintCodTipoMaterial = 0
        mblnFueraChange = False

        Me.dtpDesde.Value = Format(Today, "dd/MMM/yyyy")
        Me.dtpHasta.Value = Format(Today, "dd/MMM/yyyy")
        Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDetalle.CheckState = System.Windows.Forms.CheckState.Checked
        Me.txtMensaje.Text = ""
        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
    End Sub

    Function DevuelveQuery() As String
        On Error GoTo Err_Renamed
        Dim Sql As String
        Sql = "SELECT CodSucursal,CA.DescAlmacen,VTA.CodTipoMaterial,ISNULL(CT.DescTipoMaterial,'') AS DescTipoMaterial," & "VTA.CodMarca,ISNULL(CM.DescMarca,'') AS DescMarca,VTA.CodModelo,ISNULL(CMOD.DescModelo,'') AS DescModelo,SUM(Cantidad - CantidadDev) AS Cantidad," & IIf(chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "ROUND(SUM(PrecioReal * (Cantidad - CantidadDev)),2) AS Importe,ROUND(SUM(((Descuento * (1 + (PorcIva/100))) * (Cantidad - CantidadDev))),2) AS Descuento,", "ROUND(SUM((PrecioListaSinIva - Descuento) * (Cantidad - CantidadDev)),2) AS Importe,ROUND(SUM(Descuento * (Cantidad - CantidadDev)),2) as Descuento,") & "SUM(CASE WHEN NumPartida = 1 THEN Redondeo ELSE 0 END) AS Redondeo " & "FROM DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') VTA " & "INNER JOIN (SELECT * FROM CatAlmacen WHERE TipoAlmacen = 'P') CA ON VTA.CodSucursal = CA.CodAlmacen " & "LEFT OUTER JOIN CatTipoMaterial CT ON VTA.CodTipoMaterial = CT.CodTipoMaterial " & "LEFT OUTER JOIN CatModelos CMOD ON VTA.CodMarca = CMOD.CodMarca AND VTA.CodModelo = CMOD.CodModelo " & "LEFT OUTER JOIN CatMarcas CM ON VTA.CodMarca = CM.CodMarca " & "WHERE (Cantidad - CantidadDev) > 0 AND VTA.CodGrupo = " & gCODRELOJERIA & " " & IIf(mintCodSucursal <> 0, "AND CodSucursal = " & mintCodSucursal & " ", "") & IIf(mintCodTipoMaterial <> 0, "AND VTA.CodTipoMaterial = " & mintCodTipoMaterial & " ", "") & "GROUP BY VTA.CodTipoMaterial,ISNULL(CT.DescTipoMaterial,''),VTA.CodMarca,ISNULL(CM.DescMarca,''),Vta.CodModelo,ISNULL(CMOD.DescModelo,''),CodSucursal,CA.DescAlmacen " & "ORDER BY ISNULL(CT.DescTipoMaterial,''),VTA.CodTipoMaterial"
        DevuelveQuery = Sql
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub Imprime()

        Dim rptVentasSalidaDeMercanciaRelojMaterial As New rptVentasSalidaDeMercanciaRelojMaterial
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        'On Error GoTo Merr
        Dim lStrSql As String
        'Declarar vectores para almacenar los parámetros que se le enviarán al reporte
        Dim aParam(5) As Object
        Dim aValues(5) As Object

        If Not ValidaDatos() Then
            Exit Sub
        End If

        lStrSql = DevuelveQuery()
        gStrSql = lStrSql
        ModEstandar.BorraCmd()
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existen datos para el rango de fechas indicado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Me.BringToFront()
            Exit Sub
        Else
            rptVentasSalidaDeMercanciaRelojMaterial.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "Mensaje"
        'aValues(1) = Trim(Me.txtMensaje.Text)
        'aParam(2) = "dDesde"
        'aValues(2) = Me.dtpDesde.Value
        'aParam(3) = "dHasta"
        'aValues(3) = Me.dtpHasta.Value
        'aParam(4) = "Empresa"
        'aValues(4) = Trim(gstrNombCortoEmpresa)
        'aParam(5) = "IncluyeImpuestos"
        'aValues(5) = IIf(Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "** Las cantidades expresadas incluyen IVA.", "** Las cantidades expresadas NO incluyen IVA.")

        If (txtMensaje.Text <> Nothing) Then
            pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaRelojMaterial.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaRelojMaterial.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        End If

        If (dtpDesde.Value <> Nothing) Then
            pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaRelojMaterial.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
        End If

        If (dtpHasta.Value <> Nothing) Then
            pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaRelojMaterial.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
        End If

        If (gstrNombCortoEmpresa <> Nothing) Then
            pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaRelojMaterial.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
        End If

        If (chkImpuesto.CheckState <> Nothing) Then
            pdvNum.Value = IIf(Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "** Las cantidades expresadas incluyen IVA.", "** Las cantidades expresadas NO incluyen IVA.") : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaRelojMaterial.DataDefinition.ParameterFields("IncluyeImpuestos").ApplyCurrentValues(pvNum)
        End If



        'If chkDetalle.CheckState = System.Windows.Forms.CheckState.Unchecked Then
        '    'rptVentasSalidaDeMercanciaRelojMaterial.DeleteGroup (3)
        '    rptVentasSalidaDeMercanciaRelojMaterial.Section10.Suppress = True
        'Else
        '    rptVentasSalidaDeMercanciaRelojMaterial.Section10.Suppress = False
        'End If

        frmReportes.reporteActual = rptVentasSalidaDeMercanciaRelojMaterial 'Es el nombre del archivo que se incluyó en el proyecto
        frmReportes.Show()
        'frmReportes.Imprime(Trim(Me.Text), aParam, aValues)

        Cmd.CommandTimeout = 90

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function ValidaDatos() As Boolean
        If mblnTecleoFechaI Then
            Do While (msglTiempoCambioI) <= 2.1
            Loop
            mblnTecleoFechaI = False
        End If
        If mblnTecleoFechaF Then
            Do While (msglTiempoCambioF) <= 2.1
            Loop
            mblnTecleoFechaF = False
        End If
        System.Windows.Forms.Application.DoEvents()
        Select Case True
            Case Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Unchecked And mintCodSucursal = 0
                MsgBox("Si no quiere imprimir los resultados de todas las sucursales, seleccione una de ellas", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dbcSucursal.Focus()
            Case Me.dtpDesde.Value > Me.dtpHasta.Value
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Límite", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dtpDesde.Focus()
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Private Sub chkTodas_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodas.CheckStateChanged
        Select Case Me.chkTodas.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                Me.dbcSucursal.Text = ""
                Me.dbcSucursal.Tag = ""
                mintCodSucursal = 0
                Me.dbcSucursal.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                Me.dbcSucursal.Text = ""
                Me.dbcSucursal.Tag = ""
                mintCodSucursal = 0
                Me.dbcSucursal.Enabled = True
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub dbcMaterial_CursorChange(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMaterial.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String
        If mblnFueraChange Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcMaterial.Name Then Exit Sub
        lStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM CatTipoMaterial Where descTipoMaterial LIKE '" & Trim(Me.dbcMaterial.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcMaterial))

        If dbcMaterial.SelectedItem <> "" Then
            dbcMaterial_Leave(dbcMaterial, New System.EventArgs())
        Else
            mintCodTipoMaterial = 0
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcMaterial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMaterial.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcMaterial.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM CatTipoMaterial"
        ModDCombo.DCGotFocus(gStrSql, dbcMaterial)
    End Sub

    Private Sub dbcMaterial_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMaterial.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            If Me.dbcSucursal.Enabled Then
                Me.dbcSucursal.Focus()
            Else
                Me.chkTodas.Focus()
            End If
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcMaterial_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMaterial.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcMaterial.text)
        '''    If Me.dbcMaterial.SelectedItem <> 0 Then
        '''        dbcMaterial_LostFocus
        '''    End If
        '''    Me.dbcMaterial.text = Aux
    End Sub

    Private Sub dbcMaterial_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMaterial.Leave
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM CatTipoMaterial Where RTrim(LTrim(descTipoMaterial)) = '" & Trim(Me.dbcMaterial.Text) & "'"
        'gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descMaterial FROM CatTipoMaterial Where RTrim(LTrim(descTipoMaterial)) = '" & Trim(RsGral.Fields("descTipoMaterial").Value.ToString()) & "'"
        Aux = mintCodTipoMaterial
        mintCodTipoMaterial = 0
        If Trim(Me.dbcMaterial.Text) <> Trim(C_TODOS) Or Trim(Me.dbcMaterial.Text) = "" Then
            ModDCombo.DCLostFocus(dbcMaterial, gStrSql, mintCodTipoMaterial)
        End If
        If Aux <> mintCodTipoMaterial Then
            If mintCodTipoMaterial = 0 Then
                mblnFueraChange = True
                Me.dbcMaterial.Text = C_TODOS
                Me.dbcMaterial.Enabled = True
                mblnFueraChange = False
            End If
        End If
        If Trim(Me.dbcMaterial.Text) = "" Then Me.dbcMaterial.Text = C_TODOS
    End Sub

    Private Sub dbcMaterial_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcMaterial.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcMaterial.Text)
        '    If Me.dbcMaterial.SelectedItem <> 0 Then
        'dbcMaterial_Leave(New Object, New EventArgs)
        '    End If
        Me.dbcMaterial.Text = Aux
    End Sub

    'Private Sub dbcMaterial_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Dim Aux As String
    '    Aux = Trim(Me.dbcMaterial.text)
    '    If Me.dbcMaterial.SelectedItem <> 0 Then
    '        dbcMaterial_LostFocus
    '    End If
    '    Me.dbcMaterial.text = Aux
    'End Sub

    Private Sub dbcSucursal_CursorChange(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String
        If mblnFueraChange Then Exit Sub
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then Exit Sub
        lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcSucursal)
        If Trim(Me.dbcSucursal.Text) = "" Then
            mintCodSucursal = 0
        End If
        If dbcSucursal.SelectedItem <> "" Then
            Call dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen WHERE TipoAlmacen = 'P'"
        ModDCombo.DCGotFocus(gStrSql, dbcSucursal)
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkTodas.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursal_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcSucursal.text)
        '''    If Me.dbcSucursal.SelectedItem <> 0 Then
        '''        dbcSucursal_LostFocus
        '''    End If
        '''    Me.dbcSucursal.text = Aux
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'Else
        '    If Trim(Me.dbcSucursal.Text) = "" Or Trim(Me.dbcSucursal.Text) = C_TODAS Then Exit Sub
        'End If
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        Aux = mintCodSucursal
        mintCodSucursal = 0
        ModDCombo.DCLostFocus((Me.dbcSucursal), gStrSql, mintCodSucursal)
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcSucursal.text)
        '''    If Me.dbcSucursal.SelectedItem <> 0 Then
        '''        dbcSucursal_LostFocus
        '''    End If
        '''    Me.dbcSucursal.text = Aux
    End Sub

    'Private Sub dbcSucursal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Dim Aux As String
    '    Aux = Trim(Me.dbcSucursal.text)
    '    If Me.dbcSucursal.SelectedItem <> 0 Then
    '        dbcSucursal_LostFocus
    '    End If
    '    Me.dbcSucursal.text = Aux
    'End Sub

    Private Sub dtpDesde_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpDesde.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpDesde_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpDesde.KeyPress
        mblnTecleoFechaI = True
        'msglTiempoCambioI = VB.Timer()
    End Sub

    Private Sub dtpHasta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpHasta.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpHasta_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpHasta.KeyPress
        mblnTecleoFechaF = True
        'msglTiempoCambioF = VB.Timer()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaRelojMaterial_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaRelojMaterial_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaRelojMaterial_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "CHKTODAS" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaRelojMaterial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaRelojMaterial_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Me.dtpDesde.MinDate = C_FECHAINICIAL
        Me.dtpDesde.MaxDate = C_FECHAFINAL
        Me.dtpHasta.MinDate = C_FECHAINICIAL
        Me.dtpHasta.MaxDate = C_FECHAFINAL
        Call Me.Nuevo()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaRelojMaterial_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If mblnSalir Then
            mblnSalir = False
            Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Sale del Formulario
                    Cancel = 0
                Case MsgBoxResult.No 'No sale del formulario
                    Me.chkTodas.Focus()
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaRelojMaterial_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class