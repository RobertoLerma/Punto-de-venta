Option Strict Off
Option Explicit On
Imports System.IO
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmVtasRPTIngresosGenerales
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkDescontar As System.Windows.Forms.CheckBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblVentas_2 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_3 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_3 As System.Windows.Forms.GroupBox
    Public WithEvents chkTodas As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean

    Dim mblnFueraChange As Boolean
    Dim mintCodSucursal As Integer
    Dim mintCodVendedor As Integer
    Dim tecla As Integer

    Dim cTablaTmp As String
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim mblnSalir As Boolean


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.chkDescontar = New System.Windows.Forms.CheckBox()
        Me._fraVtas_3 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me._lblVentas_3 = New System.Windows.Forms.Label()
        Me.chkTodas = New System.Windows.Forms.CheckBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me.fraVtas = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me._fraVtas_3.SuspendLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(12, 153)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(334, 53)
        Me.txtMensaje.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'chkDescontar
        '
        Me.chkDescontar.BackColor = System.Drawing.SystemColors.Control
        Me.chkDescontar.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDescontar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDescontar.Location = New System.Drawing.Point(12, 111)
        Me.chkDescontar.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDescontar.Name = "chkDescontar"
        Me.chkDescontar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDescontar.Size = New System.Drawing.Size(164, 23)
        Me.chkDescontar.TabIndex = 8
        Me.chkDescontar.Text = "Descontar comisión bancaria"
        Me.chkDescontar.UseVisualStyleBackColor = False
        '
        '_fraVtas_3
        '
        Me._fraVtas_3.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_3.Controls.Add(Me.dtpDesde)
        Me._fraVtas_3.Controls.Add(Me.dtpHasta)
        Me._fraVtas_3.Controls.Add(Me._lblVentas_2)
        Me._fraVtas_3.Controls.Add(Me._lblVentas_3)
        Me._fraVtas_3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_3.Location = New System.Drawing.Point(6, 46)
        Me._fraVtas_3.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_3.Name = "_fraVtas_3"
        Me._fraVtas_3.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_3.Size = New System.Drawing.Size(340, 53)
        Me._fraVtas_3.TabIndex = 3
        Me._fraVtas_3.TabStop = False
        Me._fraVtas_3.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(53, 24)
        Me.dtpDesde.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(101, 20)
        Me.dtpDesde.TabIndex = 5
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(220, 24)
        Me.dtpHasta.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(95, 20)
        Me.dtpHasta.TabIndex = 7
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_2.Location = New System.Drawing.Point(182, 28)
        Me._lblVentas_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_2.TabIndex = 6
        Me._lblVentas_2.Text = "Hasta"
        '
        '_lblVentas_3
        '
        Me._lblVentas_3.AutoSize = True
        Me._lblVentas_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_3.Location = New System.Drawing.Point(12, 28)
        Me._lblVentas_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_3.Name = "_lblVentas_3"
        Me._lblVentas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_3.Size = New System.Drawing.Size(38, 13)
        Me._lblVentas_3.TabIndex = 4
        Me._lblVentas_3.Text = "Desde"
        '
        'chkTodas
        '
        Me.chkTodas.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodas.Checked = True
        Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodas.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTodas.Location = New System.Drawing.Point(12, 6)
        Me.chkTodas.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodas.Name = "chkTodas"
        Me.chkTodas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodas.Size = New System.Drawing.Size(148, 15)
        Me.chkTodas.TabIndex = 0
        Me.chkTodas.Text = "Todas las sucursales"
        Me.chkTodas.UseVisualStyleBackColor = False
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(80, 24)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(212, 21)
        Me.dbcSucursal.TabIndex = 2
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblRpt_2.Location = New System.Drawing.Point(10, 136)
        Me._lblRpt_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 9
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_0.Location = New System.Drawing.Point(28, 26)
        Me._lblVentas_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(48, 13)
        Me._lblVentas_0.TabIndex = 1
        Me._lblVentas_0.Text = "Sucursal"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(126, 221)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 79
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(11, 221)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 78
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(241, 222)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 77
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmVtasRPTIngresosGenerales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(357, 269)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.chkDescontar)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me._fraVtas_3)
        Me.Controls.Add(Me.chkTodas)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Controls.Add(Me._lblVentas_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 29)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasRPTIngresosGenerales"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Ingresos Generales"
        Me._fraVtas_3.ResumeLayout(False)
        Me._fraVtas_3.PerformLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
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

        Me.dtpDesde.Value = Format(Today, "dd/MMM/yyyy")
        Me.dtpHasta.Value = Format(Today, "dd/MMM/yyyy")

        Me.chkDescontar.CheckState = System.Windows.Forms.CheckState.Unchecked

        Me.txtMensaje.Text = ""

        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
    End Sub

    Private Sub chkTodas_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodas.CheckStateChanged
        Select Case Me.chkTodas.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                Me.dbcSucursal.Text = C_TODAS
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

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcSucursal)

        If Trim(Me.dbcSucursal.Text) = "" Then
            mintCodSucursal = 0
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
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
        Dim Aux As String
        Aux = Trim(Me.dbcSucursal.Text)
        'If Me.dbcSucursal.SelectedItem <> 0 Then
        'dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        'End If
        Me.dbcSucursal.Text = Aux
    End Sub

    Public Function DevuelveQuery() As String
        On Error GoTo Merr
        Dim I As Integer
        Dim cSELECT As String
        Dim cWHERE As String
        Dim cINGRESONETO As String

        If Me.chkDescontar.CheckState = System.Windows.Forms.CheckState.Checked Then
            cINGRESONETO = " a.IngresoNeto - (a.ComisionBancaria + a.InteresesPromocion) "
        Else
            cINGRESONETO = " a.IngresoNeto "
        End If

        cWHERE = ""
        If mintCodSucursal <> 0 Then
            cWHERE = " and a.CodSucursal = " & mintCodSucursal
        End If

        cSELECT = " select a.CodSucursal, b.DescAlmacen as DescSucursal, " & " a.TipoMovto, a.nTipoMovto, " & " sum( " & cINGRESONETO & " ) as TotalDolares, " & " sum( dbo.ConvertirCantidad('D', 'P', " & cINGRESONETO & ", a.TipoCambio, 12)) as TotalPesos " & " from vwingresos a " & " inner join CatAlmacen b on b.CodAlmacen = a.CodSucursal " & " where a.TipoMovto is not null and (a.FechaIngreso between '" & VB6.Format(Me.dtpDesde.Value, C_FORMATFECHAGUARDAR) & "' and '" & VB6.Format(Me.dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') " & cWHERE & " group by a.CodSucursal, b.DescAlmacen, a.TipoMovto, a.nTipoMovto " & " order by b.DescAlmacen, a.nTipoMovto "

        DevuelveQuery = cSELECT
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub Imprime()

        Dim rptVentasIngresosGenerales As New rptVentasIngresosGenerales

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr
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
            Exit Sub
        Else
            rptVentasIngresosGenerales.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "Mensaje"
        'aValues(1) = Trim(Me.txtMensaje.Text)
        'aParam(2) = "dDesde"
        'aValues(2) = Me.dtpDesde.Value
        'aParam(3) = "dHasta"
        'aValues(3) = Me.dtpHasta.Value
        'aParam(4) = "Empresa"
        'aValues(4) = Trim(gstrNombCortoEmpresa)
        'aParam(5) = "DescontarComision"
        'aValues(5) = IIf(Me.chkDescontar.CheckState = System.Windows.Forms.CheckState.Checked, ModVariables.C_DESCUENTOPORCOMISIONES_SI, ModVariables.C_DESCUENTOPORCOMISIONES_NO)

        If (txtMensaje.Text <> Nothing) Then
            pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
            rptVentasIngresosGenerales.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptVentasIngresosGenerales.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        End If

        If (dtpDesde.Value <> Nothing) Then
            pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
            rptVentasIngresosGenerales.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
        End If

        If (dtpHasta.Value <> Nothing) Then
            pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
            rptVentasIngresosGenerales.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
        End If

        If (gstrNombCortoEmpresa <> Nothing) Then
            pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
            rptVentasIngresosGenerales.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
        End If

        If (Me.chkDescontar.CheckState = System.Windows.Forms.CheckState.Checked Or Me.chkDescontar.CheckState = System.Windows.Forms.CheckState.Unchecked) Then
            pdvNum.Value = IIf(Me.chkDescontar.CheckState = System.Windows.Forms.CheckState.Checked, ModVariables.C_DESCUENTOPORCOMISIONES_SI, ModVariables.C_DESCUENTOPORCOMISIONES_NO) : pvNum.Add(pdvNum)
            rptVentasIngresosGenerales.DataDefinition.ParameterFields("DescontarComision").ApplyCurrentValues(pvNum)
        End If


        frmReportes.reporteActual = rptVentasIngresosGenerales 'Es el nombre del archivo que se incluyó en el proyecto
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

    Private Sub frmVtasRPTIngresosGenerales_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasRPTIngresosGenerales_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasRPTIngresosGenerales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmVtasRPTIngresosGenerales_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasRPTIngresosGenerales_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmVtasRPTIngresosGenerales_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmVtasRPTIngresosGenerales_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub
End Class