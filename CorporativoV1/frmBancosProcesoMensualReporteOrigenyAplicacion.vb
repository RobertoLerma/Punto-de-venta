Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports System
Imports System.Windows.Forms
Imports System.Data
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports System.Data.SqlClient

Public Class frmBancosProcesoMensualReporteOrigenyAplicacion
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE ORIGEN Y APLICACIÓN DE RECURSOS                                                   *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      SABADO 09 DE AGOSTO DE 2003                                                                  *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents optAnual As System.Windows.Forms.RadioButton
    Public WithEvents optMensual As System.Windows.Forms.RadioButton
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents chkPersonales As System.Windows.Forms.CheckBox
    Public WithEvents chkJoyeria As System.Windows.Forms.CheckBox
    Public WithEvents chkEgresos As System.Windows.Forms.CheckBox
    Public WithEvents chkIngresos As System.Windows.Forms.CheckBox
    Public WithEvents txtTipoCambio As System.Windows.Forms.TextBox
    Public WithEvents chkResumen As System.Windows.Forms.CheckBox
    Public WithEvents cmbMes As System.Windows.Forms.ComboBox
    Public WithEvents cmbAño As System.Windows.Forms.ComboBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents Line2 As System.Windows.Forms.Label
    Public WithEvents Line1 As System.Windows.Forms.Label
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Public WithEvents lblTipoCambio As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.optAnual = New System.Windows.Forms.RadioButton()
        Me.optMensual = New System.Windows.Forms.RadioButton()
        Me.chkPersonales = New System.Windows.Forms.CheckBox()
        Me.chkJoyeria = New System.Windows.Forms.CheckBox()
        Me.chkEgresos = New System.Windows.Forms.CheckBox()
        Me.chkIngresos = New System.Windows.Forms.CheckBox()
        Me.txtTipoCambio = New System.Windows.Forms.TextBox()
        Me.chkResumen = New System.Windows.Forms.CheckBox()
        Me.cmbMes = New System.Windows.Forms.ComboBox()
        Me.cmbAño = New System.Windows.Forms.ComboBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Line2 = New System.Windows.Forms.Label()
        Me.Line1 = New System.Windows.Forms.Label()
        Me.lblTipoCambio = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.SuspendLayout()
        '
        'optAnual
        '
        Me.optAnual.BackColor = System.Drawing.SystemColors.Control
        Me.optAnual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAnual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAnual.Location = New System.Drawing.Point(34, 60)
        Me.optAnual.Name = "optAnual"
        Me.optAnual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAnual.Size = New System.Drawing.Size(81, 21)
        Me.optAnual.TabIndex = 7
        Me.optAnual.TabStop = True
        Me.optAnual.Text = "&Anual"
        Me.ToolTip1.SetToolTip(Me.optAnual, "Muestra todos los Movimientos del Año Seleccionado")
        Me.optAnual.UseVisualStyleBackColor = False
        '
        'optMensual
        '
        Me.optMensual.BackColor = System.Drawing.SystemColors.Control
        Me.optMensual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMensual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMensual.Location = New System.Drawing.Point(34, 30)
        Me.optMensual.Name = "optMensual"
        Me.optMensual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMensual.Size = New System.Drawing.Size(81, 21)
        Me.optMensual.TabIndex = 6
        Me.optMensual.TabStop = True
        Me.optMensual.Text = "&Mensual"
        Me.ToolTip1.SetToolTip(Me.optMensual, "Solo Muestra los Movimientos del Mes Seleccionado.")
        Me.optMensual.UseVisualStyleBackColor = False
        '
        'chkPersonales
        '
        Me.chkPersonales.BackColor = System.Drawing.SystemColors.Control
        Me.chkPersonales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPersonales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPersonales.Location = New System.Drawing.Point(381, 69)
        Me.chkPersonales.Name = "chkPersonales"
        Me.chkPersonales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPersonales.Size = New System.Drawing.Size(78, 21)
        Me.chkPersonales.TabIndex = 5
        Me.chkPersonales.Text = "&Personales"
        Me.ToolTip1.SetToolTip(Me.chkPersonales, "Muestra los Ingresos y Egresos Personales")
        Me.chkPersonales.UseVisualStyleBackColor = False
        '
        'chkJoyeria
        '
        Me.chkJoyeria.BackColor = System.Drawing.SystemColors.Control
        Me.chkJoyeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkJoyeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkJoyeria.Location = New System.Drawing.Point(381, 34)
        Me.chkJoyeria.Name = "chkJoyeria"
        Me.chkJoyeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkJoyeria.Size = New System.Drawing.Size(78, 21)
        Me.chkJoyeria.TabIndex = 4
        Me.chkJoyeria.Text = "&Joyería"
        Me.ToolTip1.SetToolTip(Me.chkJoyeria, "Muestra los Ingresos y Egresos de la Joyeria")
        Me.chkJoyeria.UseVisualStyleBackColor = False
        '
        'chkEgresos
        '
        Me.chkEgresos.BackColor = System.Drawing.SystemColors.Control
        Me.chkEgresos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEgresos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEgresos.Location = New System.Drawing.Point(298, 69)
        Me.chkEgresos.Name = "chkEgresos"
        Me.chkEgresos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEgresos.Size = New System.Drawing.Size(65, 21)
        Me.chkEgresos.TabIndex = 3
        Me.chkEgresos.Text = "&Egresos"
        Me.ToolTip1.SetToolTip(Me.chkEgresos, "Muestra los Egresos")
        Me.chkEgresos.UseVisualStyleBackColor = False
        '
        'chkIngresos
        '
        Me.chkIngresos.BackColor = System.Drawing.SystemColors.Control
        Me.chkIngresos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIngresos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkIngresos.Location = New System.Drawing.Point(298, 34)
        Me.chkIngresos.Name = "chkIngresos"
        Me.chkIngresos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIngresos.Size = New System.Drawing.Size(65, 21)
        Me.chkIngresos.TabIndex = 2
        Me.chkIngresos.Text = "&Ingresos"
        Me.ToolTip1.SetToolTip(Me.chkIngresos, "Muestra los Ingresos")
        Me.chkIngresos.UseVisualStyleBackColor = False
        '
        'txtTipoCambio
        '
        Me.txtTipoCambio.AcceptsReturn = True
        Me.txtTipoCambio.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipoCambio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambio.Location = New System.Drawing.Point(507, 147)
        Me.txtTipoCambio.MaxLength = 6
        Me.txtTipoCambio.Name = "txtTipoCambio"
        Me.txtTipoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambio.Size = New System.Drawing.Size(61, 21)
        Me.txtTipoCambio.TabIndex = 11
        Me.txtTipoCambio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambio, "Tipo de Cambio del Peso frente al Dolar")
        '
        'chkResumen
        '
        Me.chkResumen.BackColor = System.Drawing.SystemColors.Control
        Me.chkResumen.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkResumen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkResumen.Location = New System.Drawing.Point(293, 138)
        Me.chkResumen.Name = "chkResumen"
        Me.chkResumen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkResumen.Size = New System.Drawing.Size(166, 25)
        Me.chkResumen.TabIndex = 10
        Me.chkResumen.Text = "Res&umen por Agrupador"
        Me.ToolTip1.SetToolTip(Me.chkResumen, "Muestra un Resumen por Agrupador")
        Me.chkResumen.UseVisualStyleBackColor = False
        '
        'cmbMes
        '
        Me.cmbMes.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMes.Items.AddRange(New Object() {"01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"})
        Me.cmbMes.Location = New System.Drawing.Point(65, 32)
        Me.cmbMes.Name = "cmbMes"
        Me.cmbMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMes.Size = New System.Drawing.Size(185, 21)
        Me.cmbMes.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.cmbMes, "Mes.")
        '
        'cmbAño
        '
        Me.cmbAño.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAño.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAño.Location = New System.Drawing.Point(65, 59)
        Me.cmbAño.Name = "cmbAño"
        Me.cmbAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAño.Size = New System.Drawing.Size(185, 21)
        Me.cmbAño.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.cmbAño, "Año.")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optDolares)
        Me.Frame2.Controls.Add(Me.optPesos)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(9, 118)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(273, 57)
        Me.Frame2.TabIndex = 17
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Moneda"
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(160, 24)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(97, 17)
        Me.optDolares.TabIndex = 9
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "&Dolares"
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(48, 24)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(97, 17)
        Me.optPesos.TabIndex = 8
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pe&sos"
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optAnual)
        Me.Frame1.Controls.Add(Me.optMensual)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(466, 7)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(137, 105)
        Me.Frame1.TabIndex = 16
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo del Reporte"
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.cmbMes)
        Me.Frame3.Controls.Add(Me.cmbAño)
        Me.Frame3.Controls.Add(Me.Label4)
        Me.Frame3.Controls.Add(Me.Label5)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(9, 7)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(273, 105)
        Me.Frame3.TabIndex = 12
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Información del Periodo"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(24, 34)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(33, 21)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Mes :"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(24, 61)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(33, 21)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Año :"
        '
        'Line2
        '
        Me.Line2.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line2.Location = New System.Drawing.Point(370, 13)
        Me.Line2.Name = "Line2"
        Me.Line2.Size = New System.Drawing.Size(1, 96)
        Me.Line2.TabIndex = 18
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(370, 13)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(1, 96)
        Me.Line1.TabIndex = 19
        '
        'lblTipoCambio
        '
        Me.lblTipoCambio.BackColor = System.Drawing.SystemColors.Control
        Me.lblTipoCambio.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTipoCambio.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTipoCambio.Location = New System.Drawing.Point(497, 127)
        Me.lblTipoCambio.Name = "lblTipoCambio"
        Me.lblTipoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTipoCambio.Size = New System.Drawing.Size(83, 14)
        Me.lblTipoCambio.TabIndex = 15
        Me.lblTipoCambio.Text = "Tipo de Cambio :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(129, 188)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 40
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(14, 188)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 39
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoMensualReporteOrigenyAplicacion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(612, 236)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.chkPersonales)
        Me.Controls.Add(Me.chkJoyeria)
        Me.Controls.Add(Me.chkEgresos)
        Me.Controls.Add(Me.chkIngresos)
        Me.Controls.Add(Me.txtTipoCambio)
        Me.Controls.Add(Me.chkResumen)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Line2)
        Me.Controls.Add(Me.Line1)
        Me.Controls.Add(Me.lblTipoCambio)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(199, 169)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoMensualReporteOrigenyAplicacion"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Origen y Aplicación de Recursos"
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub


    'Variables
    Dim mblnSALIR As Boolean
    Dim rsReporte As ADODB.Recordset

    Sub Imprime()

        Dim rptBancosProcesoMensualReportedeOrigenyAplicacion2 As New rptBancosProcesoMensualReportedeOrigenyAplicacion2
        Dim rptBancosProcesoMensualReportedeOrigenyAplicacion As New rptBancosProcesoMensualReportedeOrigenyAplicacion


        Dim sql As String
        Dim Periodo As String
        Dim Ejercicio As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim FechaFeb As String
        Dim Moneda As String
        Dim strWhere As String
        On Error GoTo ImprimeErr

        If chkIngresos.CheckState = 0 And chkEgresos.CheckState = 0 Then
            MsgBox("Favor de Confirmar si se Mostraran los Ingresos, los Egresos o Ambos ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            chkIngresos.Focus()
            Exit Sub
        End If
        If chkJoyeria.CheckState = 0 And chkPersonales.CheckState = 0 Then
            MsgBox("Favor de Confirmar si se Mostraran los Movimientos de Joyeria, Los Personales o Ambos ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            chkJoyeria.Focus()
            Exit Sub
        End If
        If CDbl(Numerico(txtTipoCambio.Text)) = 0 Then
            MsgBox("Proporcione el Tipo de Cambio, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtTipoCambio.Focus()
            Exit Sub
        End If

        txtTipoCambio.Text = VB6.Format(txtTipoCambio.Text, "###,##0.00")
        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = UCase("Reporte de Origen y Aplicación de Recursos")

        If chkResumen.CheckState = 0 And optAnual.Checked = True Then
            'rptBancosProcesoMensualReportedeOrigenyAplicacion.Section5.Suppress = False
            'rptBancosProcesoMensualReportedeOrigenyAplicacion.Section10.Suppress = False
        ElseIf chkResumen.CheckState = 1 And optAnual.Checked = True Then
            'rptBancosProcesoMensualReportedeOrigenyAplicacion.Section5.Suppress = True
            'rptBancosProcesoMensualReportedeOrigenyAplicacion.Section10.Suppress = True
        End If
        If chkResumen.CheckState = 0 And optMensual.Checked = True Then
            'rptBancosProcesoMensualReportedeOrigenyAplicacion2.Section5.Suppress = False
            'rptBancosProcesoMensualReportedeOrigenyAplicacion2.Section10.Suppress = False
        ElseIf chkResumen.CheckState = 1 And optMensual.Checked = True Then
            'rptBancosProcesoMensualReportedeOrigenyAplicacion2.Section5.Suppress = True
            'rptBancosProcesoMensualReportedeOrigenyAplicacion2.Section10.Suppress = True
        End If
        If optPesos.Checked = True Then
            Moneda = C_PESO
        ElseIf optDolares.Checked = True Then
            Moneda = C_DOLAR
        End If
        If chkIngresos.CheckState = 1 And chkEgresos.CheckState = 1 And chkJoyeria.CheckState = 1 And chkPersonales.CheckState = 1 Then
            strWhere = " MOA.ESTATUS <> 'C' "
        Else
            If chkIngresos.CheckState = 1 And chkEgresos.CheckState = 1 And chkJoyeria.CheckState = 1 And chkPersonales.CheckState = 0 Then
                strWhere = " MOA.ESTATUS <> 'C' AND MB.TIPOPAGO = 'J' "
            ElseIf chkIngresos.CheckState = 1 And chkEgresos.CheckState = 1 And chkJoyeria.CheckState = 0 And chkPersonales.CheckState = 1 Then
                strWhere = " MOA.ESTATUS <> 'C' AND MB.TIPOPAGO = 'P' "
            ElseIf chkIngresos.CheckState = 1 And chkEgresos.CheckState = 0 And chkJoyeria.CheckState = 1 And chkPersonales.CheckState = 1 Then
                strWhere = " MOA.ESTATUS <> 'C' AND MOA.APLICACION = 'E' "
            ElseIf chkIngresos.CheckState = 1 And chkEgresos.CheckState = 0 And chkJoyeria.CheckState = 1 And chkPersonales.CheckState = 0 Then
                strWhere = " MOA.ESTATUS <> 'C' AND MOA.APLICACION = 'E' AND MB.TIPOPAGO = 'J' "
            ElseIf chkIngresos.CheckState = 1 And chkEgresos.CheckState = 0 And chkJoyeria.CheckState = 0 And chkPersonales.CheckState = 1 Then
                strWhere = " MOA.ESTATUS <> 'C' AND MOA.APLICACION = 'E' AND MB.TIPOPAGO = 'P' "
            ElseIf chkIngresos.CheckState = 0 And chkEgresos.CheckState = 1 And chkJoyeria.CheckState = 0 And chkPersonales.CheckState = 1 Then
                strWhere = " MOA.ESTATUS <> 'C' AND MOA.APLICACION = 'S' AND MB.TIPOPAGO = 'P' "
            ElseIf chkIngresos.CheckState = 0 And chkEgresos.CheckState = 1 And chkJoyeria.CheckState = 1 And chkPersonales.CheckState = 0 Then
                strWhere = " MOA.ESTATUS <> 'C' AND MOA.APLICACION = 'S' AND MB.TIPOPAGO = 'J' "
            ElseIf chkIngresos.CheckState = 0 And chkEgresos.CheckState = 1 And chkJoyeria.CheckState = 1 And chkPersonales.CheckState = 1 Then
                strWhere = " MOA.ESTATUS <> 'C' AND MOA.APLICACION = 'S' "
            End If
        End If
        If optAnual.Checked = True Then
            If BICIESTO(CInt(cmbAño.Text)) Then
                FechaFeb = CStr(29)
            Else
                FechaFeb = CStr(28)
            End If
            Ejercicio = cmbAño.Text
            sql = "SELECT T.APLICACION,T.DESCRIPCIONORIGEN,T.DESCRIPCIONRUBRO,T.SALDOANTERIOR,T.IMPORTEENERO,T.IMPORTEFEBRERO,T.IMPORTEMARZO,T.IMPORTEABRIL,T.IMPORTEMAYO,T.IMPORTEJUNIO,T.IMPORTEJULIO,T.IMPORTEAGOSTO,T.IMPORTESEPTIEMBRE,T.IMPORTEOCTUBRE,T.IMPORTENOVIEMBRE,T.IMPORTEDICIEMBRE FROM " & "(SELECT (CASE MOA.APLICACION WHEN 'E' THEN '1' ELSE '2' END) AS APLICACION,COA.DESCORIGENAPLICR AS DESCRIPCIONORIGEN ,ISNULL(CRO.DESCRUBRO,' ') AS DESCRIPCIONRUBRO,ISNULL(SUM(CASE WHEN MB.FECHAMOVTO < '" & Ejercicio & "-01-01" & "' THEN CASE MOA.APLICACION WHEN 'E' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE *" & txtTipoCambio.Text & ") " & "ELSE (MOA.IMPORTE/" & txtTipoCambio.Text & ") END END ELSE CASE MOA.APLICACION WHEN 'S' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE*" & txtTipoCambio.Text & ") ELSE (MOA.IMPORTE/" & txtTipoCambio.Text & ") END END END END END),0) AS SALDOANTERIOR,ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-01-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-01-31" & "' THEN " & "CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTEENERO,ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-02-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-02-" & FechaFeb & "' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") " & "ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTEFEBRERO,ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-03-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-03-31" & "' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTEMARZO," & "ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-04-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-04-30" & "' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTEABRIL,ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-05-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-05-31" & "' THEN " & "CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTEMAYO,ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-06-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-06-30" & "' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") " & "ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTEJUNIO,ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-07-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-07-31" & "' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTEJULIO," & "ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-08-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-08-31" & "' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTEAGOSTO,ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-09-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-09-30" & "' THEN " & "CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTESEPTIEMBRE,ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-10-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-10-31" & "' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") " & "ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTEOCTUBRE,ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-11-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-11-30" & "' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTENOVIEMBRE," & "ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & Ejercicio & "-12-01" & "' AND MB.FECHAMOVTO <= '" & Ejercicio & "-12-31" & "' THEN CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTEDICIEMBRE FROM CatOrigenAplicRecursos COA INNER JOIN CatRubrosOrigenAplicRecursos " & "CRO ON COA.CodOrigenAplicR = CRO.CodOrigAplicR INNER JOIN MovimientosOrigenAplic MOA ON COA.CodOrigenAplicR = MOA.CodOrigenAplicR AND CRO.CodRubro = MOA.CodRubro INNER JOIN MovimientosBancarios MB ON MOA.FolioMovto = MB.FolioMovto WHERE " & strWhere & "GROUP BY MOA.APLICACION,COA.DESCORIGENAPLICR,CRO.DESCRUBRO) T " & "WHERE T.SALDOANTERIOR <> 0 OR T.IMPORTEENERO <> 0 OR T.IMPORTEFEBRERO <> 0 OR T.IMPORTEMARZO <> 0 OR T.IMPORTEABRIL <> 0 OR T.IMPORTEMAYO <> 0 OR T.IMPORTEJUNIO <> 0 OR T.IMPORTEJULIO <> 0 OR T.IMPORTEAGOSTO <> 0 OR T.IMPORTESEPTIEMBRE <> 0 OR T.IMPORTEOCTUBRE <> 0 OR T.IMPORTENOVIEMBRE <> 0 OR T.IMPORTEDICIEMBRE <> 0 "
        Else
            ObtenerLimitedeFechas(CInt(VB.Left(Trim(cmbMes.Text), 2)), CInt(Trim(cmbAño.Text)), FechaInicial, FechaFinal)
            Periodo = Mid(cmbMes.Text, 5, 12)
            Ejercicio = cmbAño.Text
            sql = "SELECT T.APLICACION,T.DESCRIPCIONORIGEN,T.DESCRIPCIONRUBRO,T.SALDOANTERIOR,T.IMPORTE FROM " & "(SELECT (CASE MOA.APLICACION WHEN 'E' THEN '1' ELSE '2' END) AS APLICACION," & "COA.DESCORIGENAPLICR AS DESCRIPCIONORIGEN ,ISNULL(CRO.DESCRUBRO,' ') AS DESCRIPCIONRUBRO," & "ISNULL(SUM(CASE WHEN MB.FECHAMOVTO < '" & FechaInicial & "' THEN CASE MOA.APLICACION WHEN 'E' THEN " & "CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") " & "ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END ELSE CASE MOA.APLICACION WHEN 'S' THEN " & "CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") " & "ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END  END END),0) AS SALDOANTERIOR," & "ISNULL(SUM(CASE WHEN MB.FECHAMOVTO >= '" & FechaInicial & "' AND MB.FECHAMOVTO <= '" & FechaFinal & "' THEN " & "CASE WHEN MB.MONEDA = '" & Moneda & "' THEN MOA.IMPORTE ELSE CASE MB.MONEDA WHEN '" & C_PESO & "' THEN (MOA.IMPORTE/" & txtTipoCambio.Text & ") " & "ELSE (MOA.IMPORTE*" & txtTipoCambio.Text & ") END END END),0) AS IMPORTE " & "FROM CatOrigenAplicRecursos COA INNER JOIN CatRubrosOrigenAplicRecursos " & "CRO ON COA.CodOrigenAplicR = CRO.CodOrigAplicR INNER JOIN MovimientosOrigenAplic MOA " & "ON COA.CodOrigenAplicR = MOA.CodOrigenAplicR AND CRO.CodRubro = MOA.CodRubro " & "INNER JOIN MovimientosBancarios MB ON MOA.FolioMovto = MB.FolioMovto " & "WHERE " & strWhere & "GROUP BY MOA.APLICACION,COA.DESCORIGENAPLICR,CRO.DESCRUBRO) T " & "WHERE T.SALDOANTERIOR <> 0 OR T.IMPORTE <> 0"
        End If

        BorraCmd()
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = sql
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No Existen Movimientos de Origen y Aplicación en el Periodo Especificado , Favor de Verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        Else
            If optAnual.Checked = True Then
                'frmReportes.Report = rptBancosProcesoMensualReportedeOrigenyAplicacion
                rptBancosProcesoMensualReportedeOrigenyAplicacion.SetDataSource(frmReportes.rsReport)
                frmReportes.reporteActual = rptBancosProcesoMensualReportedeOrigenyAplicacion
            Else
                'frmReportes.Report = rptBancosProcesoMensualReportedeOrigenyAplicacion2
                rptBancosProcesoMensualReportedeOrigenyAplicacion2.SetDataSource(frmReportes.rsReport)
                frmReportes.reporteActual = rptBancosProcesoMensualReportedeOrigenyAplicacion2
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.rsReport = rsReporte

        If optAnual.Checked = True Then
            '    frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "Moneda", "TipoCambio"}
            '    frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, Moneda, txtTipoCambio}
        Else
            '    frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "Moneda", "TipoCambio", "Año", "Mes"}
            '    frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, Moneda, txtTipoCambio, Ejercicio, Periodo}
        End If
        frmReportes.Text = "Reporte de Origen y Aplicación de los Recursos"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        'frmReportes.reporteActual =
        frmReportes.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error al Imprimir : " & Err.Description, MsgBoxStyle.Exclamation, "Error de Operacion")
    End Sub

    Sub ObtenerEjercicios()
        On Error GoTo MErr
        gStrSql = "SELECT DISTINCT Ejercicio FROM EjercicioPeriodo ORDER BY Ejercicio Desc"
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
        cmbMes.Focus()
    End Sub

    Sub Nuevo()
        txtTipoCambio.Text = VB6.Format(gcurCorpoTIPOCAMBIODOLAR, "###,##0.00")
        chkEgresos.CheckState = System.Windows.Forms.CheckState.Checked
        chkIngresos.CheckState = System.Windows.Forms.CheckState.Checked
        chkJoyeria.CheckState = System.Windows.Forms.CheckState.Checked
        chkPersonales.CheckState = System.Windows.Forms.CheckState.Checked
        chkResumen.CheckState = System.Windows.Forms.CheckState.Unchecked
        optPesos.Checked = True
        optDolares.Checked = False
        optMensual.Checked = True
        optAnual.Checked = False
        cmbMes.SelectedIndex = Month(Today) - 1
        cmbAño.SelectedIndex = 0
        txtTipoCambio.Text = VB6.Format(gcurCorpoTIPOCAMBIODOLAR, "###,##0.00")
    End Sub

    Private Sub chkEgresos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkEgresos.Enter
        Pon_Tool()
    End Sub

    Private Sub chkEgresos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkEgresos.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            chkIngresos.Focus()
        End If
    End Sub

    Private Sub chkIngresos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIngresos.Enter
        Pon_Tool()
    End Sub

    Private Sub chkIngresos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkIngresos.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            cmbAño.Focus()
        End If
    End Sub

    Private Sub chkJoyeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkJoyeria.Enter
        Pon_Tool()
    End Sub

    Private Sub chkJoyeria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkJoyeria.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            chkEgresos.Focus()
        End If
    End Sub

    Private Sub chkPersonales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPersonales.Enter
        Pon_Tool()
    End Sub

    Private Sub chkPersonales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkPersonales.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            chkJoyeria.Focus()
        End If
    End Sub

    Private Sub chkResumen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkResumen.Enter
        Pon_Tool()
    End Sub

    Private Sub chkResumen_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkResumen.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If optPesos.Checked Then
                optPesos.Focus()
            Else
                optDolares.Focus()
            End If
        End If
    End Sub

    Private Sub cmbAño_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAño.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbAño_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmbAño.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If cmbMes.Enabled Then
                cmbMes.Focus()
            Else
                mblnSALIR = True
                Me.Close()
            End If
        End If
    End Sub

    Private Sub cmbMes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMes.Enter
        Pon_Tool()
    End Sub

    'UPGRADE_WARNING: Form event frmBancosProcesoMensualReporteOrigenyAplicacion.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmBancosProcesoMensualReporteOrigenyAplicacion_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        'UPGRADE_WARNING: Form method frmBancosProcesoMensualReporteOrigenyAplicacion.ZOrder has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        Me.BringToFront()
    End Sub

    'UPGRADE_WARNING: Form event frmBancosProcesoMensualReporteOrigenyAplicacion.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmBancosProcesoMensualReporteOrigenyAplicacion_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoMensualReporteOrigenyAplicacion_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "cmbMes" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSALIR = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoMensualReporteOrigenyAplicacion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoMensualReporteOrigenyAplicacion_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        ObtenerEjercicios()
        Nuevo()
    End Sub

    Private Sub frmBancosProcesoMensualReporteOrigenyAplicacion_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmBancosProcesoMensualReporteOrigenyAplicacion_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'UPGRADE_NOTE: Object frmBancosProcesoMensualReporteOrigenyAplicacion may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'Me = Nothing
    End Sub

    'UPGRADE_WARNING: Event optAnual.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optAnual_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAnual.CheckedChanged
        If eventSender.Checked Then
            cmbMes.Enabled = False
        End If
    End Sub

    Private Sub optAnual_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAnual.Enter
        Pon_Tool()
    End Sub

    Private Sub optAnual_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles optAnual.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            chkPersonales.Focus()
        End If
    End Sub

    Private Sub optDolares_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles optDolares.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If optAnual.Checked Then
                optAnual.Focus()
            Else
                optMensual.Focus()
            End If
        End If
    End Sub

    'UPGRADE_WARNING: Event optMensual.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optMensual_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMensual.CheckedChanged
        If eventSender.Checked Then
            cmbMes.Enabled = True
        End If
    End Sub

    Private Sub optMensual_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMensual.Enter
        Pon_Tool()
    End Sub

    Private Sub optMensual_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles optMensual.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            chkPersonales.Focus()
        End If
    End Sub

    Private Sub optPesos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles optPesos.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If optAnual.Checked Then
                optAnual.Focus()
            Else
                optMensual.Focus()
            End If
        End If
    End Sub

    'UPGRADE_WARNING: Event txtTipoCambio.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtTipoCambio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.TextChanged
        If Trim(txtTipoCambio.Text) = "" Then
            txtTipoCambio.Text = "0.00"
        End If
    End Sub

    Private Sub txtTipoCambio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Enter
        SelTextoTxt(txtTipoCambio)
        Pon_Tool()
    End Sub

    Private Sub txtTipoCambio_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtTipoCambio.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            chkResumen.Focus()
        End If
        If KeyCode = System.Windows.Forms.Keys.Return Then
            txtTipoCambio.Text = VB6.Format(txtTipoCambio.Text, "###,##0.00")
        End If
    End Sub

    Private Sub txtTipoCambio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.MskCantidad(txtTipoCambio.Text, KeyAscii, 3, 2, (txtTipoCambio.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub
End Class