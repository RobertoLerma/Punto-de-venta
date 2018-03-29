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

Public Class frmBancosProcesoMensualConciliacionMensual
    Inherits System.Windows.Forms.Form

    Dim isLoad As Boolean = False

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             CONCILIACION MENSUAL                                                                         *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      MARTES 05 DE AGOSTO DE 2003                                                                  *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents flexCancelados As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents cmdMovimientos As System.Windows.Forms.Button
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents lblConciliadosPosteriormente As System.Windows.Forms.Label
    Public WithEvents lblConciliados As System.Windows.Forms.Label
    Public WithEvents lblNoConciliados As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents cmbAño As System.Windows.Forms.ComboBox
    Public WithEvents cmbMes As System.Windows.Forms.ComboBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Line2 As System.Windows.Forms.Label
    Public WithEvents Line1 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents dbcBanco As System.Windows.Forms.ComboBox
    Public WithEvents dbcCuentaBancaria As System.Windows.Forms.ComboBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents lblMoneda As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBancosProcesoMensualConciliacionMensual))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmbAño = New System.Windows.Forms.ComboBox()
        Me.cmbMes = New System.Windows.Forms.ComboBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.flexCancelados = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.cmdMovimientos = New System.Windows.Forms.Button()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblConciliadosPosteriormente = New System.Windows.Forms.Label()
        Me.lblConciliados = New System.Windows.Forms.Label()
        Me.lblNoConciliados = New System.Windows.Forms.Label()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Line2 = New System.Windows.Forms.Label()
        Me.Line1 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me.dbcCuentaBancaria = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblMoneda = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Frame2.SuspendLayout()
        CType(Me.flexCancelados, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbAño
        '
        Me.cmbAño.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAño.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAño.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAño.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAño.Location = New System.Drawing.Point(56, 49)
        Me.cmbAño.Name = "cmbAño"
        Me.cmbAño.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAño.Size = New System.Drawing.Size(177, 21)
        Me.cmbAño.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.cmbAño, "Año.")
        '
        'cmbMes
        '
        Me.cmbMes.BackColor = System.Drawing.SystemColors.Window
        Me.cmbMes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbMes.Items.AddRange(New Object() {"01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"})
        Me.cmbMes.Location = New System.Drawing.Point(56, 22)
        Me.cmbMes.Name = "cmbMes"
        Me.cmbMes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbMes.Size = New System.Drawing.Size(177, 21)
        Me.cmbMes.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.cmbMes, "Mes.")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.flexCancelados)
        Me.Frame2.Controls.Add(Me.cmdMovimientos)
        Me.Frame2.Controls.Add(Me.flexDetalle)
        Me.Frame2.Controls.Add(Me.Frame4)
        Me.Frame2.Controls.Add(Me.Frame3)
        Me.Frame2.Controls.Add(Me.Label7)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Controls.Add(Me.Label12)
        Me.Frame2.Controls.Add(Me.Line2)
        Me.Frame2.Controls.Add(Me.Line1)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(16, 104)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(611, 481)
        Me.Frame2.TabIndex = 13
        Me.Frame2.TabStop = False
        '
        'flexCancelados
        '
        Me.flexCancelados.DataSource = Nothing
        Me.flexCancelados.Location = New System.Drawing.Point(16, 352)
        Me.flexCancelados.Name = "flexCancelados"
        Me.flexCancelados.OcxState = CType(resources.GetObject("flexCancelados.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexCancelados.Size = New System.Drawing.Size(577, 109)
        Me.flexCancelados.TabIndex = 6
        '
        'cmdMovimientos
        '
        Me.cmdMovimientos.BackColor = System.Drawing.SystemColors.Control
        Me.cmdMovimientos.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMovimientos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMovimientos.Location = New System.Drawing.Point(16, 110)
        Me.cmdMovimientos.Name = "cmdMovimientos"
        Me.cmdMovimientos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdMovimientos.Size = New System.Drawing.Size(107, 34)
        Me.cmdMovimientos.TabIndex = 4
        Me.cmdMovimientos.Text = "Ver Movimientos"
        Me.cmdMovimientos.UseVisualStyleBackColor = False
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(16, 176)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(577, 151)
        Me.flexDetalle.TabIndex = 5
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.Label11)
        Me.Frame4.Controls.Add(Me.Label10)
        Me.Frame4.Controls.Add(Me.Label9)
        Me.Frame4.Controls.Add(Me.lblConciliadosPosteriormente)
        Me.Frame4.Controls.Add(Me.lblConciliados)
        Me.Frame4.Controls.Add(Me.lblNoConciliados)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(336, 16)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(257, 89)
        Me.Frame4.TabIndex = 17
        Me.Frame4.TabStop = False
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label11.Location = New System.Drawing.Point(48, 63)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(204, 17)
        Me.Label11.TabIndex = 23
        Me.Label11.Text = "Conciliados en una Fecha Posterior"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label10.Location = New System.Drawing.Point(48, 40)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(145, 17)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Conciliados"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label9.Location = New System.Drawing.Point(48, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(145, 17)
        Me.Label9.TabIndex = 21
        Me.Label9.Text = "No Conciliados"
        '
        'lblConciliadosPosteriormente
        '
        Me.lblConciliadosPosteriormente.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblConciliadosPosteriormente.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblConciliadosPosteriormente.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblConciliadosPosteriormente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblConciliadosPosteriormente.Location = New System.Drawing.Point(16, 60)
        Me.lblConciliadosPosteriormente.Name = "lblConciliadosPosteriormente"
        Me.lblConciliadosPosteriormente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblConciliadosPosteriormente.Size = New System.Drawing.Size(21, 21)
        Me.lblConciliadosPosteriormente.TabIndex = 20
        '
        'lblConciliados
        '
        Me.lblConciliados.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblConciliados.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblConciliados.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblConciliados.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblConciliados.Location = New System.Drawing.Point(16, 37)
        Me.lblConciliados.Name = "lblConciliados"
        Me.lblConciliados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblConciliados.Size = New System.Drawing.Size(21, 21)
        Me.lblConciliados.TabIndex = 19
        '
        'lblNoConciliados
        '
        Me.lblNoConciliados.BackColor = System.Drawing.SystemColors.Window
        Me.lblNoConciliados.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoConciliados.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNoConciliados.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNoConciliados.Location = New System.Drawing.Point(16, 14)
        Me.lblNoConciliados.Name = "lblNoConciliados"
        Me.lblNoConciliados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNoConciliados.Size = New System.Drawing.Size(21, 21)
        Me.lblNoConciliados.TabIndex = 18
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.cmbAño)
        Me.Frame3.Controls.Add(Me.cmbMes)
        Me.Frame3.Controls.Add(Me.Label5)
        Me.Frame3.Controls.Add(Me.Label4)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(16, 16)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(281, 89)
        Me.Frame3.TabIndex = 14
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Periodo de Conciliación"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(16, 51)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(33, 21)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Año :"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(16, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(33, 21)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Mes :"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label7.Location = New System.Drawing.Point(16, 336)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(193, 17)
        Me.Label7.TabIndex = 26
        Me.Label7.Text = "Movimientos Cancelados"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label6.Location = New System.Drawing.Point(16, 160)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(185, 17)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Movimientos Vigentes"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(152, 113)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(441, 31)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Presione la Barra Espaciadora o Haga Doble Click Para Marcar un Movimiento Como C" &
    "onciliado."
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Line2
        '
        Me.Line2.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Line2.Location = New System.Drawing.Point(16, 152)
        Me.Line2.Name = "Line2"
        Me.Line2.Size = New System.Drawing.Size(576, 1)
        Me.Line2.TabIndex = 27
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Line1.Location = New System.Drawing.Point(16, 152)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(576, 1)
        Me.Line1.TabIndex = 28
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFecha)
        Me.Frame1.Controls.Add(Me.dbcBanco)
        Me.Frame1.Controls.Add(Me.dbcCuentaBancaria)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.lblMoneda)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(16, 16)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(611, 80)
        Me.Frame1.TabIndex = 7
        Me.Frame1.TabStop = False
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(504, 18)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(84, 20)
        Me.dtpFecha.TabIndex = 12
        '
        'dbcBanco
        '
        Me.dbcBanco.Location = New System.Drawing.Point(114, 18)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(185, 21)
        Me.dbcBanco.TabIndex = 0
        '
        'dbcCuentaBancaria
        '
        Me.dbcCuentaBancaria.Location = New System.Drawing.Point(114, 45)
        Me.dbcCuentaBancaria.Name = "dbcCuentaBancaria"
        Me.dbcCuentaBancaria.Size = New System.Drawing.Size(185, 21)
        Me.dbcCuentaBancaria.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(456, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(49, 21)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Fecha :"
        '
        'lblMoneda
        '
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMoneda.Location = New System.Drawing.Point(317, 33)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.Size = New System.Drawing.Size(121, 21)
        Me.lblMoneda.TabIndex = 10
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(16, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(97, 21)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Cuenta Bancaria :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(49, 21)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Banco :"
        '
        'frmBancosProcesoMensualConciliacionMensual
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(644, 604)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoMensualConciliacionMensual"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Conciliación Manual"
        Me.Frame2.ResumeLayout(False)
        CType(Me.flexCancelados, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame4.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub



    'Variables
    Dim mblnSALIR As Boolean
    'Dim mblnCambios As Boolean
    Dim mblnCerrado As Boolean
    Dim FueraChange As Boolean
    Dim intCodBanco As Integer
    Dim tecla As Integer
    Dim FechaInicial As String
    Dim FechaFinal As String
    Dim FechaUltCierre As String

    Function HayCambios() As Boolean
        Dim I As Integer
        With flexDetalle
            For I = 1 To .Rows - 1
                If CDbl(Numerico(.get_TextMatrix(I, 7))) = 1 Then
                    HayCambios = True
                    Exit Function
                End If
            Next
        End With
        HayCambios = False
    End Function

    Function Guardar() As Boolean
        On Error GoTo Merr
        Dim blnTransaccion As Boolean
        Dim I As Integer
        If mblnCerrado Then
            Guardar = False
            Exit Function
        End If
        If Not HayCambios() Then
            Guardar = False
            Exit Function
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        'Desmarcar Todos los Movimientos del Periodo
        With flexDetalle
            For I = 1 To .Rows - 1
                gStrSql = "SELECT * FROM MovimientosBancarios WHERE FolioMovto = '" & Trim(.get_TextMatrix(I, 1)) & "'"
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                RsGral = Cmd.Execute
                If RsGral.RecordCount > 0 Then
                    ModStoredProcedures.PR_IMEMovimientosBancarios(.get_TextMatrix(I, 1), "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "0", "01/01/1900", "", "", "", C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                End If
            Next
        End With
        'Marcar los Movimientos Conciliados
        With flexDetalle
            For I = 1 To .Rows - 1
                If CDbl(Trim(.get_TextMatrix(I, 5))) = 1 Then
                    gStrSql = "SELECT * FROM MovimientosBancarios WHERE FolioMovto = '" & Trim(.get_TextMatrix(I, 1)) & "'"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                    RsGral = Cmd.Execute
                    If RsGral.RecordCount > 0 Then
                        ModStoredProcedures.PR_IMEMovimientosBancarios(.get_TextMatrix(I, 1), "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), "", "", "", C_MODIFICACION, CStr(0))
                        Cmd.Execute()
                    End If
                End If
            Next
        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        Guardar = True
        MsgBox("La Información se ha Guardado Exitosamente ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Limpiar()
Merr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
            Guardar = False
        End If
    End Function

    Sub ConfiguraGrid()
        Dim I As Integer
        With flexDetalle
            .set_Cols(0, 8)
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 1200)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Fecha"
            .Col = 1
            .set_ColWidth(1, 0, 1400)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Referencia"
            .Col = 2
            .set_ColWidth(2, 0, 2500)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Tipo de Movimiento"
            .Col = 3
            .set_ColWidth(3, 0, 1500)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Ingresos"
            .Col = 4
            .set_ColWidth(4, 0, 1500)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Egresos"
            .Col = 5
            .set_ColWidth(5, 0, 0)
            For I = 1 To .Rows - 1
                .set_TextMatrix(I, 0, VB6.Format(.get_TextMatrix(I, 0), "dd/mmm/yyyy"))
                .set_TextMatrix(I, 3, VB6.Format(.get_TextMatrix(I, 3), "###,##0.00"))
                .set_TextMatrix(I, 4, VB6.Format(.get_TextMatrix(I, 4), "###,##0.00"))
                .set_TextMatrix(I, 7, 0)
                If CBool(.get_TextMatrix(I, 5)) = True Then
                    .Row = I
                    If CDate(.get_TextMatrix(I, 6)) > CDate(FechaFinal) Then
                        PonerColor(lblConciliadosPosteriormente, flexDetalle)
                        .set_TextMatrix(I, 5, 1)
                    Else
                        PonerColor(lblConciliados, flexDetalle)
                        .set_TextMatrix(I, 5, 1)
                    End If
                ElseIf CBool(.get_TextMatrix(I, 5)) = False Then
                    .set_TextMatrix(I, 5, 0)
                End If
            Next
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub ConfiguraGridCancelados()
        Dim I As Integer
        With flexCancelados
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 1100)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Fecha"
            .Col = 1
            .set_ColWidth(1, 0, 1300)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Referencia"
            .Col = 2
            .set_ColWidth(2, 0, 1300)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Folio Canc."
            .Col = 3
            .set_ColWidth(3, 0, 1800)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Tipo de Movimiento"
            .Col = 4
            .set_ColWidth(4, 0, 1300)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Ingresos"
            .Col = 5
            .set_ColWidth(5, 0, 1300)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Egresos"
            .Col = 6
            .set_ColWidth(6, 0, 0)
            For I = 1 To .Rows - 1
                .set_TextMatrix(I, 0, VB6.Format(.get_TextMatrix(I, 0), "dd/mmm/yyyy"))
                .set_TextMatrix(I, 4, VB6.Format(.get_TextMatrix(I, 4), "###,##0.00"))
                .set_TextMatrix(I, 5, VB6.Format(.get_TextMatrix(I, 5), "###,##0.00"))
                .Row = I
                If CDate(.get_TextMatrix(I, 6)) > CDate(FechaFinal) Then
                    PonerColor(lblConciliadosPosteriormente, flexCancelados)
                Else
                    PonerColor(lblConciliados, flexCancelados)
                End If
            Next
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub InicializaVariables()
        mblnSALIR = False
        mblnCerrado = False
        FueraChange = False
        intCodBanco = 0
        tecla = 0
        FechaInicial = ""
        FechaFinal = ""
    End Sub

    Sub PonerColor(ByRef Control As System.Windows.Forms.Control, ByRef Grid As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid)
        Dim I As Integer
        For I = 0 To 5
            Grid.Col = I
            Grid.CellBackColor = System.Drawing.ColorTranslator.FromOle(Control.BackColor.ToKnownColor)
        Next
        If Control.BackColor.ToKnownColor = &HFF00 Then
            If Grid.Name = "flexDetalle" Then
                Grid.set_TextMatrix(flexDetalle.Row, 5, 0)
            End If
        ElseIf Control.BackColor.ToKnownColor = &HFF00 Then
            If Grid.Name = "flexDetalle" Then
                Grid.set_TextMatrix(flexDetalle.Row, 5, 1)
            End If
        End If
        Grid.Col = 0
    End Sub

    Sub ObtenerEjercicios()
        On Error GoTo Merr
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
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InsertaUltimoCierre()
        On Error GoTo Merr
        Dim blnTransaccion As Boolean
        Cnn.BeginTrans()
        blnTransaccion = True
        ModStoredProcedures.PR_IME_ConfiguracionBancos("01/01/1900", "01/01/1900", C_INSERCION, CStr(0))
        Cmd.Execute()
        Cnn.CommitTrans()
        blnTransaccion = False
Merr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Sub Encabezado()
        Dim I As Integer
        With flexDetalle
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 1200)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Fecha"
            .Col = 1
            .set_ColWidth(1, 0, 1400)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Referencia"
            .Col = 2
            .set_ColWidth(2, 0, 2500)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Tipo de Movimiento"
            .Col = 3
            .set_ColWidth(3, 0, 1500)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Ingresos"
            .Col = 4
            .set_ColWidth(4, 0, 1500)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Egresos"
            .Col = 5
            .set_ColWidth(5, 0, 0)
            .Col = 6
            .set_ColWidth(6, 0, 0)
            .Col = 7
            .set_ColWidth(7, 0, 0)
            .Rows = 15
            For I = 1 To .Rows - 1
                .set_TextMatrix(I, 3, "0.00")
                .set_TextMatrix(I, 4, "0.00")
            Next
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub EncabezadoCancelados()
        Dim I As Integer
        With flexCancelados
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 1100)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Fecha"
            .Col = 1
            .set_ColWidth(1, 0, 1300)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Referencia"
            .Col = 2
            .set_ColWidth(2, 0, 1300)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Folio Canc."
            .Col = 3
            .set_ColWidth(3, 0, 1800)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Tipo de Movimiento"
            .Col = 4
            .set_ColWidth(4, 0, 1300)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Ingresos"
            .Col = 5
            .set_ColWidth(5, 0, 1300)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Egresos"
            .Rows = 11
            For I = 1 To .Rows - 1
                .set_TextMatrix(I, 4, "0.00")
                .set_TextMatrix(I, 5, "0.00")
            Next
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub Limpiar()
        Nuevo()
        dbcBanco.Focus()
    End Sub

    Sub Nuevo()
        dbcBanco.Text = ""
        dbcCuentaBancaria.Text = ""
        'dbcCuentaBancaria.RowSource = Nothing
        dtpFecha.Value = VB6.Format(Today, "dd/mmm/yyyy")
        lblMoneda.Text = ""
        cmbMes.SelectedIndex = 0
        cmbAño.SelectedIndex = 0
        flexDetalle.Clear()
        Encabezado()
        flexCancelados.Clear()
        EncabezadoCancelados()
        InicializaVariables()
    End Sub

    Private Sub cmbAño_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAño.Enter
        Pon_Tool()
    End Sub

    Private Sub cmbMes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMes.Enter
        Pon_Tool()
    End Sub

    Private Sub cmdMovimientos_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMovimientos.Click
        On Error GoTo Merr
        Dim blnMovVigentes As Boolean
        If Trim(dbcBanco.Text) = "" Then
            MsgBox("Proporcione el Banco.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcBanco.Focus()
            Exit Sub
        End If
        If Trim(dbcCuentaBancaria.Text) = "" Then
            MsgBox("Proporcione la Cuenta Bancaria.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcCuentaBancaria.Focus()
            Exit Sub
        End If
        If Trim(cmbMes.Text) = "" Then
            MsgBox("Proporcione el Mes.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            cmbMes.Focus()
            Exit Sub
        End If
        If Trim(cmbAño.Text) = "" Then
            MsgBox("Proporcione el Año.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            cmbAño.Focus()
            Exit Sub
        End If
        mblnCerrado = False
        flexDetalle.Clear()
        Encabezado()
        flexCancelados.Clear()
        EncabezadoCancelados()
        ObtenerLimitedeFechas(CInt(VB.Left(Trim(cmbMes.Text), 2)), CInt(Trim(cmbAño.Text)), FechaInicial, FechaFinal)
        gStrSql = "SELECT UltCierreConciliacion FROM ConfiguracionBancos"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        gStrSql = "(SELECT FechaMovto AS Fecha,FolioMovto AS Referencia," & "CASE Movimiento WHEN '" & C_MOVPAGO & "' THEN 'PAGO' WHEN '" & C_MOVDEPOSITO & "' THEN 'DEPOSITO' " & "WHEN '" & C_MOVTRASPASO & "' THEN 'TRASPASO BANC.' WHEN '" & C_MOVCARGOS & "' THEN 'CARGOS DIV.' " & "WHEN '" & C_MOVANTICIPOS & "' THEN 'ANTICIPOS A PROV/ACREED' WHEN '" & C_MOVCANCELACION & "' THEN " & "'CANCELACION' WHEN '" & C_OTROSINGRESOS & "' THEN 'OTROS INGRESOS' END AS 'Tipo De Movimiento'," & "CASE WHEN TipoMovto = 'I' THEN Importe ELSE 0 END AS INGRESOS," & "CASE WHEN TipoMovto = 'E' THEN Importe ELSE 0 END AS EGRESOS,Conciliado,FechaConciliacion " & "FROM MovimientosBancarios WHERE (FechaMovto <= '" & FechaFinal & "' AND FechaMovto >= '" & FechaInicial & "') " & "AND Movimiento <> '" & C_MOVCANCELACION & "' AND  CodBanco = " & intCodBanco & " AND CtaBancaria = '" & Trim(dbcCuentaBancaria.Text) & "' " & "AND FolioMovto NOT IN(SELECT Referencia FROM MovimientosBancarios WHERE Movimiento = '" & C_MOVCANCELACION & "'))" & "UNION " & "(SELECT FechaMovto AS Fecha,FolioMovto AS Referencia," & "CASE Movimiento WHEN '" & C_MOVPAGO & "' THEN 'PAGO' WHEN '" & C_MOVDEPOSITO & "' THEN 'DEPOSITO' " & "WHEN '" & C_MOVTRASPASO & "' THEN 'TRASPASO BANC.' WHEN '" & C_MOVCARGOS & "' THEN 'CARGOS DIV.' " & "WHEN '" & C_MOVANTICIPOS & "' THEN 'ANTICIPOS A PROV/ACREED' WHEN '" & C_MOVCANCELACION & "' THEN " & "'CANCELACION' WHEN '" & C_OTROSINGRESOS & "' THEN 'OTROS INGRESOS' END AS 'Tipo De Movimiento'," & "CASE WHEN TipoMovto = 'I' THEN Importe ELSE 0 END AS INGRESOS," & "CASE WHEN TipoMovto = 'E' THEN Importe ELSE 0 END AS EGRESOS,Conciliado,FechaConciliacion " & "FROM MovimientosBancarios WHERE ((FechaMovto < '" & FechaInicial & "' AND Conciliado = 0) " & "OR FechaConciliacion = '" & FechaFinal & "') " & "AND Movimiento <> '" & C_MOVCANCELACION & "' AND CodBanco = " & intCodBanco & " AND CtaBancaria = '" & Trim(dbcCuentaBancaria.Text) & "' " & "AND FolioMovto NOT IN(SELECT Referencia FROM MovimientosBancarios WHERE Movimiento = '" & C_MOVCANCELACION & "'))"
        If RsGral.RecordCount = 0 Then
            InsertaUltimoCierre()
            mblnCerrado = False
        Else
            If CInt(VB.Left(cmbMes.Text, 2)) <= Month(RsGral.Fields("UltCierreConciliacion").Value) And CInt(cmbAño.Text) <= Year(RsGral.Fields("UltCierreConciliacion").Value) And RsGral.Fields("UltCierreConciliacion").Value > "01/01/1900" Then
                mblnCerrado = True
            Else
                mblnCerrado = False
            End If
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount >= 0 Then
            If RsGral.RecordCount > 0 Then
                flexDetalle.Recordset = RsGral
                ConfiguraGrid()
                blnMovVigentes = True
            Else
                blnMovVigentes = False
            End If
            gStrSql = "SELECT MV.FechaMovto AS Fecha,MV.FolioMovto AS Referencia,MC.FolioMovto," & "CASE MV.Movimiento WHEN '" & C_MOVPAGO & "' THEN 'PAGO' WHEN '" & C_MOVDEPOSITO & "' THEN 'DEPOSITO' " & "WHEN '" & C_MOVTRASPASO & "' THEN 'TRASPASO BANC.' WHEN '" & C_MOVCARGOS & "' THEN 'CARGOS DIV.' " & "WHEN '" & C_MOVANTICIPOS & "' THEN 'ANTICIPOS A PROV/ACREED' WHEN '" & C_MOVCANCELACION & "' THEN " & "'CANCELACION' WHEN '" & C_OTROSINGRESOS & "' THEN 'OTROS INGRESOS' END AS 'Tipo De Movimiento'," & "CASE WHEN MV.TipoMovto = 'I' THEN MV.Importe ELSE 0 END AS Ingresos," & "CASE WHEN MV.TipoMovto = 'E' THEN MV.Importe ELSE 0 END AS Egresos,MV.FechaConciliacion FROM" & "(SELECT * FROM MovimientosBancarios WHERE Movimiento <> '" & C_MOVCANCELACION & "') MV " & "INNER JOIN " & "(SELECT * FROM MovimientosBancarios WHERE Movimiento = '" & C_MOVCANCELACION & "') MC " & "ON MV.FolioMovto = MC.Referencia " & "WHERE ((MV.FechaMovto <= '" & FechaFinal & "' AND MV.FechaMovto >= '" & FechaInicial & "') OR (MV.FechaConciliacion = '" & FechaFinal & "')) " & "AND MV.CodBanco = " & intCodBanco & " AND MV.CtaBancaria = '" & Trim(dbcCuentaBancaria.Text) & "' ORDER BY MV.FechaMovto"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                flexCancelados.Recordset = RsGral
                ConfiguraGridCancelados()
            ElseIf blnMovVigentes = False Then
                MsgBox("No Existen Movimientos a Conciliar para Esta Cuenta en este Periodo" & Chr(13) & "                       Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Limpiar()
            End If
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub cmdMovimientos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMovimientos.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcBanco_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcBanco.Name Then
        '    Exit Sub
        'End If
        flexDetalle.Clear()
        Encabezado()
        flexCancelados.Clear()
        EncabezadoCancelados()
        dbcCuentaBancaria.Text = ""
        lblMoneda.Text = ""
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
        FueraChange = False
    End Sub

    Private Sub dbcBanco_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcBanco.KeyDown
        'tecla = eventSender.keyCode
        'If eventSender.keyCode = System.Windows.Forms.Keys.Escape Then
        '    mblnSALIR = True
        '    Me.Close()
        'End If
    End Sub

    Private Sub dbcBanco_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcBanco.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' ORDER BY DescBanco"
        DCLostFocus(dbcBanco, gStrSql, intCodBanco)
    End Sub

    Private Sub dbcCuentaBancaria_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcCuentaBancaria.Name Then
        '    Exit Sub
        'End If
        flexDetalle.Clear()
        Encabezado()
        flexCancelados.Clear()
        EncabezadoCancelados()
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCChange(gStrSql, tecla)
        If Trim(dbcCuentaBancaria.Text) = "" Then
            lblMoneda.Text = ""
        End If
        'intCodBanco = 0
    End Sub

    Private Sub dbcCuentaBancaria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcCuentaBancaria.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCGotFocus(gStrSql, dbcCuentaBancaria)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcCuentaBancaria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCuentaBancaria.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            dbcBanco.Focus()
        End If
    End Sub

    Private Sub dbcCuentaBancaria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcCuentaBancaria.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcCuentaBancaria_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCuentaBancaria.KeyUp
        Dim Aux As String
        Aux = dbcCuentaBancaria.Text
        'If dbcCuentaBancaria.SelectedItem <> 0 Then
        dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        'End If
        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub dbcCuentaBancaria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Leave
        On Error GoTo Err_Renamed
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCLostFocus(dbcCuentaBancaria, gStrSql, intCodBanco)
        gStrSql = "SELECT Moneda FROM CatCuentasBancarias WHERE CtaBancaria = '" & Trim(dbcCuentaBancaria.Text) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMoneda.Visible = True
                lblMoneda.Text = C_DESCPESOS
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMoneda.Visible = True
                lblMoneda.Text = C_DESCDOLARES
            End If
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcCuentaBancaria_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcCuentaBancaria.MouseUp
        Dim Aux As String
        Aux = dbcCuentaBancaria.Text
        'If dbcCuentaBancaria.SelectedItem <> 0 Then
        dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        'End If
        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.DblClick
        On Error GoTo Merr
        If Trim(flexDetalle.get_TextMatrix(flexDetalle.Row, 0)) = "" Then Exit Sub
        If (flexDetalle.Col < 3 And Trim(flexDetalle.Text) = "") Or (flexDetalle.Col >= 3 And Trim(flexDetalle.Text) = "") Then Exit Sub
        gStrSql = "SELECT UltCierreConciliacion FROM ConfiguracionBancos"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            FechaUltCierre = RsGral.Fields("UltCierreConciliacion").Value
        Else
            FechaUltCierre = "01/01/1900"
        End If
        With flexDetalle
            If CDate(.get_TextMatrix(.Row, 0)) <= CDate(FechaUltCierre) And EstaConciliado(Trim(.get_TextMatrix(.Row, 1))) Then Exit Sub
            If ((Trim(.get_TextMatrix(.Row, 0)) <> "" And Trim(.get_TextMatrix(.Row, 1)) <> "" And Trim(.get_TextMatrix(.Row, 2)) <> "" And Trim(.get_TextMatrix(.Row, 3)) <> "") And Not mblnCerrado) Then
                If System.Drawing.ColorTranslator.ToOle(.CellBackColor) = System.Drawing.ColorTranslator.ToOle(lblConciliados.BackColor) Then
                    PonerColor(lblNoConciliados, flexDetalle)
                ElseIf System.Drawing.ColorTranslator.ToOle(.CellBackColor) = System.Drawing.ColorTranslator.ToOle(lblNoConciliados.BackColor) Then
                    PonerColor(lblConciliados, flexDetalle)
                End If
                If CDbl(.get_TextMatrix(.Row, 7)) = 0 Then
                    .set_TextMatrix(.Row, 7, 1)
                ElseIf CDbl(.get_TextMatrix(.Row, 7)) = 1 Then
                    .set_TextMatrix(.Row, 7, 0)
                End If
            End If
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub flexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        Pon_Tool()
    End Sub

    Private Sub flexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexDetalle.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Space Then
            FlexDetalle_DblClick(flexDetalle, New System.EventArgs())
        End If
    End Sub

    Private Sub frmBancosProcesoMensualConciliacionMensual_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoMensualConciliacionMensual_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoMensualConciliacionMensual_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtFolioEgreso" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSALIR = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoMensualConciliacionMensual_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoMensualConciliacionMensual_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) - 400)
        ObtenerEjercicios()
        Nuevo()
    End Sub

    Private Sub frmBancosProcesoMensualConciliacionMensual_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSALIR Then
        '    If HayCambios() Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes
        '                If Guardar() = False Then
        '                    Cancel = 1
        '                End If
        '            Case MsgBoxResult.No
        '            Case MsgBoxResult.Cancel
        '                Cancel = 1
        '        End Select
        '    End If
        'Else
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0
        '        Case MsgBoxResult.No
        '            mblnSALIR = False
        '            Cancel = 1
        '            dbcBanco.Focus()
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmBancosProcesoMensualConciliacionMensual_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
        Me.Hide()
        gblnSalir = True
    End Sub
End Class