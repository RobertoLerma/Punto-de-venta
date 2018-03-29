Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosProcesoDiarioCancelaciondeMovimientosBanc
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             CANCELACION DE MOVIMIENTOS                                                                   *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      MARTES 29 DE JULIO DE 2003                                                                   *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtImporte As System.Windows.Forms.TextBox
    Public WithEvents txtConcepto As System.Windows.Forms.TextBox
    Public WithEvents txtReferencia As System.Windows.Forms.TextBox
    Public WithEvents txtCuentaBancaria As System.Windows.Forms.TextBox
    Public WithEvents txtBanco As System.Windows.Forms.TextBox
    Public WithEvents txtTipoMovimiento As System.Windows.Forms.TextBox
    Public WithEvents txtFolioMovimiento As System.Windows.Forms.TextBox
    Public WithEvents lblFechaMovimiento As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents lblMoneda As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents txtFolioCancelacion As System.Windows.Forms.TextBox
    Public WithEvents lblFechaCancelacion As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents btnNuevo As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public bandera As Boolean = False

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtImporte = New System.Windows.Forms.TextBox()
        Me.txtConcepto = New System.Windows.Forms.TextBox()
        Me.txtReferencia = New System.Windows.Forms.TextBox()
        Me.txtCuentaBancaria = New System.Windows.Forms.TextBox()
        Me.txtBanco = New System.Windows.Forms.TextBox()
        Me.txtTipoMovimiento = New System.Windows.Forms.TextBox()
        Me.txtFolioMovimiento = New System.Windows.Forms.TextBox()
        Me.txtFolioCancelacion = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.lblFechaMovimiento = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblMoneda = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.lblFechaCancelacion = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtImporte
        '
        Me.txtImporte.AcceptsReturn = True
        Me.txtImporte.BackColor = System.Drawing.SystemColors.Window
        Me.txtImporte.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImporte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImporte.Location = New System.Drawing.Point(503, 41)
        Me.txtImporte.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImporte.MaxLength = 18
        Me.txtImporte.Name = "txtImporte"
        Me.txtImporte.ReadOnly = True
        Me.txtImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImporte.Size = New System.Drawing.Size(190, 20)
        Me.txtImporte.TabIndex = 7
        Me.txtImporte.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtImporte, "Importe del Movimiento.")
        '
        'txtConcepto
        '
        Me.txtConcepto.AcceptsReturn = True
        Me.txtConcepto.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtConcepto.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConcepto.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConcepto.Location = New System.Drawing.Point(120, 139)
        Me.txtConcepto.Margin = New System.Windows.Forms.Padding(2)
        Me.txtConcepto.MaxLength = 100
        Me.txtConcepto.Name = "txtConcepto"
        Me.txtConcepto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConcepto.Size = New System.Drawing.Size(492, 20)
        Me.txtConcepto.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtConcepto, "Motivo de la Cancelación.")
        '
        'txtReferencia
        '
        Me.txtReferencia.AcceptsReturn = True
        Me.txtReferencia.BackColor = System.Drawing.SystemColors.Window
        Me.txtReferencia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReferencia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReferencia.Location = New System.Drawing.Point(120, 115)
        Me.txtReferencia.Margin = New System.Windows.Forms.Padding(2)
        Me.txtReferencia.MaxLength = 0
        Me.txtReferencia.Name = "txtReferencia"
        Me.txtReferencia.ReadOnly = True
        Me.txtReferencia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReferencia.Size = New System.Drawing.Size(235, 20)
        Me.txtReferencia.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtReferencia, "Referencias.")
        '
        'txtCuentaBancaria
        '
        Me.txtCuentaBancaria.AcceptsReturn = True
        Me.txtCuentaBancaria.BackColor = System.Drawing.SystemColors.Window
        Me.txtCuentaBancaria.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCuentaBancaria.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCuentaBancaria.Location = New System.Drawing.Point(120, 90)
        Me.txtCuentaBancaria.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCuentaBancaria.MaxLength = 0
        Me.txtCuentaBancaria.Name = "txtCuentaBancaria"
        Me.txtCuentaBancaria.ReadOnly = True
        Me.txtCuentaBancaria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCuentaBancaria.Size = New System.Drawing.Size(235, 20)
        Me.txtCuentaBancaria.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtCuentaBancaria, "Numero de Cuenta Bancaria.")
        '
        'txtBanco
        '
        Me.txtBanco.AcceptsReturn = True
        Me.txtBanco.BackColor = System.Drawing.SystemColors.Window
        Me.txtBanco.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBanco.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBanco.Location = New System.Drawing.Point(120, 66)
        Me.txtBanco.Margin = New System.Windows.Forms.Padding(2)
        Me.txtBanco.MaxLength = 0
        Me.txtBanco.Name = "txtBanco"
        Me.txtBanco.ReadOnly = True
        Me.txtBanco.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBanco.Size = New System.Drawing.Size(235, 20)
        Me.txtBanco.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtBanco, "Nombre del Banco.")
        '
        'txtTipoMovimiento
        '
        Me.txtTipoMovimiento.AcceptsReturn = True
        Me.txtTipoMovimiento.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipoMovimiento.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoMovimiento.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoMovimiento.Location = New System.Drawing.Point(120, 42)
        Me.txtTipoMovimiento.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTipoMovimiento.MaxLength = 0
        Me.txtTipoMovimiento.Name = "txtTipoMovimiento"
        Me.txtTipoMovimiento.ReadOnly = True
        Me.txtTipoMovimiento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoMovimiento.Size = New System.Drawing.Size(235, 20)
        Me.txtTipoMovimiento.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtTipoMovimiento, "Tipo de Movimiento.")
        '
        'txtFolioMovimiento
        '
        Me.txtFolioMovimiento.AcceptsReturn = True
        Me.txtFolioMovimiento.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioMovimiento.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioMovimiento.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioMovimiento.Location = New System.Drawing.Point(120, 16)
        Me.txtFolioMovimiento.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioMovimiento.MaxLength = 13
        Me.txtFolioMovimiento.Name = "txtFolioMovimiento"
        Me.txtFolioMovimiento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioMovimiento.Size = New System.Drawing.Size(152, 20)
        Me.txtFolioMovimiento.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtFolioMovimiento, "Folio del Movimiento.")
        '
        'txtFolioCancelacion
        '
        Me.txtFolioCancelacion.AcceptsReturn = True
        Me.txtFolioCancelacion.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioCancelacion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioCancelacion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioCancelacion.Location = New System.Drawing.Point(120, 12)
        Me.txtFolioCancelacion.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioCancelacion.MaxLength = 13
        Me.txtFolioCancelacion.Name = "txtFolioCancelacion"
        Me.txtFolioCancelacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioCancelacion.Size = New System.Drawing.Size(152, 20)
        Me.txtFolioCancelacion.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtFolioCancelacion, "Folio de Cancelación.")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtImporte)
        Me.Frame1.Controls.Add(Me.txtConcepto)
        Me.Frame1.Controls.Add(Me.txtReferencia)
        Me.Frame1.Controls.Add(Me.txtCuentaBancaria)
        Me.Frame1.Controls.Add(Me.txtBanco)
        Me.Frame1.Controls.Add(Me.txtTipoMovimiento)
        Me.Frame1.Controls.Add(Me.txtFolioMovimiento)
        Me.Frame1.Controls.Add(Me.lblFechaMovimiento)
        Me.Frame1.Controls.Add(Me.Label10)
        Me.Frame1.Controls.Add(Me.lblMoneda)
        Me.Frame1.Controls.Add(Me.Label9)
        Me.Frame1.Controls.Add(Me.Label8)
        Me.Frame1.Controls.Add(Me.Label7)
        Me.Frame1.Controls.Add(Me.Label6)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(12, 58)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(709, 170)
        Me.Frame1.TabIndex = 11
        Me.Frame1.TabStop = False
        '
        'lblFechaMovimiento
        '
        Me.lblFechaMovimiento.BackColor = System.Drawing.SystemColors.Window
        Me.lblFechaMovimiento.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFechaMovimiento.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFechaMovimiento.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFechaMovimiento.Location = New System.Drawing.Point(503, 18)
        Me.lblFechaMovimiento.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFechaMovimiento.Name = "lblFechaMovimiento"
        Me.lblFechaMovimiento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFechaMovimiento.Size = New System.Drawing.Size(190, 17)
        Me.lblFechaMovimiento.TabIndex = 22
        Me.lblFechaMovimiento.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(379, 20)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(120, 17)
        Me.Label10.TabIndex = 21
        Me.Label10.Text = "Fecha del Movimiento :"
        '
        'lblMoneda
        '
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMoneda.Location = New System.Drawing.Point(394, 81)
        Me.lblMoneda.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.Size = New System.Drawing.Size(105, 17)
        Me.lblMoneda.TabIndex = 19
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(379, 41)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(132, 17)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Importe del Movimiento :"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(6, 141)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(77, 17)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "Concepto :"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(6, 118)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(77, 17)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Referencias :"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(6, 93)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(102, 17)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Cuenta Bancaria :"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(6, 68)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(77, 17)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Banco :"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(6, 44)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(120, 17)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Tipo de Movimiento :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(6, 20)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(120, 17)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Folio del Movimiento :"
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.txtFolioCancelacion)
        Me.Frame4.Controls.Add(Me.lblFechaCancelacion)
        Me.Frame4.Controls.Add(Me.Label1)
        Me.Frame4.Controls.Add(Me.Label3)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(12, 13)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(709, 40)
        Me.Frame4.TabIndex = 8
        Me.Frame4.TabStop = False
        '
        'lblFechaCancelacion
        '
        Me.lblFechaCancelacion.BackColor = System.Drawing.SystemColors.Window
        Me.lblFechaCancelacion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFechaCancelacion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFechaCancelacion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFechaCancelacion.Location = New System.Drawing.Point(530, 15)
        Me.lblFechaCancelacion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFechaCancelacion.Name = "lblFechaCancelacion"
        Me.lblFechaCancelacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFechaCancelacion.Size = New System.Drawing.Size(163, 17)
        Me.lblFechaCancelacion.TabIndex = 20
        Me.lblFechaCancelacion.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(6, 15)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(135, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Folio de Cancelación :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(406, 16)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(120, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Fecha de Cancelación :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(258, 251)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 97
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(143, 251)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 96
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(373, 252)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 95
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoDiarioCancelaciondeMovimientosBanc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(732, 293)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoDiarioCancelaciondeMovimientosBanc"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Cancelación de Movimientos Bancarios"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Dim mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Dim mblnSalir As Boolean
    Dim mstrMovimiento As String
    Dim mstrNaturaleza As String
    Dim mstrFolioProgramacion As String
    Dim mcurTipoCambio As Decimal
    Dim mstrFormaPago As String
    Dim mstrTipoPago As String
    Dim mintCodBanco As Integer
    Dim mstrBeneficiario As String
    Dim mstrNoDocto As String
    Dim mdtmFechaDocto As Date
    Dim mstrFolioRetiro As String
    Dim intNumPartida As Integer
    Dim Modulo As String
    Public strControlActual As String

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim I As Integer

        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control 

        Select Case strControlActual
            Case "TXTFOLIOCANCELACION"
                strCaptionForm = "Consulta de Folios de Cancelación"
                gStrSql = "SELECT FolioMovto AS FOLIO,Concepto AS CONCEPTO," & "FechaMovto AS FECHA,ABS(Importe) AS IMPORTE FROM MovimientosBancarios " & "WHERE FolioMovto LIKE '" & txtFolioCancelacion.Text & "%' AND Movimiento = '" & C_MOVCANCELACION & "' ORDER BY FechaMovto DESC ,FolioMovto DESC"
            Case "TXTFOLIOMOVIMIENTO"
                strCaptionForm = "Consulta de Movimientos"
                gStrSql = "SELECT FOLIOMOVTO AS 'FOLIO',(CASE MOVIMIENTO WHEN '" & C_MOVPAGO & "' THEN 'PAGO' " & "WHEN '" & C_MOVDEPOSITO & "' THEN 'DEPOSITO' WHEN '" & C_MOVTRASPASO & "' THEN 'TRASPASO BANCARIO' " & "WHEN '" & C_MOVCARGOS & "' THEN 'CARGOS DIVERSOS' WHEN '" & C_OTROSINGRESOS & "' THEN 'OTROS INGRESOS' WHEN '" & C_MOVANTICIPOS & "' THEN 'ANTICIPO A PROV./ACREED.' END) AS 'MOVIMIENTO'," & "(CASE TIPOMOVTO WHEN '" & C_TIPOMOVINGRESO & "' THEN 'INGRESO' WHEN '" & C_TIPOMOVEGRESO & "' THEN 'EGRESO' END) AS 'TIPO DE MOVIMIENTO'," & "FECHAMOVTO AS 'FECHA',IMPORTE AS 'IMPORTE' FROM MOVIMIENTOSBANCARIOS WHERE FOLIOMOVTO NOT " & "IN(SELECT MB.FOLIOMOVTO FROM MOVIMIENTOSBANCARIOS MB INNER JOIN " & "(SELECT * FROM MOVIMIENTOSBANCARIOS WHERE MOVIMIENTO = '" & C_MOVCANCELACION & "') AUX ON MB.FOLIOMOVTO = AUX.REFERENCIA) " & "AND FOLIOMOVTO LIKE '" & txtFolioMovimiento.Text & "%' AND MOVIMIENTO <> '" & C_MOVCANCELACION & "' ORDER BY FECHAMOVTO DESC,FOLIOMOVTO DESC"
            Case Else
                Exit Sub
        End Select

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsjNoExiste(C_msgSINDATOS, gstrNombCortoEmpresa)
            Exit Sub
        End If

        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
        FrmConsultas.Left = VB6.TwipsToPixelsX(2500)
        ConfiguraConsultas(FrmConsultas, 400, RsGral, strTag, strCaptionForm)

        'Carga el formulario de consulta 
        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTFOLIOCANCELACION"
                    'ConfiguraConsultas(FrmConsultas, 400, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 1400) 'Columna del Folio
                    .set_ColWidth(1, 0, 5000) 'Columna del Concepto del Movimiento
                    .set_ColWidth(2, 0, 1200) 'Columna de la Fecha del Movimiento
                    .set_ColWidth(3, 0, 1800) 'Columna del Importe del Movimiento
                    .set_ColAlignment(2, 4)
                    For I = 1 To FrmConsultas.Flexdet.Rows - 1
                        FrmConsultas.Flexdet.set_TextMatrix(I, 2, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 2), "dd/MMM/yyyy"))
                        FrmConsultas.Flexdet.set_TextMatrix(I, 3, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 3), "###,##0.00"))
                    Next I
                    'FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
                    'FrmConsultas.Left = VB6.TwipsToPixelsX(2500)
                Case "TXTFOLIOMOVIMIENTO"
                    'ConfiguraConsultas(FrmConsultas, 400, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 1600) 'Columna del Folio
                    .set_ColWidth(1, 0, 3500) 'Columna del Movimiento
                    .set_ColWidth(2, 0, 3000) 'Columna del Tipo de Movimiento
                    .set_ColWidth(3, 0, 1400) 'Columna de la Fecha del Movimiento
                    .set_ColWidth(4, 0, 1800) 'Columna del Importe del Movimiento
                    .set_ColAlignment(3, 4)
                    For I = 1 To FrmConsultas.Flexdet.Rows - 1
                        FrmConsultas.Flexdet.set_TextMatrix(I, 3, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 3), "dd/MMM/yyyy"))
                        FrmConsultas.Flexdet.set_TextMatrix(I, 4, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 4), "###,##0.00"))
                    Next I
                    'FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
                    'FrmConsultas.Left = VB6.TwipsToPixelsX(1500)
            End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function Cancelar() As Boolean
        Dim blnTransaccion As Boolean
        Dim strFolioCancelacion As String
        Dim Ejercicio As Integer
        Dim Periodo As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim RsAux As ADODB.Recordset
        Dim CodBanco As Integer
        Dim FechaPasoBancos As String
        Dim Autorizacion As String
        Dim CodSucursal As Integer
        On Error GoTo Err_Renamed
        If Not mblnNuevo Then
            Exit Function
        End If
        If ValidaDatos() = False Then
            Exit Function
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        'Generar Folio del Movimiento
        Ejercicio = CInt(VB6.Format(Year(CDate(lblFechaCancelacion.Text)), "0000"))
        Periodo = VB6.Format(Month(CDate(lblFechaCancelacion.Text)), "00")
        BuscaEjercicio(lblFechaCancelacion.Text)
        ObtenerLimitedeFechas(CInt(Periodo), Ejercicio, FechaInicial, FechaFinal)
        FechaFinal = VB6.Format(FechaFinal, "dd/mmm/yyyy")
        gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVCANCELACION & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFolioCancelacion.Text = C_TIPOMOVCANCELACION & VB6.Format(Year(CDate(lblFechaCancelacion.Text)), "0000") & VB6.Format(Month(CDate(lblFechaCancelacion.Text)), "00") & VB6.Format(VB.Day(CDate(lblFechaCancelacion.Text)), "00") & VB6.Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
            ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVCANCELACION, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        Select Case mstrMovimiento
            Case C_MOVPAGO
                'Guardar el Movimiento Bancario de Cancelación
                ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioCancelacion.Text, VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MOVCANCELACION, C_TIPOMOVEGRESO, IIf(mstrNaturaleza = C_NATURALEZACOMERCIAL, C_NATURALEZACOMERCIAL, C_NATURALEZAINTERNA), IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(mcurTipoCambio), mstrFormaPago, mstrTipoPago, CStr(mintCodBanco), txtCuentaBancaria.Text, mstrBeneficiario, txtConcepto.Text, IIf(Trim(mstrFolioProgramacion) = "", "0", "1"), mstrFolioProgramacion, "0", VB6.Format(CStr(mdtmFechaDocto), C_FORMATFECHAGUARDAR), mstrNoDocto, CStr(CDbl(Numerico(txtImporte.Text)) * -1), "V", "01/01/1900", mstrFolioRetiro, "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), Modulo, txtFolioMovimiento.Text, "", C_INSERCION, CStr(0))
                Cmd.Execute()
                If Trim(Modulo) = C_MODULOCXP Then
                    gStrSql = "SELECT * FROM Pagos WHERE FolioPagoBancos = '" & txtFolioMovimiento.Text & "'"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                    RsGral = Cmd.Execute
                    If RsGral.RecordCount > 0 Then
                        Do While Not RsGral.EOF
                            ModStoredProcedures.PR_IMEPagos(RsGral.Fields("FolioProgramacionP").Value, CStr(RsGral.Fields("NumPartida").Value), "0", "", "", "", "01/01/1900", "01/01/1900", "0", "", "0", "0", "0", "0", "C", VB6.Format(Today, C_FORMATFECHAGUARDAR), "", "0", "0", "01/01/1900", txtFolioMovimiento.Text, CStr(RsGral.Fields("PartidaPago").Value), C_MODIFICACION, CStr(1))
                            Cmd.Execute()
                            ModStoredProcedures.PR_IMEProgramacionPagos(RsGral.Fields("FolioProgramacionP").Value, CStr(RsGral.Fields("NumPartida").Value), "0", "", "", "", "01/01/1900", "01/01/1900", "0", "", "0", "0", "0", "0", "0", "V", "01/01/1900", "", "0", "0", "01/01/1900", C_MODIFICACION, CStr(3))
                            Cmd.Execute()
                            RsGral.MoveNext()
                        Loop
                    End If
                    gStrSql = "SELECT * FROM NotasCreditoCab WHERE FolioPagoBancos = '" & txtFolioMovimiento.Text & "'"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                    RsGral = Cmd.Execute
                    If RsGral.RecordCount > 0 Then
                        Do While Not RsGral.EOF
                            ModStoredProcedures.PR_IMENotasCreditoCab(RsGral.Fields("FolioNotaCredito").Value, "01/01/1900", "", "0", "", "", "", "0", "0", "0", "0", "V", "01/01/1900", "0", "0", "0", "01/01/1900", "", "", "0", C_MODIFICACION, CStr(4))
                            Cmd.Execute()
                            RsGral.MoveNext()
                        Loop
                    End If
                    gStrSql = "SELECT * FROM Anticipos WHERE FolioPagoBancos = '" & txtFolioMovimiento.Text & "'"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                    RsGral = Cmd.Execute
                    If RsGral.RecordCount > 0 Then
                        Do While Not RsGral.EOF
                            ModStoredProcedures.PR_IME_Anticipos(RsGral.Fields("FolioAnticipo").Value, "01/01/1900", "", "0", "", "", "0", "0", "0", "0", "V", "01/01/1900", "0", "0", "01/01/1900", "", C_MODIFICACION, CStr(3))
                            Cmd.Execute()
                            RsGral.MoveNext()
                        Loop
                    End If
                End If
                'Cancelar los Movimientos de Origen y Aplicación
                ModStoredProcedures.PR_IMEMovimientosOrigenAplic(txtFolioMovimiento.Text, "0", "0", "0", "0", "", "0", "C", VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MODIFICACION, CStr(0))
                Cmd.Execute()
                'Conciliar el Movimiento Cancelado
                ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioMovimiento.Text, "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), "", "", "", C_MODIFICACION, CStr(0))
                Cmd.Execute()
            Case C_MOVDEPOSITO
                If mstrNaturaleza = C_NATURALEZACOMERCIAL Then
                    'Guardar el Movimiento Bancario de Cancelación
                    ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioCancelacion.Text, VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MOVCANCELACION, C_TIPOMOVINGRESO, C_NATURALEZACOMERCIAL, IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(mcurTipoCambio), mstrFormaPago, mstrTipoPago, CStr(mintCodBanco), txtCuentaBancaria.Text, mstrBeneficiario, txtConcepto.Text, "0", mstrFolioProgramacion, "0", VB6.Format(CStr(mdtmFechaDocto), C_FORMATFECHAGUARDAR), mstrNoDocto, CStr(CDbl(Numerico(txtImporte.Text)) * -1), "V", "01/01/1900", mstrFolioRetiro, "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), C_MODULOBANCOS, txtFolioMovimiento.Text, "", C_INSERCION, CStr(0))
                    Cmd.Execute()
                    'Cancelar los Movimientos de Origen y Aplicación
                    ModStoredProcedures.PR_IMEMovimientosOrigenAplic(txtFolioMovimiento.Text, "0", "0", "0", "0", "", "0", "C", VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                    'Cancelar los Movimientos de Referencias
                    ModStoredProcedures.PR_IMEMovimientosReferencias(txtFolioMovimiento.Text, "0", "0", "", "0", "C", "", C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                    'Checar Los Movimientos de Referencias
                    gStrSql = "SELECT * FROM MovimientosReferencias WHERE FolioMovto = '" & txtFolioMovimiento.Text & "'"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                    RsGral = Cmd.Execute
                    If RsGral.RecordCount > 0 Then
                        If RsGral.Fields("TipoReferencia").Value = "V" Then
                            Do While Not RsGral.EOF
                                CodBanco = mintCodBanco
                                FechaPasoBancos = Mid(txtFolioMovimiento.Text, 2, 4) & "-" & Mid(txtFolioMovimiento.Text, 6, 2) & "-" & Mid(txtFolioMovimiento.Text, 8, 2)
                                CodSucursal = CInt(Mid(RsGral.Fields("ReferenciaBanco").Value, 5, 2))
                                Autorizacion = Trim(Mid(RsGral.Fields("ReferenciaBanco").Value, 12, 10))
                                gStrSql = "SELECT ING.FOLIOINGRESO,INGFP.AUTORIZACION,INGFP.PASOBANCOS,INGFP.FECHAPASOBANCOS,INGFP.NUMPARTIDA " & "FROM INGRESOS ING INNER JOIN INGRESOSFORMADEPAGO INGFP " & "ON ING.FOLIOINGRESO = INGFP.FOLIOINGRESO " & "WHERE ING.CODSUCURSAL = " & CodSucursal & " AND INGFP.CODBANCO = " & CodBanco & " AND INGFP.FECHAPASOBANCOS = '" & FechaPasoBancos & "' " & "AND INGFP.PASOBANCOS = 1 AND INGFP.AUTORIZACION = '" & Autorizacion & "' " & "AND ING.ESTATUS <> 'C'"
                                ModEstandar.BorraCmd()
                                Cmd.CommandText = "dbo.Up_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                                RsAux = Cmd.Execute
                                If RsAux.RecordCount > 0 Then
                                    ModStoredProcedures.PR_IEIngresosFormasdePago(RsAux.Fields("FolioIngreso").Value, CStr(RsAux.Fields("NumPartida").Value), "01/01/1900", "", "0", "0", "0", "0", "", "", "", "", "0", "0", "0", "", "01/01/1900", "0", "01/01/1900", CStr(0), C_MODIFICACION, CStr(0))
                                    Cmd.Execute()
                                End If
                                RsGral.MoveNext()
                            Loop
                        End If
                    End If
                    'Conciliar el Movimiento Cancelado
                    ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioMovimiento.Text, "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), "", "", "", C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                ElseIf mstrNaturaleza = C_NATURALEZAINTERNA Then
                    'Guardar el Movimiento Bancario de Cancelación
                    ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioCancelacion.Text, VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MOVCANCELACION, C_TIPOMOVINGRESO, C_NATURALEZAINTERNA, IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(mcurTipoCambio), mstrFormaPago, mstrTipoPago, CStr(mintCodBanco), txtCuentaBancaria.Text, mstrBeneficiario, txtConcepto.Text, "0", mstrFolioProgramacion, "0", VB6.Format(CStr(mdtmFechaDocto), C_FORMATFECHAGUARDAR), mstrNoDocto, CStr(CDbl(Numerico(txtImporte.Text)) * -1), "V", "01/01/1900", mstrFolioRetiro, "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), C_MODULOBANCOS, txtFolioMovimiento.Text, "", C_INSERCION, CStr(0))
                    Cmd.Execute()
                    'Restaurar el Folio de Retiro
                    If lblMoneda.Text = C_DESCPESOS Then
                        gStrSql = "SELECT R.CodFormaPago FROM Retiros R,CatFormasPago FP WHERE R.FolioRetiro = '" & mstrFolioRetiro & "' AND " & "FP.CodFormaPago = R.CodFormaPago AND FP.EsDolar = 0 GROUP BY R.CodFormaPago"
                    ElseIf lblMoneda.Text = C_DESCDOLARES Then
                        gStrSql = "SELECT R.CodFormaPago FROM Retiros R,CatFormasPago FP WHERE R.FolioRetiro = '" & mstrFolioRetiro & "' AND " & "FP.CodFormaPago = R.CodFormaPago AND FP.EsDolar = 1 GROUP BY R.CodFormaPago"
                    End If
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                    RsGral = Cmd.Execute
                    If RsGral.RecordCount > 0 Then
                        Do While Not RsGral.EOF
                            ModStoredProcedures.PR_IE_Retiros(mstrFolioRetiro, "01/01/1900", "0", "0", "", "", "", "0", "", "0", "", "01/01/1900", "0", "01/01/1900", "0", RsGral.Fields("CodFormaPago").Value, "0", C_MODIFICACION, CStr(0))
                            Cmd.Execute()
                            RsGral.MoveNext()
                        Loop
                    End If
                    'Conciliar el Movimiento Cancelado
                    ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioMovimiento.Text, "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), "", "", "", C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                End If
            Case C_OTROSINGRESOS
                'Guardar el Movimiento Bancario de Cancelación
                ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioCancelacion.Text, VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MOVCANCELACION, C_TIPOMOVINGRESO, mstrNaturaleza, IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(mcurTipoCambio), mstrFormaPago, mstrTipoPago, CStr(mintCodBanco), txtCuentaBancaria.Text, mstrBeneficiario, txtConcepto.Text, "0", mstrFolioProgramacion, "0", VB6.Format(CStr(mdtmFechaDocto), C_FORMATFECHAGUARDAR), mstrNoDocto, CStr(CDbl(Numerico(txtImporte.Text)) * -1), "V", "01/01/1900", mstrFolioRetiro, "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), C_MODULOBANCOS, txtFolioMovimiento.Text, "", C_INSERCION, CStr(0))
                Cmd.Execute()
                'Cancelar los Movimientos de Origen y Aplicación
                ModStoredProcedures.PR_IMEMovimientosOrigenAplic(txtFolioMovimiento.Text, "0", "0", "0", "0", "", "0", "C", VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MODIFICACION, CStr(0))
                Cmd.Execute()
                'Cancelar los Movimientos de Referencias
                ModStoredProcedures.PR_IMEMovimientosReferencias(txtFolioMovimiento.Text, "0", "0", "", "0", "C", "", C_MODIFICACION, CStr(0))
                Cmd.Execute()
                'Conciliar el Movimiento Cancelado
                ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioMovimiento.Text, "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), "", "", "", C_MODIFICACION, CStr(0))
                Cmd.Execute()
            Case C_MOVCARGOS
                'Guardar el Movimiento Bancario de Cancelación
                ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioCancelacion.Text, VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MOVCANCELACION, C_TIPOMOVEGRESO, IIf(mstrNaturaleza = C_NATURALEZACOMERCIAL, C_NATURALEZACOMERCIAL, C_NATURALEZAINTERNA), IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(mcurTipoCambio), mstrFormaPago, mstrTipoPago, CStr(mintCodBanco), txtCuentaBancaria.Text, mstrBeneficiario, txtConcepto.Text, "0", mstrFolioProgramacion, "0", VB6.Format(CStr(mdtmFechaDocto), C_FORMATFECHAGUARDAR), mstrNoDocto, CStr(CDbl(Numerico(txtImporte.Text)) * -1), "V", "01/01/1900", mstrFolioRetiro, "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), C_MODULOBANCOS, txtFolioMovimiento.Text, "", C_INSERCION, CStr(0))
                Cmd.Execute()
                'Cancelar los Movimientos de Origen y Aplicación
                ModStoredProcedures.PR_IMEMovimientosOrigenAplic(txtFolioMovimiento.Text, "0", "0", "0", "0", "", "0", "C", VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MODIFICACION, CStr(0))
                Cmd.Execute()
                'Cancelar los Movimientos de Referencias
                ModStoredProcedures.PR_IMEMovimientosReferencias(txtFolioMovimiento.Text, "0", "0", "", "0", "C", "", C_MODIFICACION, CStr(0))
                Cmd.Execute()
                'Conciliar el Movimiento Cancelado
                ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioMovimiento.Text, "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), "", "", "", C_MODIFICACION, CStr(0))
                Cmd.Execute()
            Case C_MOVTRASPASO
                'Guardar el Movimiento Bancario de Cancelación
                ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioCancelacion.Text, VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MOVCANCELACION, IIf(VB.Left(txtFolioMovimiento.Text, 1) = C_TIPOMOVINGRESO, C_TIPOMOVINGRESO, C_TIPOMOVEGRESO), IIf(mstrNaturaleza = C_NATURALEZACOMERCIAL, C_NATURALEZACOMERCIAL, C_NATURALEZAINTERNA), IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(mcurTipoCambio), mstrFormaPago, mstrTipoPago, CStr(mintCodBanco), txtCuentaBancaria.Text, mstrBeneficiario, txtConcepto.Text, "0", mstrFolioProgramacion, "0", VB6.Format(CStr(mdtmFechaDocto), C_FORMATFECHAGUARDAR), mstrNoDocto, CStr(CDbl(Numerico(txtImporte.Text)) * -1), "V", "01/01/1900", mstrFolioRetiro, "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), C_MODULOBANCOS, txtFolioMovimiento.Text, "", C_INSERCION, CStr(0))
                Cmd.Execute()
                'Conciliar el Movimiento Cancelado
                ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioMovimiento.Text, "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), "", "", "", C_MODIFICACION, CStr(0))
                Cmd.Execute()
                'Generar el Segundo Folio del Movimiento
                Ejercicio = CInt(VB6.Format(Year(CDate(lblFechaCancelacion.Text)), "0000"))
                Periodo = VB6.Format(Month(CDate(lblFechaCancelacion.Text)), "00")
                BuscaEjercicio(lblFechaCancelacion.Text)
                gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVCANCELACION & "'"
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute
                If RsGral.RecordCount > 0 Then
                    strFolioCancelacion = C_TIPOMOVCANCELACION & VB6.Format(Year(CDate(lblFechaCancelacion.Text)), "0000") & VB6.Format(Month(CDate(lblFechaCancelacion.Text)), "00") & VB6.Format(VB.Day(CDate(lblFechaCancelacion.Text)), "00") & VB6.Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
                    ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVCANCELACION, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                End If
                gStrSql = "SELECT * FROM MovimientosBancarios WHERE FolioMovto = '" & txtReferencia.Text & "'"
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute
                If RsGral.RecordCount > 0 Then
                    'Guardar el Segundo Movimiento Bancario de Cancelación
                    ModStoredProcedures.PR_IMEMovimientosBancarios(strFolioCancelacion, VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MOVCANCELACION, IIf(VB.Left(txtReferencia.Text, 1) = C_TIPOMOVINGRESO, C_TIPOMOVINGRESO, C_TIPOMOVEGRESO), RsGral.Fields("Naturaleza").Value, RsGral.Fields("Moneda").Value, CStr(RsGral.Fields("TipoCambio").Value), RsGral.Fields("FormaPago").Value, RsGral.Fields("TipoPago").Value, CStr(RsGral.Fields("CodBanco").Value), RsGral.Fields("CtaBancaria").Value, RsGral.Fields("Beneficiario").Value, txtConcepto.Text, "0", "", "0", VB6.Format(CStr(RsGral.Fields("FechaDocto").Value), C_FORMATFECHAGUARDAR), RsGral.Fields("NoDocto").Value, CStr(RsGral.Fields("importe").Value * -1), "V", "01/01/1900", "", "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), C_MODULOBANCOS, txtReferencia.Text, "", C_INSERCION, CStr(0))
                    Cmd.Execute()
                    'Conciliar el Movimiento Cancelado
                    ModStoredProcedures.PR_IMEMovimientosBancarios(txtReferencia.Text, "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), "", "", "", C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                End If
            Case C_MOVANTICIPOS
                'Guardar el Movimiento Bancario de Cancelación
                ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioCancelacion.Text, VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MOVCANCELACION, C_TIPOMOVEGRESO, IIf(mstrNaturaleza = C_NATURALEZACOMERCIAL, C_NATURALEZACOMERCIAL, C_NATURALEZAINTERNA), IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(mcurTipoCambio), mstrFormaPago, mstrTipoPago, CStr(mintCodBanco), txtCuentaBancaria.Text, mstrBeneficiario, txtConcepto.Text, "0", mstrFolioProgramacion, "0", VB6.Format(CStr(mdtmFechaDocto), C_FORMATFECHAGUARDAR), mstrNoDocto, CStr(CDbl(Numerico(txtImporte.Text)) * -1), "V", "01/01/1900", mstrFolioRetiro, "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), C_MODULOBANCOS, txtFolioMovimiento.Text, "", C_INSERCION, CStr(0))
                Cmd.Execute()
                'Cancelar los Movimientos de Origen y Aplicación
                ModStoredProcedures.PR_IMEMovimientosOrigenAplic(txtFolioMovimiento.Text, "0", "0", "0", "0", "", "0", "C", VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), C_MODIFICACION, CStr(0))
                Cmd.Execute()
                'Cancelar el Anticipo
                ModStoredProcedures.PR_IME_Anticipos("", "01/01/1900", txtFolioMovimiento.Text, "0", "", "", "0", "0", "0", "0", "C", VB6.Format(lblFechaCancelacion.Text, C_FORMATFECHAGUARDAR), "0", "0", "01/01/1900", "", C_MODIFICACION, CStr(0))
                Cmd.Execute()
                'Conciliar el Movimiento Cancelado
                ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioMovimiento.Text, "01/01/1900", "", "", "", "", "0", "", "", "0", "", "", "", "0", "", "0", "01/01/1900", "", "0", "", "01/01/1900", "", "1", VB6.Format(FechaFinal, C_FORMATFECHAGUARDAR), "", "", "", C_MODIFICACION, CStr(0))
                Cmd.Execute()
        End Select
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        Select Case mstrMovimiento
            Case C_MOVPAGO, C_MOVDEPOSITO, C_MOVCARGOS, C_MOVANTICIPOS, C_OTROSINGRESOS
                MsgBox("Los Datos se Han Guardado con Exito." & Chr(13) & "Se ha Generado el Folio de Cancelación " & txtFolioCancelacion.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Case C_MOVTRASPASO
                MsgBox("Los Datos se Han Guardado con Exito." & Chr(13) & "Se Han Generado los Siguientes Folios de Cancelación" & Chr(13) & "El Folio " & txtFolioCancelacion.Text & " y el " & strFolioCancelacion, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        End Select
        Limpiar()
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Sub Limpiar()
        Nuevo()
        InicializaVariables()
        txtFolioCancelacion.Text = ""
        txtFolioMovimiento.Text = ""
        txtFolioCancelacion.Focus()
    End Sub

    Sub LlenaDatos()

        If (bandera = True) Then
            Exit Sub
        End If

        On Error GoTo Merr
        If Trim(txtFolioCancelacion.Text) = "" Then
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT MB1.FolioMovto AS FolioCancelacion,MB1.FechaMovto AS FechaCancelacion,MB1.Concepto," & "MB2.FolioMovto AS FolioMovimiento,MB2.FechaMovto AS FechaMovimiento,MB2.Movimiento,MB2.CtaBancaria," & "MB2.Moneda,MB2.Referencia,MB2.Importe,CB.DescBanco " & "FROM (SELECT * FROM MovimientosBancarios WHERE FolioMovto = '" & txtFolioCancelacion.Text & "') MB1 INNER JOIN MovimientosBancarios MB2 ON MB2.FolioMovto = MB1.Referencia," & "CatBancos CB WHERE MB1.FolioMovto = '" & txtFolioCancelacion.Text & "' AND MB1.Movimiento = '" & C_MOVCANCELACION & "' " & "AND CB.CodBanco = MB2.CodBanco"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            mblnNuevo = False
            Dim fechaCancelacion As String = AgregarHoraAFecha(RsGral.Fields("FechaCancelacion").Value.ToString())
            Dim fechaMovimiento As String = AgregarHoraAFecha(RsGral.Fields("FechaMovto").Value.ToString())
            'lblFechaCancelacion.Text = VB6.Format(RsGral.Fields("FechaCancelacion").Value, C_FORMATFECHAMOSTRAR)
            lblFechaCancelacion.Text = fechaCancelacion
            'lblFechaMovimiento.Text = VB6.Format(RsGral.Fields("FechaMovimiento").Value, C_FORMATFECHAMOSTRAR)
            lblFechaMovimiento.Text = fechaMovimiento
            If RsGral.Fields("Movimiento").Value = C_MOVPAGO Then
                txtTipoMovimiento.Text = "PAGO"
            ElseIf RsGral.Fields("Movimiento").Value = C_MOVDEPOSITO Then
                txtTipoMovimiento.Text = "DEPOSITO"
            ElseIf RsGral.Fields("Movimiento").Value = C_MOVTRASPASO Then
                txtTipoMovimiento.Text = "TRASPASO BANCARIO"
            ElseIf RsGral.Fields("Movimiento").Value = C_MOVCARGOS Then
                txtTipoMovimiento.Text = "CARGOS DIVERSOS"
            ElseIf RsGral.Fields("Movimiento").Value = C_MOVANTICIPOS Then
                txtTipoMovimiento.Text = "ANTICIPO A PROV./ACREED."
            ElseIf RsGral.Fields("Movimiento").Value = C_OTROSINGRESOS Then
                txtTipoMovimiento.Text = "OTROS INGRESOS"
            End If
            txtFolioMovimiento.Text = RsGral.Fields("FolioMovimiento").Value
            txtBanco.Text = Trim(RsGral.Fields("DescBanco").Value)
            txtCuentaBancaria.Text = Trim(RsGral.Fields("CtaBancaria").Value)
            txtReferencia.Text = Trim(RsGral.Fields("Referencia").Value)
            txtImporte.Text = VB6.Format(RsGral.Fields("importe").Value, "###,##0.00")
            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMoneda.Text = C_DESCPESOS
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMoneda.Text = C_DESCDOLARES
            End If
            txtConcepto.Text = RsGral.Fields("Concepto").Value
            txtFolioMovimiento.ReadOnly = True
        Else
            MsgBox("Folio de Cancelación no Existe, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtFolioCancelacion.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatosMovimientos()
        On Error GoTo Merr
        If mblnNuevo Then
            If Trim(txtFolioMovimiento.Text) = "" Then
                Exit Sub
            End If
            gStrSql = "SELECT MB.FolioMovto,MB.FechaMovto,MB.Movimiento,MB.Naturaleza,MB.Moneda,MB.TipoCambio,MB.Importe," & "MB.Referencia,MB.CtaBancaria,MB.FolioProgramacion,MB.PartidaPP,CB.DescBanco,MB.FormaPago,MB.TipoPago,MB.CodBanco,MB.Beneficiario," & "MB.Concepto,MB.FechaDocto,MB.NoDocto,MB.FolioRetiro,ISNULL(AUX.Referencia,'') AS Referen,MB.Modulo " & "FROM MovimientosBancarios MB LEFT OUTER JOIN (SELECT * FROM MOVIMIENTOSBANCARIOS WHERE REFERENCIA = '" & txtFolioMovimiento.Text & "' AND MOVIMIENTO = '" & C_MOVCANCELACION & "') AUX " & "ON MB.FolioMovto = AUX.Referencia,CatBancos CB WHERE MB.FolioMovto = '" & txtFolioMovimiento.Text & "' AND " & "MB.Movimiento <> '" & C_MOVCANCELACION & "' AND CB.CodBanco = MB.CodBanco"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                If txtFolioMovimiento.Text <> RsGral.Fields("Referen").Value Then
                    Dim fechaMovimiento As String = AgregarHoraAFecha(RsGral.Fields("FechaMovto").Value.ToString())
                    lblFechaMovimiento.Text = fechaMovimiento
                    txtFolioMovimiento.Text = RsGral.Fields("FolioMovto").Value
                    If RsGral.Fields("Movimiento").Value = C_MOVPAGO Then
                        txtTipoMovimiento.Text = "PAGO"
                        mstrMovimiento = C_MOVPAGO
                        If RsGral.Fields("FolioProgramacion").Value <> "" Then
                            mstrFolioProgramacion = Trim(RsGral.Fields("FolioProgramacion").Value)
                        End If
                    ElseIf RsGral.Fields("Movimiento").Value = C_MOVDEPOSITO Then
                        txtTipoMovimiento.Text = "DEPOSITO"
                        mstrMovimiento = C_MOVDEPOSITO
                    ElseIf RsGral.Fields("Movimiento").Value = C_MOVTRASPASO Then
                        txtTipoMovimiento.Text = "TRASPASO BANCARIO"
                        mstrMovimiento = C_MOVTRASPASO
                    ElseIf RsGral.Fields("Movimiento").Value = C_MOVCARGOS Then
                        txtTipoMovimiento.Text = "CARGOS DIVERSOS"
                        mstrMovimiento = C_MOVCARGOS
                    ElseIf RsGral.Fields("Movimiento").Value = C_MOVANTICIPOS Then
                        txtTipoMovimiento.Text = "ANTICIPO A PROV./ACREED."
                        mstrMovimiento = C_MOVANTICIPOS
                    ElseIf RsGral.Fields("Movimiento").Value = C_OTROSINGRESOS Then
                        txtTipoMovimiento.Text = "OTROS INGRESOS"
                        mstrMovimiento = C_OTROSINGRESOS
                    End If
                    txtBanco.Text = Trim(RsGral.Fields("DescBanco").Value)
                    txtCuentaBancaria.Text = Trim(RsGral.Fields("CtaBancaria").Value)
                    txtReferencia.Text = Trim(RsGral.Fields("Referencia").Value)
                    txtImporte.Text = VB6.Format(RsGral.Fields("importe").Value, "###,##0.00")
                    If RsGral.Fields("Moneda").Value = C_PESO Then
                        lblMoneda.Text = C_DESCPESOS
                    ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                        lblMoneda.Text = C_DESCDOLARES
                    End If
                    mstrNaturaleza = Trim(RsGral.Fields("Naturaleza").Value)
                    mcurTipoCambio = RsGral.Fields("TipoCambio").Value
                    mstrFormaPago = RsGral.Fields("FormaPago").Value
                    mstrTipoPago = RsGral.Fields("TipoPago").Value
                    mintCodBanco = RsGral.Fields("CodBanco").Value
                    mstrBeneficiario = RsGral.Fields("Beneficiario").Value
                    mdtmFechaDocto = RsGral.Fields("FechaDocto").Value
                    mstrNoDocto = RsGral.Fields("NoDocto").Value
                    mstrFolioRetiro = RsGral.Fields("FolioRetiro").Value
                    intNumPartida = RsGral.Fields("PartidaPP").Value
                    Modulo = RsGral.Fields("Modulo").Value
                Else
                    MsgBox("Este Folio de Movimiento ya fue Cancelado, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    txtFolioMovimiento.Focus()
                End If
            Else
                MsgBox("Folio de Movimiento No Existe, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtFolioMovimiento.Focus()
            End If
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Nuevo()

        'If (bandera = True) Then
        '    Exit Sub
        'End If

        If Not mblnNuevo Then
            txtFolioMovimiento.Text = ""
        End If
        txtTipoMovimiento.Text = ""
        txtFolioMovimiento.ReadOnly = False
        txtBanco.Text = ""
        txtCuentaBancaria.Text = ""
        txtReferencia.Text = ""
        txtConcepto.Text = ""
        txtImporte.Text = "0.00"
        lblMoneda.Text = ""
        'lblFechaCancelacion.Text = VB6.Format(Now, "dd/mmm/yyyy")
        'lblFechaMovimiento.Text = VB6.Format(Now, "dd/mmm/yyyy")
        Dim fechaCancelacion As String = AgregarHoraAFecha(Now)
        Dim fechaMovimiento As String = AgregarHoraAFecha(Now)
        lblFechaCancelacion.Text = fechaCancelacion
        lblFechaMovimiento.Text = fechaMovimiento
        InicializaVariables()
    End Sub

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnSalir = False
        mstrMovimiento = ""
        mstrFolioProgramacion = ""
        mstrNaturaleza = ""
        mcurTipoCambio = 0
        mstrFormaPago = ""
        mstrTipoPago = ""
        mintCodBanco = 0
        mstrBeneficiario = ""
        mstrNoDocto = ""
        mdtmFechaDocto = Now
        mstrFolioRetiro = ""
    End Sub

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Not BuscaUltimoCierre(CDate(lblFechaCancelacion.Text)) Then
            Exit Function
        End If
        If mstrMovimiento = C_MOVANTICIPOS Then
            If EstaAplicado(txtFolioMovimiento.Text) Then
                Exit Function
            End If
        End If
        If mstrMovimiento = C_MOVTRASPASO Then
            If EstaConciliado(txtFolioMovimiento.Text) Or EstaConciliado(txtReferencia.Text) Then
                MsgBox("¡¡¡ATENCION!!! Puede que Uno o los Dos Movimientos que Involucran" & Chr(13) & "a Este Traspaso ya Esten Conciliados, Favor de Verificar ....", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Function
            End If
        Else
            If EstaConciliado(txtFolioMovimiento.Text) Then
                MsgBox("¡¡¡ATENCIÓN!!! Este Movimiento ya fue Conciliado, No se Puede Cancelar.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Function
            End If
        End If
        If Len(Trim(txtFolioMovimiento.Text)) = 0 Or (Len(Trim(txtTipoMovimiento.Text)) = 0 And Len(Trim(txtBanco.Text)) = 0 And Len(Trim(txtCuentaBancaria.Text)) = 0 And CDbl(Numerico(txtImporte.Text)) = 0) Then
            MsgBox("Proporcione un Folio de Movimiento", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtFolioMovimiento.Focus()
            Exit Function
        End If
        If Len(Trim(txtConcepto.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Concepto", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtConcepto.Focus()
            Exit Function
        End If
        '    If (CDate(lblFechaMovimiento) > CDate(lblFechaCancelacion)) Or (CDate(lblFechaCancelacion) > Date) Then
        '        MsgBox "La Fecha de Cancelación es Incorrecta, Favor de Verificar.", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '        Exit Function
        '    End If
        ValidaDatos = True
    End Function

    Private Sub frmBancosProcesoDiarioCancelaciondeMovimientosBanc_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoDiarioCancelaciondeMovimientosBanc_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoDiarioCancelaciondeMovimientosBanc_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtFolioCancelacion" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoDiarioCancelaciondeMovimientosBanc_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioCancelaciondeMovimientosBanc_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        'FrmConsultas.InitializeComponent()
        ModEstandar.CentrarForma(Me)
        bandera = True
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        InicializaVariables()
        Nuevo()
        BuscaEjercicio(lblFechaCancelacion.Text)
    End Sub

    Private Sub frmBancosProcesoDiarioCancelaciondeMovimientosBanc_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmBancosProcesoDiarioCancelaciondeMovimientosBanc_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        IsNothing(Me)
    End Sub

    Private Sub txtBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBanco.Enter
        Pon_Tool()
    End Sub

    Private Sub txtConcepto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.Enter
        Pon_Tool()
    End Sub

    Private Sub txtConcepto_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConcepto.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "!""#$%&/()=?'¡¿*,;.:<>@+-_")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCuentaBancaria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCuentaBancaria.Enter
        Pon_Tool()
    End Sub

    Private Sub txtFolioCancelacion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioCancelacion.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtFolioCancelacion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioCancelacion.Enter
        strControlActual = UCase("txtFolioCancelacion")
        SelTextoTxt(txtFolioCancelacion)
        Pon_Tool()
    End Sub

    Private Sub txtFolioCancelacion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioCancelacion.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, C_MOVCANCELACION)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioCancelacion_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioCancelacion.Leave

        If Me.ActiveControl.Name = "btnBuscar" Then
            Exit Sub
        End If

        If Trim(txtFolioCancelacion.Text) = "" Then
            txtFolioCancelacion.Text = C_TIPOMOVCANCELACION & VB6.Format(Year(CDate(lblFechaCancelacion.Text)), "0000") & VB6.Format(Month(CDate(lblFechaCancelacion.Text)), "00") & VB6.Format(VB.Day(CDate(lblFechaCancelacion.Text)), "00") & "0000"
            Exit Sub
        End If

        If mblnCambiosEnCodigo = True And txtFolioCancelacion.Text <> "" And VB.Right(txtFolioCancelacion.Text, 4) <> "0000" Then
            LlenaDatos()
        End If

    End Sub
    Private Sub txtFolioMovimiento_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioMovimiento.TextChanged
        If mblnNuevo Then
            Nuevo()
        End If
    End Sub

    Private Sub txtFolioMovimiento_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioMovimiento.Enter
        strControlActual = UCase("txtFolioMovimiento")
        SelTextoTxt(txtFolioMovimiento)
        Pon_Tool()
    End Sub

    Private Sub txtFolioMovimiento_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioMovimiento.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, C_TIPOMOVINGRESO & C_TIPOMOVEGRESO)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioMovimiento_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioMovimiento.Leave

        If Me.ActiveControl.Name = "btnBuscar" Then
            Exit Sub
        End If

        If Trim(txtFolioMovimiento.Text) <> "" Then
            LlenaDatosMovimientos()
        End If
    End Sub

    Private Sub txtImporte_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Enter
        Pon_Tool()
    End Sub

    Private Sub txtReferencia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferencia.Enter
        Pon_Tool()
    End Sub

    Private Sub txtTipoMovimiento_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoMovimiento.Enter
        Pon_Tool()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click

    End Sub
End Class