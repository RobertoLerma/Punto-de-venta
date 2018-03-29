Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms
Imports System.Data
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosProcesoDiarioTraspasosBancarios
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REGISTRO DE TRASPASOS BANCARIOS                                                              *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      VIERNES 25 DE JULIO DE 2003                                                                  *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtBancoDestino As System.Windows.Forms.TextBox
    Public WithEvents txtBancoOrigen As System.Windows.Forms.TextBox
    Public WithEvents txtImporteCtaDestino As System.Windows.Forms.TextBox
    Public WithEvents txtCuentaBancariaDestino As System.Windows.Forms.TextBox
    Public WithEvents txtCuentaBancariaOrigen As System.Windows.Forms.TextBox
    Public WithEvents txtBeneficiario As System.Windows.Forms.TextBox
    Public WithEvents txtConcepto As System.Windows.Forms.TextBox
    Public WithEvents txtNumeroCheque As System.Windows.Forms.TextBox
    Public WithEvents dtpFechaCheque As System.Windows.Forms.DateTimePicker
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents txtImporteCtaOrigen As System.Windows.Forms.TextBox
    Public WithEvents _optFormaPago_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optFormaPago_0 As System.Windows.Forms.RadioButton
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents lblMovIngreso As System.Windows.Forms.Label
    Public WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents lblCancelacion As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents lblMonedaDestino As System.Windows.Forms.Label
    Public WithEvents lblMonedaOrigen As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents txtTipoCambio As System.Windows.Forms.TextBox
    Public WithEvents txtFolioEgreso As System.Windows.Forms.TextBox
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents optFormaPago As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtBancoDestino = New System.Windows.Forms.TextBox()
        Me.txtBancoOrigen = New System.Windows.Forms.TextBox()
        Me.txtImporteCtaDestino = New System.Windows.Forms.TextBox()
        Me.txtCuentaBancariaDestino = New System.Windows.Forms.TextBox()
        Me.txtCuentaBancariaOrigen = New System.Windows.Forms.TextBox()
        Me.txtBeneficiario = New System.Windows.Forms.TextBox()
        Me.txtConcepto = New System.Windows.Forms.TextBox()
        Me.txtNumeroCheque = New System.Windows.Forms.TextBox()
        Me.txtImporteCtaOrigen = New System.Windows.Forms.TextBox()
        Me._optFormaPago_1 = New System.Windows.Forms.RadioButton()
        Me._optFormaPago_0 = New System.Windows.Forms.RadioButton()
        Me.txtTipoCambio = New System.Windows.Forms.TextBox()
        Me.txtFolioEgreso = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaCheque = New System.Windows.Forms.DateTimePicker()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.lblMovIngreso = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.lblCancelacion = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblMonedaDestino = New System.Windows.Forms.Label()
        Me.lblMonedaOrigen = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.optFormaPago = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frame4.SuspendLayout()
        CType(Me.optFormaPago, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtBancoDestino
        '
        Me.txtBancoDestino.AcceptsReturn = True
        Me.txtBancoDestino.BackColor = System.Drawing.SystemColors.Window
        Me.txtBancoDestino.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBancoDestino.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBancoDestino.Location = New System.Drawing.Point(99, 85)
        Me.txtBancoDestino.Margin = New System.Windows.Forms.Padding(2)
        Me.txtBancoDestino.MaxLength = 3
        Me.txtBancoDestino.Name = "txtBancoDestino"
        Me.txtBancoDestino.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBancoDestino.Size = New System.Drawing.Size(52, 20)
        Me.txtBancoDestino.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtBancoDestino, "Clave del Banco Destino.")
        '
        'txtBancoOrigen
        '
        Me.txtBancoOrigen.AcceptsReturn = True
        Me.txtBancoOrigen.BackColor = System.Drawing.SystemColors.Window
        Me.txtBancoOrigen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBancoOrigen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBancoOrigen.Location = New System.Drawing.Point(99, 59)
        Me.txtBancoOrigen.Margin = New System.Windows.Forms.Padding(2)
        Me.txtBancoOrigen.MaxLength = 3
        Me.txtBancoOrigen.Name = "txtBancoOrigen"
        Me.txtBancoOrigen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBancoOrigen.Size = New System.Drawing.Size(52, 20)
        Me.txtBancoOrigen.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtBancoOrigen, "Clave del Banco Origen.")
        '
        'txtImporteCtaDestino
        '
        Me.txtImporteCtaDestino.AcceptsReturn = True
        Me.txtImporteCtaDestino.BackColor = System.Drawing.SystemColors.Window
        Me.txtImporteCtaDestino.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImporteCtaDestino.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImporteCtaDestino.Location = New System.Drawing.Point(406, 164)
        Me.txtImporteCtaDestino.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImporteCtaDestino.MaxLength = 18
        Me.txtImporteCtaDestino.Name = "txtImporteCtaDestino"
        Me.txtImporteCtaDestino.ReadOnly = True
        Me.txtImporteCtaDestino.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImporteCtaDestino.Size = New System.Drawing.Size(76, 20)
        Me.txtImporteCtaDestino.TabIndex = 12
        Me.txtImporteCtaDestino.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtImporteCtaDestino, "Importe que se va a Transferir a la Cuenta Destino.")
        '
        'txtCuentaBancariaDestino
        '
        Me.txtCuentaBancariaDestino.AcceptsReturn = True
        Me.txtCuentaBancariaDestino.BackColor = System.Drawing.SystemColors.Window
        Me.txtCuentaBancariaDestino.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCuentaBancariaDestino.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCuentaBancariaDestino.Location = New System.Drawing.Point(295, 81)
        Me.txtCuentaBancariaDestino.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCuentaBancariaDestino.MaxLength = 16
        Me.txtCuentaBancariaDestino.Name = "txtCuentaBancariaDestino"
        Me.txtCuentaBancariaDestino.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCuentaBancariaDestino.Size = New System.Drawing.Size(95, 20)
        Me.txtCuentaBancariaDestino.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtCuentaBancariaDestino, "Cuenta Bancaria Destino.")
        '
        'txtCuentaBancariaOrigen
        '
        Me.txtCuentaBancariaOrigen.AcceptsReturn = True
        Me.txtCuentaBancariaOrigen.BackColor = System.Drawing.SystemColors.Window
        Me.txtCuentaBancariaOrigen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCuentaBancariaOrigen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCuentaBancariaOrigen.Location = New System.Drawing.Point(295, 56)
        Me.txtCuentaBancariaOrigen.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCuentaBancariaOrigen.MaxLength = 16
        Me.txtCuentaBancariaOrigen.Name = "txtCuentaBancariaOrigen"
        Me.txtCuentaBancariaOrigen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCuentaBancariaOrigen.Size = New System.Drawing.Size(95, 20)
        Me.txtCuentaBancariaOrigen.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtCuentaBancariaOrigen, "Cuenta Bancaria Origen.")
        '
        'txtBeneficiario
        '
        Me.txtBeneficiario.AcceptsReturn = True
        Me.txtBeneficiario.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtBeneficiario.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBeneficiario.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBeneficiario.Location = New System.Drawing.Point(99, 110)
        Me.txtBeneficiario.Margin = New System.Windows.Forms.Padding(2)
        Me.txtBeneficiario.MaxLength = 50
        Me.txtBeneficiario.Name = "txtBeneficiario"
        Me.txtBeneficiario.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBeneficiario.Size = New System.Drawing.Size(402, 20)
        Me.txtBeneficiario.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.txtBeneficiario, "Persona que Recibira el Traspaso.")
        '
        'txtConcepto
        '
        Me.txtConcepto.AcceptsReturn = True
        Me.txtConcepto.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtConcepto.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConcepto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtConcepto.Location = New System.Drawing.Point(99, 134)
        Me.txtConcepto.Margin = New System.Windows.Forms.Padding(2)
        Me.txtConcepto.MaxLength = 100
        Me.txtConcepto.Name = "txtConcepto"
        Me.txtConcepto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConcepto.Size = New System.Drawing.Size(402, 20)
        Me.txtConcepto.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txtConcepto, "Concepto del Traspaso.")
        '
        'txtNumeroCheque
        '
        Me.txtNumeroCheque.AcceptsReturn = True
        Me.txtNumeroCheque.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumeroCheque.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumeroCheque.Enabled = False
        Me.txtNumeroCheque.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumeroCheque.Location = New System.Drawing.Point(295, 17)
        Me.txtNumeroCheque.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNumeroCheque.MaxLength = 10
        Me.txtNumeroCheque.Name = "txtNumeroCheque"
        Me.txtNumeroCheque.ReadOnly = True
        Me.txtNumeroCheque.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumeroCheque.Size = New System.Drawing.Size(72, 20)
        Me.txtNumeroCheque.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.txtNumeroCheque, "Numero de Cheque.")
        '
        'txtImporteCtaOrigen
        '
        Me.txtImporteCtaOrigen.AcceptsReturn = True
        Me.txtImporteCtaOrigen.BackColor = System.Drawing.SystemColors.Window
        Me.txtImporteCtaOrigen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImporteCtaOrigen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImporteCtaOrigen.Location = New System.Drawing.Point(212, 165)
        Me.txtImporteCtaOrigen.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImporteCtaOrigen.MaxLength = 18
        Me.txtImporteCtaOrigen.Name = "txtImporteCtaOrigen"
        Me.txtImporteCtaOrigen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImporteCtaOrigen.Size = New System.Drawing.Size(76, 20)
        Me.txtImporteCtaOrigen.TabIndex = 11
        Me.txtImporteCtaOrigen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtImporteCtaOrigen, "Importe de la Cuenta Origen.")
        '
        '_optFormaPago_1
        '
        Me._optFormaPago_1.BackColor = System.Drawing.SystemColors.Control
        Me._optFormaPago_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optFormaPago_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optFormaPago_1.Location = New System.Drawing.Point(83, 13)
        Me._optFormaPago_1.Margin = New System.Windows.Forms.Padding(2)
        Me._optFormaPago_1.Name = "_optFormaPago_1"
        Me._optFormaPago_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optFormaPago_1.Size = New System.Drawing.Size(87, 21)
        Me._optFormaPago_1.TabIndex = 4
        Me._optFormaPago_1.TabStop = True
        Me._optFormaPago_1.Text = "Electrónico"
        Me.ToolTip1.SetToolTip(Me._optFormaPago_1, "Pago Electrónico.")
        Me._optFormaPago_1.UseVisualStyleBackColor = False
        '
        '_optFormaPago_0
        '
        Me._optFormaPago_0.BackColor = System.Drawing.SystemColors.Control
        Me._optFormaPago_0.Checked = True
        Me._optFormaPago_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optFormaPago_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optFormaPago_0.Location = New System.Drawing.Point(12, 13)
        Me._optFormaPago_0.Margin = New System.Windows.Forms.Padding(2)
        Me._optFormaPago_0.Name = "_optFormaPago_0"
        Me._optFormaPago_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optFormaPago_0.Size = New System.Drawing.Size(68, 21)
        Me._optFormaPago_0.TabIndex = 3
        Me._optFormaPago_0.TabStop = True
        Me._optFormaPago_0.Text = "Cheque"
        Me.ToolTip1.SetToolTip(Me._optFormaPago_0, "Pago con Cheque.")
        Me._optFormaPago_0.UseVisualStyleBackColor = False
        '
        'txtTipoCambio
        '
        Me.txtTipoCambio.AcceptsReturn = True
        Me.txtTipoCambio.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipoCambio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambio.Location = New System.Drawing.Point(312, 14)
        Me.txtTipoCambio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTipoCambio.MaxLength = 6
        Me.txtTipoCambio.Name = "txtTipoCambio"
        Me.txtTipoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambio.Size = New System.Drawing.Size(72, 20)
        Me.txtTipoCambio.TabIndex = 1
        Me.txtTipoCambio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambio, "Tipo de Cambio.")
        '
        'txtFolioEgreso
        '
        Me.txtFolioEgreso.AcceptsReturn = True
        Me.txtFolioEgreso.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioEgreso.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioEgreso.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioEgreso.Location = New System.Drawing.Point(99, 14)
        Me.txtFolioEgreso.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioEgreso.MaxLength = 13
        Me.txtFolioEgreso.Name = "txtFolioEgreso"
        Me.txtFolioEgreso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioEgreso.Size = New System.Drawing.Size(115, 20)
        Me.txtFolioEgreso.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtFolioEgreso, "Folio del Egreso.")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtBancoDestino)
        Me.Frame1.Controls.Add(Me.txtBancoOrigen)
        Me.Frame1.Controls.Add(Me.txtImporteCtaDestino)
        Me.Frame1.Controls.Add(Me.txtCuentaBancariaDestino)
        Me.Frame1.Controls.Add(Me.txtCuentaBancariaOrigen)
        Me.Frame1.Controls.Add(Me.txtBeneficiario)
        Me.Frame1.Controls.Add(Me.txtConcepto)
        Me.Frame1.Controls.Add(Me.Frame3)
        Me.Frame1.Controls.Add(Me.txtImporteCtaOrigen)
        Me.Frame1.Controls.Add(Me.Frame5)
        Me.Frame1.Controls.Add(Me.lblMovIngreso)
        Me.Frame1.Controls.Add(Me.Label13)
        Me.Frame1.Controls.Add(Me.Label12)
        Me.Frame1.Controls.Add(Me.lblCancelacion)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.lblMonedaDestino)
        Me.Frame1.Controls.Add(Me.lblMonedaOrigen)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label6)
        Me.Frame1.Controls.Add(Me.Label7)
        Me.Frame1.Controls.Add(Me.Label11)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 49)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(576, 313)
        Me.Frame1.TabIndex = 18
        Me.Frame1.TabStop = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.txtNumeroCheque)
        Me.Frame3.Controls.Add(Me.dtpFechaCheque)
        Me.Frame3.Controls.Add(Me.Label9)
        Me.Frame3.Controls.Add(Me.Label10)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(99, 211)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(402, 49)
        Me.Frame3.TabIndex = 20
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Cheque"
        '
        'dtpFechaCheque
        '
        Me.dtpFechaCheque.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaCheque.Location = New System.Drawing.Point(92, 17)
        Me.dtpFechaCheque.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaCheque.Name = "dtpFechaCheque"
        Me.dtpFechaCheque.Size = New System.Drawing.Size(87, 20)
        Me.dtpFechaCheque.TabIndex = 13
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(41, 20)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(51, 13)
        Me.Label9.TabIndex = 22
        Me.Label9.Text = "Fecha :"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(244, 21)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(59, 13)
        Me.Label10.TabIndex = 21
        Me.Label10.Text = "Numero :"
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me._optFormaPago_1)
        Me.Frame5.Controls.Add(Me._optFormaPago_0)
        Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame5.Location = New System.Drawing.Point(12, 11)
        Me.Frame5.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(179, 38)
        Me.Frame5.TabIndex = 19
        Me.Frame5.TabStop = False
        '
        'lblMovIngreso
        '
        Me.lblMovIngreso.BackColor = System.Drawing.SystemColors.Control
        Me.lblMovIngreso.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMovIngreso.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblMovIngreso.Location = New System.Drawing.Point(157, 271)
        Me.lblMovIngreso.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMovIngreso.Name = "lblMovIngreso"
        Me.lblMovIngreso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMovIngreso.Size = New System.Drawing.Size(271, 23)
        Me.lblMovIngreso.TabIndex = 35
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(12, 89)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(97, 17)
        Me.Label13.TabIndex = 34
        Me.Label13.Text = "Banco Destino :"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(12, 67)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(97, 17)
        Me.Label12.TabIndex = 33
        Me.Label12.Text = "Banco Origen :"
        '
        'lblCancelacion
        '
        Me.lblCancelacion.BackColor = System.Drawing.SystemColors.Control
        Me.lblCancelacion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCancelacion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblCancelacion.Location = New System.Drawing.Point(211, 25)
        Me.lblCancelacion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCancelacion.Name = "lblCancelacion"
        Me.lblCancelacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCancelacion.Size = New System.Drawing.Size(271, 20)
        Me.lblCancelacion.TabIndex = 32
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(292, 167)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(122, 15)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Importe Cta Destino :"
        '
        'lblMonedaDestino
        '
        Me.lblMonedaDestino.BackColor = System.Drawing.SystemColors.Control
        Me.lblMonedaDestino.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMonedaDestino.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMonedaDestino.Location = New System.Drawing.Point(403, 81)
        Me.lblMonedaDestino.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMonedaDestino.Name = "lblMonedaDestino"
        Me.lblMonedaDestino.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMonedaDestino.Size = New System.Drawing.Size(95, 17)
        Me.lblMonedaDestino.TabIndex = 29
        Me.lblMonedaDestino.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblMonedaOrigen
        '
        Me.lblMonedaOrigen.BackColor = System.Drawing.SystemColors.Control
        Me.lblMonedaOrigen.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMonedaOrigen.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMonedaOrigen.Location = New System.Drawing.Point(403, 56)
        Me.lblMonedaOrigen.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMonedaOrigen.Name = "lblMonedaOrigen"
        Me.lblMonedaOrigen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMonedaOrigen.Size = New System.Drawing.Size(95, 17)
        Me.lblMonedaOrigen.TabIndex = 28
        Me.lblMonedaOrigen.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(157, 62)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(134, 17)
        Me.Label4.TabIndex = 27
        Me.Label4.Text = "Cuenta Bancaria Origen :"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(157, 84)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(134, 17)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "Cuenta Bancaria Destino :"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(12, 110)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(80, 17)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Beneficiario :"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(12, 132)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(67, 17)
        Me.Label7.TabIndex = 24
        Me.Label7.Text = "Concepto :"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(107, 167)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(107, 18)
        Me.Label11.TabIndex = 23
        Me.Label11.Text = "Importe Cta Origen :"
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.dtpFecha)
        Me.Frame4.Controls.Add(Me.txtTipoCambio)
        Me.Frame4.Controls.Add(Me.txtFolioEgreso)
        Me.Frame4.Controls.Add(Me.Label8)
        Me.Frame4.Controls.Add(Me.Label3)
        Me.Frame4.Controls.Add(Me.Label1)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(8, 3)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(576, 40)
        Me.Frame4.TabIndex = 15
        Me.Frame4.TabStop = False
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(457, 14)
        Me.dtpFecha.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(97, 20)
        Me.dtpFecha.TabIndex = 2
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(223, 18)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(92, 11)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Tipo de Cambio :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(413, 17)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(51, 12)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Fecha :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(9, 17)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Folio de Egreso :"
        '
        'optFormaPago
        '
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackColor = System.Drawing.SystemColors.Control
        Me.btnLimpiar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLimpiar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLimpiar.Location = New System.Drawing.Point(125, 384)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLimpiar.Size = New System.Drawing.Size(109, 36)
        Me.btnLimpiar.TabIndex = 44
        Me.btnLimpiar.Text = "&Nuevo"
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(10, 384)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 43
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoDiarioTraspasosBancarios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(594, 432)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(184, 203)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoDiarioTraspasosBancarios"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Traspasos Bancarios"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame5.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        CType(Me.optFormaPago, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    'Variables
    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Dim mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Dim mblnSalir As Boolean 'Para Salir Con el Esc
    Dim LetraFolio As String
    Dim ConsecutivoCheque As Integer
    Dim FueraChange As Boolean
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim strControlActual As String 'Nombre del control actual
        Dim I As Integer
        strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control
        Select Case strControlActual
            Case "TXTFOLIOEGRESO"
                strCaptionForm = "Consulta de Registro de Pagos"
                gStrSql = "SELECT FolioMovto AS FOLIO,Concepto AS CONCEPTO,Beneficiario AS BENEFICIARIO," & "FechaMovto AS FECHA,Importe AS IMPORTE FROM MovimientosBancarios " & "WHERE FolioMovto LIKE '" & txtFolioEgreso.Text & "%' AND Movimiento = '" & C_MOVTRASPASO & "' AND TipoMovto = '" & C_TIPOMOVEGRESO & "' ORDER BY FechaMovto DESC ,FolioMovto DESC"
            Case "TXTCUENTABANCARIAORIGEN"
                strCaptionForm = "Consulta de Cuentas Bancarias"
                If CDbl(Numerico(txtBancoOrigen.Text)) <> 0 Then
                    gStrSql = "SELECT CB.CodBanco AS CODIGO, CB.DescBanco AS 'DESCRIPCION DEL BANCO', CC.CtaBancaria AS 'CUENTA BANCARIA' " & "FROM CatBancos CB, CatCuentasBancarias CC " & "WHERE CC.CtaBancaria LIKE '" & Trim(txtCuentaBancariaOrigen.Text) & "%' AND CB.CodBanco = " & txtBancoOrigen.Text & " AND " & "CB.CodBanco = CC.CodBanco AND CB.ControlInterno = 0 AND CB.Sucursal = 0 ORDER BY CB.DescBanco"
                Else
                    MsgBox("No ha Capturado el Codigo del Banco, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    txtBancoOrigen.Focus()
                    Exit Sub
                End If
            Case "TXTCUENTABANCARIADESTINO"
                strCaptionForm = "Consulta de Cuentas Bancarias"
                If CDbl(Numerico(txtBancoDestino.Text)) <> 0 Then
                    gStrSql = "SELECT CB.CodBanco AS CODIGO, CB.DescBanco AS 'DESCRIPCION DEL BANCO', CC.CtaBancaria AS 'CUENTA BANCARIA' " & "FROM CatBancos CB, CatCuentasBancarias CC " & "WHERE CC.CtaBancaria LIKE '" & Trim(txtCuentaBancariaDestino.Text) & "%' AND CB.CodBanco = " & txtBancoDestino.Text & " AND " & "CB.CodBanco = CC.CodBanco AND CB.ControlInterno = 0 AND CB.Sucursal = 0 ORDER BY CB.DescBanco"
                Else
                    MsgBox("No ha Capturado el Codigo del Banco, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    txtBancoDestino.Focus()
                    Exit Sub
                End If
            Case "TXTBANCOORIGEN"
                strCaptionForm = "Consulta de Bancos"
                gStrSql = "SELECT RIGHT('000'+LTRIM(Codbanco),3) AS CODIGO,Descbanco AS DESCRIPCION FROM Catbancos WHERE ControlInterno = 0  ORDER BY CodBanco"
            Case "TXTBANCODESTINO"
                strCaptionForm = "Consulta de Bancos"
                gStrSql = "SELECT RIGHT('000'+LTRIM(Codbanco),3) AS CODIGO,Descbanco AS DESCRIPCION FROM Catbancos WHERE ControlInterno = 0 ORDER BY CodBanco"
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
        'Carga el formulario de consulta

        'si jala todo      Load(FrmConsultas)
        '      With FrmConsultas.Flexdet
        '	Select Case strControlActual
        '		Case "TXTFOLIOEGRESO"
        '			Call ConfiguraConsultas(FrmConsultas, 12700, RsGral, strTag, strCaptionForm)
        '			.set_ColWidth(0,  , 1400) 'Columna del Folio
        '			.set_ColWidth(1,  , 4000) 'Columna del Concepto del Movimiento
        '			.set_ColWidth(2,  , 3900) 'Columna del Beneficiario del Movimiento
        '			.set_ColWidth(3,  , 1200) 'Columna de la Fecha del Movimiento
        '			.set_ColWidth(4,  , 1800) 'Columna del Importe del Movimiento
        '			.set_ColAlignment(3, 4)
        '			For I = 1 To FrmConsultas.Flexdet.Rows - 1
        '				FrmConsultas.Flexdet.set_TextMatrix(I, 3, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 3), "dd/MMM/yyyy"))
        '				FrmConsultas.Flexdet.set_TextMatrix(I, 4, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 4), "###,##0.00"))
        '			Next 
        '			FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
        '			FrmConsultas.Left = VB6.TwipsToPixelsX(1150)
        '		Case "TXTCUENTABANCARIAORIGEN"
        '			Call ConfiguraConsultas(FrmConsultas, 9000, RsGral, strTag, strCaptionForm)
        '			.set_ColWidth(0,  , 1000)
        '			.set_ColWidth(1,  , 4500)
        '			.set_ColWidth(2,  , 3000)
        '			.set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
        '			For I = 1 To FrmConsultas.Flexdet.Rows - 1
        '				FrmConsultas.Flexdet.set_TextMatrix(I, 0, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 0), "000"))
        '				'FrmConsultas.Flexdet.TextMatrix(I, 2) = Trim(Str(FrmConsultas.Flexdet.TextMatrix(I, 2)))
        '			Next 
        '			FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
        '			FrmConsultas.Left = VB6.TwipsToPixelsX(2970)
        '		Case "TXTCUENTABANCARIADESTINO"
        '			Call ConfiguraConsultas(FrmConsultas, 9000, RsGral, strTag, strCaptionForm)
        '			.set_ColWidth(0,  , 1000)
        '			.set_ColWidth(1,  , 4500)
        '			.set_ColWidth(2,  , 3000)
        '			.set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
        '			For I = 1 To FrmConsultas.Flexdet.Rows - 1
        '				FrmConsultas.Flexdet.set_TextMatrix(I, 0, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 0), "000"))
        '				'FrmConsultas.Flexdet.TextMatrix(I, 2) = Trim(Str(FrmConsultas.Flexdet.TextMatrix(I, 2)))
        '			Next 
        '			FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
        '			FrmConsultas.Left = VB6.TwipsToPixelsX(2970)
        '		Case "TXTBANCODESTINO", "TXTBANCOORIGEN"
        '			Call ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)
        '			.set_ColWidth(0,  , 900) 'Columna del Código
        '			.set_ColWidth(1,  , 4800) 'Columna de la Descripción
        '	End Select
        'End With
        'FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub BuscaBancoOrigen()
        On Error GoTo Merr
        gStrSql = "SELECT CodBanco FROM CatBancos WHERE CodBanco = " & txtBancoOrigen.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtBancoOrigen.Text = VB6.Format(txtBancoOrigen.Text, "000")
        Else
            MsgBox("Este Codigo de Banco no Existe, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtBancoOrigen.Text = "000"
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub BuscaBancoDestino()
        On Error GoTo Merr
        gStrSql = "SELECT CodBanco FROM CatBancos WHERE CodBanco = " & txtBancoDestino.Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtBancoDestino.Text = VB6.Format(txtBancoDestino.Text, "000")
        Else
            MsgBox("Este Codigo de Banco no Existe, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtBancoDestino.Text = "000"
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub CalculaImporteDestino()
        If lblMonedaOrigen.Text = lblMonedaDestino.Text Then
            txtImporteCtaDestino.Text = VB6.Format(txtImporteCtaOrigen.Text, "###,##0.00")
        ElseIf lblMonedaOrigen.Text = C_DESCPESOS And lblMonedaDestino.Text = C_DESCDOLARES Then
            txtImporteCtaDestino.Text = VB6.Format(CDec(Numerico(VB6.Format(txtImporteCtaOrigen.Text, "#####0.00"))) / CDec(Numerico(VB6.Format(txtTipoCambio.Text, "#####0.00"))), "###,##0.00")
        ElseIf lblMonedaOrigen.Text = C_DESCDOLARES And lblMonedaDestino.Text = C_DESCPESOS Then
            txtImporteCtaDestino.Text = VB6.Format(CDec(Numerico(VB6.Format(txtImporteCtaOrigen.Text, "#####0.00"))) * CDec(Numerico(VB6.Format(txtTipoCambio.Text, "#####0.00"))), "###,##0.00")
        End If
    End Sub

    Sub ChecaCuentaDestino()
        On Error GoTo Merr
        If CDbl(Numerico(txtBancoDestino.Text)) = 0 Then
            MsgBox("No ha Capturado el Codigo del Banco, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtBancoDestino.Focus()
            Exit Sub
        End If
        gStrSql = "SELECT CB.CodBanco, CC.Moneda FROM CatBancos CB Inner Join CatCuentasBancarias CC On CB.CodBanco = CC.CodBanco " & "WHERE (CB.CodBanco = " & CInt(Numerico((txtBancoDestino.Text))) & " AND CB.ControlInterno = 0 AND CB.Sucursal = 0) And CC.CtaBancaria = '" & Trim(txtCuentaBancariaDestino.Text) & "' "

        '''gStrSql = "SELECT CB.CodBanco,CC.Moneda FROM CatBancos CB,CatCuentasBancarias CC WHERE CC.CtaBancaria = '" & txtCuentaBancariaDestino & "' " & _
        '"AND CB.ControlInterno = 0 AND CB.Sucursal = 0 AND CB.CodBanco = CC.CodBanco"

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If Trim(txtCuentaBancariaOrigen.Text) <> "" And CDbl(Numerico(txtBancoOrigen.Text)) <> 0 Then
                If Trim(txtCuentaBancariaOrigen.Text) = Trim(txtCuentaBancariaDestino.Text) And txtBancoOrigen.Text = txtBancoDestino.Text Then
                    MsgBox("No es Posible Hacer un Traspaso a una Misma Cuenta, Favor de Verificar.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    txtCuentaBancariaDestino.Text = ""
                    ModEstandar.RetrocederTab(Me)
                    Exit Sub
                End If
            End If
            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMonedaDestino.Text = C_DESCPESOS
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMonedaDestino.Text = C_DESCDOLARES
            End If
        Else
            MsgBox("Cuenta Bancaria no Existe para este Banco, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCuentaBancariaDestino.Text = ""
            txtCuentaBancariaDestino.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub ChecaCuentaOrigen()
        On Error GoTo Merr
        If CInt(Numerico(txtBancoOrigen.Text)) = 0 Then
            MsgBox("No ha Capturado el Codigo del Banco, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtBancoOrigen.Focus()
            Exit Sub
        End If
        gStrSql = "SELECT CB.CodBanco, CC.Moneda FROM CatBancos CB Inner Join CatCuentasBancarias CC On CB.CodBanco = CC.CodBanco " & "WHERE (CB.CodBanco = " & CInt(Numerico((txtBancoOrigen.Text))) & " AND CB.ControlInterno = 0 AND CB.Sucursal = 0) And CC.CtaBancaria = '" & Trim(txtCuentaBancariaOrigen.Text) & "' "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If Trim(txtCuentaBancariaDestino.Text) <> "" And CDbl(Numerico(txtBancoDestino.Text)) <> 0 Then
                If Trim(txtCuentaBancariaOrigen.Text) = Trim(txtCuentaBancariaDestino.Text) And txtBancoOrigen.Text = txtBancoDestino.Text Then
                    MsgBox("No es Posible Hacer un Traspaso a una Misma Cuenta, Favor de Verificar.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    txtCuentaBancariaOrigen.Text = ""
                    ModEstandar.RetrocederTab(Me)
                    Exit Sub
                End If
            End If
            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMonedaOrigen.Text = C_DESCPESOS
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMonedaOrigen.Text = C_DESCDOLARES
            End If
            If _optFormaPago_0.Checked Then
                ConsecutivoCheque = ObtieneNumCheque(CInt(txtBancoOrigen.Text), txtCuentaBancariaOrigen.Text, LetraFolio)
                txtNumeroCheque.Text = Trim(LetraFolio) & VB6.Format(CStr(ConsecutivoCheque), "000000")
            End If
        Else
            MsgBox("Cuenta Bancaria no Existe para este Banco, Favor de Verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtCuentaBancariaOrigen.Text = ""
            txtCuentaBancariaOrigen.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function Guardar() As Boolean
        Dim blnTransaccion As Boolean
        Dim strFolioIngreso As String
        Dim Ejercicio As Integer
        Dim Periodo As String
        On Error GoTo Err_Renamed

        'Do While (VB.Timer() - sglTiempoCambio) <= 2.1
        'Loop
        'System.Windows.Forms.Application.DoEvents()

        If Not mblnNuevo Then
            Exit Function
        End If
        If ValidaDatos() = False Then
            Exit Function
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        'Generar Folio del Movimiento de Egreso
        Ejercicio = CInt(VB6.Format(Year(CDate(dtpFecha.Value)), "0000"))
        Periodo = VB6.Format(Month(CDate(dtpFecha.Value)), "00")
        BuscaEjercicio(dtpFecha.Value)
        gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVEGRESO & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFolioEgreso.Text = C_TIPOMOVEGRESO & VB6.Format(Year(CDate(dtpFecha.Value)), "0000") & VB6.Format(Month(CDate(dtpFecha.Value)), "00") & VB6.Format(VB.Day(CDate(dtpFecha.Value)), "00") & VB6.Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
            ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVEGRESO, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        'Generar el Folio del Movimiento de Ingreso
        Ejercicio = CInt(VB6.Format(Year(CDate(dtpFecha.Value)), "0000"))
        Periodo = VB6.Format(Month(CDate(dtpFecha.Value)), "00")
        BuscaEjercicio(dtpFecha.Value)
        gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVINGRESO & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            strFolioIngreso = C_TIPOMOVINGRESO & VB6.Format(Year(CDate(dtpFecha.Value)), "0000") & VB6.Format(Month(CDate(dtpFecha.Value)), "00") & VB6.Format(VB.Day(CDate(dtpFecha.Value)), "00") & VB6.Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
            ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVINGRESO, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        'Obtener el Consecutivo de Cheque si es que se Genero Cheque y Actualizar el Consecutivo
        If _optFormaPago_0.Checked Then
            ConsecutivoCheque = ObtieneNumCheque(CInt(txtBancoOrigen.Text), txtCuentaBancariaOrigen.Text, LetraFolio)
            txtNumeroCheque.Text = Trim(LetraFolio) & VB6.Format(CStr(ConsecutivoCheque), "000000")
            ModStoredProcedures.PR_IMECatCuentasBancarias(CStr(txtBancoOrigen.Text), txtCuentaBancariaOrigen.Text, "", "", "", "", "0", VB6.Format(CStr(ConsecutivoCheque), "000000"), "", C_MODIFICACION, CStr(1))
            Cmd.Execute()
        End If
        'Guardar el Movimiento Bancario de Egreso
        ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioEgreso.Text, VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_MOVTRASPASO, C_TIPOMOVEGRESO, C_NATURALEZACOMERCIAL, IIf(lblMonedaOrigen.Text = C_DESCPESOS, C_PESO, C_DOLAR), txtTipoCambio.Text, IIf(_optFormaPago_0.Checked, C_FORMAPAGOCHEQUE, C_FORMAPAGOELECTRONICO), C_TIPOPAGOJOYERIA, txtBancoOrigen.Text, txtCuentaBancariaOrigen.Text, txtBeneficiario.Text, txtConcepto.Text, "0", "", "0", IIf(_optFormaPago_0.Checked, VB6.Format(dtpFechaCheque.Value, C_FORMATFECHAGUARDAR), "01/01/1900"), txtNumeroCheque.Text, txtImporteCtaOrigen.Text, "V", "01/01/1900", "", "0", "01/01/1900", C_MODULOBANCOS, strFolioIngreso, "", C_INSERCION, CStr(0))
        Cmd.Execute()
        'Guardar el Movimiento Bancario de Ingreso
        ModStoredProcedures.PR_IMEMovimientosBancarios(strFolioIngreso, VB6.Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_MOVTRASPASO, C_TIPOMOVINGRESO, C_NATURALEZACOMERCIAL, IIf(lblMonedaDestino.Text = C_DESCPESOS, C_PESO, C_DOLAR), txtTipoCambio.Text, "", C_TIPOPAGOJOYERIA, txtBancoDestino.Text, txtCuentaBancariaDestino.Text, "", txtConcepto.Text, "0", "", "0", "01/01/1900", "", txtImporteCtaDestino.Text, "V", "01/01/1900", "", "0", "01/01/1900", C_MODULOBANCOS, txtFolioEgreso.Text, "", C_INSERCION, CStr(0))
        Cmd.Execute()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("Los datos se han guardado con éxito" & vbNewLine & vbNewLine & "Folio Egreso  :  " & txtFolioEgreso.Text & Chr(13) & "Folio Ingreso :    " & strFolioIngreso, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Limpiar()
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnSalir = False
    End Sub

    Sub Limpiar()
        Nuevo()
        txtFolioEgreso.Text = ""
        txtFolioEgreso.Focus()
    End Sub

    Sub LlenaDatos()
        On Error GoTo Merr
        Dim I As Integer
        Dim Total As Decimal
        Dim RsAux As New ADODB.Recordset
        If Trim(txtFolioEgreso.Text) = "" Then
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT * FROM MovimientosBancarios WHERE FolioMovto = '" & txtFolioEgreso.Text & "' AND Movimiento = '" & C_MOVTRASPASO & "' AND " & "TipoMovto = '" & C_TIPOMOVEGRESO & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            gStrSql = "SELECT FolioMovto FROM MovimientosBancarios WHERE Referencia = '" & txtFolioEgreso.Text & "' AND Movimiento = '" & C_MOVCANCELACION & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsAux = Cmd.Execute
            If RsAux.RecordCount > 0 Then
                lblCancelacion.Text = "Movimiento de Cancelación : " & RsAux.Fields("FolioMovto").Value
            End If
            txtFolioEgreso.Text = Trim(RsGral.Fields("FolioMovto").Value)
            txtTipoCambio.Text = VB6.Format(RsGral.Fields("TipoCambio").Value, "###,##0.00")
            txtTipoCambio.ReadOnly = True
            dtpFecha.Value = VB6.Format(RsGral.Fields("FechaMovto").Value, C_FORMATFECHAMOSTRAR)
            If RsGral.Fields("FormaPago").Value = C_FORMAPAGOCHEQUE Then
                _optFormaPago_0.Checked = True
            ElseIf RsGral.Fields("FormaPago").Value = C_FORMAPAGOELECTRONICO Then
                optFormaPago(1).Checked = True
            End If
            txtBancoOrigen.Text = VB6.Format(RsGral.Fields("CodBanco").Value, "000")
            txtCuentaBancariaOrigen.Text = Trim(RsGral.Fields("CtaBancaria").Value)
            txtBeneficiario.Text = Trim(RsGral.Fields("Beneficiario").Value)
            txtConcepto.Text = Trim(RsGral.Fields("Concepto").Value)
            If Trim(RsGral.Fields("NoDocto").Value) <> "" Then
                txtNumeroCheque.Text = RsGral.Fields("NoDocto").Value
                dtpFechaCheque.Value = VB6.Format(RsGral.Fields("FechaDocto").Value, C_FORMATFECHAMOSTRAR)
            End If
            FueraChange = True
            txtImporteCtaOrigen.Text = VB6.Format(RsGral.Fields("importe").Value, "###,##0.00")
            FueraChange = False
            Frame1.Enabled = False
            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMonedaOrigen.Text = C_DESCPESOS
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMonedaOrigen.Text = C_DESCDOLARES
            End If
            gStrSql = "SELECT * FROM MovimientosBancarios WHERE Referencia = '" & txtFolioEgreso.Text & "' AND Movimiento = '" & C_MOVTRASPASO & "' AND " & "TipoMovto = '" & C_TIPOMOVINGRESO & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                txtBancoDestino.Text = VB6.Format(RsGral.Fields("CodBanco").Value, "000")
                txtCuentaBancariaDestino.Text = Trim(RsGral.Fields("CtaBancaria").Value)
                txtImporteCtaDestino.Text = VB6.Format(RsGral.Fields("importe").Value, "###,##0.00")
                lblMovIngreso.Text = "Folio de Ingreso: " & RsGral.Fields("FolioMovto").Value
                If RsGral.Fields("Moneda").Value = C_PESO Then
                    lblMonedaDestino.Text = C_DESCPESOS
                ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                    lblMonedaDestino.Text = C_DESCDOLARES
                End If
            End If
            mblnNuevo = False
            dtpFecha.Enabled = False
        Else
            MsgBox("Folio de Movimiento de Egreso no Existe ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFolioEgreso.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Nuevo()
        lblMonedaOrigen.Text = ""
        lblMonedaDestino.Text = ""
        lblCancelacion.Text = ""
        lblMovIngreso.Text = ""
        _optFormaPago_0.Checked = True
        dtpFecha.Value = VB6.Format(Now, "dd/mmm/yyyy")
        txtTipoCambio.ReadOnly = False
        txtTipoCambio.Text = VB6.Format(gcurCorpoTIPOCAMBIODOLAR, "###,##0.00")
        txtBancoOrigen.Text = ""
        txtCuentaBancariaOrigen.Text = ""
        txtBancoDestino.Text = ""
        txtCuentaBancariaDestino.Text = ""
        txtConcepto.Text = ""
        txtBeneficiario.Text = ""
        FueraChange = True
        txtImporteCtaOrigen.Text = "0.00"
        txtImporteCtaDestino.Text = "0.00"
        FueraChange = False
        dtpFecha.Enabled = False
        dtpFechaCheque.Enabled = True
        dtpFechaCheque.Value = VB6.Format(Now, "dd/mmm/yyyy")
        txtNumeroCheque.Enabled = True
        txtNumeroCheque.Text = ""
        Frame1.Enabled = True
        InicializaVariables()
    End Sub

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Not BuscaUltimoCierre(dtpFecha.Value) Then
            Exit Function
        End If
        If CDbl(Numerico(txtTipoCambio.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Tipo de Cambio", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtTipoCambio.Focus()
            Exit Function
        End If
        If Len(Trim(txtBancoOrigen.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Banco Origen", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtBancoOrigen.Focus()
            Exit Function
        End If
        If Len(Trim(txtCuentaBancariaOrigen.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Cuenta Bancaria Origen", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCuentaBancariaOrigen.Focus()
            Exit Function
        End If
        If Len(Trim(txtBancoDestino.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Banco Destino", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            txtBancoDestino.Focus()
            Exit Function
        End If
        If Len(Trim(txtCuentaBancariaDestino.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Cuenta Bancaria Destino", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtCuentaBancariaDestino.Focus()
            Exit Function
        End If
        If Len(Trim(txtBeneficiario.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Beneficiario", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtBeneficiario.Focus()
            Exit Function
        End If
        If Len(Trim(txtConcepto.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Concepto", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtConcepto.Focus()
            Exit Function
        End If
        If CDbl(Numerico(txtImporteCtaOrigen.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Importe de la Cuenta Origen", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtImporteCtaOrigen.Focus()
            Exit Function
        End If
        If Not ChecaSaldo(CInt(txtBancoOrigen.Text), Trim(txtCuentaBancariaOrigen.Text), CDec(txtImporteCtaOrigen.Text)) Then
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub dtpFecha_ValueChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFecha.ValueChanged
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFecha_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFecha.Click
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFecha_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFecha.KeyPress
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaCheque_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaCheque.Enter
        Pon_Tool()
    End Sub

    Private Sub frmBancosProcesoDiarioTraspasosBancarios_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoDiarioTraspasosBancarios_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoDiarioTraspasosBancarios_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If Me.ActiveControl.Name = "txtFolioEgreso" Then
                    If Len(Trim(txtFolioEgreso.Text)) = 13 And VB.Right(txtFolioEgreso.Text, 4) <> "0000" Then
                        Frame1.Enabled = False
                    End If
                End If
                If Me.ActiveControl.Name = "txtTipoCambio" Then
                    If CDbl(Numerico(txtTipoCambio.Text)) = 0 Then
                        MsgBox("El Tipo de Cambio debe ser Mayor que Cero, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        Exit Sub
                    End If
                End If
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtFolioEgreso" Then
                    If Me.ActiveControl.Name = "txtTipoCambio" Then
                        If CDbl(Numerico(txtTipoCambio.Text)) = 0 Then
                            MsgBox("El Tipo de Cambio debe ser Mayor que Cero, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            Exit Sub
                        End If
                    End If
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoDiarioTraspasosBancarios_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioTraspasosBancarios_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        'gstrMovimiento = "S"
        InicializaVariables()
        Nuevo()
        BuscaEjercicio(dtpFecha.Value)
    End Sub

    Private Sub frmBancosProcesoDiarioTraspasosBancarios_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmBancosProcesoDiarioTraspasosBancarios_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
        'Me.Close()
        gblnSalir = True
    End Sub
    Private Sub optFormaPago_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFormaPago.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = optFormaPago.GetIndex(eventSender)
            Select Case Index
                Case 0
                    dtpFechaCheque.Enabled = True
                    txtNumeroCheque.Enabled = True
                    If txtCuentaBancariaOrigen.Text <> "" Then
                        ConsecutivoCheque = ObtieneNumCheque(CInt(txtBancoOrigen.Text), txtCuentaBancariaOrigen.Text, LetraFolio)
                        txtNumeroCheque.Text = Trim(LetraFolio) & VB6.Format(CStr(ConsecutivoCheque), "000000")
                    End If
                Case 1
                    dtpFechaCheque.Enabled = False
                    txtNumeroCheque.Text = ""
                    txtNumeroCheque.Enabled = False
            End Select
        End If
    End Sub

    Private Sub optFormaPago_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFormaPago.Enter
        Dim Index As Integer = optFormaPago.GetIndex(eventSender)
        Select Case Index
            Case 0
                Pon_Tool()
            Case 1
                Pon_Tool()
        End Select
    End Sub

    Private Sub txtBancoDestino_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBancoDestino.TextChanged
        txtCuentaBancariaDestino.Text = ""
        lblMonedaDestino.Text = ""
    End Sub

    Private Sub txtBancoDestino_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBancoDestino.Enter
        SelTextoTxt(txtBancoDestino)
        Pon_Tool()
    End Sub

    Private Sub txtBancoDestino_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBancoDestino.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtBancoDestino_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBancoDestino.Leave
        If CDbl(Numerico(txtBancoDestino.Text)) = 0 Then
            txtBancoDestino.Text = "000"
        Else
            BuscaBancoDestino()
        End If
    End Sub

    Private Sub txtBancoOrigen_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBancoOrigen.TextChanged
        txtCuentaBancariaOrigen.Text = ""
        lblMonedaOrigen.Text = ""
    End Sub

    Private Sub txtBancoOrigen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBancoOrigen.Enter
        SelTextoTxt(txtBancoOrigen)
        Pon_Tool()
    End Sub

    Private Sub txtBancoOrigen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBancoOrigen.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtBancoOrigen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBancoOrigen.Leave
        If CDbl(Numerico(txtBancoOrigen.Text)) = 0 Then
            txtBancoOrigen.Text = "000"
        Else
            BuscaBancoOrigen()
        End If
    End Sub

    Private Sub txtBeneficiario_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBeneficiario.Enter
        SelTextoTxt(txtBeneficiario)
        Pon_Tool()
    End Sub

    Private Sub txtBeneficiario_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBeneficiario.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "!""#$%&/()=?'¡¿*,;.:<>@+-_")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtConcepto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.Enter
        SelTextoTxt(txtConcepto)
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
    Private Sub txtCuentaBancariaDestino_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCuentaBancariaDestino.TextChanged
        If mblnNuevo Then
            lblMonedaDestino.Text = ""
            txtImporteCtaDestino.Text = "0.00"
        End If
    End Sub

    Private Sub txtCuentaBancariaDestino_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCuentaBancariaDestino.Enter
        SelTextoTxt(txtCuentaBancariaDestino)
        Pon_Tool()
    End Sub

    Private Sub txtCuentaBancariaDestino_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCuentaBancariaDestino.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCuentaBancariaDestino_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCuentaBancariaDestino.Leave
        If Trim(txtCuentaBancariaDestino.Text) <> "" Then
            ChecaCuentaDestino()
        End If
        CalculaImporteDestino()
    End Sub

    Private Sub txtCuentaBancariaOrigen_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCuentaBancariaOrigen.TextChanged
        If mblnNuevo Then
            txtNumeroCheque.Text = ""
            lblMonedaOrigen.Text = ""
        End If
    End Sub

    Private Sub txtCuentaBancariaOrigen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCuentaBancariaOrigen.Enter
        SelTextoTxt(txtCuentaBancariaOrigen)
        Pon_Tool()
    End Sub

    Private Sub txtCuentaBancariaOrigen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCuentaBancariaOrigen.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCuentaBancariaOrigen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCuentaBancariaOrigen.Leave
        If Trim(txtCuentaBancariaOrigen.Text) <> "" Then
            ChecaCuentaOrigen()
        End If
        CalculaImporteDestino()
    End Sub
    Private Sub txtFolioEgreso_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEgreso.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtFolioEgreso_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEgreso.Enter
        SelTextoTxt(txtFolioEgreso)
        Pon_Tool()
    End Sub

    Private Sub txtFolioEgreso_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioEgreso.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, C_TIPOMOVEGRESO)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioEgreso_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEgreso.Leave
        'If System.Windows.Forms.Form.ActiveForm.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If Trim(txtFolioEgreso.Text) = "" Then
            txtFolioEgreso.Text = C_TIPOMOVEGRESO & VB6.Format(Year(CDate(dtpFecha.Value)), "0000") & VB6.Format(Month(CDate(dtpFecha.Value)), "00") & VB6.Format(VB.Day(CDate(dtpFecha.Value)), "00") & "0000"
            Exit Sub
        End If
        If mblnCambiosEnCodigo = True And txtFolioEgreso.Text <> "" And VB.Right(txtFolioEgreso.Text, 4) <> "0000" Then
            LlenaDatos()
        End If
    End Sub

    Private Sub txtImporteCtaDestino_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporteCtaDestino.Enter
        SelTextoTxt(txtImporteCtaDestino)
        Pon_Tool()
    End Sub
    Private Sub txtImporteCtaOrigen_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporteCtaOrigen.TextChanged
        If FueraChange = True Then Exit Sub
        If Trim(txtCuentaBancariaOrigen.Text) = "" Or Trim(txtCuentaBancariaDestino.Text) = "" Then
            MsgBox("No se han Capturado Correctamente las Cuentas Bancarias." & Chr(13) & "     Es Necesario Capturar las 2 Cuentas Bancarias.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            FueraChange = True
            txtImporteCtaOrigen.Text = "0.00"
            txtImporteCtaDestino.Text = "0.00"
            FueraChange = False
            Exit Sub
        End If
        If Trim(txtImporteCtaOrigen.Text) = "" Then
            txtImporteCtaOrigen.Text = "0.00"
        End If
        CalculaImporteDestino()
    End Sub

    Private Sub txtImporteCtaOrigen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporteCtaOrigen.Enter
        SelTextoTxt(txtImporteCtaOrigen)
        Pon_Tool()
    End Sub

    Private Sub txtImporteCtaOrigen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImporteCtaOrigen.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.MskCantidad(txtImporteCtaOrigen.Text, KeyAscii, 15, 2, (txtImporteCtaOrigen.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtImporteCtaOrigen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporteCtaOrigen.Leave
        txtImporteCtaOrigen.Text = VB6.Format(txtImporteCtaOrigen.Text, "###,##0.00")
    End Sub

    Private Sub txtNumeroCheque_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumeroCheque.Enter
        SelTextoTxt(txtNumeroCheque)
        Pon_Tool()
    End Sub

    Private Sub txtTipoCambio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.TextChanged
        If Trim(txtTipoCambio.Text) = "" Then
            txtTipoCambio.Text = "0.00"
        End If
    End Sub

    Private Sub txtTipoCambio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Enter
        SelTextoTxt(txtTipoCambio)
        Pon_Tool()
    End Sub

    Private Sub txtTipoCambio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.MskCantidad(txtTipoCambio.Text, KeyAscii, 3, 2, (txtTipoCambio.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Leave
        txtTipoCambio.Text = VB6.Format(txtTipoCambio.Text, "###,##0.00")
        CalculaImporteDestino()
    End Sub

    Private Sub txtTipoCambio_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtTipoCambio.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If CDbl(Numerico(txtTipoCambio.Text)) = 0 Then
            MsgBox("El Tipo de Cambio debe ser Mayor que Cero, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Cancel = True
        Else
            Cancel = False
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub
End Class