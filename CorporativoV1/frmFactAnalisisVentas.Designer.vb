<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFactAnalisisVentas
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        'Me.SuspendLayout()
        ''
        ''frmFactAnalisisVentas
        ''
        'Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        'Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        'Me.ClientSize = New System.Drawing.Size(723, 259)
        'Me.Name = "frmFactAnalisisVentas"
        'Me.Text = "frmFactAnalisisVentas"
        'Me.ResumeLayout(False)
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFactAnalisisVentas))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me.DtpDesde = New System.Windows.Forms.DateTimePicker
        Me.Frame7 = New System.Windows.Forms.GroupBox
        Me.lblFacturar = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.lblExcluido = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Frame6 = New System.Windows.Forms.GroupBox
        Me.cmdDatosFiscales = New System.Windows.Forms.Button
        Me.chkDoctoCliente = New System.Windows.Forms.CheckBox
        Me.chkDesglosarIva = New System.Windows.Forms.CheckBox
        Me.cmdGenerarFactura = New System.Windows.Forms.Button
        Me.cmdImpresionTickets = New System.Windows.Forms.Button
        Me.cmdImprimirFactura = New System.Windows.Forms.Button
        Me.Frame5 = New System.Windows.Forms.GroupBox
        Me.txtFacturaAdicional = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtFlex = New System.Windows.Forms.TextBox
        Me.Frame4 = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.lblSubTotal = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.lblRedondeo = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.lblTotal = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.lblTotalPesos = New System.Windows.Forms.Label
        Me.txtDescripcion = New System.Windows.Forms.TextBox
        Me.Frame3 = New System.Windows.Forms.GroupBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.lblImporteRedondeo = New System.Windows.Forms.Label
        Me.lblImporteTotal = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.lblImporteSubTotal = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.lblFactura = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.flexVentasPendientes = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
        Me.flexDetalleVenta = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
        Me.flexVentas = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
        Me.Frame2 = New System.Windows.Forms.GroupBox
        Me.txtFolioFactura = New System.Windows.Forms.TextBox
        Me.dtpFechaRegistro = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.txtPorcentaje = New System.Windows.Forms.TextBox
        Me.optPorcentual = New System.Windows.Forms.RadioButton
        Me.optManual = New System.Windows.Forms.RadioButton
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtCodSucursal = New System.Windows.Forms.TextBox
        Me.dtpFechaVenta = New System.Windows.Forms.DateTimePicker
        Me.dbcSucursal = New AxMSDataListLib.AxDataCombo
        Me.DtpHasta = New System.Windows.Forms.DateTimePicker
        Me.lblMoneda = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.lblSubTot = New System.Windows.Forms.Label
        Me.lblIva = New System.Windows.Forms.Label
        Me.lblDescuento = New System.Windows.Forms.Label
        Me.lblCantidad = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.lblDesc = New System.Windows.Forms.Label
        Me.lblDescripcion = New System.Windows.Forms.Label
        Me.lblEstadoFolio = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Frame7.SuspendLayout()
        Me.Frame6.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        CType(Me.DtpDesde, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.flexVentasPendientes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.flexDetalleVenta, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.flexVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtpFechaRegistro, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtpFechaVenta, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dbcSucursal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DtpHasta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Text = "Análisis de las Ventas"
        Me.ClientSize = New System.Drawing.Size(1009, 579)
        Me.Location = New System.Drawing.Point(8, 113)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ControlBox = True
        Me.Enabled = True
        Me.MinimizeBox = True
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = True
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmFactAnalisisVentas"
        'DtpDesde.OcxState = CType(resources.GetObject("DtpDesde.OcxState"), System.Windows.Forms.AxHost.State)
        Me.DtpDesde.Size = New System.Drawing.Size(97, 21)
        Me.DtpDesde.Location = New System.Drawing.Point(8, 324)
        Me.DtpDesde.TabIndex = 20
        Me.DtpDesde.Name = "DtpDesde"
        Me.Frame7.Size = New System.Drawing.Size(224, 66)
        Me.Frame7.Location = New System.Drawing.Point(10, 62)
        Me.Frame7.TabIndex = 59
        Me.Frame7.BackColor = System.Drawing.SystemColors.Control
        Me.Frame7.Enabled = True
        Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame7.Visible = True
        Me.Frame7.Name = "Frame7"
        Me.lblFacturar.BackColor = System.Drawing.Color.FromArgb(192, 255, 192)
        Me.lblFacturar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblFacturar.Size = New System.Drawing.Size(17, 17)
        Me.lblFacturar.Location = New System.Drawing.Point(10, 37)
        Me.lblFacturar.TabIndex = 63
        Me.lblFacturar.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblFacturar.Enabled = True
        Me.lblFacturar.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFacturar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFacturar.UseMnemonic = True
        Me.lblFacturar.Visible = True
        Me.lblFacturar.AutoSize = False
        Me.lblFacturar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFacturar.Name = "lblFacturar"
        Me.Label14.Text = "Folios Seleccionados Para Facturación del Pto. de Venta"
        Me.Label14.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Size = New System.Drawing.Size(178, 29)
        Me.Label14.Location = New System.Drawing.Point(36, 32)
        Me.Label14.TabIndex = 62
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Enabled = True
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.UseMnemonic = True
        Me.Label14.Visible = True
        Me.Label14.AutoSize = False
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label14.Name = "Label14"
        Me.lblExcluido.BackColor = System.Drawing.Color.FromArgb(192, 192, 255)
        Me.lblExcluido.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblExcluido.Size = New System.Drawing.Size(17, 17)
        Me.lblExcluido.Location = New System.Drawing.Point(10, 12)
        Me.lblExcluido.TabIndex = 61
        Me.lblExcluido.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblExcluido.Enabled = True
        Me.lblExcluido.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblExcluido.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblExcluido.UseMnemonic = True
        Me.lblExcluido.Visible = True
        Me.lblExcluido.AutoSize = False
        Me.lblExcluido.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblExcluido.Name = "lblExcluido"
        Me.Label18.Text = "Partidas Excluidas"
        Me.Label18.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Size = New System.Drawing.Size(113, 17)
        Me.Label18.Location = New System.Drawing.Point(36, 15)
        Me.Label18.TabIndex = 60
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Enabled = True
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label18.UseMnemonic = True
        Me.Label18.Visible = True
        Me.Label18.AutoSize = False
        Me.Label18.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label18.Name = "Label18"
        Me.Frame6.Text = "Documento..."
        Me.Frame6.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.Frame6.Size = New System.Drawing.Size(257, 126)
        Me.Frame6.Location = New System.Drawing.Point(744, 448)
        Me.Frame6.TabIndex = 52
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Enabled = True
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Visible = True
        Me.Frame6.Name = "Frame6"
        Me.cmdDatosFiscales.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdDatosFiscales.Text = "Datos Fiscales"
        Me.cmdDatosFiscales.Size = New System.Drawing.Size(94, 25)
        Me.cmdDatosFiscales.Location = New System.Drawing.Point(153, 50)
        Me.cmdDatosFiscales.TabIndex = 57
        Me.cmdDatosFiscales.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDatosFiscales.CausesValidation = True
        Me.cmdDatosFiscales.Enabled = True
        Me.cmdDatosFiscales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDatosFiscales.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDatosFiscales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDatosFiscales.TabStop = True
        Me.cmdDatosFiscales.Name = "cmdDatosFiscales"
        Me.chkDoctoCliente.Text = "Docto. Cliente"
        Me.chkDoctoCliente.Size = New System.Drawing.Size(103, 21)
        Me.chkDoctoCliente.Location = New System.Drawing.Point(141, 24)
        Me.chkDoctoCliente.TabIndex = 56
        Me.chkDoctoCliente.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.chkDoctoCliente.FlatStyle = System.Windows.Forms.FlatStyle.Standard
        Me.chkDoctoCliente.BackColor = System.Drawing.SystemColors.Control
        Me.chkDoctoCliente.CausesValidation = True
        Me.chkDoctoCliente.Enabled = True
        Me.chkDoctoCliente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDoctoCliente.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDoctoCliente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDoctoCliente.Appearance = System.Windows.Forms.Appearance.Normal
        Me.chkDoctoCliente.TabStop = True
        Me.chkDoctoCliente.CheckState = System.Windows.Forms.CheckState.Unchecked
        Me.chkDoctoCliente.Visible = True
        Me.chkDoctoCliente.Name = "chkDoctoCliente"
        Me.chkDesglosarIva.Text = "Desglosar Iva"
        Me.chkDesglosarIva.Size = New System.Drawing.Size(89, 13)
        Me.chkDesglosarIva.Location = New System.Drawing.Point(160, 96)
        Me.chkDesglosarIva.TabIndex = 58
        Me.chkDesglosarIva.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.chkDesglosarIva.FlatStyle = System.Windows.Forms.FlatStyle.Standard
        Me.chkDesglosarIva.BackColor = System.Drawing.SystemColors.Control
        Me.chkDesglosarIva.CausesValidation = True
        Me.chkDesglosarIva.Enabled = True
        Me.chkDesglosarIva.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDesglosarIva.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDesglosarIva.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDesglosarIva.Appearance = System.Windows.Forms.Appearance.Normal
        Me.chkDesglosarIva.TabStop = True
        Me.chkDesglosarIva.CheckState = System.Windows.Forms.CheckState.Unchecked
        Me.chkDesglosarIva.Visible = True
        Me.chkDesglosarIva.Name = "chkDesglosarIva"
        Me.cmdGenerarFactura.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdGenerarFactura.Text = "Generar Factura"
        Me.cmdGenerarFactura.Size = New System.Drawing.Size(113, 25)
        Me.cmdGenerarFactura.Location = New System.Drawing.Point(12, 24)
        Me.cmdGenerarFactura.TabIndex = 53
        Me.cmdGenerarFactura.BackColor = System.Drawing.SystemColors.Control
        Me.cmdGenerarFactura.CausesValidation = True
        Me.cmdGenerarFactura.Enabled = True
        Me.cmdGenerarFactura.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdGenerarFactura.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdGenerarFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdGenerarFactura.TabStop = True
        Me.cmdGenerarFactura.Name = "cmdGenerarFactura"
        Me.cmdImpresionTickets.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdImpresionTickets.Text = "Impresión de Tickets"
        Me.cmdImpresionTickets.Size = New System.Drawing.Size(113, 25)
        Me.cmdImpresionTickets.Location = New System.Drawing.Point(12, 56)
        Me.cmdImpresionTickets.TabIndex = 54
        Me.cmdImpresionTickets.BackColor = System.Drawing.SystemColors.Control
        Me.cmdImpresionTickets.CausesValidation = True
        Me.cmdImpresionTickets.Enabled = True
        Me.cmdImpresionTickets.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdImpresionTickets.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdImpresionTickets.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdImpresionTickets.TabStop = True
        Me.cmdImpresionTickets.Name = "cmdImpresionTickets"
        Me.cmdImprimirFactura.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdImprimirFactura.Text = "Imprimir Factura"
        Me.cmdImprimirFactura.Size = New System.Drawing.Size(113, 25)
        Me.cmdImprimirFactura.Location = New System.Drawing.Point(12, 88)
        Me.cmdImprimirFactura.TabIndex = 55
        Me.cmdImprimirFactura.BackColor = System.Drawing.SystemColors.Control
        Me.cmdImprimirFactura.CausesValidation = True
        Me.cmdImprimirFactura.Enabled = True
        Me.cmdImprimirFactura.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdImprimirFactura.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdImprimirFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdImprimirFactura.TabStop = True
        Me.cmdImprimirFactura.Name = "cmdImprimirFactura"
        Me.Frame5.Text = "Dato Adicional"
        Me.Frame5.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.Frame5.Size = New System.Drawing.Size(205, 71)
        Me.Frame5.Location = New System.Drawing.Point(543, 11)
        Me.Frame5.TabIndex = 50
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Enabled = True
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Visible = True
        Me.Frame5.Name = "Frame5"
        Me.txtFacturaAdicional.AutoSize = False
        Me.txtFacturaAdicional.Size = New System.Drawing.Size(132, 19)
        Me.txtFacturaAdicional.Location = New System.Drawing.Point(60, 30)
        Me.txtFacturaAdicional.MaxLength = 17
        Me.txtFacturaAdicional.TabIndex = 6
        Me.txtFacturaAdicional.AcceptsReturn = True
        Me.txtFacturaAdicional.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtFacturaAdicional.BackColor = System.Drawing.SystemColors.Window
        Me.txtFacturaAdicional.CausesValidation = True
        Me.txtFacturaAdicional.Enabled = True
        Me.txtFacturaAdicional.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFacturaAdicional.HideSelection = True
        Me.txtFacturaAdicional.ReadOnly = False
        Me.txtFacturaAdicional.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFacturaAdicional.Multiline = False
        Me.txtFacturaAdicional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFacturaAdicional.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtFacturaAdicional.TabStop = True
        Me.txtFacturaAdicional.Visible = True
        Me.txtFacturaAdicional.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtFacturaAdicional.Name = "txtFacturaAdicional"
        Me.Label7.Text = "Factura:"
        Me.Label7.Size = New System.Drawing.Size(43, 13)
        Me.Label7.Location = New System.Drawing.Point(11, 33)
        Me.Label7.TabIndex = 51
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Enabled = True
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.UseMnemonic = True
        Me.Label7.Visible = True
        Me.Label7.AutoSize = False
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label7.Name = "Label7"
        Me.txtFlex.AutoSize = False
        Me.txtFlex.Size = New System.Drawing.Size(73, 21)
        Me.txtFlex.Location = New System.Drawing.Point(376, 178)
        Me.txtFlex.TabIndex = 43
        Me.txtFlex.Visible = False
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.CausesValidation = True
        Me.txtFlex.Enabled = True
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.HideSelection = True
        Me.txtFlex.ReadOnly = False
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.Multiline = False
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtFlex.TabStop = True
        Me.txtFlex.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtFlex.Name = "txtFlex"
        Me.Frame4.Text = "Totales por Folio de Venta"
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.Frame4.Size = New System.Drawing.Size(257, 129)
        Me.Frame4.Location = New System.Drawing.Point(744, 312)
        Me.Frame4.TabIndex = 32
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Enabled = True
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Visible = True
        Me.Frame4.Name = "Frame4"
        Me.Label6.Text = "SubTotal :"
        Me.Label6.Size = New System.Drawing.Size(57, 21)
        Me.Label6.Location = New System.Drawing.Point(32, 18)
        Me.Label6.TabIndex = 40
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Enabled = True
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.UseMnemonic = True
        Me.Label6.Visible = True
        Me.Label6.AutoSize = False
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label6.Name = "Label6"
        Me.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblSubTotal.BackColor = System.Drawing.SystemColors.Window
        Me.lblSubTotal.Size = New System.Drawing.Size(121, 21)
        Me.lblSubTotal.Location = New System.Drawing.Point(107, 16)
        Me.lblSubTotal.TabIndex = 39
        Me.lblSubTotal.Enabled = True
        Me.lblSubTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSubTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSubTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSubTotal.UseMnemonic = True
        Me.lblSubTotal.Visible = True
        Me.lblSubTotal.AutoSize = False
        Me.lblSubTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSubTotal.Name = "lblSubTotal"
        Me.Label8.Text = "Redondeo :"
        Me.Label8.Size = New System.Drawing.Size(57, 21)
        Me.Label8.Location = New System.Drawing.Point(32, 45)
        Me.Label8.TabIndex = 38
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Enabled = True
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.UseMnemonic = True
        Me.Label8.Visible = True
        Me.Label8.AutoSize = False
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label8.Name = "Label8"
        Me.lblRedondeo.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblRedondeo.BackColor = System.Drawing.SystemColors.Window
        Me.lblRedondeo.Size = New System.Drawing.Size(121, 21)
        Me.lblRedondeo.Location = New System.Drawing.Point(107, 43)
        Me.lblRedondeo.TabIndex = 37
        Me.lblRedondeo.Enabled = True
        Me.lblRedondeo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRedondeo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRedondeo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRedondeo.UseMnemonic = True
        Me.lblRedondeo.Visible = True
        Me.lblRedondeo.AutoSize = False
        Me.lblRedondeo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRedondeo.Name = "lblRedondeo"
        Me.Label10.Text = "Total :"
        Me.Label10.Size = New System.Drawing.Size(57, 21)
        Me.Label10.Location = New System.Drawing.Point(32, 71)
        Me.Label10.TabIndex = 36
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Enabled = True
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.UseMnemonic = True
        Me.Label10.Visible = True
        Me.Label10.AutoSize = False
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label10.Name = "Label10"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblTotal.BackColor = System.Drawing.SystemColors.Window
        Me.lblTotal.Size = New System.Drawing.Size(121, 21)
        Me.lblTotal.Location = New System.Drawing.Point(107, 69)
        Me.lblTotal.TabIndex = 35
        Me.lblTotal.Enabled = True
        Me.lblTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotal.UseMnemonic = True
        Me.lblTotal.Visible = True
        Me.lblTotal.AutoSize = False
        Me.lblTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotal.Name = "lblTotal"
        Me.Label12.Text = "Total Pesos :"
        Me.Label12.Size = New System.Drawing.Size(65, 21)
        Me.Label12.Location = New System.Drawing.Point(32, 98)
        Me.Label12.TabIndex = 34
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Enabled = True
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.UseMnemonic = True
        Me.Label12.Visible = True
        Me.Label12.AutoSize = False
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label12.Name = "Label12"
        Me.lblTotalPesos.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblTotalPesos.BackColor = System.Drawing.SystemColors.Window
        Me.lblTotalPesos.Size = New System.Drawing.Size(121, 21)
        Me.lblTotalPesos.Location = New System.Drawing.Point(107, 96)
        Me.lblTotalPesos.TabIndex = 33
        Me.lblTotalPesos.Enabled = True
        Me.lblTotalPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotalPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotalPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotalPesos.UseMnemonic = True
        Me.lblTotalPesos.Visible = True
        Me.lblTotalPesos.AutoSize = False
        Me.lblTotalPesos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalPesos.Name = "lblTotalPesos"
        Me.txtDescripcion.AutoSize = False
        Me.txtDescripcion.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtDescripcion.Size = New System.Drawing.Size(225, 21)
        Me.txtDescripcion.Location = New System.Drawing.Point(776, 287)
        Me.txtDescripcion.MaxLength = 50
        Me.txtDescripcion.TabIndex = 19
        Me.txtDescripcion.AcceptsReturn = True
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescripcion.CausesValidation = True
        Me.txtDescripcion.Enabled = True
        Me.txtDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescripcion.HideSelection = True
        Me.txtDescripcion.ReadOnly = False
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescripcion.Multiline = False
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtDescripcion.TabStop = True
        Me.txtDescripcion.Visible = True
        Me.txtDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.Frame3.Text = "Totales de la Factura"
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.Frame3.Size = New System.Drawing.Size(721, 49)
        Me.Frame3.Location = New System.Drawing.Point(8, 525)
        Me.Frame3.TabIndex = 25
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Enabled = True
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Visible = True
        Me.Frame3.Name = "Frame3"
        Me.Label11.Text = "Redondeo :"
        Me.Label11.Size = New System.Drawing.Size(57, 17)
        Me.Label11.Location = New System.Drawing.Point(388, 20)
        Me.Label11.TabIndex = 46
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Enabled = True
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.UseMnemonic = True
        Me.Label11.Visible = True
        Me.Label11.AutoSize = False
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label11.Name = "Label11"
        Me.lblImporteRedondeo.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblImporteRedondeo.BackColor = System.Drawing.SystemColors.Window
        Me.lblImporteRedondeo.Size = New System.Drawing.Size(85, 21)
        Me.lblImporteRedondeo.Location = New System.Drawing.Point(448, 18)
        Me.lblImporteRedondeo.TabIndex = 45
        Me.lblImporteRedondeo.Enabled = True
        Me.lblImporteRedondeo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblImporteRedondeo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblImporteRedondeo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblImporteRedondeo.UseMnemonic = True
        Me.lblImporteRedondeo.Visible = True
        Me.lblImporteRedondeo.AutoSize = False
        Me.lblImporteRedondeo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImporteRedondeo.Name = "lblImporteRedondeo"
        Me.lblImporteTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblImporteTotal.BackColor = System.Drawing.SystemColors.Window
        Me.lblImporteTotal.Size = New System.Drawing.Size(129, 21)
        Me.lblImporteTotal.Location = New System.Drawing.Point(580, 18)
        Me.lblImporteTotal.TabIndex = 31
        Me.lblImporteTotal.Enabled = True
        Me.lblImporteTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblImporteTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblImporteTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblImporteTotal.UseMnemonic = True
        Me.lblImporteTotal.Visible = True
        Me.lblImporteTotal.AutoSize = False
        Me.lblImporteTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImporteTotal.Name = "lblImporteTotal"
        Me.Label15.Text = "Total :"
        Me.Label15.Size = New System.Drawing.Size(37, 17)
        Me.Label15.Location = New System.Drawing.Point(548, 20)
        Me.Label15.TabIndex = 30
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Enabled = True
        Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label15.UseMnemonic = True
        Me.Label15.Visible = True
        Me.Label15.AutoSize = False
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label15.Name = "Label15"
        Me.lblImporteSubTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblImporteSubTotal.BackColor = System.Drawing.SystemColors.Window
        Me.lblImporteSubTotal.Size = New System.Drawing.Size(129, 21)
        Me.lblImporteSubTotal.Location = New System.Drawing.Point(240, 18)
        Me.lblImporteSubTotal.TabIndex = 29
        Me.lblImporteSubTotal.Enabled = True
        Me.lblImporteSubTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblImporteSubTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblImporteSubTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblImporteSubTotal.UseMnemonic = True
        Me.lblImporteSubTotal.Visible = True
        Me.lblImporteSubTotal.AutoSize = False
        Me.lblImporteSubTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImporteSubTotal.Name = "lblImporteSubTotal"
        Me.Label13.Text = "SubTotal :"
        Me.Label13.Size = New System.Drawing.Size(57, 17)
        Me.Label13.Location = New System.Drawing.Point(188, 20)
        Me.Label13.TabIndex = 28
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Enabled = True
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.UseMnemonic = True
        Me.Label13.Visible = True
        Me.Label13.AutoSize = False
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label13.Name = "Label13"
        Me.lblFactura.BackColor = System.Drawing.SystemColors.Window
        Me.lblFactura.Size = New System.Drawing.Size(113, 21)
        Me.lblFactura.Location = New System.Drawing.Point(56, 18)
        Me.lblFactura.TabIndex = 27
        Me.lblFactura.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblFactura.Enabled = True
        Me.lblFactura.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFactura.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFactura.UseMnemonic = True
        Me.lblFactura.Visible = True
        Me.lblFactura.AutoSize = False
        Me.lblFactura.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFactura.Name = "lblFactura"
        Me.Label9.Text = "Factura :"
        Me.Label9.Size = New System.Drawing.Size(49, 17)
        Me.Label9.Location = New System.Drawing.Point(12, 20)
        Me.Label9.TabIndex = 26
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Enabled = True
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.UseMnemonic = True
        Me.Label9.Visible = True
        Me.Label9.AutoSize = False
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label9.Name = "Label9"
        flexVentasPendientes.OcxState = CType(resources.GetObject("flexVentasPendientes.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexVentasPendientes.Size = New System.Drawing.Size(721, 151)
        Me.flexVentasPendientes.Location = New System.Drawing.Point(8, 349)
        Me.flexVentasPendientes.TabIndex = 22
        Me.flexVentasPendientes.Name = "flexVentasPendientes"
        flexDetalleVenta.OcxState = CType(resources.GetObject("flexDetalleVenta.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalleVenta.Size = New System.Drawing.Size(761, 151)
        Me.flexDetalleVenta.Location = New System.Drawing.Point(240, 136)
        Me.flexDetalleVenta.TabIndex = 18
        Me.flexDetalleVenta.Name = "flexDetalleVenta"
        flexVentas.OcxState = CType(resources.GetObject("flexVentas.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexVentas.Size = New System.Drawing.Size(228, 129)
        Me.flexVentas.Location = New System.Drawing.Point(8, 157)
        Me.flexVentas.TabIndex = 9
        Me.flexVentas.Name = "flexVentas"
        Me.Frame2.Text = "Facturas Registradas"
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.Frame2.Size = New System.Drawing.Size(241, 71)
        Me.Frame2.Location = New System.Drawing.Point(759, 11)
        Me.Frame2.TabIndex = 14
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Enabled = True
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Visible = True
        Me.Frame2.Name = "Frame2"
        Me.txtFolioFactura.AutoSize = False
        Me.txtFolioFactura.Size = New System.Drawing.Size(115, 21)
        Me.txtFolioFactura.Location = New System.Drawing.Point(111, 16)
        Me.txtFolioFactura.MaxLength = 17
        Me.txtFolioFactura.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtFolioFactura, "Folio de la Factura.")
        Me.txtFolioFactura.AcceptsReturn = True
        Me.txtFolioFactura.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtFolioFactura.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioFactura.CausesValidation = True
        Me.txtFolioFactura.Enabled = True
        Me.txtFolioFactura.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioFactura.HideSelection = True
        Me.txtFolioFactura.ReadOnly = False
        Me.txtFolioFactura.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioFactura.Multiline = False
        Me.txtFolioFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioFactura.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtFolioFactura.TabStop = True
        Me.txtFolioFactura.Visible = True
        Me.txtFolioFactura.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtFolioFactura.Name = "txtFolioFactura"
        'dtpFechaRegistro.OcxState = CType(resources.GetObject("dtpFechaRegistro.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dtpFechaRegistro.Size = New System.Drawing.Size(115, 21)
        Me.dtpFechaRegistro.Location = New System.Drawing.Point(112, 41)
        Me.dtpFechaRegistro.TabIndex = 8
        Me.dtpFechaRegistro.Name = "dtpFechaRegistro"
        Me.Label5.Text = "Fecha de Registro :"
        Me.Label5.Size = New System.Drawing.Size(97, 16)
        Me.Label5.Location = New System.Drawing.Point(8, 45)
        Me.Label5.TabIndex = 16
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Enabled = True
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.UseMnemonic = True
        Me.Label5.Visible = True
        Me.Label5.AutoSize = False
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label5.Name = "Label5"
        Me.Label4.Text = "Folio Factura :"
        Me.Label4.Size = New System.Drawing.Size(73, 15)
        Me.Label4.Location = New System.Drawing.Point(8, 21)
        Me.Label4.TabIndex = 15
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Enabled = True
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.UseMnemonic = True
        Me.Label4.Visible = True
        Me.Label4.AutoSize = False
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label4.Name = "Label4"
        Me.Frame1.Text = "Método"
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.Frame1.Size = New System.Drawing.Size(185, 70)
        Me.Frame1.Location = New System.Drawing.Point(347, 11)
        Me.Frame1.TabIndex = 12
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Enabled = True
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Visible = True
        Me.Frame1.Name = "Frame1"
        Me.txtPorcentaje.AutoSize = False
        Me.txtPorcentaje.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPorcentaje.Size = New System.Drawing.Size(46, 21)
        Me.txtPorcentaje.Location = New System.Drawing.Point(108, 37)
        Me.txtPorcentaje.MaxLength = 3
        Me.txtPorcentaje.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtPorcentaje, "Porcentaje a Aplicar.")
        Me.txtPorcentaje.AcceptsReturn = True
        Me.txtPorcentaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtPorcentaje.CausesValidation = True
        Me.txtPorcentaje.Enabled = True
        Me.txtPorcentaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPorcentaje.HideSelection = True
        Me.txtPorcentaje.ReadOnly = False
        Me.txtPorcentaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPorcentaje.Multiline = False
        Me.txtPorcentaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPorcentaje.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtPorcentaje.TabStop = True
        Me.txtPorcentaje.Visible = True
        Me.txtPorcentaje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtPorcentaje.Name = "txtPorcentaje"
        Me.optPorcentual.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.optPorcentual.Text = "Porcentual"
        Me.optPorcentual.Size = New System.Drawing.Size(81, 17)
        Me.optPorcentual.Location = New System.Drawing.Point(16, 42)
        Me.optPorcentual.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.optPorcentual, "Mediante este método, se Proporciona un porcentaje y se hacen los calculos automaticamente.")
        Me.optPorcentual.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.optPorcentual.BackColor = System.Drawing.SystemColors.Control
        Me.optPorcentual.CausesValidation = True
        Me.optPorcentual.Enabled = True
        Me.optPorcentual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPorcentual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPorcentual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPorcentual.Appearance = System.Windows.Forms.Appearance.Normal
        Me.optPorcentual.TabStop = True
        Me.optPorcentual.Checked = False
        Me.optPorcentual.Visible = True
        Me.optPorcentual.Name = "optPorcentual"
        Me.optManual.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.optManual.Text = "Manual"
        Me.optManual.Size = New System.Drawing.Size(81, 17)
        Me.optManual.Location = New System.Drawing.Point(16, 19)
        Me.optManual.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.optManual, "Mediante este método, los calculos se hacen manualmente.")
        Me.optManual.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.optManual.BackColor = System.Drawing.SystemColors.Control
        Me.optManual.CausesValidation = True
        Me.optManual.Enabled = True
        Me.optManual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optManual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optManual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optManual.Appearance = System.Windows.Forms.Appearance.Normal
        Me.optManual.TabStop = True
        Me.optManual.Checked = False
        Me.optManual.Visible = True
        Me.optManual.Name = "optManual"
        Me.Label3.Text = "%"
        Me.Label3.Size = New System.Drawing.Size(8, 13)
        Me.Label3.Location = New System.Drawing.Point(162, 42)
        Me.Label3.TabIndex = 13
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Enabled = True
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.UseMnemonic = True
        Me.Label3.Visible = True
        Me.Label3.AutoSize = True
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label3.Name = "Label3"
        Me.txtCodSucursal.AutoSize = False
        Me.txtCodSucursal.Enabled = False
        Me.txtCodSucursal.Size = New System.Drawing.Size(41, 21)
        Me.txtCodSucursal.Location = New System.Drawing.Point(260, 63)
        Me.txtCodSucursal.MaxLength = 3
        Me.txtCodSucursal.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCodSucursal, "Codigo de la Sucursal.")
        Me.txtCodSucursal.Visible = False
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.CausesValidation = True
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.HideSelection = True
        Me.txtCodSucursal.ReadOnly = False
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.Multiline = False
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtCodSucursal.TabStop = True
        Me.txtCodSucursal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtCodSucursal.Name = "txtCodSucursal"
        'dtpFechaVenta.OcxState = CType(resources.GetObject("dtpFechaVenta.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dtpFechaVenta.Size = New System.Drawing.Size(97, 21)
        Me.dtpFechaVenta.Location = New System.Drawing.Point(96, 10)
        Me.dtpFechaVenta.TabIndex = 0
        Me.dtpFechaVenta.Name = "dtpFechaVenta"
        dbcSucursal.OcxState = CType(resources.GetObject("dbcSucursal.OcxState"), System.Windows.Forms.AxHost.State)
        Me.dbcSucursal.Size = New System.Drawing.Size(203, 21)
        Me.dbcSucursal.Location = New System.Drawing.Point(95, 38)
        Me.dbcSucursal.TabIndex = 2
        Me.dbcSucursal.Name = "dbcSucursal"
        'DtpHasta.OcxState = CType(resources.GetObject("DtpHasta.OcxState"), System.Windows.Forms.AxHost.State)
        Me.DtpHasta.Size = New System.Drawing.Size(97, 21)
        Me.DtpHasta.Location = New System.Drawing.Point(139, 324)
        Me.DtpHasta.TabIndex = 21
        Me.DtpHasta.Name = "DtpHasta"
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Window
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(192, 0, 0)
        Me.lblMoneda.Size = New System.Drawing.Size(102, 21)
        Me.lblMoneda.Location = New System.Drawing.Point(117, 287)
        Me.lblMoneda.TabIndex = 67
        Me.lblMoneda.Enabled = True
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.UseMnemonic = True
        Me.lblMoneda.Visible = True
        Me.lblMoneda.AutoSize = False
        Me.lblMoneda.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMoneda.Name = "lblMoneda"
        Me.Label21.Text = "Hasta"
        Me.Label21.Size = New System.Drawing.Size(28, 13)
        Me.Label21.Location = New System.Drawing.Point(176, 310)
        Me.Label21.TabIndex = 66
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label21.BackColor = System.Drawing.SystemColors.Control
        Me.Label21.Enabled = True
        Me.Label21.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label21.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label21.UseMnemonic = True
        Me.Label21.Visible = True
        Me.Label21.AutoSize = True
        Me.Label21.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label21.Name = "Label21"
        Me.Label20.Text = "Desde"
        Me.Label20.Size = New System.Drawing.Size(31, 13)
        Me.Label20.Location = New System.Drawing.Point(41, 310)
        Me.Label20.TabIndex = 65
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label20.BackColor = System.Drawing.SystemColors.Control
        Me.Label20.Enabled = True
        Me.Label20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label20.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label20.UseMnemonic = True
        Me.Label20.Visible = True
        Me.Label20.AutoSize = True
        Me.Label20.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label20.Name = "Label20"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label17.Text = "Ventas"
        Me.Label17.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label17.Size = New System.Drawing.Size(228, 22)
        Me.Label17.Location = New System.Drawing.Point(8, 136)
        Me.Label17.TabIndex = 64
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Enabled = True
        Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label17.UseMnemonic = True
        Me.Label17.Visible = True
        Me.Label17.AutoSize = False
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Name = "Label17"
        Me.lblSubTot.Size = New System.Drawing.Size(65, 17)
        Me.lblSubTot.Location = New System.Drawing.Point(800, 488)
        Me.lblSubTot.TabIndex = 49
        Me.lblSubTot.Visible = False
        Me.lblSubTot.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblSubTot.BackColor = System.Drawing.SystemColors.Control
        Me.lblSubTot.Enabled = True
        Me.lblSubTot.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSubTot.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSubTot.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSubTot.UseMnemonic = True
        Me.lblSubTot.AutoSize = False
        Me.lblSubTot.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lblSubTot.Name = "lblSubTot"
        Me.lblIva.Size = New System.Drawing.Size(33, 25)
        Me.lblIva.Location = New System.Drawing.Point(840, 438)
        Me.lblIva.TabIndex = 48
        Me.lblIva.Visible = False
        Me.lblIva.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblIva.BackColor = System.Drawing.SystemColors.Control
        Me.lblIva.Enabled = True
        Me.lblIva.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblIva.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblIva.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblIva.UseMnemonic = True
        Me.lblIva.AutoSize = False
        Me.lblIva.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lblIva.Name = "lblIva"
        Me.lblDescuento.Size = New System.Drawing.Size(65, 17)
        Me.lblDescuento.Location = New System.Drawing.Point(752, 464)
        Me.lblDescuento.TabIndex = 47
        Me.lblDescuento.Visible = False
        Me.lblDescuento.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblDescuento.BackColor = System.Drawing.SystemColors.Control
        Me.lblDescuento.Enabled = True
        Me.lblDescuento.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDescuento.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDescuento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDescuento.UseMnemonic = True
        Me.lblDescuento.AutoSize = False
        Me.lblDescuento.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lblDescuento.Name = "lblDescuento"
        Me.lblCantidad.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblCantidad.BackColor = System.Drawing.SystemColors.Window
        Me.lblCantidad.Size = New System.Drawing.Size(47, 21)
        Me.lblCantidad.Location = New System.Drawing.Point(728, 287)
        Me.lblCantidad.TabIndex = 44
        Me.lblCantidad.Enabled = True
        Me.lblCantidad.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCantidad.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCantidad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCantidad.UseMnemonic = True
        Me.lblCantidad.Visible = True
        Me.lblCantidad.AutoSize = False
        Me.lblCantidad.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCantidad.Name = "lblCantidad"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label19.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
        Me.Label19.Text = "Presione <Supr> o Haga Doble Click Para Excluir una Partida, Presione <F4> para Regresar una Partida al Grid de Ventas Pendientes, Presione <F6> En el Grid de Ventas para Marcar o Desmarcar un Folio de Venta para ser Incluido en la Factura del Pto. de Venta"
        Me.Label19.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.Label19.Size = New System.Drawing.Size(759, 33)
        Me.Label19.Location = New System.Drawing.Point(241, 92)
        Me.Label19.TabIndex = 42
        Me.Label19.Enabled = True
        Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label19.UseMnemonic = True
        Me.Label19.Visible = True
        Me.Label19.AutoSize = False
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label19.Name = "Label19"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label16.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
        Me.Label16.Text = "Presione <F5> o Doble Click Para Incluir una Partida en el Grid de Detalle de Ventas"
        Me.Label16.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.Label16.Size = New System.Drawing.Size(489, 21)
        Me.Label16.Location = New System.Drawing.Point(240, 318)
        Me.Label16.TabIndex = 41
        Me.Label16.Enabled = True
        Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label16.UseMnemonic = True
        Me.Label16.Visible = True
        Me.Label16.AutoSize = False
        Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label16.Name = "Label16"
        Me.lblDesc.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblDesc.BackColor = System.Drawing.SystemColors.Window
        Me.lblDesc.ForeColor = System.Drawing.Color.FromArgb(0, 128, 0)
        Me.lblDesc.Size = New System.Drawing.Size(721, 21)
        Me.lblDesc.Location = New System.Drawing.Point(8, 501)
        Me.lblDesc.TabIndex = 24
        Me.lblDesc.Enabled = True
        Me.lblDesc.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDesc.UseMnemonic = True
        Me.lblDesc.Visible = True
        Me.lblDesc.AutoSize = False
        Me.lblDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDesc.Name = "lblDesc"
        Me.lblDescripcion.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.lblDescripcion.ForeColor = System.Drawing.Color.FromArgb(0, 0, 192)
        Me.lblDescripcion.Size = New System.Drawing.Size(487, 21)
        Me.lblDescripcion.Location = New System.Drawing.Point(240, 287)
        Me.lblDescripcion.TabIndex = 23
        Me.lblDescripcion.Enabled = True
        Me.lblDescripcion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDescripcion.UseMnemonic = True
        Me.lblDescripcion.Visible = True
        Me.lblDescripcion.AutoSize = False
        Me.lblDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDescripcion.Name = "lblDescripcion"
        Me.lblEstadoFolio.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblEstadoFolio.BackColor = System.Drawing.SystemColors.Window
        Me.lblEstadoFolio.ForeColor = System.Drawing.Color.FromArgb(192, 0, 0)
        Me.lblEstadoFolio.Size = New System.Drawing.Size(109, 22)
        Me.lblEstadoFolio.Location = New System.Drawing.Point(8, 287)
        Me.lblEstadoFolio.TabIndex = 17
        Me.lblEstadoFolio.Enabled = True
        Me.lblEstadoFolio.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEstadoFolio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEstadoFolio.UseMnemonic = True
        Me.lblEstadoFolio.Visible = True
        Me.lblEstadoFolio.AutoSize = False
        Me.lblEstadoFolio.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEstadoFolio.Name = "lblEstadoFolio"
        Me.Label2.Text = "Sucursal :"
        Me.Label2.Size = New System.Drawing.Size(62, 15)
        Me.Label2.Location = New System.Drawing.Point(8, 40)
        Me.Label2.TabIndex = 11
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Enabled = True
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.UseMnemonic = True
        Me.Label2.Visible = True
        Me.Label2.AutoSize = False
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label2.Name = "Label2"
        Me.Label1.Text = "Fecha Venta :"
        Me.Label1.Size = New System.Drawing.Size(82, 21)
        Me.Label1.Location = New System.Drawing.Point(8, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Enabled = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.UseMnemonic = True
        Me.Label1.Visible = True
        Me.Label1.AutoSize = False
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label1.Name = "Label1"
        CType(Me.DtpHasta, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dbcSucursal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtpFechaVenta, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtpFechaRegistro, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.flexVentas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.flexDetalleVenta, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.flexVentasPendientes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DtpDesde, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Controls.Add(DtpDesde)
        Me.Controls.Add(Frame7)
        Me.Controls.Add(Frame6)
        Me.Controls.Add(Frame5)
        Me.Controls.Add(txtFlex)
        Me.Controls.Add(Frame4)
        Me.Controls.Add(txtDescripcion)
        Me.Controls.Add(Frame3)
        Me.Controls.Add(flexVentasPendientes)
        Me.Controls.Add(flexDetalleVenta)
        Me.Controls.Add(flexVentas)
        Me.Controls.Add(Frame2)
        Me.Controls.Add(Frame1)
        Me.Controls.Add(txtCodSucursal)
        Me.Controls.Add(dtpFechaVenta)
        Me.Controls.Add(dbcSucursal)
        Me.Controls.Add(DtpHasta)
        Me.Controls.Add(lblMoneda)
        Me.Controls.Add(Label21)
        Me.Controls.Add(Label20)
        Me.Controls.Add(Label17)
        Me.Controls.Add(lblSubTot)
        Me.Controls.Add(lblIva)
        Me.Controls.Add(lblDescuento)
        Me.Controls.Add(lblCantidad)
        Me.Controls.Add(Label19)
        Me.Controls.Add(Label16)
        Me.Controls.Add(lblDesc)
        Me.Controls.Add(lblDescripcion)
        Me.Controls.Add(lblEstadoFolio)
        Me.Controls.Add(Label2)
        Me.Controls.Add(Label1)
        Me.Frame7.Controls.Add(lblFacturar)
        Me.Frame7.Controls.Add(Label14)
        Me.Frame7.Controls.Add(lblExcluido)
        Me.Frame7.Controls.Add(Label18)
        Me.Frame6.Controls.Add(cmdDatosFiscales)
        Me.Frame6.Controls.Add(chkDoctoCliente)
        Me.Frame6.Controls.Add(chkDesglosarIva)
        Me.Frame6.Controls.Add(cmdGenerarFactura)
        Me.Frame6.Controls.Add(cmdImpresionTickets)
        Me.Frame6.Controls.Add(cmdImprimirFactura)
        Me.Frame5.Controls.Add(txtFacturaAdicional)
        Me.Frame5.Controls.Add(Label7)
        Me.Frame4.Controls.Add(Label6)
        Me.Frame4.Controls.Add(lblSubTotal)
        Me.Frame4.Controls.Add(Label8)
        Me.Frame4.Controls.Add(lblRedondeo)
        Me.Frame4.Controls.Add(Label10)
        Me.Frame4.Controls.Add(lblTotal)
        Me.Frame4.Controls.Add(Label12)
        Me.Frame4.Controls.Add(lblTotalPesos)
        Me.Frame3.Controls.Add(Label11)
        Me.Frame3.Controls.Add(lblImporteRedondeo)
        Me.Frame3.Controls.Add(lblImporteTotal)
        Me.Frame3.Controls.Add(Label15)
        Me.Frame3.Controls.Add(lblImporteSubTotal)
        Me.Frame3.Controls.Add(Label13)
        Me.Frame3.Controls.Add(lblFactura)
        Me.Frame3.Controls.Add(Label9)
        Me.Frame2.Controls.Add(txtFolioFactura)
        Me.Frame2.Controls.Add(dtpFechaRegistro)
        Me.Frame2.Controls.Add(Label5)
        Me.Frame2.Controls.Add(Label4)
        Me.Frame1.Controls.Add(txtPorcentaje)
        Me.Frame1.Controls.Add(optPorcentual)
        Me.Frame1.Controls.Add(optManual)
        Me.Frame1.Controls.Add(Label3)
        Me.Frame7.ResumeLayout(False)
        Me.Frame6.ResumeLayout(False)
        Me.Frame5.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()


    End Sub
End Class
