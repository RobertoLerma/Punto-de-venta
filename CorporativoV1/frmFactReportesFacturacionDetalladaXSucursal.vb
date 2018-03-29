Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmFactReportesFacturacionDetalladaXSucursal
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE FACTURACION DETALLADA POR SUCURSAL                                                *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :                                                                                                   *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents chkV As System.Windows.Forms.CheckBox
    Public WithEvents txtTextoAdicional As System.Windows.Forms.TextBox
    Public WithEvents chkTodaslasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim intCodSucursal As Integer
    Dim tecla As Integer
    Dim rsReporte As ADODB.Recordset
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo




    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtTextoAdicional = New System.Windows.Forms.TextBox()
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.chkV = New System.Windows.Forms.CheckBox()
        Me.chkTodaslasSucursales = New System.Windows.Forms.CheckBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me._Label2_0 = New System.Windows.Forms.Label()
        Me._Label2_1 = New System.Windows.Forms.Label()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtTextoAdicional
        '
        Me.txtTextoAdicional.AcceptsReturn = True
        Me.txtTextoAdicional.BackColor = System.Drawing.SystemColors.Window
        Me.txtTextoAdicional.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTextoAdicional.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTextoAdicional.Location = New System.Drawing.Point(14, 207)
        Me.txtTextoAdicional.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTextoAdicional.MaxLength = 120
        Me.txtTextoAdicional.Multiline = True
        Me.txtTextoAdicional.Name = "txtTextoAdicional"
        Me.txtTextoAdicional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTextoAdicional.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtTextoAdicional.Size = New System.Drawing.Size(353, 77)
        Me.txtTextoAdicional.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.txtTextoAdicional, "Texto Adicional.")
        '
        'txtCodSucursal
        '
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.Enabled = False
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.Location = New System.Drawing.Point(68, 33)
        Me.txtCodSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucursal.MaxLength = 3
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(55, 20)
        Me.txtCodSucursal.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCodSucursal, "Codgo de Sucursal.")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optDolares)
        Me.Frame2.Controls.Add(Me.optPesos)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(15, 119)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(352, 40)
        Me.Frame2.TabIndex = 14
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Moneda"
        '
        'optDolares
        '
        Me.optDolares.BackColor = System.Drawing.SystemColors.Control
        Me.optDolares.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDolares.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDolares.Location = New System.Drawing.Point(150, 15)
        Me.optDolares.Margin = New System.Windows.Forms.Padding(2)
        Me.optDolares.Name = "optDolares"
        Me.optDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDolares.Size = New System.Drawing.Size(67, 17)
        Me.optDolares.TabIndex = 6
        Me.optDolares.TabStop = True
        Me.optDolares.Text = "Dolares"
        Me.optDolares.UseVisualStyleBackColor = False
        '
        'optPesos
        '
        Me.optPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPesos.Location = New System.Drawing.Point(54, 15)
        Me.optPesos.Margin = New System.Windows.Forms.Padding(2)
        Me.optPesos.Name = "optPesos"
        Me.optPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPesos.Size = New System.Drawing.Size(67, 17)
        Me.optPesos.TabIndex = 5
        Me.optPesos.TabStop = True
        Me.optPesos.Text = "Pesos"
        Me.optPesos.UseVisualStyleBackColor = False
        '
        'chkV
        '
        Me.chkV.BackColor = System.Drawing.SystemColors.Control
        Me.chkV.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkV.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkV.Location = New System.Drawing.Point(349, 172)
        Me.chkV.Margin = New System.Windows.Forms.Padding(2)
        Me.chkV.Name = "chkV"
        Me.chkV.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkV.Size = New System.Drawing.Size(18, 20)
        Me.chkV.TabIndex = 13
        Me.chkV.TabStop = False
        Me.chkV.Text = "Check1"
        Me.chkV.UseVisualStyleBackColor = False
        '
        'chkTodaslasSucursales
        '
        Me.chkTodaslasSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodaslasSucursales.Checked = True
        Me.chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodaslasSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodaslasSucursales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodaslasSucursales.Location = New System.Drawing.Point(12, 13)
        Me.chkTodaslasSucursales.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodaslasSucursales.Name = "chkTodaslasSucursales"
        Me.chkTodaslasSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodaslasSucursales.Size = New System.Drawing.Size(136, 17)
        Me.chkTodaslasSucursales.TabIndex = 0
        Me.chkTodaslasSucursales.Text = "Todas las Sucursales"
        Me.chkTodaslasSucursales.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFechaInicial)
        Me.Frame1.Controls.Add(Me.dtpFechaFinal)
        Me.Frame1.Controls.Add(Me._Label2_0)
        Me.Frame1.Controls.Add(Me._Label2_1)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(12, 54)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(355, 46)
        Me.Frame1.TabIndex = 7
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(74, 17)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(96, 20)
        Me.dtpFechaInicial.TabIndex = 3
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(240, 17)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(95, 20)
        Me.dtpFechaFinal.TabIndex = 4
        '
        '_Label2_0
        '
        Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_0.Location = New System.Drawing.Point(20, 20)
        Me._Label2_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_0.Size = New System.Drawing.Size(59, 17)
        Me._Label2_0.TabIndex = 9
        Me._Label2_0.Text = "Desde el :"
        '
        '_Label2_1
        '
        Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_1.Location = New System.Drawing.Point(183, 20)
        Me._Label2_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_1.Size = New System.Drawing.Size(53, 17)
        Me._Label2_1.TabIndex = 8
        Me._Label2_1.Text = "Hasta el :"
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(127, 32)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(179, 21)
        Me.dbcSucursal.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(11, 191)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(91, 14)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Texto Adicional"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 36)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(61, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Sucursal :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(129, 298)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 117
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(14, 298)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 116
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmFactReportesFacturacionDetalladaXSucursal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(380, 344)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.chkV)
        Me.Controls.Add(Me.txtTextoAdicional)
        Me.Controls.Add(Me.chkTodaslasSucursales)
        Me.Controls.Add(Me.txtCodSucursal)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(368, 287)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmFactReportesFacturacionDetalladaXSucursal"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Facturación Detallada por Sucursal"
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub


    '''Modificación.-  Agregar Facturación Especial
    Sub Imprime()
        Dim RptFactFacturacionDetalladaXSucursal As New RptFactFacturacionDetalladaXSucursal

        Dim sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim PeriodoReporte As String
        Dim TextoAdicional As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim Moneda As String
        Dim TextoMoneda As String
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

        NombreEmpresa = UCase(gstrCorpoNOMBREEMPRESA)
        NombreReporte = UCase("Facturacion Detallada por Sucursal")
        Dim fechaInicial1 As String = AgregarHoraAFecha(dtpFechaInicial.Value)
        Dim fechaFinal2 As String = AgregarHoraAFecha(dtpFechaFinal.Value)
        PeriodoReporte = "Del " & fechaInicial1 & " al " & fechaFinal2
        'PeriodoReporte = "Del " & Format(dtpFechaInicial.Value, "dd/MMM/yyyy") & " al " & Format(dtpFechaFinal.Value, "dd/MMM/yyyy")
        txtTextoAdicional.Text = ModEstandar.QuitaEnter(txtTextoAdicional.Text)
        TextoAdicional = txtTextoAdicional.Text
        'FechaInicial = Format(Month(dtpFechaInicial.Value), "00") & "/" & Format((dtpFechaInicial.Value), "00") & "/" & Format(Year(dtpFechaInicial.Value), "0000")
        'FechaFinal = Format(Month(dtpFechaFinal.Value), "00") & "/" & Format((dtpFechaFinal.Value), "00") & "/" & Format(Year(dtpFechaFinal.Value), "0000")
        FechaInicial = AgregarHoraAFecha(dtpFechaInicial.Value)
        FechaFinal = AgregarHoraAFecha(dtpFechaFinal.Value)

        Moneda = IIf(optPesos.Checked = True, "P", "D")

        If chkTodaslasSucursales.CheckState = 1 And Moneda = "P" Then
            sql = "(select suc.descalmacen as sucursal,f.foliofactura as foliofactura,f.fechafactura as fecha,(det.cantidadadicional) as cantidad," & "case when f.moneda = 'P' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) * f.tipocambio,1) end as importe,case when f.moneda = 'P' then round(f.descuento,2) else round(f.descuento * f.tipocambio,1) end as descuento," & "case when f.moneda = 'P' then round(((f.subtotal + f.redondeo) - f.descuento),2) else round(((f.subtotal + f.redondeo) - f.descuento) * f.tipocambio,1) end as subtotal," & "case when f.moneda = 'P' then round(f.iva,2) else round(f.iva * f.tipocambio,1) end as iva,case when f.moneda = 'P' then round(f.total + f.redondeo,2) else round((f.total + f.redondeo) * f.tipocambio,1) end as total " & "from facturas f inner join movimientosventascab cab on f.foliofactura = cab.foliofactura Inner Join (select foliofactura,sum(cantidadadicional) as cantidadadicional from movimientosventasdet group by foliofactura) det on f.foliofactura = det.foliofactura and cab.foliofactura = det.foliofactura " & "inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen and cab.codsucursal = suc.codalmacen " & "where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' and cab.estatus <> 'C' and f.tipofactura = 'N' group by suc.descalmacen,f.foliofactura,f.fechafactura,det.cantidadadicional,f.subtotal,f.redondeo,f.descuento,f.iva,f.total,f.moneda,f.tipocambio) " & "Union " & "(Select DescAlmacen AS Sucursal,FolioFactura as Factura, FechaFactura as Fecha, sum(Cantidad) as Cantidad,sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal," & "sum(Iva) as Iva, sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen,sum(F.Cantidad) as Cantidad,case when f.moneda = 'P' then round(F.SubTotal,2) else round(f.subtotal * f.tipocambio,1) end as subtotal," & "case when f.moneda = 'P' then round(F.Descuento,2) else round(f.descuento * f.tipocambio,1) end as descuento,case when f.moneda = 'P' then round(F.Iva,2) else round(f.iva * f.tipocambio,1) end as iva," & "case when f.moneda = 'P' then round(F.Total,2) else round(f.total * f.tipocambio,1) end as total,case when f.moneda = 'P' then round(F.Redondeo,2) else round(f.redondeo * f.tipocambio,1) end as redondeo," & "sum(case when f.moneda = 'P' then round(F.Importe,2) else round(f.importe * f.tipocambio,1) end) as Importe " & "From Facturas F Inner Join CatAlmacen A On F.CodSucursal = A.CodAlmacen " & "Where F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' and f.estatus <> 'C' " & "Group By F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total," & "F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group By DescAlmacen, FolioFactura, FechaFactura)"

            '        sql = "(SELECT SucFac.DescAlmacen AS Sucursal,Cab.FolioFactura AS Factura,SucFac.FechaFactura AS Fecha,SUM(Det.Cantidad) AS Cantidad,SUM(round((Cab.SubTotalAdicional + Cab.RedondeoAdicional) * sucfac.tipocambio,1)) AS Importe,SUM(round(Cab.DescuentoAdicional * sucfac.tipocambio,1)) AS Descuento,SUM(round(((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional) * sucfac.tipocambio,1)) AS SubTotal,SUM(round(Cab.IvaAdicional * sucfac.tipocambio,1)) AS Iva,SUM(round((Cab.TotalAdicional + Cab.RedondeoAdicional) * sucfac.tipocambio,1)) As Total FROM (SELECT CatAlmacen.CodAlmacen,CatAlmacen.DescAlmacen,Facturas.FechaFactura,Facturas.moneda,facturas.tipocambio,facturas.foliofactura FROM CatAlmacen," & _
            ''        "Facturas WHERE CatAlmacen.TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen,FechaFactura,Facturas.moneda,facturas.tipocambio,facturas.foliofactura) SucFac INNER JOIN MovimientosVentasCab Cab on Cab.FechaVenta = SucFac.FechaFactura AND SucFac.CodAlmacen = Cab.CodSucursal and sucfac.foliofactura = cab.foliofactura INNER JOIN (SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM MovimientosVentasDet GROUP BY FolioVenta) Det ON Cab.FolioVenta = Det.FolioVenta Where Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' " & _
            ''        "AND Cab.Estatus <> 'C' GROUP BY SucFac.DescAlmacen,Cab.FolioFactura,SucFac.FechaFactura) UNION (SELECT SucFac.DescAlmacen AS Sucursal,Det.FolioFactura AS Factura,Det.FechaFactura AS Fecha, SUM(Det.CantidadAdicional) AS Cantidad,SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe, SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal, " & _
            ''        "SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total FROM (SELECT DISTINCT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen) SucFac INNER JOIN (SELECT Det.FolioFactura,ISNULL(Det.CodSucursalAdicional,0) AS CodSucursal, SUM(Det.CantidadAdicional) AS CantidadAdicional, SUM(round((Det.PrecioListaSinIvaAdicional * Det.CantidadAdicional) * f.tipocambio,1)) AS SubTotalAdicional,round(Det.RedondeoAdicional * f.tipocambio,1) as redondeoadicional, SUM(round(((Det.ImptePromocionesAdicional + " & _
            ''        "Det.ImpteDescuentosAdicional) * Det.CantidadAdicional) * f.tipocambio,1)) AS DescuentoAdicional, SUM(round((Det.IvaRealAdicional * Det.CantidadAdicional) * f.tipocambio,1)) AS IvaAdicional,SUM(round((Det.PrecioRealAdicional * Det.CantidadAdicional) * f.tipocambio,1)) AS TotalAdicional, F.FechaFactura FROM MovimientosVentasDet Det INNER JOIN Facturas F ON Det.FolioFactura = F.FolioFactura WHERE Det.FolioAdicional <> '' AND Det.FolioFactura <> '' AND Det.EstatusAdicional <> 'O' GROUP BY Det.FolioFactura,Det.CodSucursalAdicional,Det.RedondeoAdicional,F.FechaFactura,f.tipocambio) " & _
            ''        "Det ON SucFac.CodAlmacen = Det.CodSucursal WHERE Det.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' GROUP BY SucFac.DescAlmacen,Det.FolioFactura,Det.FechaFactura) UNION  (Select DescAlmacen AS Sucursal, FolioFactura as Factura, FechaFactura as Fecha, sum(Cantidad) as Cantidad, sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal, sum(Iva) as Iva, " & _
            ''        "sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen, sum(F.Cantidad) as Cantidad,case when f.moneda = 'P' then round(F.SubTotal,2) else round(f.subtotal * f.tipocambio,1) end as subtotal,case when f.moneda = 'P' then round(F.Descuento,2) else round(f.descuento * f.tipocambio,1) end as descuento,case when f.moneda = 'P' then round(F.Iva,2) else round(f.iva * f.tipocambio,1) end as iva,case when f.moneda = 'P' then round(F.Total,2) else round(f.total * f.tipocambio,1) end as total,case when f.moneda = 'P' then round(F.Redondeo,2) else round(f.redondeo * f.tipocambio,1) end as redondeo,sum(case when f.moneda = 'P' then round(F.Importe,2) else round(f.importe * f.tipocambio,1) end) as Importe From Facturas F Inner Join CatAlmacen A On F.CodSucursal = A.CodAlmacen Where F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' Group By F.FolioFactura, F.FechaFactura, F.CodSucursal, " & _
            ''        "A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total, F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group By DescAlmacen, FolioFactura, FechaFactura ) "

            '''        sql = "(SELECT SucFac.DescAlmacen AS Sucursal,Cab.FolioFactura AS Factura,SucFac.FechaFactura AS Fecha,SUM(Det.Cantidad) AS Cantidad,SUM(Cab.SubTotalAdicional + Cab.RedondeoAdicional) AS Importe,SUM(Cab.DescuentoAdicional) AS Descuento,SUM((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional) AS SubTotal,SUM(Cab.IvaAdicional) AS Iva,SUM(Cab.TotalAdicional + Cab.RedondeoAdicional) As Total FROM " & _
            ''''        "(SELECT CatAlmacen.CodAlmacen,CatAlmacen.DescAlmacen,Facturas.FechaFactura FROM CatAlmacen,Facturas WHERE CatAlmacen.TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen,FechaFactura) SucFac INNER JOIN MovimientosVentasCab Cab on Cab.FechaVenta = SucFac.FechaFactura AND SucFac.CodAlmacen = Cab.CodSucursal INNER JOIN (SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM MovimientosVentasDet GROUP BY FolioVenta) Det " & _
            ''''        "ON Cab.FolioVenta = Det.FolioVenta Where Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' AND Cab.Estatus <> 'C' GROUP BY SucFac.DescAlmacen,Cab.FolioFactura,SucFac.FechaFactura) UNION (SELECT SucFac.DescAlmacen AS Sucursal,Det.FolioFactura AS Factura,Det.FechaFactura AS Fecha,SUM(Det.CantidadAdicional) AS Cantidad,SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe," & _
            ''''        "SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal,SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total FROM (SELECT DISTINCT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen) SucFac INNER JOIN (SELECT Det.FolioFactura,ISNULL(Det.CodSucursalAdicional,0) AS CodSucursal,SUM(Det.CantidadAdicional) AS CantidadAdicional," & _
            ''''        "SUM(Det.PrecioListaSinIvaAdicional * Det.CantidadAdicional) AS SubTotalAdicional,Det.RedondeoAdicional,SUM((Det.ImptePromocionesAdicional + Det.ImpteDescuentosAdicional) * Det.CantidadAdicional) AS DescuentoAdicional,SUM(Det.IvaRealAdicional * Det.CantidadAdicional) AS IvaAdicional,SUM(Det.PrecioRealAdicional * Det.CantidadAdicional) AS TotalAdicional,F.FechaFactura FROM MovimientosVentasDet Det INNER JOIN Facturas F ON Det.FolioFactura = F.FolioFactura WHERE Det.FolioAdicional <> '' AND Det.FolioFactura <> '' AND Det.EstatusAdicional <> 'O' " & _
            ''''        "GROUP BY Det.FolioFactura,Det.CodSucursalAdicional,Det.RedondeoAdicional,F.FechaFactura) Det ON SucFac.CodAlmacen = Det.CodSucursal WHERE Det.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' GROUP BY SucFac.DescAlmacen,Det.FolioFactura,Det.FechaFactura)"
            TextoMoneda = "Los importes estan expresados en pesos"
        ElseIf chkTodaslasSucursales.CheckState = 1 And Moneda = "D" Then

            sql = "(select suc.descalmacen as sucursal,f.foliofactura as foliofactura,f.fechafactura as fecha,(det.cantidadadicional) as cantidad," & "case when f.moneda = 'D' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) / f.tipocambio,2) end as importe,case when f.moneda = 'D' then round(f.descuento,2) else round(f.descuento / f.tipocambio,2) end as descuento," & "case when f.moneda = 'D' then round(((f.subtotal + f.redondeo) - f.descuento),2) else round(((f.subtotal + f.redondeo) - f.descuento) / f.tipocambio,2) end as subtotal," & "case when f.moneda = 'D' then round(f.iva,2) else round(f.iva / f.tipocambio,2) end as iva,case when f.moneda = 'D' then round(f.total + f.redondeo,2) else round((f.total + f.redondeo) / f.tipocambio,2) end as total " & "from facturas f inner join movimientosventascab cab on f.foliofactura = cab.foliofactura Inner Join (select foliofactura,sum(cantidadadicional) as cantidadadicional from movimientosventasdet group by foliofactura) det on f.foliofactura = det.foliofactura and cab.foliofactura = det.foliofactura " & "inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen and cab.codsucursal = suc.codalmacen " & "where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' and cab.estatus <> 'C' and f.tipofactura = 'N' group by suc.descalmacen,f.foliofactura,f.fechafactura,det.cantidadadicional,f.subtotal,f.redondeo,f.descuento,f.iva,f.total,f.moneda,f.tipocambio) " & "Union " & "(Select DescAlmacen AS Sucursal,FolioFactura as Factura, FechaFactura as Fecha, sum(Cantidad) as Cantidad,sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal," & "sum(Iva) as Iva, sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen,sum(F.Cantidad) as Cantidad,case when f.moneda = 'D' then round(F.SubTotal,2) else round(f.subtotal / f.tipocambio,2) end as subtotal," & "case when f.moneda = 'D' then round(F.Descuento,2) else round(f.descuento / f.tipocambio,2) end as descuento,case when f.moneda = 'D' then round(F.Iva,2) else round(f.iva / f.tipocambio,2) end as iva," & "case when f.moneda = 'D' then round(F.Total,2) else round(f.total / f.tipocambio,2) end as total,case when f.moneda = 'D' then round(F.Redondeo,2) else round(f.redondeo / f.tipocambio,2) end as redondeo," & "sum(case when f.moneda = 'D' then round(F.Importe,2) else round(f.importe / f.tipocambio,2) end) as Importe " & "From Facturas F Inner Join CatAlmacen A On F.CodSucursal = A.CodAlmacen " & "Where F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' and f.estatus <> 'C' " & "Group By F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total," & "F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group By DescAlmacen, FolioFactura, FechaFactura)"

            '        sql = "(SELECT SucFac.DescAlmacen AS Sucursal,Cab.FolioFactura AS Factura,SucFac.FechaFactura AS Fecha,SUM(Det.Cantidad) AS Cantidad,SUM((Cab.SubTotalAdicional + Cab.RedondeoAdicional)) AS Importe,SUM(Cab.DescuentoAdicional) AS Descuento,SUM(((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional)) AS SubTotal,SUM(Cab.IvaAdicional) AS Iva,SUM((Cab.TotalAdicional + Cab.RedondeoAdicional)) As Total FROM (SELECT CatAlmacen.CodAlmacen,CatAlmacen.DescAlmacen,Facturas.FechaFactura,Facturas.moneda,facturas.tipocambio FROM CatAlmacen," & _
            ''        "Facturas WHERE CatAlmacen.TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen,FechaFactura,Facturas.moneda,facturas.tipocambio) SucFac INNER JOIN MovimientosVentasCab Cab on Cab.FechaVenta = SucFac.FechaFactura AND SucFac.CodAlmacen = Cab.CodSucursal INNER JOIN (SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM MovimientosVentasDet GROUP BY FolioVenta) Det ON Cab.FolioVenta = Det.FolioVenta Where Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' " & _
            ''        "AND Cab.Estatus <> 'C' GROUP BY SucFac.DescAlmacen,Cab.FolioFactura,SucFac.FechaFactura) UNION (SELECT SucFac.DescAlmacen AS Sucursal,Det.FolioFactura AS Factura,Det.FechaFactura AS Fecha, SUM(Det.CantidadAdicional) AS Cantidad,SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe, SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal, " & _
            ''        "SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total FROM (SELECT DISTINCT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen) SucFac INNER JOIN (SELECT Det.FolioFactura,ISNULL(Det.CodSucursalAdicional,0) AS CodSucursal, SUM(Det.CantidadAdicional) AS CantidadAdicional, SUM((Det.PrecioListaSinIvaAdicional * Det.CantidadAdicional)) AS SubTotalAdicional,Det.RedondeoAdicional,SUM(((Det.ImptePromocionesAdicional + " & _
            ''        "Det.ImpteDescuentosAdicional) * Det.CantidadAdicional)) AS DescuentoAdicional, SUM((Det.IvaRealAdicional * Det.CantidadAdicional)) AS IvaAdicional,SUM((Det.PrecioRealAdicional * Det.CantidadAdicional)) AS TotalAdicional, F.FechaFactura FROM MovimientosVentasDet Det INNER JOIN Facturas F ON Det.FolioFactura = F.FolioFactura WHERE Det.FolioAdicional <> '' AND Det.FolioFactura <> '' AND Det.EstatusAdicional <> 'O' GROUP BY Det.FolioFactura,Det.CodSucursalAdicional,Det.RedondeoAdicional,F.FechaFactura,f.tipocambio) " & _
            ''        "Det ON SucFac.CodAlmacen = Det.CodSucursal WHERE Det.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' GROUP BY SucFac.DescAlmacen,Det.FolioFactura,Det.FechaFactura) UNION  (Select DescAlmacen AS Sucursal, FolioFactura as Factura, FechaFactura as Fecha, sum(Cantidad) as Cantidad, sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal, sum(Iva) as Iva, " & _
            ''        "sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen, sum(F.Cantidad) as Cantidad,case when f.moneda = 'D' then F.SubTotal else round(f.subtotal / f.tipocambio,2) end as subtotal,case when f.moneda = 'D' then F.Descuento else round(f.descuento / f.tipocambio,2) end as descuento,case when f.moneda = 'D' then F.Iva else round(f.iva / f.tipocambio,2) end as iva,case when f.moneda = 'D' then F.Total else round(f.total / f.tipocambio,2) end as total,case when f.moneda = 'D' then F.Redondeo else round(f.redondeo / f.tipocambio,2) end as redondeo,sum(case when f.moneda = 'D' then F.Importe else round(f.importe / f.tipocambio,2) end) as Importe From Facturas F Inner Join CatAlmacen A On F.CodSucursal = A.CodAlmacen Where F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' Group By F.FolioFactura, F.FechaFactura, F.CodSucursal, " & _
            ''        "A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total, F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group By DescAlmacen, FolioFactura, FechaFactura ) "
            TextoMoneda = "Los importes estan expresados en dólares"
        ElseIf chkTodaslasSucursales.CheckState = 0 And Moneda = "P" Then
            If CShort(Numerico(txtCodSucursal.Text)) = 0 Then
                MsgBox("Proporcione el Codigo de la Sucursal ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtCodSucursal.Focus()
                Exit Sub
            End If

            sql = "(select suc.descalmacen as sucursal,f.foliofactura as foliofactura,f.fechafactura as fecha,(det.cantidadadicional) as cantidad," & "case when f.moneda = 'P' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) * f.tipocambio,1) end as importe,case when f.moneda = 'P' then round(f.descuento,2) else round(f.descuento * f.tipocambio,1) end as descuento," & "case when f.moneda = 'P' then round(((f.subtotal + f.redondeo) - f.descuento),2) else round(((f.subtotal + f.redondeo) - f.descuento) * f.tipocambio,1) end as subtotal," & "case when f.moneda = 'P' then round(f.iva,2) else round(f.iva * f.tipocambio,1) end as iva,case when f.moneda = 'P' then round(f.total + f.redondeo,2) else round((f.total + f.redondeo) * f.tipocambio,1) end as total " & "from facturas f inner join movimientosventascab cab on f.foliofactura = cab.foliofactura Inner Join (select foliofactura,sum(cantidadadicional) as cantidadadicional from movimientosventasdet group by foliofactura) det on f.foliofactura = det.foliofactura and cab.foliofactura = det.foliofactura " & "inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen and cab.codsucursal = suc.codalmacen " & "where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' and cab.estatus <> 'C' and f.tipofactura = 'N' and suc.codalmacen = " & Numerico(txtCodSucursal.Text) & "  group by suc.descalmacen,f.foliofactura,f.fechafactura,det.cantidadadicional,f.subtotal,f.redondeo,f.descuento,f.iva,f.total,f.moneda,f.tipocambio) " & "Union " & "(Select DescAlmacen AS Sucursal,FolioFactura as Factura, FechaFactura as Fecha, sum(Cantidad) as Cantidad,sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal," & "sum(Iva) as Iva, sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen,sum(F.Cantidad) as Cantidad,case when f.moneda = 'P' then round(F.SubTotal,2) else round(f.subtotal * f.tipocambio,1) end as subtotal," & "case when f.moneda = 'P' then round(F.Descuento,2) else round(f.descuento * f.tipocambio,1) end as descuento,case when f.moneda = 'P' then round(F.Iva,2) else round(f.iva * f.tipocambio,1) end as iva," & "case when f.moneda = 'P' then round(F.Total,2) else round(f.total * f.tipocambio,1) end as total,case when f.moneda = 'P' then round(F.Redondeo,2) else round(f.redondeo * f.tipocambio,1) end as redondeo," & "sum(case when f.moneda = 'P' then round(F.Importe,2) else round(f.importe * f.tipocambio,1) end) as Importe " & "From Facturas F Inner Join CatAlmacen A On F.CodSucursal = A.CodAlmacen " & "Where F.codsucursal = " & Numerico(txtCodSucursal.Text) & " and F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' and f.estatus <> 'C' " & "Group By F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total," & "F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group By DescAlmacen, FolioFactura, FechaFactura)"

            '        sql = "(SELECT SucFac.DescAlmacen AS Sucursal,Cab.FolioFactura AS Factura,SucFac.FechaFactura AS Fecha,SUM(Det.Cantidad) AS Cantidad,SUM(round((Cab.SubTotalAdicional + Cab.RedondeoAdicional) * sucfac.tipocambio,1)) AS Importe,SUM(round(Cab.DescuentoAdicional * sucfac.tipocambio,1)) AS Descuento,SUM(round(((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional) * sucfac.tipocambio,1)) AS SubTotal,SUM(round(Cab.IvaAdicional * sucfac.tipocambio,1)) AS Iva,SUM(round((Cab.TotalAdicional + Cab.RedondeoAdicional) * sucfac.tipocambio,1)) As Total FROM (SELECT CatAlmacen.CodAlmacen,CatAlmacen.DescAlmacen,Facturas.FechaFactura,facturas.tipocambio FROM CatAlmacen," & _
            ''        "Facturas WHERE CatAlmacen.TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen,FechaFactura,facturas.tipocambio) SucFac INNER JOIN MovimientosVentasCab Cab on Cab.FechaVenta = SucFac.FechaFactura AND SucFac.CodAlmacen = Cab.CodSucursal INNER JOIN (SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM MovimientosVentasDet GROUP BY FolioVenta) Det ON Cab.FolioVenta = Det.FolioVenta Where Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' " & _
            ''        "AND Cab.Estatus <> 'C' AND SucFac.CodAlmacen = " & CInt(Numerico(txtCodSucursal)) & " GROUP BY SucFac.DescAlmacen,Cab.FolioFactura,SucFac.FechaFactura) UNION (SELECT SucFac.DescAlmacen AS Sucursal,Det.FolioFactura AS Factura,Det.FechaFactura AS Fecha, SUM(Det.CantidadAdicional) AS Cantidad,SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe, SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal, " & _
            ''        "SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total FROM (SELECT DISTINCT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen) SucFac INNER JOIN (SELECT Det.FolioFactura,ISNULL(Det.CodSucursalAdicional,0) AS CodSucursal, SUM(Det.CantidadAdicional) AS CantidadAdicional, SUM(round((Det.PrecioListaSinIvaAdicional * Det.CantidadAdicional) * f.tipocambio,1)) AS SubTotalAdicional,round(Det.RedondeoAdicional * f.tipocambio,1) as redondeoadicional,SUM(round(((Det.ImptePromocionesAdicional + " & _
            ''        "Det.ImpteDescuentosAdicional) * Det.CantidadAdicional) * f.tipocambio,1)) AS DescuentoAdicional, SUM(round((Det.IvaRealAdicional * Det.CantidadAdicional) * f.tipocambio,1)) AS IvaAdicional,SUM(round((Det.PrecioRealAdicional * Det.CantidadAdicional) * f.tipocambio,1)) AS TotalAdicional, F.FechaFactura FROM MovimientosVentasDet Det INNER JOIN Facturas F ON Det.FolioFactura = F.FolioFactura WHERE Det.FolioAdicional <> '' AND Det.FolioFactura <> '' AND Det.EstatusAdicional <> 'O' GROUP BY Det.FolioFactura,Det.CodSucursalAdicional,Det.RedondeoAdicional,F.FechaFactura,f.tipocambio) " & _
            ''        "Det ON SucFac.CodAlmacen = Det.CodSucursal WHERE Det.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND SucFac.CodAlmacen = " & CInt(Numerico(txtCodSucursal)) & " GROUP BY SucFac.DescAlmacen,Det.FolioFactura,Det.FechaFactura) UNION  (Select DescAlmacen AS Sucursal, FolioFactura as Factura, FechaFactura as Fecha, sum(Cantidad) as Cantidad, sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal, sum(Iva) as Iva, " & _
            ''        "sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen, sum(F.Cantidad) as Cantidad,case when f.moneda = 'P' then F.SubTotal else round(f.subtotal * f.tipocambio,1) end as subtotal,case when f.moneda = 'P' then F.Descuento else round(f.descuento * f.tipocambio,1) end as descuento,case when f.moneda = 'P' then F.Iva else round(f.iva * f.tipocambio,1) end as iva,case when f.moneda = 'P' then F.Total else round(f.total * f.tipocambio,1) end as total,case when f.moneda = 'P' then F.Redondeo else round(f.redondeo * f.tipocambio,1) end as redondeo,sum(case when f.moneda = 'P' then F.Importe else round(f.importe * f.tipocambio,1) end) as Importe From Facturas F Inner Join CatAlmacen A On F.CodSucursal = A.CodAlmacen Where F.CodSucursal = " & CInt(Numerico(txtCodSucursal)) & " And F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' Group By F.FolioFactura, F.FechaFactura, F.CodSucursal, " & _
            ''        "A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total, F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group By DescAlmacen, FolioFactura, FechaFactura ) "

            '''        sql = "(SELECT SucFac.DescAlmacen AS Sucursal,Cab.FolioFactura AS Factura,SucFac.FechaFactura AS Fecha,SUM(Det.Cantidad) AS Cantidad,SUM(Cab.SubTotalAdicional + Cab.RedondeoAdicional) AS Importe,SUM(Cab.DescuentoAdicional) AS Descuento,SUM((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional) AS SubTotal,SUM(Cab.IvaAdicional) AS Iva,SUM(Cab.TotalAdicional + Cab.RedondeoAdicional) As Total FROM " & _
            ''''        "(SELECT CatAlmacen.CodAlmacen,CatAlmacen.DescAlmacen,Facturas.FechaFactura FROM CatAlmacen,Facturas WHERE CatAlmacen.TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen,FechaFactura) SucFac INNER JOIN MovimientosVentasCab Cab on Cab.FechaVenta = SucFac.FechaFactura AND SucFac.CodAlmacen = Cab.CodSucursal INNER JOIN (SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM MovimientosVentasDet GROUP BY FolioVenta) Det " & _
            ''''        "ON Cab.FolioVenta = Det.FolioVenta Where Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' AND Cab.Estatus <> 'C' AND SucFac.CodAlmacen = " & Numerico(txtCodSucursal) & " GROUP BY SucFac.DescAlmacen,Cab.FolioFactura,SucFac.FechaFactura) UNION (SELECT SucFac.DescAlmacen AS Sucursal,Det.FolioFactura AS Factura,Det.FechaFactura AS Fecha,SUM(Det.CantidadAdicional) AS Cantidad,SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe," & _
            ''''        "SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal,SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total FROM (SELECT DISTINCT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen) SucFac INNER JOIN (SELECT Det.FolioFactura,ISNULL(Det.CodSucursalAdicional,0) AS CodSucursal,SUM(Det.CantidadAdicional) AS CantidadAdicional," & _
            ''''        "SUM(Det.PrecioListaSinIvaAdicional * Det.CantidadAdicional) AS SubTotalAdicional,Det.RedondeoAdicional,SUM((Det.ImptePromocionesAdicional + Det.ImpteDescuentosAdicional) * Det.CantidadAdicional) AS DescuentoAdicional,SUM(Det.IvaRealAdicional * Det.CantidadAdicional) AS IvaAdicional,SUM(Det.PrecioRealAdicional * Det.CantidadAdicional) AS TotalAdicional,F.FechaFactura FROM MovimientosVentasDet Det INNER JOIN Facturas F ON Det.FolioFactura = F.FolioFactura WHERE Det.FolioAdicional <> '' AND Det.FolioFactura <> '' AND Det.EstatusAdicional <> 'O' " & _
            ''''        "GROUP BY Det.FolioFactura,Det.CodSucursalAdicional,Det.RedondeoAdicional,F.FechaFactura) Det ON SucFac.CodAlmacen = Det.CodSucursal WHERE Det.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND SucFac.CodAlmacen = " & Numerico(txtCodSucursal) & " GROUP BY SucFac.DescAlmacen,Det.FolioFactura,Det.FechaFactura)"
            TextoMoneda = "Los importes estan expresados en pesos"
        ElseIf chkTodaslasSucursales.CheckState = 0 And Moneda = "D" Then
            If CShort(Numerico(txtCodSucursal.Text)) = 0 Then
                MsgBox("Proporcione el Codigo de la Sucursal ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtCodSucursal.Focus()
                Exit Sub
            End If

            sql = "(select suc.descalmacen as sucursal,f.foliofactura as foliofactura,f.fechafactura as fecha,(det.cantidadadicional) as cantidad," & "case when f.moneda = 'D' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) / f.tipocambio,2) end as importe,case when f.moneda = 'D' then round(f.descuento,2) else round(f.descuento / f.tipocambio,2) end as descuento," & "case when f.moneda = 'D' then round(((f.subtotal + f.redondeo) - f.descuento),2) else round(((f.subtotal + f.redondeo) - f.descuento) / f.tipocambio,2) end as subtotal," & "case when f.moneda = 'D' then round(f.iva,2) else round(f.iva / f.tipocambio,2) end as iva,case when f.moneda = 'D' then round(f.total + f.redondeo,2) else round((f.total + f.redondeo) / f.tipocambio,2) end as total " & "from facturas f inner join movimientosventascab cab on f.foliofactura = cab.foliofactura Inner Join (select foliofactura,sum(cantidadadicional) as cantidadadicional from movimientosventasdet group by foliofactura) det on f.foliofactura = det.foliofactura and cab.foliofactura = det.foliofactura " & "inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen and cab.codsucursal = suc.codalmacen " & "where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' and cab.estatus <> 'C' and f.tipofactura = 'N' and suc.codalmacen = " & Numerico(txtCodSucursal.Text) & "  group by suc.descalmacen,f.foliofactura,f.fechafactura,det.cantidadadicional,f.subtotal,f.redondeo,f.descuento,f.iva,f.total,f.moneda,f.tipocambio) " & "Union " & "(Select DescAlmacen AS Sucursal,FolioFactura as Factura, FechaFactura as Fecha, sum(Cantidad) as Cantidad,sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal," & "sum(Iva) as Iva, sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen,sum(F.Cantidad) as Cantidad,case when f.moneda = 'D' then round(F.SubTotal,2) else round(f.subtotal / f.tipocambio,2) end as subtotal," & "case when f.moneda = 'D' then round(F.Descuento,2) else round(f.descuento / f.tipocambio,2) end as descuento,case when f.moneda = 'D' then round(F.Iva,2) else round(f.iva / f.tipocambio,2) end as iva," & "case when f.moneda = 'D' then round(F.Total,2) else round(f.total / f.tipocambio,2) end as total,case when f.moneda = 'D' then round(F.Redondeo,2) else round(f.redondeo / f.tipocambio,2) end as redondeo," & "sum(case when f.moneda = 'D' then round(F.Importe,2) else round(f.importe / f.tipocambio,2) end) as Importe " & "From Facturas F Inner Join CatAlmacen A On F.CodSucursal = A.CodAlmacen " & "Where F.codsucursal = " & Numerico(txtCodSucursal.Text) & " and F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' and f.estatus <> 'C' " & "Group By F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total," & "F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group By DescAlmacen, FolioFactura, FechaFactura)"

            '        sql = "(SELECT SucFac.DescAlmacen AS Sucursal,Cab.FolioFactura AS Factura,SucFac.FechaFactura AS Fecha,SUM(Det.Cantidad) AS Cantidad,SUM((Cab.SubTotalAdicional + Cab.RedondeoAdicional)) AS Importe,SUM(Cab.DescuentoAdicional) AS Descuento,SUM(((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional)) AS SubTotal,SUM(Cab.IvaAdicional) AS Iva,SUM((Cab.TotalAdicional + Cab.RedondeoAdicional)) As Total FROM (SELECT CatAlmacen.CodAlmacen,CatAlmacen.DescAlmacen,Facturas.FechaFactura,facturas.tipocambio FROM CatAlmacen," & _
            ''        "Facturas WHERE CatAlmacen.TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen,FechaFactura,facturas.tipocambio) SucFac INNER JOIN MovimientosVentasCab Cab on Cab.FechaVenta = SucFac.FechaFactura AND SucFac.CodAlmacen = Cab.CodSucursal INNER JOIN (SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM MovimientosVentasDet GROUP BY FolioVenta) Det ON Cab.FolioVenta = Det.FolioVenta Where Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' " & _
            ''        "AND Cab.Estatus <> 'C' AND SucFac.CodAlmacen = " & CInt(Numerico(txtCodSucursal)) & " GROUP BY SucFac.DescAlmacen,Cab.FolioFactura,SucFac.FechaFactura) UNION (SELECT SucFac.DescAlmacen AS Sucursal,Det.FolioFactura AS Factura,Det.FechaFactura AS Fecha, SUM(Det.CantidadAdicional) AS Cantidad,SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe, SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal, " & _
            ''        "SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total FROM (SELECT DISTINCT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' Group By CodAlmacen,DescAlmacen) SucFac INNER JOIN (SELECT Det.FolioFactura,ISNULL(Det.CodSucursalAdicional,0) AS CodSucursal, SUM(Det.CantidadAdicional) AS CantidadAdicional, SUM((Det.PrecioListaSinIvaAdicional * Det.CantidadAdicional)) AS SubTotalAdicional,Det.RedondeoAdicional as redondeoadicional,SUM(((Det.ImptePromocionesAdicional + " & _
            ''        "Det.ImpteDescuentosAdicional) * Det.CantidadAdicional)) AS DescuentoAdicional, SUM((Det.IvaRealAdicional * Det.CantidadAdicional)) AS IvaAdicional,SUM((Det.PrecioRealAdicional * Det.CantidadAdicional)) AS TotalAdicional, F.FechaFactura FROM MovimientosVentasDet Det INNER JOIN Facturas F ON Det.FolioFactura = F.FolioFactura WHERE Det.FolioAdicional <> '' AND Det.FolioFactura <> '' AND Det.EstatusAdicional <> 'O' GROUP BY Det.FolioFactura,Det.CodSucursalAdicional,Det.RedondeoAdicional,F.FechaFactura,f.tipocambio) " & _
            ''        "Det ON SucFac.CodAlmacen = Det.CodSucursal WHERE Det.FechaFactura BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND SucFac.CodAlmacen = " & CInt(Numerico(txtCodSucursal)) & " GROUP BY SucFac.DescAlmacen,Det.FolioFactura,Det.FechaFactura) UNION  (Select DescAlmacen AS Sucursal, FolioFactura as Factura, FechaFactura as Fecha, sum(Cantidad) as Cantidad, sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal, sum(Iva) as Iva, " & _
            ''        "sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.FechaFactura, F.CodSucursal, A.DescAlmacen, sum(F.Cantidad) as Cantidad,case when f.moneda = 'D' then F.SubTotal else round(f.subtotal / f.tipocambio,2) end as subtotal,case when f.moneda = 'D' then F.Descuento else round(f.descuento / f.tipocambio,2) end as descuento,case when f.moneda = 'D' then F.Iva else round(f.iva / f.tipocambio,2) end as iva,case when f.moneda = 'D' then F.Total else round(f.total / f.tipocambio,2) end as total,case when f.moneda = 'D' then F.Redondeo else round(f.redondeo / f.tipocambio,2) end as redondeo,sum(case when f.moneda = 'D' then F.Importe else round(f.importe / f.tipocambio,2) end) as Importe From Facturas F Inner Join CatAlmacen A On F.CodSucursal = A.CodAlmacen Where F.CodSucursal = " & CInt(Numerico(txtCodSucursal)) & " And F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' Group By F.FolioFactura, F.FechaFactura, F.CodSucursal, " & _
            ''        "A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total, F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group By DescAlmacen, FolioFactura, FechaFactura ) "
            TextoMoneda = "Los importes estan expresados en dólares"
        End If
        BorraCmd()
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = sql
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No Existe Información En Este Rango de Fechas...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        Else
            'frmReportes.Report = RptFactFacturacionDetalladaXSucursal
            RptFactFacturacionDetalladaXSucursal.SetDataSource(frmReportes.rsReport)
        End If
        'H
        If chkV.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            With frmReportes.rsReport
                '.Text6.Suppress = True
                '.Field17.Suppress = True
                '.Field24.Suppress = True
                '.Field30.Suppress = True
            End With
        Else
            With frmReportes.rsReport
                '.Text6.Suppress = False
                '.Field17.Suppress = False
                '.Field24.Suppress = False
                '.Field30.Suppress = False
            End With
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.rsReport = rsReporte
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "PeriodoReporte", "TextoAdicional", "Moneda"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte, TextoAdicional, TextoMoneda}
        frmReportes.Text = "Facturación Detallada por Sucursal"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.reporteActual = RptFactFacturacionDetalladaXSucursal
        frmReportes.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
ImprimeErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Error al Imprimir : " & Err.Description, MsgBoxStyle.Exclamation, "Error de Operacion")
    End Sub

    Sub Limpiar()
        chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        txtCodSucursal.Text = ""
        txtCodSucursal.Enabled = False
        dbcSucursal.Text = ""
        dbcSucursal.Enabled = False
        dbcSucursal.Text = Nothing
        dtpFechaInicial.Value = Now
        dtpFechaFinal.Value = Now
        txtTextoAdicional.Text = ""
        chkTodaslasSucursales.Focus()
        optPesos.Checked = True
        optDolares.Checked = False
    End Sub

    Private Sub chkV_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkV.Enter
        On Error Resume Next
        txtTextoAdicional.Focus()
    End Sub

    Private Sub chkV_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkV.MouseUp
        Dim Button As Integer = eventArgs.Button \ &H100000
        Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        chkV.CheckState = IIf(Shift = VB6.ShiftConstants.CtrlMask, System.Windows.Forms.CheckState.Checked, System.Windows.Forms.CheckState.Unchecked)
    End Sub

    Private Sub chkTodaslasSucursales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.CheckStateChanged
        If chkTodaslasSucursales.CheckState = 1 Then
            txtCodSucursal.Text = ""
            txtCodSucursal.Enabled = False
            dbcSucursal.Text = ""
            dbcSucursal.Text = Nothing
            dbcSucursal.Enabled = False
        ElseIf chkTodaslasSucursales.CheckState = 0 Then
            txtCodSucursal.Enabled = True
            dbcSucursal.Enabled = True
        End If
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        FueraChange = True
        txtCodSucursal.Text = Format(String.Concat(intCodSucursal, "000"))
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.SelectedValue Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursal)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        tecla = eventArgs.KeyCode
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                txtCodSucursal.Focus()
        End Select
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
        gStrSql = "SELECT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
        'txtCodSucursal.Text = Format(String.Concat(intCodSucursal, "000"))
        txtCodSucursal.Text = (intCodSucursal)
        For i = 0 To 2 - txtCodSucursal.TextLength
            txtCodSucursal.Text = String.Concat("0" + txtCodSucursal.Text)
        Next i

        FueraChange = False
    End Sub

    Private Sub chkTodaslasSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        'If dbcSucursal.SelectedValue <> Nothing Then
        '    dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
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

    Private Sub dtpFechaInicial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.CursorChanged
        'sglTiempoCambio = VB.Timer()
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

    Private Sub frmFactReportesFacturacionDetalladaXSucursal_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmFactReportesFacturacionDetalladaXSucursal_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmFactReportesFacturacionDetalladaXSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "chkTodaslasSucursales" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmFactReportesFacturacionDetalladaXSucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmFactReportesFacturacionDetalladaXSucursal_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        dtpFechaInicial.MinDate = C_FECHAINICIAL
        dtpFechaInicial.MaxDate = C_FECHAFINAL
        dtpFechaFinal.MinDate = C_FECHAINICIAL
        dtpFechaFinal.MaxDate = C_FECHAFINAL
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
        optPesos.Checked = True
        optDolares.Checked = False
    End Sub

    Private Sub frmFactReportesFacturacionDetalladaXSucursal_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
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

    Private Sub frmFactReportesFacturacionDetalladaXSucursal_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
        'MDIMenuPrincipalCorpo.mnuFacturacionRptFactOpc(1).Enabled = True
    End Sub

    Private Sub txtCodSucursal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.TextChanged
        If FueraChange = True Then Exit Sub
        dbcSucursal.Text = ""
        dbcSucursal.Text = Nothing
    End Sub

    Private Sub txtCodSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtCodSucursal)
    End Sub

    Private Sub txtCodsucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodSucursal.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Leave
        If CDbl(Numerico(txtCodSucursal.Text)) = 0 Then
            'txtCodSucursal.Text = "000"
            For i = 1 To 3 - (txtCodSucursal.TextLength)
                txtCodSucursal.Text = String.Concat("0" + txtCodSucursal.Text)
            Next i
            Exit Sub
        End If
        FueraChange = True
        'txtCodSucursal.Text = Format(txtCodSucursal.Text, "000")
        For i = 1 To 3 - (txtCodSucursal.TextLength)
            txtCodSucursal.Text = String.Concat("0" + txtCodSucursal.Text)
        Next i
        FueraChange = False
        gStrSql = "SELECT * FROM CatAlmacen WHERE CodAlmacen=" & "'" & txtCodSucursal.Text & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("TipoAlmacen").Value = "P" Then
                FueraChange = True
                dbcSucursal.Text = Trim(RsGral.Fields("DescAlmacen").Value)
                FueraChange = False
            ElseIf RsGral.Fields("TipoAlmacen").Value = "V" Then
                MsgBox("Este Almacen es de Tipo Vendedor Externo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtCodSucursal.Text = ""
                txtCodSucursal.Focus()
            End If
        Else
            MsjNoExiste("La Sucursal", gstrNombCortoEmpresa)
            txtCodSucursal.Text = ""
            txtCodSucursal.Focus()
        End If
    End Sub

    Private Sub txtTextoAdicional_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTextoAdicional.Enter
        Pon_Tool()
        SelTextoTxt(txtTextoAdicional)
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click

    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class