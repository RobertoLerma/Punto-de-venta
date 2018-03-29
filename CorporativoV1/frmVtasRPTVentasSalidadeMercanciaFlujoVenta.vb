Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasRPTVentasSalidadeMercanciaFlujoVenta
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkDescendente As System.Windows.Forms.CheckBox
    Public WithEvents chkProvs As System.Windows.Forms.CheckBox
    Public WithEvents chkSucs As System.Windows.Forms.CheckBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents fraFiltro As System.Windows.Forms.GroupBox
    Public WithEvents chkImpuesto As System.Windows.Forms.CheckBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_2 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_1 As System.Windows.Forms.GroupBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Const C_TODAS As String = "[ TODAS ... ]"
    Const C_TODOS As String = "[ TODOS ... ]"
    Const C_XXXXX As String = "[ ......... ]"

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean

    Dim mblnFueraChange As Boolean
    Dim mintCodSucursal As Integer
    Dim mintCodProveedor As Integer
    Dim tecla As Integer
    Dim cTablaTmp As String
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim mblnSalir As Boolean


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.fraFiltro = New System.Windows.Forms.GroupBox()
        Me.chkDescendente = New System.Windows.Forms.CheckBox()
        Me.chkProvs = New System.Windows.Forms.CheckBox()
        Me.chkSucs = New System.Windows.Forms.CheckBox()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.chkImpuesto = New System.Windows.Forms.CheckBox()
        Me._fraVtas_1 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.fraVtas = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.fraFiltro.SuspendLayout()
        Me._fraVtas_1.SuspendLayout()
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
        Me.txtMensaje.Location = New System.Drawing.Point(8, 287)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(366, 77)
        Me.txtMensaje.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'fraFiltro
        '
        Me.fraFiltro.BackColor = System.Drawing.SystemColors.Control
        Me.fraFiltro.Controls.Add(Me.chkDescendente)
        Me.fraFiltro.Controls.Add(Me.chkProvs)
        Me.fraFiltro.Controls.Add(Me.chkSucs)
        Me.fraFiltro.Controls.Add(Me.dbcProveedor)
        Me.fraFiltro.Controls.Add(Me.dbcSucursal)
        Me.fraFiltro.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraFiltro.Location = New System.Drawing.Point(6, 6)
        Me.fraFiltro.Margin = New System.Windows.Forms.Padding(2)
        Me.fraFiltro.Name = "fraFiltro"
        Me.fraFiltro.Padding = New System.Windows.Forms.Padding(2)
        Me.fraFiltro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraFiltro.Size = New System.Drawing.Size(368, 148)
        Me.fraFiltro.TabIndex = 0
        Me.fraFiltro.TabStop = False
        Me.fraFiltro.Text = " Por Proveedor / Sucursal"
        '
        'chkDescendente
        '
        Me.chkDescendente.BackColor = System.Drawing.SystemColors.Control
        Me.chkDescendente.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDescendente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDescendente.Location = New System.Drawing.Point(248, 119)
        Me.chkDescendente.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDescendente.Name = "chkDescendente"
        Me.chkDescendente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDescendente.Size = New System.Drawing.Size(120, 25)
        Me.chkDescendente.TabIndex = 5
        Me.chkDescendente.Text = "Importe Vtas Desc"
        Me.chkDescendente.UseVisualStyleBackColor = False
        '
        'chkProvs
        '
        Me.chkProvs.BackColor = System.Drawing.SystemColors.Control
        Me.chkProvs.Checked = True
        Me.chkProvs.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkProvs.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkProvs.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkProvs.Location = New System.Drawing.Point(6, 17)
        Me.chkProvs.Margin = New System.Windows.Forms.Padding(2)
        Me.chkProvs.Name = "chkProvs"
        Me.chkProvs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkProvs.Size = New System.Drawing.Size(157, 19)
        Me.chkProvs.TabIndex = 1
        Me.chkProvs.Text = "Todas los proveedores"
        Me.chkProvs.UseVisualStyleBackColor = False
        '
        'chkSucs
        '
        Me.chkSucs.BackColor = System.Drawing.SystemColors.Control
        Me.chkSucs.Checked = True
        Me.chkSucs.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSucs.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSucs.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkSucs.Location = New System.Drawing.Point(101, 65)
        Me.chkSucs.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSucs.Name = "chkSucs"
        Me.chkSucs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSucs.Size = New System.Drawing.Size(131, 21)
        Me.chkSucs.TabIndex = 3
        Me.chkSucs.Text = "Por sucursal ..."
        Me.chkSucs.UseVisualStyleBackColor = False
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(101, 40)
        Me.dbcProveedor.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(240, 21)
        Me.dbcProveedor.TabIndex = 2
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(101, 91)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(240, 21)
        Me.dbcSucursal.TabIndex = 4
        '
        'chkImpuesto
        '
        Me.chkImpuesto.BackColor = System.Drawing.SystemColors.Control
        Me.chkImpuesto.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkImpuesto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkImpuesto.Location = New System.Drawing.Point(272, 232)
        Me.chkImpuesto.Margin = New System.Windows.Forms.Padding(2)
        Me.chkImpuesto.Name = "chkImpuesto"
        Me.chkImpuesto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkImpuesto.Size = New System.Drawing.Size(102, 25)
        Me.chkImpuesto.TabIndex = 11
        Me.chkImpuesto.Text = "Incluir Impuesto"
        Me.chkImpuesto.UseVisualStyleBackColor = False
        '
        '_fraVtas_1
        '
        Me._fraVtas_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_1.Controls.Add(Me.dtpDesde)
        Me._fraVtas_1.Controls.Add(Me.dtpHasta)
        Me._fraVtas_1.Controls.Add(Me._lblVentas_1)
        Me._fraVtas_1.Controls.Add(Me._lblVentas_2)
        Me._fraVtas_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_1.Location = New System.Drawing.Point(9, 171)
        Me._fraVtas_1.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.Name = "_fraVtas_1"
        Me._fraVtas_1.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_1.Size = New System.Drawing.Size(365, 57)
        Me._fraVtas_1.TabIndex = 6
        Me._fraVtas_1.TabStop = False
        Me._fraVtas_1.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(73, 23)
        Me.dtpDesde.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(102, 20)
        Me.dtpDesde.TabIndex = 8
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(254, 23)
        Me.dtpHasta.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(98, 20)
        Me.dtpHasta.TabIndex = 10
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_1.Location = New System.Drawing.Point(17, 27)
        Me._lblVentas_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(52, 13)
        Me._lblVentas_1.TabIndex = 7
        Me._lblVentas_1.Text = "Desde el "
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_2.Location = New System.Drawing.Point(196, 27)
        Me._lblVentas_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(46, 13)
        Me._lblVentas_2.TabIndex = 9
        Me._lblVentas_2.Text = "Hasta el"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblRpt_2.Location = New System.Drawing.Point(9, 261)
        Me._lblRpt_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 12
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(122, 382)
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
        Me.btnImprimir.Location = New System.Drawing.Point(7, 382)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 35
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasRPTVentasSalidadeMercanciaFlujoVenta
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(387, 431)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.fraFiltro)
        Me.Controls.Add(Me.chkImpuesto)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me._fraVtas_1)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(321, 199)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasRPTVentasSalidadeMercanciaFlujoVenta"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Flujo de Venta por Proveedor"
        Me.fraFiltro.ResumeLayout(False)
        Me._fraVtas_1.ResumeLayout(False)
        Me._fraVtas_1.PerformLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Public Sub Nuevo()
        mblnFueraChange = True

        chkProvs.CheckState = System.Windows.Forms.CheckState.Checked
        dbcProveedor.Text = C_TODOS
        dbcProveedor.Enabled = False

        chkSucs.CheckState = System.Windows.Forms.CheckState.Unchecked
        dbcSucursal.Text = C_XXXXX
        dbcSucursal.Enabled = False
        chkDescendente.Enabled = False
        chkDescendente.CheckState = System.Windows.Forms.CheckState.Checked

        mintCodSucursal = 0
        mintCodProveedor = 0
        mblnFueraChange = False

        dtpDesde.Value = Format(Today, C_FORMATFECHAMOSTRAR)
        dtpHasta.Value = Format(Today, C_FORMATFECHAMOSTRAR)
        chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked
        txtMensaje.Text = ""
    End Sub

    Public Sub Limpiar()
        Nuevo()
        chkProvs.Focus()
    End Sub

    ''' se agregó ventas, utilidad, valor del inv, filtro por ventas e inv > 0
    ''' Ult Modific.- 28MAR2006
    Function DevuelveQuery(ByRef lTipo As Integer) As String
        On Error GoTo Err_Renamed
        Dim Sql As String
        Dim lProv As String
        Dim lSuc As String
        Dim lProvE As String
        Dim lSucE As String
        Dim lOrder As String

        lProv = ""
        lSuc = ""
        lProvE = ""
        lSucE = ""
        lOrder = " Asc "
        If chkDescendente.CheckState = System.Windows.Forms.CheckState.Checked Then lOrder = " Desc "

        If lTipo = 0 Then '''reporte nuevo - agrupado por Sucursal/Proveedor

            If Trim(dbcProveedor.Text) = C_TODOS Then
                lProv = " codProveedor "
                lProvE = " A.CodProveedor "
            Else
                lProv = Trim(CStr(mintCodProveedor))
                lProvE = Trim(CStr(mintCodProveedor))
            End If

            If Trim(dbcSucursal.Text) = C_TODAS Then
                lSuc = " codSucursal "
                lSucE = " Inv.CodAlmacen "
            Else
                lSuc = Trim(CStr(mintCodSucursal))
                lSucE = Trim(CStr(mintCodSucursal))
            End If

            Sql = "Select  Info.CodAlmacen, Info.DescAlmacen, Info.CodProveedor, Info.DescProvAcreed, Info.ImporteVentas, Info.CostoVentas, (Info.ImporteVentas - Info.CostoVentas) as Utilidad, Info.VtaTotal, Info.ValorInv " & "From    ( " & "        SELECT  Invent.CodAlmacen, CS.DescAlmacen, Invent.CodProveedor, CP.DescProvAcreed, sum(IsNull(Vta.ImporteVentas,0)) As ImporteVentas, sum(ISNull(Vta.CostoVentas,0)) As CostoVentas, Tot.VtaTotal, Invent.ValorInv " & "        FROM    ( " & "                 SELECT   Inv.CodAlmacen, A.CodProveedor, Sum(((Inv.ExistenciaInicial+Inv.Entradas) - (Inv.Salidas+Inv.Apartados))) as Existencia, sum(((Inv.ExistenciaInicial + Inv.Entradas) - (Inv.Salidas + Inv.Apartados)) * Round(A.CostoReal,0)) As ValorInv " & "                 From     Inventario Inv (Nolock) " & "                 Inner    Join CatArticulos A (Nolock) On Inv.CodArticulo = A.CodArticulo " & "                 Inner    Join CatAlmacen Alm (Nolock) On Inv.CodAlmacen = Alm.CodAlmacen And Alm.TipoAlmacen = 'P' " & "                 Where    A.CodProveedor = " & lProvE & " " & "                 And      Inv.CodAlmacen = " & lSucE & " " & "                 Group    by Inv.CodAlmacen, A.CodProveedor " & "                 ) Invent Left Outer Join ( " & "                 SELECT   CodSucursal, CodProveedor, ROUND(SUM(ISNULL(PrecioReal * (Cantidad - CantidadDev),0)),2) AS ImporteVentas, Round(sum(IsNull(CostoVenta * (Cantidad - CantidadDev), 0)), 2) As CostoVentas " & "                 FROM     DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') " & "                 Where    (Cantidad - CantidadDev) > 0 " & "                 AND      CodProveedor =  " & lProv & " " & "                 AND      CodSucursal  =  " & lSuc & " " & "                 GROUP    BY CodSucursal, CodProveedor " & "                 ) Vta    On Invent.CodAlmacen = Vta.codSucursal And Invent.CodProveedor = Vta.codProveedor " & "                 Left     Outer Join ( "
            Sql = Sql & "                 SELECT   CodSucursal, ROUND(SUM(ISNULL(PrecioReal * (Cantidad - CantidadDev),0)),2) AS VtaTotal " & "                 FROM     DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') " & "                 Where    (Cantidad - CantidadDev) > 0 " & "                 AND      CodProveedor =  " & lProv & " " & "                 AND      CodSucursal  =  " & lSuc & " " & "                 GROUP    BY CodSucursal " & "                 ) Tot    On Invent.CodAlmacen = Tot.CodSucursal " & "        Inner Join CatProvAcreed CP (Nolock) On Invent.CodProveedor = CP.CodProvAcreed " & "        Inner Join CatAlmacen CS (Nolock) On Invent.CodAlmacen = CS.CodAlmacen And CS.TipoAlmacen = 'P' " & "        Group    by Invent.CodAlmacen, CS.DescAlmacen, Invent.codProveedor, CP.DescProvAcreed, Tot.VtaTotal, Invent.ValorInv " & "        ) as Info " & "Where   (ImporteVentas > 0 or ValorInv > 0) " & "Order   by VtaTotal Desc, ImporteVentas Desc, DescAlmacen, DescProvAcreed "

        ElseIf lTipo = 1 Then  '''Reporte original - agrupado por Proveedor

            Sql = "SELECT Ventas.CodProvAcreed, Ventas.DescProvAcreed, Ventas.ImporteVentas, Compras.ImporteCompras, " & "CASE WHEN Compras.ImporteCompras <> 0 THEN (Ventas.ImporteVentas/Compras.ImporteCompras) ELSE 0 END AS IndiceRotacion " & "FROM (SELECT CP.CodProvAcreed,CP.DescProvAcreed," & IIf(chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "ROUND(SUM(ISNULL(PrecioReal * (Cantidad - CantidadDev),0)),2)", "ROUND(SUM(ISNULL((PrecioListaSinIva - Descuento) * (Cantidad - CantidadDev),0)),2)") & " AS ImporteVentas " & "FROM (SELECT * FROM CatProvAcreed (Nolock) WHERE Tipo = 'P') CP LEFT OUTER JOIN (SELECT * FROM DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') WHERE (Cantidad - CantidadDev) > 0) VTA " & "ON CP.CodProvAcreed = VTA.CodProveedor GROUP BY CP.CodProvAcreed,CP.DescProvAcreed) Ventas INNER JOIN " & "(SELECT CP.CodProvAcreed,CP.DescProvAcreed,ROUND(SUM(ISNULL(CASE WHEN OC.Moneda = 'D' THEN " & IIf(chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "OC.Total", "OC.SubTotal - OC.Descuento") & " WHEN OC.Moneda = 'P' THEN DBO.ConvertirCantidad('P','D'," & IIf(chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "OC.Total", "OC.SubTotal - OC.Descuento") & ",OC.TipoCambioC,OC.TipoCambioEuroC) WHEN OC.Moneda = 'E' THEN DBO.ConvertirCantidad('E','D'," & IIf(chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "OC.Total", "OC.SubTotal - OC.Descuento") & ",OC.TipoCambioC,OC.TipoCambioEuroC) END,0)),2) AS ImporteCompras " & "FROM (SELECT * FROM CatProvAcreed (Nolock) WHERE Tipo = 'P') CP LEFT OUTER JOIN (SELECT * FROM OrdenesCompra (Nolock) WHERE (Estatus = 'R' OR Estatus = 'G') AND FechaCompraEI BETWEEN '" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "' AND '" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') OC ON CP.CodProvAcreed = OC.CodProvAcreed " & "GROUP BY CP.CodProvAcreed,CP.DescProvAcreed) Compras ON Ventas.CodProvAcreed = Compras.CodProvAcreed Where (VENTAS.IMPORTEVENTAS <> 0 Or COMPRAS.IMPORTECOMPRAS <> 0) " & IIf(mintCodProveedor <> 0, "AND Ventas.CodProvAcreed = " & mintCodProveedor, "") & " Order by Ventas.ImporteVentas desc "

        End If
        DevuelveQuery = Sql

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub Imprime()
        Dim rptVentasSalidaDeMercanciaFlujoVenta As New rptVentasSalidaDeMercanciaFlujoVenta
        Dim rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order As New rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        'On Error GoTo Merr
        Dim lStrSql As String
        'Declarar vectores para almacenar los parámetros que se le enviarán al reporte
        Dim aParam(6) As Object
        Dim aValues(6) As Object

        If Not ValidaDatos() Then Exit Sub

        'POR SUCURSAL
        If ((chkSucs.CheckState = System.Windows.Forms.CheckState.Checked) And Trim(dbcSucursal.Text) <> "") And ((chkSucs.CheckState = System.Windows.Forms.CheckState.Checked) Or (Trim(dbcSucursal.Text) <> "")) Then

            lStrSql = DevuelveQuery(0)
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
                rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order.SetDataSource(frmReportes.rsReport)
            End If

            'aParam(1) = "Empresa"
            'aValues(1) = Trim(gstrNombCortoEmpresa)
            'aParam(2) = "dDesde"
            'aValues(2) = Me.dtpDesde.Value
            'aParam(3) = "dHasta"
            'aValues(3) = Me.dtpHasta.Value
            'aParam(4) = "Mensaje"
            'aValues(4) = Trim(Me.txtMensaje.Text)
            'aParam(5) = "MonedaDeCantidades"
            'aParam(6) = "IncluyeImpuestos"
            'aValues(6) = IIf(Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "** Las cantidades expresadas incluyen IVA.", "** Las cantidades expresadas NO incluyen IVA.")


            If (txtMensaje.Text <> Nothing) Then
                pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
            End If

            If (dtpDesde.Value <> Nothing) Then
                pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
            End If

            If (dtpHasta.Value <> Nothing) Then
                pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
            End If

            If (gstrNombCortoEmpresa <> Nothing) Then
                pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
            End If

            'If (MonedaDeCantidades <> Nothing) Then
            '    pdvNum.Value = MonedaDeCantidades : pvNum.Add(pdvNum)
            '    rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order.DataDefinition.ParameterFields("MonedaDeCantidades").ApplyCurrentValues(pvNum)
            'End If

            If (chkImpuesto.CheckState <> Nothing) Then
                pdvNum.Value = IIf(Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "** Las cantidades expresadas incluyen IVA.", "** Las cantidades expresadas NO incluyen IVA.") : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order.DataDefinition.ParameterFields("IncluyeImpuestos").ApplyCurrentValues(pvNum)
            End If



            frmReportes.reporteActual = rptVentasSalidaDeMercanciaFlujoVenta_xSuc_Order 'Es el nombre del archivo que se incluyó en el proyecto
            'frmReportes.Imprime(Trim(Me.Text), aParam, aValues)
            frmReportes.Show()
            Cmd.CommandTimeout = 90

        End If

        'POR PROVEEDOR
        If (((chkProvs.CheckState = System.Windows.Forms.CheckState.Checked) Or (Trim(dbcProveedor.Text) <> "")) And (chkSucs.CheckState = System.Windows.Forms.CheckState.Unchecked)) Then
            lStrSql = DevuelveQuery(1)
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
                dtpDesde.Focus()
                Exit Sub
            Else
                rptVentasSalidaDeMercanciaFlujoVenta.SetDataSource(frmReportes.rsReport)
            End If


            'aParam(1) = "Empresa"
            'aValues(1) = Trim(gstrNombCortoEmpresa)
            'aParam(2) = "dDesde"
            'aValues(2) = Me.dtpDesde.Value
            'aParam(3) = "dHasta"
            'aValues(3) = Me.dtpHasta.Value
            'aParam(4) = "Mensaje"
            'aValues(4) = Trim(Me.txtMensaje.Text)
            'aParam(5) = "MonedaDeCantidades"
            'aParam(6) = "IncluyeImpuestos"
            'aValues(6) = IIf(Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "** Las cantidades expresadas incluyen IVA.", "** Las cantidades expresadas NO incluyen IVA.")



            If (txtMensaje.Text <> Nothing) Then
                pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
            End If

            If (dtpDesde.Value <> Nothing) Then
                pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
            End If

            If (dtpHasta.Value <> Nothing) Then
                pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
            End If

            If (gstrNombCortoEmpresa <> Nothing) Then
                pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
            End If

            'If (MonedaDeCantidades <> Nothing) Then
            '    pdvNum.Value = MonedaDeCantidades : pvNum.Add(pdvNum)
            '    rptVentasSalidaDeMercanciaFlujoVenta.DataDefinition.ParameterFields("MonedaDeCantidades").ApplyCurrentValues(pvNum)
            'End If

            If (chkImpuesto.CheckState <> Nothing) Then
                pdvNum.Value = IIf(Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked, "** Las cantidades expresadas incluyen IVA.", "** Las cantidades expresadas NO incluyen IVA.") : pvNum.Add(pdvNum)
                rptVentasSalidaDeMercanciaFlujoVenta.DataDefinition.ParameterFields("IncluyeImpuestos").ApplyCurrentValues(pvNum)
            End If


            frmReportes.reporteActual = rptVentasSalidaDeMercanciaFlujoVenta 'Es el nombre del archivo que se incluyó en el proyecto
            frmReportes.Show()
            'frmReportes.Imprime(Trim(Me.Text), aParam, aValues)

            Cmd.CommandTimeout = 90
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function ValidaDatos() As Boolean
        On Error GoTo Merr

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

        If chkProvs.CheckState = System.Windows.Forms.CheckState.Unchecked Then

            If Trim(dbcProveedor.Text) = "" Then
                MsgBox("Debe seleccionar un proveedor...", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                dbcProveedor.Focus()
            End If
        End If

        If chkSucs.CheckState = System.Windows.Forms.CheckState.Checked Then

            If Trim(dbcSucursal.Text) = "" Then
                MsgBox("Debe seleccionar una sucursal...", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                dbcSucursal.Focus()
            End If
        End If
        Select Case True
            Case dtpDesde.Value > dtpHasta.Value
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Límite", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                dtpDesde.Focus()
            Case Else
                ValidaDatos = True
        End Select

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Private Sub chkProvs_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkProvs.CheckStateChanged
        If chkProvs.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            dbcProveedor.Enabled = True

            dbcProveedor.Text = ""
        Else
            dbcProveedor.Enabled = False

            dbcProveedor.Text = C_TODOS
        End If
        mintCodProveedor = 0
    End Sub

    Private Sub chkProvs_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkProvs.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
        End If
    End Sub

    Private Sub chkSucs_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSucs.CheckStateChanged
        If chkSucs.CheckState = System.Windows.Forms.CheckState.Unchecked Then

            dbcSucursal.Text = C_XXXXX
            dbcSucursal.Enabled = False
            chkDescendente.Enabled = False
        Else
            dbcSucursal.Enabled = True

            dbcSucursal.Text = C_TODAS
            chkDescendente.Enabled = True
            chkDescendente.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        mintCodSucursal = 0
    End Sub

    Private Sub chkSucs_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkSucs.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If chkProvs.CheckState = System.Windows.Forms.CheckState.Checked Then
                chkProvs.Focus()
            Else
                dbcProveedor.Focus()
            End If
        End If
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed (Nolock) Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%' Order by descProvAcreed "
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcProveedor))


        If Trim(Me.dbcProveedor.Text) = "" Then
            mintCodProveedor = 0
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        Pon_Tool()
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed (Nolock) WHERE Tipo = '" & C_TPROVEEDOR & "' ORDER BY descProvAcreed "
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcProveedor))
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcProveedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkProvs.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
        Dim Aux As Integer

        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub

        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed (Nolock) Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%' Order by descProvAcreed "
        Aux = mintCodProveedor
        mintCodProveedor = 0

        If Trim(Me.dbcProveedor.Text) <> Trim(C_TODOS) Or Trim(Me.dbcProveedor.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcProveedor), gStrSql, mintCodProveedor)
        End If

        If Aux <> mintCodProveedor Then
            If mintCodProveedor = 0 Then
                mblnFueraChange = True

                Me.dbcProveedor.Text = C_TODOS
                Me.dbcProveedor.Enabled = True
                mblnFueraChange = False
            End If
        End If

        If Trim(Me.dbcProveedor.Text) = "" Then Me.dbcProveedor.Text = C_TODOS
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen (Nolock) Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%' Order by codAlmacen "
        ModDCombo.DCChange(lStrSql, tecla, dbcSucursal)


        If Trim(Me.dbcSucursal.Text) = "" Then
            mintCodSucursal = 0
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen (Nolock) WHERE TipoAlmacen = 'P' Order by codAlmacen "
        ModDCombo.DCGotFocus(gStrSql, dbcSucursal)
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkSucs.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        On Error GoTo Merr
        Dim Aux As Integer

        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub


        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen (Nolock) Where TipoAlmacen = 'P' And DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' Order by codAlmacen "
        Aux = mintCodSucursal
        mintCodSucursal = 0

        If Trim(Me.dbcSucursal.Text) <> Trim(C_TODAS) Or Trim(Me.dbcSucursal.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcSucursal), gStrSql, mintCodSucursal)
        End If

        If Aux <> mintCodSucursal Then
            If mintCodSucursal = 0 Then
                mblnFueraChange = True

                dbcSucursal.Text = C_TODAS
                dbcSucursal.Enabled = True
                mblnFueraChange = False
            End If
        End If

        If Trim(dbcSucursal.Text) = "" Then dbcSucursal.Text = C_TODAS

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

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
    Private Sub frmVtasRPTVentasSalidadeMercanciaFlujoVenta_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaFlujoVenta_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaFlujoVenta_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaFlujoVenta_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaFlujoVenta_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        dtpDesde.MinDate = C_FECHAINICIAL
        dtpDesde.MaxDate = C_FECHAFINAL
        dtpHasta.MinDate = C_FECHAINICIAL
        dtpHasta.MaxDate = C_FECHAFINAL
        Nuevo()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaFlujoVenta_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If mblnSalir Then
            mblnSalir = False
            Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Sale del Formulario
                    Cancel = 0
                Case MsgBoxResult.No 'No sale del formulario
                    'If chkSucs.Value = vbChecked Then
                    '   chkSucs.SetFocus
                    'Else
                    '   chkProvs.SetFocus
                    'End If
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaFlujoVenta_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub OptSucProv_Click(ByRef Index As Integer)
        '''   Select Case Index
        '''          Case 0 ''' Sucursal
        '''               dbcSucursal.Enabled = True
        '''               dbcSucursal.text = C_TODAS
        '''               dbcProveedor.Enabled = False
        '''               dbcProveedor.text = C_TODOS
        '''          Case 1 ''' Proveedor
        '''               dbcProveedor.Enabled = True
        '''               dbcProveedor.text = C_TODOS
        '''               dbcSucursal.Enabled = False
        '''               dbcSucursal.text = C_TODAS
        '''   End Select
    End Sub

    Private Sub OptSucProv_KeyDown(ByRef Index As Integer, ByRef KeyCode As Integer, ByRef Shift As Integer)
        '''   If KeyCode = vbKeyEscape Then
        '''      Select Case Index
        '''             Case 0
        '''                  mblnSalir = True
        '''                  Unload Me
        '''             Case 1
        '''                  OptSucProv(0).SetFocus
        '''      End Select
        '''   End If
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
    ''' original
    '''Sql = "SELECT Ventas.CodProvAcreed,Ventas.DescProvAcreed,Ventas.ImporteVentas,Compras.ImporteCompras," & _
    '"CASE WHEN Compras.ImporteCompras <> 0 THEN (Ventas.ImporteVentas/Compras.ImporteCompras) ELSE 0 END AS IndiceRotacion " & _
    '"FROM (SELECT CP.CodProvAcreed,CP.DescProvAcreed," & IIf(chkImpuesto.Value = vbChecked, "ROUND(SUM(ISNULL(PrecioReal * (Cantidad - CantidadDev),0)),2)", "ROUND(SUM(ISNULL((PrecioListaSinIva - Descuento) * (Cantidad - CantidadDev),0)),2)") & " AS ImporteVentas " & _
    '"FROM (SELECT * FROM CatProvAcreed WHERE Tipo = 'P') CP LEFT OUTER JOIN (SELECT * FROM DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta, C_FORMATFECHAGUARDAR) & "') WHERE (Cantidad - CantidadDev) > 0) VTA " & _
    '"ON CP.CodProvAcreed = VTA.CodProveedor GROUP BY CP.CodProvAcreed,CP.DescProvAcreed) Ventas INNER JOIN " & _
    '"(SELECT CP.CodProvAcreed,CP.DescProvAcreed,ROUND(SUM(ISNULL(CASE WHEN OC.Moneda = 'D' THEN " & IIf(chkImpuesto.Value = vbChecked, "OC.Total", "OC.SubTotal - OC.Descuento") & " WHEN OC.Moneda = 'P' THEN DBO.ConvertirCantidad('P','D'," & IIf(chkImpuesto.Value = vbChecked, "OC.Total", "OC.SubTotal - OC.Descuento") & ",OC.TipoCambioC,OC.TipoCambioEuroC) WHEN OC.Moneda = 'E' THEN DBO.ConvertirCantidad('E','D'," & IIf(chkImpuesto.Value = vbChecked, "OC.Total", "OC.SubTotal - OC.Descuento") & ",OC.TipoCambioC,OC.TipoCambioEuroC) END,0)),2) AS ImporteCompras " & _
    '"FROM (SELECT * FROM CatProvAcreed WHERE Tipo = 'P') CP LEFT OUTER JOIN (SELECT * FROM OrdenesCompra WHERE (Estatus = 'R' OR Estatus = 'G') AND FechaCompraEI BETWEEN '" & Format(dtpDesde, C_FORMATFECHAGUARDAR) & "' AND '" & Format(dtpHasta, C_FORMATFECHAGUARDAR) & "') OC ON CP.CodProvAcreed = OC.CodProvAcreed " & _
    '"GROUP BY CP.CodProvAcreed,CP.DescProvAcreed) Compras ON Ventas.CodProvAcreed = Compras.CodProvAcreed Where (VENTAS.IMPORTEVENTAS <> 0 Or COMPRAS.IMPORTECOMPRAS <> 0) " & IIf(mintCodProveedor <> 0, "AND Ventas.CodProvAcreed = " & mintCodProveedor, "")

    ''' QUERY ANTERIOR - SOLO CONSIDERABA INV INICIAL PARA LOS ARTICULOS VENDIDOS - 03MAR2006
    '''Sql = "Select   Info.CodSucursal, Info.DescAlmacen, Info.CodProveedor, Info.DescProvAcreed, Info.ImporteVentas, Info.CostoVentas, (Info.ImporteVentas - Info.CostoVentas) as Utilidad, Info.VtaTotal, ValorInv " & _
    '"From     ( " & _
    '"         SELECT   Vta.CodSucursal, CS.DescAlmacen, Vta.CodProveedor, CP.DescProvAcreed, sum(Vta.ImporteVentas) As ImporteVentas, sum(Vta.CostoVentas) As CostoVentas, Tot.VtaTotal, Invent.ValorInv " & _
    '"         FROM     ( " & _
    '"                  SELECT   CodSucursal, CodProveedor, ROUND(SUM(ISNULL(PrecioReal * (Cantidad - CantidadDev),0)),2) AS ImporteVentas, Round(sum(IsNull(CostoVenta * (Cantidad - CantidadDev), 0)), 2) As CostoVentas " & _
    '"                  FROM     DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta, C_FORMATFECHAGUARDAR) & "') " & _
    '"                  Where    (Cantidad - CantidadDev) > 0 " & _
    '"                  AND      CodProveedor = " & lProv & " " & _
    '"                  AND      CodSucursal =  " & lSuc & " " & _
    '"                  GROUP    BY CodSucursal, CodProveedor " & _
    '"                  ) Vta Inner Join ( " & _
    '"                  SELECT   CodSucursal, ROUND(SUM(ISNULL(PrecioReal * (Cantidad - CantidadDev),0)),2) AS VtaTotal " & _
    '"                  FROM     DBO.VTAS_SALIDAMCIA('" & Format(dtpDesde, C_FORMATFECHAGUARDAR) & "','" & Format(dtpHasta, C_FORMATFECHAGUARDAR) & "') " & _
    '"                  Where    (Cantidad - CantidadDev) > 0 " & _
    '"                  AND      CodProveedor = " & lProv & " " & _
    '"                  AND      CodSucursal =  " & lSuc & " " & _
    '"                  GROUP    BY CodSucursal " & _
    '"                  ) Tot On Vta.CodSucursal = Tot.CodSucursal Inner Join ( " & _
    '"                  SELECT   Inv.CodAlmacen, A.CodProveedor, sum(((Inv.ExistenciaInicial+Inv.Entradas) - (Inv.Salidas+Inv.Apartados))) as Existencia, Round(sum(((Inv.ExistenciaInicial + Inv.Entradas) - (Inv.Salidas + Inv.Apartados)) * A.CostoReal), 2) As ValorInv " & _
    '"                  From     Inventario Inv (Nolock) " & _
    '"                  Inner    Join CatArticulos A (Nolock) On Inv.CodArticulo = A.CodArticulo " & _
    '"                  Inner    Join CatAlmacen Alm (Nolock) On Inv.CodAlmacen = Alm.CodAlmacen And Alm.TipoAlmacen = 'P' " & _
    '"                  Where    A.CodProveedor = " & lProvE & " " & _
    '"                  And      Inv.CodAlmacen = " & lSucE & " " & _
    '"                  Group    by Inv.CodAlmacen, A.CodProveedor "
    '''Sql = Sql & _
    '"                  ) Invent On Vta.CodSucursal = Invent.CodAlmacen And Vta.CodProveedor = Invent.CodProveedor " & _
    '"                  Inner Join CatProvAcreed CP (Nolock) On Vta.CodProveedor = CP.CodProvAcreed " & _
    '"                  Inner Join CatAlmacen CS (Nolock) On Vta.CodSucursal = CS.CodAlmacen And CS.TipoAlmacen = 'P' " & _
    '"         Where    (Vta.ImporteVentas <> 0) " & _
    '"         Group     by Vta.CodSucursal, CS.DescAlmacen, Vta.CodProveedor, CP.DescProvAcreed, Tot.VtaTotal, Invent.ValorInv " & _
    '" ) as Info " & _
    '"Order   by VtaTotal  " & lOrder & " , ImporteVentas " & lOrder & ", DescAlmacen, DescProvAcreed "
End Class