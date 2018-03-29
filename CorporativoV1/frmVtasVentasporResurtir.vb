Option Strict Off
Option Explicit On
Imports System.IO
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.Office.Interop
Public Class frmVtasVentasporResurtir
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE VENTAS PARA RESURTIR                                                              *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      JUEVES 20 DE MAYO DE 2004                                                                    *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' MODIFIC.-  COLUMNAS NUEVAS:  ORIGENANT, COSTOFACTURAPESOS
    '            NVA AGRUPACION POR SUCURSAL-PROVEEDOR
    '            31 MAYO 2005
    '
    ' MODIFIC.-  COLUMNA UTILIDAD SE AGREGO VOLUMEN - TOMANDO EN CUENTA LA CANTIDAD (VTA-DEV)
    '**********************************************************************************************************************'
    ' MODIFIC.-  CONSIDERA COLUMNAS X VOLUMEN ( PVTA-DESCTO-CTO )  SE ELIMINARON TOTALES DE COLS UNITARIAS ( PPUB-CTOFACT )
    '            22JUN2005
    '**********************************************************************************************************************'
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents chkMostrarCFP As System.Windows.Forms.CheckBox
    Public WithEvents chkMostrarCodArtProv As System.Windows.Forms.CheckBox
    Public WithEvents chkMostrarCodAnt As System.Windows.Forms.CheckBox
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents chkTodaslasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox

    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim intCodSucursal As Integer
    Dim tecla As Integer
    Dim RsAux As ADODB.Recordset
    Dim sglTiempoCambio As Single
    Dim ObjExcel As Object
    Dim objLibro As Excel.Workbook
    Dim objHoja As Excel.Worksheet
    Dim Renglon As Integer
    Dim Columna As Integer
    Dim MostrarCostoyUtilidad As Boolean



    Dim SubTotalPrecioPub As Decimal
    Dim SubTotalPrecioVta As Decimal
    Dim SubTotalDescuento As Decimal
    Dim SubTotalCosto As Decimal
    Dim SubTotalUtilidad As Decimal
    Dim SubTotalCFP As Decimal
    Dim Margen As Decimal
    Dim TotalPrecioPub As Decimal
    Dim TotalPrecioVta As Decimal
    Dim TotalDescuento As Decimal
    Dim TotalCosto As Decimal
    Dim TotalUtilidad As Decimal
    Dim TotalCFP As Decimal
    Dim Totales As String
    Dim Cantidad As String
    Dim CodSuc As Integer
    Dim TTotalPrecioPub As Decimal
    Dim TTotalPrecioVta As Decimal
    Dim TTotalDescuento As Decimal
    Dim TTotalCosto As Decimal
    Dim TTotalUtilidad As Decimal
    Dim TTotalCFP As Decimal
    Dim mcurImptexProv As Decimal
    Dim mcurImptexSuc As Decimal
    Dim mcurUtilxSuc As Decimal
    Dim mintRenProv As Integer
    Dim mintRenSuc As Integer
    Dim mintColSuc As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim mintColUtSuc As Integer

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.chkMostrarCFP = New System.Windows.Forms.CheckBox()
        Me.chkMostrarCodArtProv = New System.Windows.Forms.CheckBox()
        Me.chkMostrarCodAnt = New System.Windows.Forms.CheckBox()
        Me.chkTodaslasSucursales = New System.Windows.Forms.CheckBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(6, 13)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(330, 64)
        Me.txtMensaje.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'chkMostrarCFP
        '
        Me.chkMostrarCFP.BackColor = System.Drawing.SystemColors.Control
        Me.chkMostrarCFP.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMostrarCFP.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkMostrarCFP.Location = New System.Drawing.Point(223, 26)
        Me.chkMostrarCFP.Margin = New System.Windows.Forms.Padding(2)
        Me.chkMostrarCFP.Name = "chkMostrarCFP"
        Me.chkMostrarCFP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMostrarCFP.Size = New System.Drawing.Size(84, 15)
        Me.chkMostrarCFP.TabIndex = 6
        Me.chkMostrarCFP.Text = "CFP"
        Me.ToolTip1.SetToolTip(Me.chkMostrarCFP, "Muestra CFP")
        Me.chkMostrarCFP.UseVisualStyleBackColor = False
        '
        'chkMostrarCodArtProv
        '
        Me.chkMostrarCodArtProv.BackColor = System.Drawing.SystemColors.Control
        Me.chkMostrarCodArtProv.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMostrarCodArtProv.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkMostrarCodArtProv.Location = New System.Drawing.Point(12, 41)
        Me.chkMostrarCodArtProv.Margin = New System.Windows.Forms.Padding(2)
        Me.chkMostrarCodArtProv.Name = "chkMostrarCodArtProv"
        Me.chkMostrarCodArtProv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMostrarCodArtProv.Size = New System.Drawing.Size(206, 32)
        Me.chkMostrarCodArtProv.TabIndex = 5
        Me.chkMostrarCodArtProv.Text = "Mostrar Código Articulo del Proveedor"
        Me.ToolTip1.SetToolTip(Me.chkMostrarCodArtProv, "Muestra el Código del Articulo del Proveedor")
        Me.chkMostrarCodArtProv.UseVisualStyleBackColor = False
        '
        'chkMostrarCodAnt
        '
        Me.chkMostrarCodAnt.BackColor = System.Drawing.SystemColors.Control
        Me.chkMostrarCodAnt.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMostrarCodAnt.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkMostrarCodAnt.Location = New System.Drawing.Point(12, 13)
        Me.chkMostrarCodAnt.Margin = New System.Windows.Forms.Padding(2)
        Me.chkMostrarCodAnt.Name = "chkMostrarCodAnt"
        Me.chkMostrarCodAnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMostrarCodAnt.Size = New System.Drawing.Size(194, 28)
        Me.chkMostrarCodAnt.TabIndex = 4
        Me.chkMostrarCodAnt.Text = "Mostrar Código Anterior"
        Me.ToolTip1.SetToolTip(Me.chkMostrarCodAnt, "Muestra el Código Anterior del Articulo")
        Me.chkMostrarCodAnt.UseVisualStyleBackColor = False
        '
        'chkTodaslasSucursales
        '
        Me.chkTodaslasSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodaslasSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodaslasSucursales.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTodaslasSucursales.Location = New System.Drawing.Point(16, 13)
        Me.chkTodaslasSucursales.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodaslasSucursales.Name = "chkTodaslasSucursales"
        Me.chkTodaslasSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodaslasSucursales.Size = New System.Drawing.Size(211, 23)
        Me.chkTodaslasSucursales.TabIndex = 0
        Me.chkTodaslasSucursales.Text = "Todas las Sucursales"
        Me.ToolTip1.SetToolTip(Me.chkTodaslasSucursales, "Muestra Todas las Sucursales")
        Me.chkTodaslasSucursales.UseVisualStyleBackColor = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.txtMensaje)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(18, 236)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(346, 91)
        Me.Frame4.TabIndex = 14
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Texto Adicional"
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.chkMostrarCFP)
        Me.Frame3.Controls.Add(Me.chkMostrarCodArtProv)
        Me.Frame3.Controls.Add(Me.chkMostrarCodAnt)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(18, 153)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(346, 78)
        Me.Frame3.TabIndex = 13
        Me.Frame3.TabStop = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.dtpFechaInicial)
        Me.Frame2.Controls.Add(Me.dtpFechaFinal)
        Me.Frame2.Controls.Add(Me.Label2)
        Me.Frame2.Controls.Add(Me.Label3)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(18, 89)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(346, 46)
        Me.Frame2.TabIndex = 10
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(73, 18)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(97, 20)
        Me.dtpFechaInicial.TabIndex = 2
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(240, 18)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(96, 20)
        Me.dtpFechaFinal.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(26, 21)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(49, 17)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Desde"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(196, 23)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(40, 17)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Hasta"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dbcSucursal)
        Me.Frame1.Controls.Add(Me.chkTodaslasSucursales)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(18, 10)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(346, 74)
        Me.Frame1.TabIndex = 8
        Me.Frame1.TabStop = False
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(73, 38)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(223, 21)
        Me.dbcSucursal.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(14, 41)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(62, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Sucursal : "
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(136, 344)
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
        Me.btnImprimir.Location = New System.Drawing.Point(21, 344)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 78
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(251, 345)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 77
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmVtasVentasporResurtir
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(380, 401)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(310, 174)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasVentasporResurtir"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Ventas por Resurtir"
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Sub CierraInstanciasdeExcel(ByRef Tipo As Integer)
        If Tipo = 1 Then
            objLibro.Close()
            ObjExcel.Quit()
        End If
        If ObjExcel Is Nothing Then ObjExcel = Nothing
        If objLibro Is Nothing Then objLibro = Nothing
        If objHoja Is Nothing Then objHoja = Nothing
    End Sub

    Function DevuelveQuery() As String
        On Error GoTo Err_Renamed
        Dim Sql As String

        '''Sql = "SELECT Vta.CodSucursal,S.DescAlmacen,Vta.CodGrupo,G.DescGrupo,Vta.Familia,Vta.DescFamilia," & _
        '"Vta.CodArticulo,Substring(Vta.DescArticulo,1,50) as DescArticulo," & _
        '"Convert(char(1), Vta.OrigenAnt) + '-' + right('00000'+ltrim(rtrim(Convert(char(5), Vta.CodigoAnt))),5) as CodigoAnt," & _
        '"Vta.CodigoArticuloProv, Vta.Cantidad - Vta.CantidadDev as CantidadVta, Vta.PrecioLista as PrecioPublico, Vta.PrecioReal as PrecioVenta," & _
        '"Vta.Descuento as Descuento, Vta.CostoVenta as Costo, Vta.UtilidadxArt as Utilidad, (Vta.UtilidadxArt/Vta.PrecioReal)*100 as Margen," & _
        '"Vta.FechaVenta, Vta.TipoCambio, Case When Vta.PesosFijos = 1 Then 'P' Else 'D' End as MonedaArt " & _
        '"FROM VENTAS_SALIDAMCIA('" & Format(dtpFechaInicial, C_FORMATFECHAGUARDAR) & "','" & Format(dtpFechaFinal, C_FORMATFECHAGUARDAR) & "') Vta " & _
        '"Inner Join CatAlmacen S (Nolock) On Vta.CodSucursal = S.CodAlmacen " & _
        '"Left Outer Join CatGrupos G (Nolock) On Vta.CodGrupo = G.CodGrupo " & _
        '"Where Vta.Tipo <> 'R' AND (Vta.Cantidad - Vta.CantidadDev) <> 0 " & IIf(chkTodaslasSucursales.Value = vbChecked, "", "AND Vta.CodSucursal = " & intCodSucursal & " ") & _
        '"Order By Vta.CodGrupo,Vta.DescFamilia,Vta.CodArticulo"

        Sql = "SELECT Vta.CodSucursal,S.DescAlmacen,Vta.CodGrupo,G.DescGrupo,Vta.Familia,Vta.DescFamilia," & "Vta.CodArticulo,Substring(Vta.DescArticulo,1,50) as DescArticulo," & "Convert(char(1), Vta.OrigenAnt) + '-' + right('00000'+ltrim(rtrim(Convert(char(5), Vta.CodigoAnt))),5) as CodigoAnt," & "Vta.CodigoArticuloProv, Vta.Cantidad - Vta.CantidadDev as CantidadVta, Vta.PrecioLista as PrecioPublico, (Vta.PrecioReal * (Vta.Cantidad - Vta.CantidadDev)) as PrecioVenta," & "(Vta.Descuento*(1+(Vta.PorcIva/100)) * (Vta.Cantidad - Vta.CantidadDev)) as Descuento, (Vta.CostoVenta * (Vta.Cantidad - Vta.CantidadDev)) as Costo, " & "((Vta.UtilidadxArt) * (Vta.Cantidad - Vta.CantidadDev)) as Utilidad, (Vta.UtilidadxArt/Vta.PrecioReal)*100 as Margen," & "Vta.FechaVenta, Vta.TipoCambio, Case When Vta.PesosFijos = 1 Then 'P' Else 'D' End as MonedaArt, " & "Vta.CodProveedor, P.DescProvAcreed, A.CodAlmacenOrigen as OrigenAnt, A.CostoFacturaPesos as CFP " & "FROM VENTAS_SALIDAMCIA('" & Format(dtpFechaInicial.Value, C_FORMATFECHAGUARDAR) & "','" & Format(dtpFechaFinal.Value, C_FORMATFECHAGUARDAR) & "') Vta " & "Inner Join CatAlmacen S (Nolock) On Vta.CodSucursal = S.CodAlmacen " & "Left Outer Join CatGrupos G (Nolock) On Vta.CodGrupo = G.CodGrupo " & "Inner Join CatArticulos A (Nolock) On Vta.CodArticulo = A.CodArticulo " & "Inner Join CatProvAcreed P (Nolock) On Vta.CodProveedor = P.CodProvAcreed " & "Where Vta.Tipo <> 'R' AND (Vta.Cantidad - Vta.CantidadDev) <> 0 " & IIf(chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked, "", "AND Vta.CodSucursal = " & intCodSucursal & " ") & "Order By Vta.CodSucursal, P.DescProvAcreed, Vta.CodArticulo"

        DevuelveQuery = Sql

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    ''' MODIFIC.-  SE CAMBIO EL GRUPO POR LA DESCRIPCION DEL PROVEEDOR
    ''' REESTRUCTURACION DEL REPORTE 31MAYO2005 - MAVF
    Sub EncabezadoGrupo()
        On Error GoTo Err_Renamed

        With objHoja
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "Prov: "
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 8.14
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).MergeCells = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1))._Default = Trim(RsGral.Fields("DescProvACreed").Value)
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            Renglon = Renglon + 1
            Columna = 2
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "CODIGO"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 9
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "DESCRIPCION"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 35
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            If chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "CODIGO PROVEEDOR"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 12.43
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
            End If
            If chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked Then
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "CODIGO ANTERIOR"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 10.29
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
            End If

            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "ORIGEN"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 7.86
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With

            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "CANT"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 5.57
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "PRECIO PUB"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon + 1, Columna), .Cells._Default(Renglon + 1, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon + 2, Columna), .Cells._Default(Renglon + 2, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 11.5
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "PRECIO VTA"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            If Not MostrarCostoyUtilidad Then
                .Range(.Cells._Default(Renglon - 1, Columna), .Cells._Default(Renglon - 1, Columna)).Select()
                .Range(.Cells._Default(Renglon - 1, Columna), .Cells._Default(Renglon - 1, Columna))._Default = "U N I T A R I O S"
                .Range(.Cells._Default(Renglon - 1, Columna), .Cells._Default(Renglon - 1, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                With .Range(.Cells._Default(Renglon - 1, Columna), .Cells._Default(Renglon - 1, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
            End If
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 11.5
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "DESCUENTO"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            If Not MostrarCostoyUtilidad Then
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon + 1, Columna), .Cells._Default(Renglon + 1, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon + 2, Columna), .Cells._Default(Renglon + 2, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            End If
            If MostrarCostoyUtilidad Then
                .Range(.Cells._Default(Renglon - 1, Columna), .Cells._Default(Renglon - 1, Columna)).Select()
                .Range(.Cells._Default(Renglon - 1, Columna), .Cells._Default(Renglon - 1, Columna)).MergeCells = True
                .Range(.Cells._Default(Renglon - 1, Columna), .Cells._Default(Renglon - 1, Columna))._Default = "U N I T A R I O S"
                .Range(.Cells._Default(Renglon - 1, Columna), .Cells._Default(Renglon - 1, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                With .Range(.Cells._Default(Renglon - 1, Columna), .Cells._Default(Renglon - 1, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
            End If
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 11.5
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            If MostrarCostoyUtilidad Then
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "COSTO"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous

                If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon + 1, Columna), .Cells._Default(Renglon + 1, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon + 2, Columna), .Cells._Default(Renglon + 2, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End If

                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 11.5
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With

                '''COLUMNA NUEVA - COSTO FACTURA PESOS
                If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Columna = Columna + 1
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "CTO FACT $"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon + 1, Columna), .Cells._Default(Renglon + 1, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon + 2, Columna), .Cells._Default(Renglon + 2, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 11
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                End If

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "UTILIDAD"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 10
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = "MARGEN"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).RowHeight = 24
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).WrapText = True
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Interior.ColorIndex = 15
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).ColumnWidth = 10
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
            End If
        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Encabezado()
        On Error GoTo Err_Renamed
        Dim Columna As Integer
        With objHoja
            .Range("C1").FormulaR1C1 = Trim(gstrCorpoNOMBREEMPRESA)
            .Range("C1:G1").Select()
            .Range("C1:G1").MergeCells = True
            .Range("C1:G1").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C1:G1").Font
                .Bold = True
                .Size = 12
                .Name = "Arial"
            End With
            .Range("C2").FormulaR1C1 = "Reporte de Ventas para Resurtir"
            .Range("C2:G2").Select()
            .Range("C2:G2").MergeCells = True
            .Range("C2:G2").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C2:G2").Font
                .Bold = False
                .Size = 11
                .Name = "Arial"
            End With
            .Range("C3").FormulaR1C1 = "Desde el " & Format(dtpFechaInicial.Value, "dd/mmm/yyyy") & " Hasta el " & Format(dtpFechaFinal.Value, "dd/mmm/yyyy")
            .Range("C3:G3").Select()
            .Range("C3:G3").MergeCells = True
            .Range("C3:G3").HorizontalAlignment = Excel.Constants.xlCenter
            With .Range("C3:G3").Font
                .Bold = False
                .Size = 10
                .Name = "Arial"
            End With
            .Range("A4").FormulaR1C1 = "Fecha: " & Format(Today, "dd/mmm/yyyy")
            .Range("A4:B4").Select()
            .Range("A4:B4").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A4:B4").Font
                .Bold = False
                .Size = 9
                .Name = "Arial"
            End With
            .Range("A5").FormulaR1C1 = "Mensaje: "
            .Range("A5").Select()
            .Range("A5").HorizontalAlignment = Excel.Constants.xlLeft
            With .Range("A5").Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            If Trim(txtMensaje.Text) <> "" Then
                .Range("B5").FormulaR1C1 = Trim(QuitaEnter(txtMensaje.Text))
                .Range("B5:J5").Select()
                .Range("B5:J5").MergeCells = True
                .Range("B5:J5").HorizontalAlignment = Excel.Constants.xlLeft
                With .Range("B5:J5").Font
                    .Bold = False
                    .Size = 9
                    .Name = "Arial"
                End With
            End If

            '''        .Range("A7").FormulaR1C1 = "Sucursal: "
            '''        .Range("A7").Select
            '''        .Range("A7").HorizontalAlignment = xlLeft
            '''        With .Range("A7").Font
            '''            .Bold = True
            '''            .Size = 9
            '''            .Name = "Arial"
            '''        End With

            '''        .Range("B7").FormulaR1C1 = IIf(chkTodaslasSucursales.Value = vbChecked, "Todas", UCase(Left(dbcSucursal.text, 1)) & LCase(Mid(dbcSucursal.text, 2, 39)))
            '''        .Range("B7:C7").Select
            '''        .Range("B7:C7").MergeCells = True
            '''        .Range("B7").HorizontalAlignment = xlLeft
            '''        With .Range("B7").Font
            '''            .Bold = True
            '''            .Size = 9
            '''            .Name = "Arial"
            '''        End With

        End With
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    'Function ArchivoAbierto() As Boolean
    '    On Error GoTo Err
    '    Dim Archivo As String
    '    If Dir(gstrCorpoDriveLocal & "\Sistema\", vbDirectory + vbHidden) = "" Then
    '        MsgBox "No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
    '        ArchivoAbierto = True
    '        Exit Function
    '    End If
    '    Archivo = "VR" & CStr(Format(Month(Date), "00")) & CStr(Format(Day(Date), "00")) & Right(CStr(Format(Year(Date), "00")), 2) & ".xls"
    '    If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\", vbDirectory) = "" Then
    '        MkDir gstrCorpoDriveLocal & "\Sistema\Informes\"
    '    End If
    '    If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo, vbArchive) <> "" Then
    '        Kill gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo
    '    End If
    '    Set ObjExcel = CreateObject("Excel.Application")
    '    Set objLibro = ObjExcel.Workbooks.Add
    '    Set objHoja = objLibro.ActiveSheet
    '    ObjExcel.Visible = False
    '    objLibro.Sheets(1).Select
    '    Set objHoja = objLibro.ActiveSheet
    '    objLibro.ActiveSheet.Name = "Ventas para Resurtir"
    '    objLibro.SaveAs gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", _
    ''    FileFormat:=xlNormal, Password:="", writerespassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    '    CierraInstanciasdeExcel
    '    ArchivoAbierto = False
    'Err:
    '    If Err.Number = 70 Then
    '        MsgBox "No se puede generar un nuevo archivo hasta que el anterior este cerrado.", vbCritical + vbOKOnly, gstrNombCortoEmpresa
    '        CierraInstanciasdeExcel
    '        ArchivoAbierto = True
    '    ElseIf Err.Number <> 0 Then
    '        ModEstandar.MostrarError
    '        CierraInstanciasdeExcel
    '        ArchivoAbierto = True
    '    End If
    'End Function

    Sub EnviaExcel()
        On Error GoTo Err_Renamed
        Dim Archivo As String
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        System.Windows.Forms.Application.DoEvents()
        If Dir(gstrCorpoDriveLocal & "\Sistema\", FileAttribute.Directory + FileAttribute.Hidden) = "" Then
            MsgBox("No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        Archivo = "RV" & CStr(Format(Month(Today), "00")) & CStr(Format((Today), "00")) & (CStr(Format(Year(Today), "00"))) & ".xls"
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\", FileAttribute.Directory) = "" Then
            MkDir(gstrCorpoDriveLocal & "\Sistema\Informes\")
        End If
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo, FileAttribute.Archive) <> "" Then
            Kill(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo)
        End If

        ObjExcel = CreateObject("Excel.Application")
        objLibro = ObjExcel.Workbooks.Add
        objHoja = objLibro.ActiveSheet
        ObjExcel.Visible = False
        objLibro.Sheets(1).Select()
        objHoja = objLibro.ActiveSheet
        objLibro.ActiveSheet.Name = "Ventas para Resurtir"
        Encabezado()
        LlenaDatos()
        objLibro.SaveAs(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        System.Windows.Forms.Application.DoEvents()

        Select Case MsgBox("Se ha creado el archivo " & Archivo & " ¿Desea abrirlo?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
            Case MsgBoxResult.Yes
                ObjExcel.Visible = True
                ObjExcel = Nothing
                objLibro = Nothing
                objHoja = Nothing
            Case MsgBoxResult.No Or MsgBoxResult.Cancel
                CierraInstanciasdeExcel(1)
        End Select

Err_Renamed:
        If Err.Number = 70 Then
            MsgBox("No se puede generar un nuevo archivo hasta que el anterior este cerrado.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            CierraInstanciasdeExcel(2)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        ElseIf Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Imprime()
        On Error GoTo ImprimeErr
        If Not ValidaDatos() Then Exit Sub
        gStrSql = DevuelveQuery()
        ModEstandar.BorraCmd()
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            EnviaExcel()
        Else
            MsgBox("No existe información por mostrar en este periodo de fechas, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        End If
        Cmd.CommandTimeout = 90

ImprimeErr:
        If Err.Number <> 0 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
            FueraChange = False
        End If
    End Sub

    Sub Limpiar()
        Nuevo()
        chkTodaslasSucursales.Focus()
    End Sub

    Private Sub EncabezadoSucursal(ByRef Col As Integer, ByRef Ren As Integer, ByRef objH As Excel.Worksheet, ByRef ColProv As Integer, ByRef ColProvUt As Integer)
        With objH
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = "Sucursal: "
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With

            .Range(.Cells._Default(Ren, Col + 1), .Cells._Default(Ren, Col + 1))._Default = RsGral.Fields("DescAlmacen").Value
            .Range(.Cells._Default(Ren, Col + 1), .Cells._Default(Ren, Col + 1)).Select()
            .Range(.Cells._Default(Ren, Col + 1), .Cells._Default(Ren, Col + 1)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Ren, Col + 1), .Cells._Default(Ren, Col + 1)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With

            '''simula importe total por sucursal para PrecioPublico
            .Range(.Cells._Default(Ren, ColProv), .Cells._Default(Ren, ColProv))._Default = "1"
            .Range(.Cells._Default(Ren, ColProv), .Cells._Default(Ren, ColProv)).Select()
            .Range(.Cells._Default(Ren, ColProv), .Cells._Default(Ren, ColProv)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Ren, ColProv), .Cells._Default(Ren, ColProv)).NumberFormat = "###,##0.00"
            With .Range(.Cells._Default(Ren, ColProv), .Cells._Default(Ren, ColProv)).Font
                .Bold = True
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
                .Size = 9
                .Name = "Arial"
            End With
            If MostrarCostoyUtilidad Then
                '''simula importe total por sucursal para Utilidad
                .Range(.Cells._Default(Ren, ColProvUt), .Cells._Default(Ren, ColProvUt))._Default = "1"
                .Range(.Cells._Default(Ren, ColProvUt), .Cells._Default(Ren, ColProvUt)).Select()
                .Range(.Cells._Default(Ren, ColProvUt), .Cells._Default(Ren, ColProvUt)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Ren, ColProvUt), .Cells._Default(Ren, ColProvUt)).NumberFormat = "###,##0.00"
                With .Range(.Cells._Default(Ren, ColProvUt), .Cells._Default(Ren, ColProvUt)).Font
                    .Bold = True
                    .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
                    .Size = 9
                    .Name = "Arial"
                End With
            End If

        End With

    End Sub

    Private Sub TotalesxProveedor(ByRef Col As Integer, ByRef Ren As Integer, ByRef objH As Excel.Worksheet)
        Dim lCol As Integer
        Dim lRen As Integer
        Dim lColUt As Integer
        Dim lRenUt As Integer
        Dim lRenglonP As Integer

        mintRenProv = Ren
        With objH
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = "Total: "
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With

            Col = Col + 2
            '''columna y renglon para el porcentaje de vta con respecto a la sucursal
            lCol = Col + 1
            lRen = Ren

            '.Range(.Cells(Ren, Col), .Cells(Ren, Col)) = SubTotalPrecioPub
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = ""
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Ren + 2, Col), .Cells._Default(Ren + 2, Col)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                .Size = 8
                .Name = "Arial"
            End With
            Col = Col + 1
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = SubTotalPrecioVta
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                .Size = 8
                .Name = "Arial"
            End With
            Col = Col + 1
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = SubTotalDescuento
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            If Not MostrarCostoyUtilidad Then
                .Range(.Cells._Default(Ren + 2, Col), .Cells._Default(Ren + 2, Col)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            End If
            With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                .Size = 8
                .Name = "Arial"
            End With

            If MostrarCostoyUtilidad Then
                Col = Col + 1
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = SubTotalCosto
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                    .Range(.Cells._Default(Ren + 2, Col), .Cells._Default(Ren + 2, Col)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End If
                With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                '''NVA col CFP
                If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Col = Col + 1
                    '.Range(.Cells(Ren, Col), .Cells(Ren, Col)) = SubTotalCFP
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = ""
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Ren + 2, Col), .Cells._Default(Ren + 2, Col)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                End If

                Col = Col + 1
                lColUt = Col
                lRenUt = Ren

                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = SubTotalUtilidad
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Col = Col + 1
                Margen = System.Math.Round((SubTotalUtilidad / SubTotalPrecioVta) * 100, 2)
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = VB6.Format(Margen, "###,##0.00") & "%"
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
            End If

            '''formula para margen de ventas
            lRenglonP = (lRen - mintRenSuc) + 1
            .Range(.Cells._Default(lRen + 1, lCol), .Cells._Default(lRen + 1, lCol)).FormulaR1C1 = "=(R[-1]C[0] / R[-" & lRenglonP & "]C[0])"
            .Range(.Cells._Default(lRen + 1, lCol), .Cells._Default(lRen + 1, lCol)).Select()
            .Range(.Cells._Default(lRen + 1, lCol), .Cells._Default(lRen + 1, lCol)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(lRen + 1, lCol), .Cells._Default(lRen + 1, lCol)).NumberFormat = "##0.00%"
            With .Range(.Cells._Default(lRen + 1, lCol), .Cells._Default(lRen + 1, lCol)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            If MostrarCostoyUtilidad Then
                '''formula para margen de utilidad
                lRenglonP = (lRenUt - mintRenSuc) + 1
                .Range(.Cells._Default(lRenUt + 1, lColUt), .Cells._Default(lRenUt + 1, lColUt)).FormulaR1C1 = "=(R[-1]C[0] / R[-" & lRenglonP & "]C[0])"
                .Range(.Cells._Default(lRenUt + 1, lColUt), .Cells._Default(lRenUt + 1, lColUt)).Select()
                .Range(.Cells._Default(lRenUt + 1, lColUt), .Cells._Default(lRenUt + 1, lColUt)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(lRenUt + 1, lColUt), .Cells._Default(lRenUt + 1, lColUt)).NumberFormat = "##0.00%"
                With .Range(.Cells._Default(lRen + 1, lColUt), .Cells._Default(lRenUt + 1, lColUt)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
            End If

        End With

    End Sub

    Private Sub TotalesxSucursal(ByRef Col As Integer, ByRef Ren As Integer, ByRef objH As Excel.Worksheet)
        Dim lRenglonS As Integer
        Dim lColumnaS As Integer

        With objH
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col + 1)).Select()
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col + 1)).MergeCells = True
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col + 1))._Default = "Total x Suc: "
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col + 1)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col + 1)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            Col = Col + 2
            '.Range(.Cells(Ren, Col), .Cells(Ren, Col)) = TotalPrecioPub
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = ""
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
            With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Col = Col + 1
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = TotalPrecioVta
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
            With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Col = Col + 1
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = TotalDescuento
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
            With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            If MostrarCostoyUtilidad Then
                Col = Col + 1
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = TotalCosto
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
                With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With

                '''NVA col COSTO FACTURA PESOS
                If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Col = Col + 1
                    '.Range(.Cells(Ren, Col), .Cells(Ren, Col)) = TotalCFP
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = ""
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
                    With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                End If

                Col = Col + 1
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = TotalUtilidad
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).NumberFormat = "###,##0.00"
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
                With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With

                Col = Col + 1
                Margen = System.Math.Round((TotalUtilidad / TotalPrecioVta) * 100, 2)
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col))._Default = VB6.Format(Margen, "###,##0.00") & "%"
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Select()
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
                With .Range(.Cells._Default(Ren, Col), .Cells._Default(Ren, Col)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
            Else
                '''          '''NVA col COSTO FACTURA PESOS
                '''          If chkMostrarCFP.Value = vbChecked Then
                '''              Col = Col + 1
                '''              .Range(.Cells(Ren, Col), .Cells(Ren, Col)) = TotalCFP
                '''              .Range(.Cells(Ren, Col), .Cells(Ren, Col)).Select
                '''              .Range(.Cells(Ren, Col), .Cells(Ren, Col)).NumberFormat = "###,##0.00"
                '''              .Range(.Cells(Ren, Col), .Cells(Ren, Col)).HorizontalAlignment = xlRight
                '''              .Range(.Cells(Ren, Col), .Cells(Ren, Col)).Borders(xlEdgeTop).LineStyle = xlContinuous
                '''              .Range(.Cells(Ren, Col), .Cells(Ren, Col)).Borders(xlEdgeBottom).Weight = xlMedium
                '''              With .Range(.Cells(Ren, Col), .Cells(Ren, Col)).Font
                '''                  .Bold = True
                '''                  .Size = 8
                '''                  .Name = "Arial"
                '''              End With
                '''          End If
            End If
            '''fija el total de ventas de la sucursal a nivel de la descripcion de la sucursal para que las operaciones de %
            '''se calculen automaticamente con la formula establecida
            lColumnaS = (Col - mintColSuc)
            'UPGRADE_WARNING: Couldn't resolve default property of object objH.Range().Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Range(.Cells._Default(mintRenSuc, mintColSuc), .Cells._Default(mintRenSuc, mintColSuc)).Value = .Range(.Cells._Default(Ren, Col - lColumnaS), .Cells._Default(Ren, Col - lColumnaS)).Value
            .Range(.Cells._Default(mintRenSuc, mintColSuc), .Cells._Default(mintRenSuc, mintColSuc)).Select()
            .Range(.Cells._Default(mintRenSuc, mintColSuc), .Cells._Default(mintRenSuc, mintColSuc)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(mintRenSuc, mintColSuc), .Cells._Default(mintRenSuc, mintColSuc)).NumberFormat = "###,##0.00"
            With .Range(.Cells._Default(mintRenSuc, mintColSuc), .Cells._Default(mintRenSuc, mintColSuc)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With

            If MostrarCostoyUtilidad Then
                '''fija la utilidad de la sucursal a nivel de la descripcion de la sucursal para que las operaciones de %
                '''se calculen automaticamente con la formula establecida
                lColumnaS = (Col - 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object objH.Range().Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range(.Cells._Default(mintRenSuc, Col - 1), .Cells._Default(mintRenSuc, Col - 1)).Value = .Range(.Cells._Default(Ren, Col - 1), .Cells._Default(Ren, Col - 1)).Value
                .Range(.Cells._Default(mintRenSuc, Col - 1), .Cells._Default(mintRenSuc, Col - 1)).Select()
                .Range(.Cells._Default(mintRenSuc, Col - 1), .Cells._Default(mintRenSuc, Col - 1)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(mintRenSuc, Col - 1), .Cells._Default(mintRenSuc, Col - 1)).NumberFormat = "###,##0.00"
                With .Range(.Cells._Default(mintRenSuc, Col - 1), .Cells._Default(mintRenSuc, Col - 1)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
            End If
        End With

    End Sub

    Private Sub GranTotal(ByRef Col As Integer, ByRef Ren As Integer, ByRef objH As Excel.Worksheet)

        With objH
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).MergeCells = True
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1))._Default = "Gran Total: "
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).HorizontalAlignment = Excel.Constants.xlLeft
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna + 1)).Font
                .Bold = True
                .Size = 9
                .Name = "Arial"
            End With
            Columna = Columna + 2
            '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = TTotalPrecioPub
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = ""
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlDouble
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = TTotalPrecioVta
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlDouble
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            Columna = Columna + 1
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = TTotalDescuento
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlDouble
            .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
            With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                .Bold = True
                .Size = 8
                .Name = "Arial"
            End With
            If MostrarCostoyUtilidad Then
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = TTotalCosto
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlDouble
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With

                '''NVA COLUMNA COSTO FACTURA PESOS
                If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Columna = Columna + 1
                    '.Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = TTotalCFP
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = ""
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlDouble
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Bold = True
                        .Size = 8
                        .Name = "Arial"
                    End With
                End If

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = TTotalUtilidad
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlDouble
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                Margen = System.Math.Round((TTotalUtilidad / TTotalPrecioVta) * 100, 2)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = VB6.Format(Margen, "###,##0.00") & "%"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlDouble
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Name = "Arial"
                End With
                '''      Else
                '''          '''NVA COLUMNA COSTO FACTURA PESOS
                '''          If chkMostrarCFP.Value = vbChecked Then
                '''              Columna = Columna + 1
                '''              .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = TTotalCFP
                '''              .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Select
                '''              .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).NumberFormat = "###,##0.00"
                '''              .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlRight
                '''              .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeTop).LineStyle = xlDouble
                '''              .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).Weight = xlMedium
                '''              With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
                '''                  .Bold = True
                '''                  .Size = 8
                '''                  .Name = "Arial"
                '''              End With
                '''          End If
            End If
        End With

    End Sub

    Sub LlenaDatos()
        On Error GoTo Err_Renamed
        Dim RenRecorridos As Integer
        Dim CodGrupo As Integer
        Dim I As Integer
        Dim Rango As String
        Dim Familia As String

        Renglon = 9
        Familia = "/*"
        Columna = 1
        CodGrupo = 0
        CodSuc = 0

        SubTotalPrecioPub = 0
        SubTotalPrecioVta = 0
        SubTotalDescuento = 0
        SubTotalCosto = 0
        SubTotalUtilidad = 0
        SubTotalCFP = 0
        Margen = 0

        TotalPrecioPub = 0
        TotalPrecioVta = 0
        TotalDescuento = 0
        TotalCosto = 0
        TotalUtilidad = 0
        TotalCFP = 0

        TTotalPrecioPub = 0
        TTotalPrecioVta = 0
        TTotalDescuento = 0
        TTotalCosto = 0
        TTotalUtilidad = 0
        TTotalCFP = 0

        mcurImptexProv = 0
        mcurImptexSuc = 0
        mcurUtilxSuc = 0
        mintRenProv = 0
        mintColSuc = 0
        mintRenSuc = 0
        mintColUtSuc = 0

        With objHoja
            RsGral.MoveFirst()
            Do While Not RsGral.EOF
                If CodSuc = 0 Then
                    CodSuc = RsGral.Fields("CodSucursal").Value
                    Columna = 1
                    Renglon = 7

                    If chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        mintColSuc = 7 '''ojo 6
                        If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then mintColUtSuc = 11 Else mintColUtSuc = 10
                    ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        mintColSuc = 8 '''ojo 7
                        If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then mintColUtSuc = 12 Else mintColUtSuc = 11
                    ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                        mintColSuc = 8 ''' ojo 7
                        If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then mintColUtSuc = 12 Else mintColUtSuc = 11
                    ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                        mintColSuc = 9 '''ojo 8
                        If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then mintColUtSuc = 13 Else mintColUtSuc = 12
                    End If

                    mintRenSuc = Renglon
                    EncabezadoSucursal(Columna, Renglon, objHoja, mintColSuc, mintColUtSuc)
                    Columna = 1
                    Renglon = 9
                End If

                If CodGrupo = 0 Then
                    CodGrupo = RsGral.Fields("CodProveedor").Value
                    Columna = 1
                    EncabezadoGrupo()
                ElseIf CodGrupo <> RsGral.Fields("CodProveedor").Value Then
                    If chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        Columna = 4
                    ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        Columna = 5
                    ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                        Columna = 5
                    ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                        Columna = 6
                    End If

                    Renglon = Renglon + 1
                    TotalesxProveedor(Columna, Renglon, objHoja)

                    TotalPrecioPub = TotalPrecioPub + SubTotalPrecioPub
                    TotalPrecioVta = TotalPrecioVta + SubTotalPrecioVta
                    TotalDescuento = TotalDescuento + SubTotalDescuento
                    TotalCosto = TotalCosto + SubTotalCosto
                    TotalUtilidad = TotalUtilidad + SubTotalUtilidad
                    TotalCFP = TotalCFP + SubTotalCFP

                    TTotalPrecioPub = TTotalPrecioPub + SubTotalPrecioPub
                    TTotalPrecioVta = TTotalPrecioVta + SubTotalPrecioVta
                    TTotalDescuento = TTotalDescuento + SubTotalDescuento
                    TTotalCosto = TTotalCosto + SubTotalCosto
                    TTotalUtilidad = TTotalUtilidad + SubTotalUtilidad
                    TTotalCFP = TTotalCFP + SubTotalCFP

                    SubTotalPrecioPub = 0
                    SubTotalPrecioVta = 0
                    SubTotalDescuento = 0
                    SubTotalCosto = 0
                    SubTotalUtilidad = 0
                    SubTotalCFP = 0

                    CodGrupo = RsGral.Fields("CodProveedor").Value
                    Renglon = Renglon + 3
                    Columna = 1

                    If CodSuc <> RsGral.Fields("CodSucursal").Value Then
                        CodSuc = RsGral.Fields("CodSucursal").Value
                        Columna = 1
                        Renglon = Renglon + 1

                        TTotalPrecioPub = TTotalPrecioPub + SubTotalPrecioPub
                        TTotalPrecioVta = TTotalPrecioVta + SubTotalPrecioVta
                        TTotalDescuento = TTotalDescuento + SubTotalDescuento
                        TTotalCosto = TTotalCosto + SubTotalCosto
                        TTotalUtilidad = TTotalUtilidad + SubTotalUtilidad
                        TTotalCFP = TTotalCFP + SubTotalCFP

                        '*********************************************************
                        '*********************************************************
                        '''Totales por Sucursal
                        If chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                            Columna = 4
                        ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                            Columna = 5
                        ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                            Columna = 5
                        ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                            Columna = 6
                        End If

                        TotalesxSucursal(Columna, Renglon, objHoja)
                        Columna = 1
                        Renglon = Renglon + 2
                        '*********************************************************
                        '*********************************************************

                        TotalPrecioPub = 0
                        TotalPrecioVta = 0
                        TotalDescuento = 0
                        TotalCosto = 0
                        TotalUtilidad = 0
                        TotalCFP = 0

                        If chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                            mintColSuc = 7 ''' ojo 6
                            If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then mintColUtSuc = 11 Else mintColUtSuc = 10
                        ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                            mintColSuc = 8 ''' ojo 7
                            If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then mintColUtSuc = 12 Else mintColUtSuc = 11
                        ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                            mintColSuc = 8 ''' ojo 7
                            If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then mintColUtSuc = 12 Else mintColUtSuc = 11
                        ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                            mintColSuc = 9 ''' ojo 8
                            If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then mintColUtSuc = 13 Else mintColUtSuc = 12
                        End If

                        mintRenSuc = Renglon
                        EncabezadoSucursal(Columna, Renglon, objHoja, mintColSuc, mintColUtSuc)
                        Columna = 1
                        Renglon = Renglon + 2
                        CodSuc = RsGral.Fields("CodSucursal").Value
                    End If

                    EncabezadoGrupo()
                End If

                '''DETALLE
                Renglon = Renglon + 1
                Columna = 2
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("CodArticulo").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Trim(RsGral.Fields("DescArticulo").Value)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                If chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Columna = Columna + 1
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("CodigoArticuloProv").Value
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                End If
                If chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked Then
                    Columna = Columna + 1
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("CodigoAnt").Value
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                End If

                '''COLUMNA NUEVA - ORIGEN
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("OrigenAnt").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("cantidadvta").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("PrecioPublico").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon + 1, Columna), .Cells._Default(Renglon + 1, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(Renglon + 2, Columna), .Cells._Default(Renglon + 2, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                SubTotalPrecioPub = SubTotalPrecioPub + RsGral.Fields("PrecioPublico").Value
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("PrecioVenta").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                SubTotalPrecioVta = SubTotalPrecioVta + RsGral.Fields("PrecioVenta").Value
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("Descuento").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                If Not MostrarCostoyUtilidad Then
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon + 1, Columna), .Cells._Default(Renglon + 1, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon + 2, Columna), .Cells._Default(Renglon + 2, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End If
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                SubTotalDescuento = SubTotalDescuento + RsGral.Fields("Descuento").Value

                If MostrarCostoyUtilidad Then
                    Columna = Columna + 1
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("Costo").Value
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                    If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                        .Range(.Cells._Default(Renglon + 1, Columna), .Cells._Default(Renglon + 1, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                        .Range(.Cells._Default(Renglon + 2, Columna), .Cells._Default(Renglon + 2, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    End If
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    SubTotalCosto = SubTotalCosto + RsGral.Fields("Costo").Value

                    '''NVA COLUMNA COSTO FACTURA PESOS
                    If chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked Then
                        Columna = Columna + 1
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("CFP").Value
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                        .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                        .Range(.Cells._Default(Renglon + 1, Columna), .Cells._Default(Renglon + 1, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                        .Range(.Cells._Default(Renglon + 2, Columna), .Cells._Default(Renglon + 2, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                        With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                            .Size = 8
                            .Name = "Arial"
                        End With
                        SubTotalCFP = SubTotalCFP + RsGral.Fields("CFP").Value
                    End If

                    Columna = Columna + 1
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("utilidad").Value
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                    SubTotalUtilidad = SubTotalUtilidad + RsGral.Fields("utilidad").Value
                    Columna = Columna + 1
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = VB6.Format(RsGral.Fields("Margen").Value, "###,##0.00") & "%"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                Else
                    '''                '''NVA COLUMNA COSTO FACTURA PESOS
                    '''                If chkMostrarCFP.Value = vbChecked Then
                    '''                   Columna = Columna + 1
                    '''                   .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = RsGral!CFP
                    '''                   .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Select
                    '''                   .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).NumberFormat = "###,##0.00"
                    '''                   .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlRight
                    '''                   .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
                    '''                   .Range(.Cells(Renglon + 1, Columna), .Cells(Renglon + 1, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
                    '''                   .Range(.Cells(Renglon + 2, Columna), .Cells(Renglon + 2, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
                    '''                   With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
                    '''                       .Size = 8
                    '''                       .Name = "Arial"
                    '''                   End With
                    '''                End If
                End If

                RsGral.MoveNext()
                If RsGral.EOF Then

                    TTotalPrecioPub = TTotalPrecioPub + SubTotalPrecioPub
                    TTotalPrecioVta = TTotalPrecioVta + SubTotalPrecioVta
                    TTotalDescuento = TTotalDescuento + SubTotalDescuento
                    TTotalCosto = TTotalCosto + SubTotalCosto
                    TTotalUtilidad = TTotalUtilidad + SubTotalUtilidad
                    TTotalCFP = TTotalCFP + SubTotalCFP

                    Renglon = Renglon + 1
                    If chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        Columna = 4
                    ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                        Columna = 5
                    ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                        Columna = 5
                    ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                        Columna = 6
                    End If

                    TotalesxProveedor(Columna, Renglon, objHoja)
                    Columna = 1

                    TotalPrecioPub = TotalPrecioPub + SubTotalPrecioPub
                    TotalPrecioVta = TotalPrecioVta + SubTotalPrecioVta
                    TotalDescuento = TotalDescuento + SubTotalDescuento
                    TotalUtilidad = TotalUtilidad + SubTotalUtilidad
                    TotalCosto = TotalCosto + SubTotalCosto
                    TotalCFP = TotalCFP + SubTotalCFP
                End If
            Loop

            '''Totales por Sucursal
            Renglon = Renglon + 3
            If chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                Columna = 4
            ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                Columna = 5
            ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                Columna = 5
            ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                Columna = 6
            End If

            TotalesxSucursal(Columna, Renglon, objHoja)
            Columna = 1

            '''Gran Total
            Renglon = Renglon + 3
            If chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                Columna = 4
            ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                Columna = 5
            ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                Columna = 5
            ElseIf chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked And chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked Then
                Columna = 6
            End If

            GranTotal(Columna, Renglon, objHoja)
            Columna = 1

            .Application.ActiveWindow.Zoom = 85
            .Range("A1").Select()
        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Sub Nuevo()
        chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        FueraChange = True
        dbcSucursal.Text = ""
        dbcSucursal.Enabled = False
        FueraChange = False
        dtpFechaInicial.Value = Today
        dtpFechaFinal.Value = Today
        txtMensaje.Text = ""
        chkMostrarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked
        chkMostrarCodArtProv.CheckState = System.Windows.Forms.CheckState.Checked

        If MostrarCostoyUtilidad Then
            chkMostrarCFP.Visible = True
            chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked
        Else
            chkMostrarCFP.CheckState = System.Windows.Forms.CheckState.Checked
            chkMostrarCFP.Visible = False
        End If
        mblnSalir = False

    End Sub

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            If intCodSucursal = 0 Then
                MsgBox("Proporcione una Sucursal, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                dbcSucursal.Focus()
                Exit Function
            End If
        End If
        Do While (sglTiempoCambio) <= 2.1
        Loop
        System.Windows.Forms.Application.DoEvents()
        If dtpFechaInicial.Value > dtpFechaFinal.Value Then
            MsgBox("La Fecha Inicial no Puede ser Mayor que la Fecha Final.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaInicial.Focus()
            Exit Function
        End If
        If dtpFechaInicial.Value > Now Then
            MsgBox("la Fecha Inicial no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaInicial.Focus()
            Exit Function
        End If
        If dtpFechaFinal.Value > Now Then
            MsgBox("la Fecha Final no Puede ser Mayor que la Fecha Actual.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dtpFechaFinal.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub chkMostrarCFP_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkMostrarCFP.Enter
        Pon_Tool()
    End Sub

    Private Sub chkMostrarCodAnt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkMostrarCodAnt.Enter
        Pon_Tool()
    End Sub

    Private Sub chkMostrarCodArtProv_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkMostrarCodArtProv.Enter
        Pon_Tool()
    End Sub

    Private Sub chkTodaslasSucursales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.CheckStateChanged
        If chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            FueraChange = True
            dbcSucursal.Text = "        [TODAS LAS SUCURSALES]..."
            dbcSucursal.Enabled = False
            FueraChange = False
        Else
            FueraChange = True
            dbcSucursal.Text = ""
            dbcSucursal.Enabled = True
            FueraChange = False
        End If
    End Sub

    Private Sub chkTodaslasSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursal" Then
        '    Exit Sub
        'End If
        'Nuevo
        gStrSql = "SELECT CodAlmacen,Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        If dbcSucursal.SelectedItem <> "" Then
            Call dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        End If
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen  FROM CatAlmacen   Where TipoAlmacen ='P'  ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursal)
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        tecla = eventArgs.KeyCode
        'If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
        chkTodaslasSucursales.Focus()
        'End If
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        intCodSucursal = 0
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%'  and TipoAlmacen ='P'  ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursal, gStrSql, intCodSucursal)
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

    Private Sub dtpFechaInicial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInicial.CursorChanged
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaInicial.Click
        'glTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaInicial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaInicial.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaInicial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaInicial.KeyPress
        'sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub frmVtasVentasporResurtir_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasVentasporResurtir_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasVentasporResurtir_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmVtasVentasporResurtir_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasVentasporResurtir_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)

        gStrSql = "Select * From CatUsuarios (Nolock) Where CodUsuario = " & gIntCodUsuario
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Tipo").Value = C_TADMIN Then
                MostrarCostoyUtilidad = True
            Else
                MostrarCostoyUtilidad = False
            End If
        End If
        Nuevo()

    End Sub

    Private Sub frmVtasVentasporResurtir_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        If mblnSalir Then
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSalir = False
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasVentasporResurtir_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class