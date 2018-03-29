Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmFactReportesFacturacionGlobalXSucursal
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REPORTE DE FACTURACION GLOBAL POR SUCURSAL                                                   *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :                                                                                                   *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents optPesos As System.Windows.Forms.RadioButton
    Public WithEvents optDolares As System.Windows.Forms.RadioButton
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents chkV As System.Windows.Forms.CheckBox
    Public WithEvents txtTextoAdicional As System.Windows.Forms.TextBox
    Public WithEvents dtpFechaInicial As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFinal As System.Windows.Forms.DateTimePicker
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents chkTodaslasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim intCodSucursal As Integer
    Dim tecla As Integer
    Dim rsReporte As ADODB.Recordset
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim Moneda As String

    '''Modificación.-  Agregar Facturación Especial
    Sub Imprime()

        Dim RptFactFacturacionGlobalXSucursal As New RptFactFacturacionGlobalXSucursal

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        Dim sql As String
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim PeriodoReporte As String
        Dim TextoAdicional As String
        Dim FechaInicial As String
        Dim FechaFinal As String
        Dim Servidor As String
        Dim BasedeDatos As String
        Dim TextoMoneda As String
        Dim I As Object
        Dim Pos As Integer
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
        NombreReporte = UCase("Facturacion Global por Sucursal")
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

            sql = "Select Case When rtrim(ltrim(isnull(norm.sucursal,''))) = '' and rtrim(ltrim(isnull(esp.sucursal,''))) <> '' then rtrim(ltrim(esp.sucursal)) when rtrim(ltrim(isnull(norm.sucursal,''))) <> '' and rtrim(ltrim(isnull(esp.sucursal,''))) = '' " & "then rtrim(ltrim(norm.sucursal)) when rtrim(ltrim(isnull(norm.sucursal,''))) <> '' and rtrim(ltrim(isnull(esp.sucursal,''))) <> '' then rtrim(ltrim(norm.sucursal)) end as sucursal,isnull(norm.cantidad,0) + isnull(esp.cantidad,0) as cantidad," & "isnull(norm.importe,0) + isnull(esp.importe,0) as importe,isnull(norm.descuento,0) + isnull(esp.descuento,0) as descuento,isnull(norm.subtotal,0) + isnull(esp.subtotal,0) as subtotal,isnull(norm.iva,0) + isnull(esp.iva,0) as iva,isnull(norm.total,0) + isnull(esp.total,0) as total " & "From " & "(select RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,suc.codalmacen,sum(det.cantidadadicional) as cantidad,sum(case when f.moneda = 'P' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) * f.tipocambio,1) end) as importe," & "sum(case when f.moneda = 'P' then round(f.descuento,2) else round(f.descuento * f.tipocambio,1) end) as descuento,sum(case when f.moneda = 'P' then round((f.subtotal + f.redondeo) - f.descuento,2) else round(((f.subtotal + f.redondeo) - f.descuento) * f.tipocambio,1) end) as subtotal,sum(case when f.moneda = 'P' then round(f.iva,2) else round(f.iva * f.tipocambio,1) end) as iva," & "sum(case when f.moneda = 'P' then  round(f.total + f.redondeo,2) else round((f.total + f.redondeo) * f.tipocambio,1) end) as total from (select foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,sum(cantidad) as cantidad,subtotal,redondeo,descuento,iva,total from facturas where tipofactura = 'N' group by foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,subtotal,redondeo,descuento,iva,total) f " & "inner join (select distinct foliofactura,codsucursal,estatus from movimientosventascab) cab on f.foliofactura = cab.foliofactura Inner Join (select foliofactura,sum(cantidadadicional) as cantidadadicional from movimientosventasdet group by foliofactura) det on f.foliofactura = det.foliofactura and cab.foliofactura = det.foliofactura inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen and cab.codsucursal = suc.codalmacen " & "where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' and cab.estatus <> 'C' group by suc.codalmacen,suc.descalmacen) Norm " & "full Join " & "(select RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,suc.codalmacen,sum(f.cantidad) as cantidad,sum(case when f.moneda = 'P' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) * f.tipocambio,1) end) as importe,sum(case when f.moneda = 'P' then round(f.descuento,2) else round(f.descuento * f.tipocambio,1) end) as descuento,sum(case when f.moneda = 'P' then round((f.subtotal + f.redondeo) - f.descuento,2) else round(((f.subtotal + f.redondeo) - f.descuento) * f.tipocambio,1) end) as subtotal," & "sum(case when f.moneda = 'P' then round(f.iva,2) else round(f.iva * f.tipocambio,1) end) as iva,sum(case when f.moneda = 'P' then  round(f.total + f.redondeo,2) else round((f.total + f.redondeo) * f.tipocambio,1) end) as total from (select foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,sum(cantidad) as cantidad,subtotal,redondeo,descuento,iva,total from facturas where tipofactura = 'E' group by foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,subtotal,redondeo,descuento,iva,total) f " & "inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' group by suc.codalmacen,suc.descalmacen) Esp on norm.codalmacen = esp.codalmacen"
            TextoMoneda = "Los importes estan expresados en pesos"
            '    sql = "SELECT CASE WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) = '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  FactE.Sucursal  WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  Vta2.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' " &
            '            "THEN  Vta1.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,'')))  = '' THEN  Vta1.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(FactE.Sucursal,'')))  = '' THEN  Vta1.Sucursal END AS Sucursal, (ISNULL(Vta1.CantArticulos,0) + ISNULL(Vta2.CantArticulos,0) + ISNULL(FactE.CantArticulos,0)) AS CantArticulos, " &
            '            "(ISNULL(Vta1.Importe,0) +  ISNULL(Vta2.Importe,0) + ISNULL(FactE.Importe,0)) AS Importe, (ISNULL(Vta1.Descuento,0) + ISNULL(Vta2.Descuento,0) + ISNULL(FactE.Descuento,0)) AS Descuento, (ISNULL(Vta1.SubTotal,0) + ISNULL(Vta2.SubTotal,0) + ISNULL(FactE.SubTotal,0)) AS SubTotal, (ISNULL(Vta1.Iva,0) + ISNULL(Vta2.Iva,0) + ISNULL(FactE.Iva,0)) AS Iva, (ISNULL(Vta1.Total,0) + ISNULL(Vta2.Total,0) + ISNULL(FactE.Total,0)) AS Total From (SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen " &
            '            "AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal, SUM(Det.Cantidad) AS CantArticulos, SUM(round((Cab.SubTotalAdicional + Cab.RedondeoAdicional) * F.TIPOCAMBIO,1)) AS Importe, SUM(round(Cab.DescuentoAdicional * f.tipocambio,1)) AS Descuento, SUM(round(((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional) * f.tipocambio,1)) AS SubTotal, SUM(round(Cab.IvaAdicional * f.tipocambio,1)) AS Iva, SUM(round((Cab.TotalAdicional + Cab.RedondeoAdicional) * f.tipocambio,1)) As Total FROM CatAlmacen Suc, MovimientosVentasCab Cab INNER JOIN ( SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM " &
            '            "MovimientosVentasDet GROUP BY FolioVenta) Det ON Cab.FolioVenta = Det.FolioVenta inner join facturas f on cab.foliofactura = f.foliofactura WHERE Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Suc.CodAlmacen = Cab.CodSucursal AND Suc.TipoAlmacen = 'P' AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' AND Cab.Estatus <> 'C' GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta1 FULL OUTER JOIN  (SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen " &
            '            "AS Varchar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,SUM(Det.CantidadAdicional) AS CantArticulos, SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe,SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal,SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total From (SELECT FolioAdicional,det.foliofactura,ISNULL(CodSucursalAdicional,0) AS CodSucursal,SUM(CantidadAdicional) AS " &
            '            "CantidadAdicional,SUM(round((PrecioListaSinIvaAdicional * CantidadAdicional) * f.tipocambio,1)) AS SubTotalAdicional,round(RedondeoAdicional * f.tipocambio,1) as redondeoadicional,SUM(round(((ImptePromocionesAdicional + ImpteDescuentosAdicional) * CantidadAdicional) * f.tipocambio,1)) AS DescuentoAdicional,SUM(round((IvaRealAdicional * CantidadAdicional) * f.tipocambio,1)) AS IvaAdicional,SUM(round((PrecioRealAdicional * CantidadAdicional) * f.tipocambio,1)) AS TotalAdicional, FechaVentaAdicional From MovimientosVentasDet det inner join facturas f on det.foliofactura = f.foliofactura WHERE FolioAdicional <> '' AND det.FolioFactura <> '' AND EstatusAdicional <> 'O' GROUP BY " &
            '            "FolioAdicional,det.foliofactura,CodSucursalAdicional,RedondeoAdicional,FechaVentaAdicional,f.tipocambio) Det INNER JOIN CatAlmacen Suc ON Det.CodSucursal = Suc.CodAlmacen WHERE Det.FechaVentaAdicional Between '" & FechaInicial & "' AND '" & FechaFinal & "' AND Suc.TipoAlmacen = 'P' GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta2 ON Vta1.Sucursal = Vta2.Sucursal FULL OUTER JOIN (Select Right(Replicate('0',3) + Cast(CodSucursal AS " &
            '            "VarChar(3)),3) + ' ' + DescAlmacen AS Sucursal, sum(Cantidad) as CantArticulos, sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal, sum(Iva) as Iva, sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.CodSucursal, A.DescAlmacen, sum(F.Cantidad) as Cantidad,case when f.moneda = 'P' then F.SubTotal else round(f.subtotal * f.tipocambio,1) end as subtotal,case when f.moneda = 'P' then F.Descuento else round(f.descuento * f.tipocambio,1) end as descuento,case when f.moneda = 'P' then F.Iva else round(f.iva * f.tipocambio,1) end as iva,case when f.moneda = 'P' then F.Total else round(f.total * f.tipocambio,1) end as total,case when f.moneda = 'P' then F.Redondeo else round(f.redondeo * f.tipocambio,1) end as redondeo,sum(case when f.moneda = 'P' then F.Importe else round(f.importe * f.tipocambio,1) end) as Importe From Facturas F Inner Join CatAlmacen A On F.CodSucursal = " &
            '            "A.CodAlmacen Where F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' Group By F.FolioFactura, F.CodSucursal, A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total, F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group by Right(Replicate('0',3) + Cast(CodSucursal AS VarChar(3)),3) + ' ' + DescAlmacen ) as FactE on Vta2.Sucursal = FactE.Sucursal "

            '    sql = "SELECT   CASE WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  FactE.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  Vta2.Sucursal " &
            '                  "WHEN     ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  Vta1.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,'')))  = '' THEN  Vta1.Sucursal " &
            '                  "WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(FactE.Sucursal,'')))  = '' THEN  Vta1.Sucursal END AS Sucursal, " &
            '            ",(ISNULL(Vta1.CantArticulos,0) + ISNULL(Vta2.CantArticulos,0)) AS CantArticulos,(ISNULL(Vta1.Importe,0) + ISNULL(Vta2.Importe,0)) AS Importe,(ISNULL(Vta1.Descuento,0) + ISNULL(Vta2.Descuento,0)) AS Descuento,(ISNULL(Vta1.SubTotal,0) + ISNULL(Vta2.SubTotal,0)) AS SubTotal,(ISNULL(Vta1.Iva,0) + ISNULL(Vta2.Iva,0)) AS Iva,(ISNULL(Vta1.Total,0) + ISNULL(Vta2.Total,0)) AS Total From " &
            '            "(SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,SUM(Det.Cantidad) AS CantArticulos,SUM(Cab.SubTotalAdicional + Cab.RedondeoAdicional) AS Importe,SUM(Cab.DescuentoAdicional) AS Descuento,SUM((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional) AS SubTotal,SUM(Cab.IvaAdicional) AS Iva,SUM(Cab.TotalAdicional + Cab.RedondeoAdicional) As Total FROM " &
            '            "CatAlmacen Suc,MovimientosVentasCab Cab INNER JOIN (SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM MovimientosVentasDet GROUP BY FolioVenta) Det ON Cab.FolioVenta = Det.FolioVenta WHERE Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Suc.CodAlmacen = Cab.CodSucursal AND Suc.TipoAlmacen = 'P' " &
            '            "AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' AND Cab.Estatus <> 'C' GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta1 FULL OUTER JOIN " &
            '            "(SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS Varchar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,SUM(Det.CantidadAdicional) AS CantArticulos,SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe,SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal,SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total From " &
            '            "(SELECT FolioAdicional,ISNULL(CodSucursalAdicional,0) AS CodSucursal,SUM(CantidadAdicional) AS CantidadAdicional,SUM(PrecioListaSinIvaAdicional * CantidadAdicional) AS SubTotalAdicional,RedondeoAdicional,SUM((ImptePromocionesAdicional + ImpteDescuentosAdicional) * CantidadAdicional) AS DescuentoAdicional,SUM(IvaRealAdicional * CantidadAdicional) AS IvaAdicional,SUM(PrecioRealAdicional * CantidadAdicional) AS TotalAdicional,FechaVentaAdicional From MovimientosVentasDet " &
            '            "WHERE FolioAdicional <> '' AND FolioFactura <> '' AND EstatusAdicional <> 'O' GROUP BY FolioAdicional,CodSucursalAdicional,RedondeoAdicional,FechaVentaAdicional) Det INNER JOIN CatAlmacen Suc ON Det.CodSucursal = Suc.CodAlmacen WHERE Det.FechaVentaAdicional BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Suc.TipoAlmacen = 'P' GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta2 ON Vta1.Sucursal = Vta2.Sucursal"

        ElseIf chkTodaslasSucursales.CheckState = 1 And Moneda = "D" Then

            sql = "select case when rtrim(ltrim(isnull(norm.sucursal,''))) = '' and rtrim(ltrim(isnull(esp.sucursal,''))) <> '' then rtrim(ltrim(esp.sucursal)) when rtrim(ltrim(isnull(norm.sucursal,''))) <> '' and rtrim(ltrim(isnull(esp.sucursal,''))) = '' " & "then rtrim(ltrim(norm.sucursal)) when rtrim(ltrim(isnull(norm.sucursal,''))) <> '' and rtrim(ltrim(isnull(esp.sucursal,''))) <> '' then rtrim(ltrim(norm.sucursal)) end as sucursal,isnull(norm.cantidad,0) + isnull(esp.cantidad,0) as cantidad," & "isnull(norm.importe,0) + isnull(esp.importe,0) as importe,isnull(norm.descuento,0) + isnull(esp.descuento,0) as descuento,isnull(norm.subtotal,0) + isnull(esp.subtotal,0) as subtotal,isnull(norm.iva,0) + isnull(esp.iva,0) as iva,isnull(norm.total,0) + isnull(esp.total,0) as total " & "From " & "(select RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,suc.codalmacen,sum(det.cantidadadicional) as cantidad,sum(case when f.moneda = 'D' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) / f.tipocambio,2) end) as importe," & "sum(case when f.moneda = 'D' then round(f.descuento,2) else round(f.descuento / f.tipocambio,2) end) as descuento,sum(case when f.moneda = 'D' then round((f.subtotal + f.redondeo) - f.descuento,2) else round(((f.subtotal + f.redondeo) - f.descuento) / f.tipocambio,2) end) as subtotal,sum(case when f.moneda = 'D' then round(f.iva,2) else round(f.iva / f.tipocambio,2) end) as iva," & "sum(case when f.moneda = 'D' then  round(f.total + f.redondeo,2) else round((f.total + f.redondeo) / f.tipocambio,2) end) as total from (select foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,sum(cantidad) as cantidad,subtotal,redondeo,descuento,iva,total from facturas where tipofactura = 'N' group by foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,subtotal,redondeo,descuento,iva,total) f " & "inner join (select distinct foliofactura,codsucursal,estatus from movimientosventascab) cab on f.foliofactura = cab.foliofactura Inner Join (select foliofactura,sum(cantidadadicional) as cantidadadicional from movimientosventasdet group by foliofactura) det on f.foliofactura = det.foliofactura and cab.foliofactura = det.foliofactura inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen and cab.codsucursal = suc.codalmacen " & "where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' and cab.estatus <> 'C' group by suc.codalmacen,suc.descalmacen) Norm " & "full Join " & "(select RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,suc.codalmacen,sum(f.cantidad) as cantidad,sum(case when f.moneda = 'D' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) / f.tipocambio,2) end) as importe,sum(case when f.moneda = 'D' then round(f.descuento,2) else round(f.descuento / f.tipocambio,2) end) as descuento,sum(case when f.moneda = 'D' then round((f.subtotal + f.redondeo) - f.descuento,2) else round(((f.subtotal + f.redondeo) - f.descuento) / f.tipocambio,2) end) as subtotal," & "sum(case when f.moneda = 'D' then round(f.iva,2) else round(f.iva / f.tipocambio,2) end) as iva,sum(case when f.moneda = 'D' then round(f.total + f.redondeo,2) else round((f.total + f.redondeo) / f.tipocambio,2) end) as total from (select foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,sum(cantidad) as cantidad,subtotal,redondeo,descuento,iva,total from facturas where tipofactura = 'E' group by foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,subtotal,redondeo,descuento,iva,total) f " & "inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' group by suc.codalmacen,suc.descalmacen) Esp on norm.codalmacen = esp.codalmacen"

            'sql = "SELECT CASE WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) = '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  FactE.Sucursal  WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  Vta2.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' " &
            '        "THEN  Vta1.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,'')))  = '' THEN  Vta1.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(FactE.Sucursal,'')))  = '' THEN  Vta1.Sucursal END AS Sucursal, (ISNULL(Vta1.CantArticulos,0) + ISNULL(Vta2.CantArticulos,0) + ISNULL(FactE.CantArticulos,0)) AS CantArticulos, " &
            '        "(ISNULL(Vta1.Importe,0) +  ISNULL(Vta2.Importe,0) + ISNULL(FactE.Importe,0)) AS Importe, (ISNULL(Vta1.Descuento,0) + ISNULL(Vta2.Descuento,0) + ISNULL(FactE.Descuento,0)) AS Descuento, (ISNULL(Vta1.SubTotal,0) + ISNULL(Vta2.SubTotal,0) + ISNULL(FactE.SubTotal,0)) AS SubTotal, (ISNULL(Vta1.Iva,0) + ISNULL(Vta2.Iva,0) + ISNULL(FactE.Iva,0)) AS Iva, (ISNULL(Vta1.Total,0) + ISNULL(Vta2.Total,0) + ISNULL(FactE.Total,0)) AS Total From (SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen " &
            '        "AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal, SUM(Det.Cantidad) AS CantArticulos, SUM((Cab.SubTotalAdicional + Cab.RedondeoAdicional)) AS Importe, SUM(Cab.DescuentoAdicional) AS Descuento, SUM(((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional)) AS SubTotal, SUM(Cab.IvaAdicional) AS Iva, SUM((Cab.TotalAdicional + Cab.RedondeoAdicional)) As Total FROM CatAlmacen Suc, MovimientosVentasCab Cab INNER JOIN ( SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM " &
            '        "MovimientosVentasDet GROUP BY FolioVenta) Det ON Cab.FolioVenta = Det.FolioVenta inner join facturas f on cab.foliofactura = f.foliofactura WHERE Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Suc.CodAlmacen = Cab.CodSucursal AND Suc.TipoAlmacen = 'P' AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' AND Cab.Estatus <> 'C' GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta1 FULL OUTER JOIN  (SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen " &
            '        "AS Varchar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,SUM(Det.CantidadAdicional) AS CantArticulos, SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe,SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal,SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total From (SELECT FolioAdicional,det.foliofactura,ISNULL(CodSucursalAdicional,0) AS CodSucursal,SUM(CantidadAdicional) AS " &
            '        "CantidadAdicional,SUM((PrecioListaSinIvaAdicional * CantidadAdicional)) AS SubTotalAdicional,(RedondeoAdicional) as redondeoadicional,SUM(((ImptePromocionesAdicional + ImpteDescuentosAdicional) * CantidadAdicional)) AS DescuentoAdicional,SUM((IvaRealAdicional * CantidadAdicional)) AS IvaAdicional,SUM((PrecioRealAdicional * CantidadAdicional)) AS TotalAdicional, FechaVentaAdicional From MovimientosVentasDet det inner join facturas f on det.foliofactura = f.foliofactura WHERE FolioAdicional <> '' AND det.FolioFactura <> '' AND EstatusAdicional <> 'O' GROUP BY " &
            '        "FolioAdicional,det.foliofactura,CodSucursalAdicional,RedondeoAdicional,FechaVentaAdicional,f.tipocambio) Det INNER JOIN CatAlmacen Suc ON Det.CodSucursal = Suc.CodAlmacen WHERE Det.FechaVentaAdicional Between '" & FechaInicial & "' AND '" & FechaFinal & "' AND Suc.TipoAlmacen = 'P' GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta2 ON Vta1.Sucursal = Vta2.Sucursal FULL OUTER JOIN (Select Right(Replicate('0',3) + Cast(CodSucursal AS " &
            '        "VarChar(3)),3) + ' ' + DescAlmacen AS Sucursal, sum(Cantidad) as CantArticulos, sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal, sum(Iva) as Iva, sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.CodSucursal, A.DescAlmacen, sum(F.Cantidad) as Cantidad,case when f.moneda = 'D' then F.SubTotal else round(f.subtotal / f.tipocambio,2) end as subtotal,case when f.moneda = 'D' then F.Descuento else round(f.descuento / f.tipocambio,2) end as descuento,case when f.moneda = 'D' then F.Iva else round(f.iva / f.tipocambio,2) end as iva,case when f.moneda = 'D' then F.Total else round(f.total / f.tipocambio,2) end as total,case when f.moneda = 'D' then F.Redondeo else round(f.redondeo / f.tipocambio,2) end as redondeo,sum(case when f.moneda = 'D' then F.Importe else round(f.importe / f.tipocambio,2) end) as Importe From Facturas F Inner Join CatAlmacen A On F.CodSucursal = " &
            '        "A.CodAlmacen Where F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' Group By F.FolioFactura, F.CodSucursal, A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total, F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group by Right(Replicate('0',3) + Cast(CodSucursal AS VarChar(3)),3) + ' ' + DescAlmacen ) as FactE on Vta2.Sucursal = FactE.Sucursal "
            TextoMoneda = "Los importes estan expresados en dólares"
        ElseIf chkTodaslasSucursales.CheckState = 0 And Moneda = "P" Then
            If CShort(Numerico(txtCodSucursal.Text)) = 0 Then
                MsgBox("Proporcione el Codigo de la Sucursal ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtCodSucursal.Focus()
                Exit Sub
            End If

            sql = "select case when rtrim(ltrim(isnull(norm.sucursal,''))) = '' and rtrim(ltrim(isnull(esp.sucursal,''))) <> '' then rtrim(ltrim(esp.sucursal)) when rtrim(ltrim(isnull(norm.sucursal,''))) <> '' and rtrim(ltrim(isnull(esp.sucursal,''))) = '' " & "then rtrim(ltrim(norm.sucursal)) when rtrim(ltrim(isnull(norm.sucursal,''))) <> '' and rtrim(ltrim(isnull(esp.sucursal,''))) <> '' then rtrim(ltrim(norm.sucursal)) end as sucursal,isnull(norm.cantidad,0) + isnull(esp.cantidad,0) as cantidad," & "isnull(norm.importe,0) + isnull(esp.importe,0) as importe,isnull(norm.descuento,0) + isnull(esp.descuento,0) as descuento,isnull(norm.subtotal,0) + isnull(esp.subtotal,0) as subtotal,isnull(norm.iva,0) + isnull(esp.iva,0) as iva,isnull(norm.total,0) + isnull(esp.total,0) as total " & "From " & "(select RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,suc.codalmacen,sum(det.cantidadadicional) as cantidad,sum(case when f.moneda = 'P' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) * f.tipocambio,1) end) as importe," & "sum(case when f.moneda = 'P' then round(f.descuento,2) else round(f.descuento * f.tipocambio,1) end) as descuento,sum(case when f.moneda = 'P' then round((f.subtotal + f.redondeo) - f.descuento,2) else round(((f.subtotal + f.redondeo) - f.descuento) * f.tipocambio,1) end) as subtotal,sum(case when f.moneda = 'P' then round(f.iva,2) else round(f.iva * f.tipocambio,1) end) as iva," & "sum(case when f.moneda = 'P' then  round(f.total + f.redondeo,2) else round((f.total + f.redondeo) * f.tipocambio,1) end) as total from (select foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,sum(cantidad) as cantidad,subtotal,redondeo,descuento,iva,total from facturas where tipofactura = 'N' group by foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,subtotal,redondeo,descuento,iva,total) f " & "inner join (select distinct foliofactura,codsucursal,estatus from movimientosventascab) cab on f.foliofactura = cab.foliofactura Inner Join (select foliofactura,sum(cantidadadicional) as cantidadadicional from movimientosventasdet group by foliofactura) det on f.foliofactura = det.foliofactura and cab.foliofactura = det.foliofactura inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen and cab.codsucursal = suc.codalmacen " & "where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' and cab.estatus <> 'C' AND suc.codalmacen = " & Numerico(txtCodSucursal.Text) & " group by suc.codalmacen,suc.descalmacen) Norm " & "full Join " & "(select RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,suc.codalmacen,sum(f.cantidad) as cantidad,sum(case when f.moneda = 'P' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) * f.tipocambio,1) end) as importe,sum(case when f.moneda = 'P' then round(f.descuento,2) else round(f.descuento * f.tipocambio,1) end) as descuento,sum(case when f.moneda = 'P' then round((f.subtotal + f.redondeo) - f.descuento,2) else round(((f.subtotal + f.redondeo) - f.descuento) * f.tipocambio,1) end) as subtotal," & "sum(case when f.moneda = 'P' then round(f.iva,2) else round(f.iva * f.tipocambio,1) end) as iva,sum(case when f.moneda = 'P' then  round(f.total + f.redondeo,2) else round((f.total + f.redondeo) * f.tipocambio,1) end) as total from (select foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,sum(cantidad) as cantidad,subtotal,redondeo,descuento,iva,total from facturas where tipofactura = 'E' group by foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,subtotal,redondeo,descuento,iva,total) f " & "inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' and suc.codalmacen = " & Numerico(txtCodSucursal.Text) & " group by suc.codalmacen,suc.descalmacen) Esp on norm.codalmacen = esp.codalmacen"

            TextoMoneda = "Los importes estan expresados en pesos"
            'sql = "SELECT CASE WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  FactE.Sucursal  WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  Vta2.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' " &
            '        "THEN  Vta1.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,'')))  = '' THEN  Vta1.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(FactE.Sucursal,'')))  = '' THEN  Vta1.Sucursal END AS Sucursal, (ISNULL(Vta1.CantArticulos,0) + ISNULL(Vta2.CantArticulos,0) + ISNULL(FactE.CantArticulos,0)) AS CantArticulos, " &
            '        "(ISNULL(Vta1.Importe,0) +  ISNULL(Vta2.Importe,0) + ISNULL(FactE.Importe,0)) AS Importe, (ISNULL(Vta1.Descuento,0) + ISNULL(Vta2.Descuento,0) + ISNULL(FactE.Descuento,0)) AS Descuento, (ISNULL(Vta1.SubTotal,0) + ISNULL(Vta2.SubTotal,0) + ISNULL(FactE.SubTotal,0)) AS SubTotal, (ISNULL(Vta1.Iva,0) + ISNULL(Vta2.Iva,0) + ISNULL(FactE.Iva,0)) AS Iva, (ISNULL(Vta1.Total,0) + ISNULL(Vta2.Total,0) + ISNULL(FactE.Total,0)) AS Total From ( SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen " &
            '        "AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal, SUM(Det.Cantidad) AS CantArticulos, SUM(round((Cab.SubTotalAdicional + Cab.RedondeoAdicional) * F.TIPOCAMBIO,1)) AS Importe, SUM(round(Cab.DescuentoAdicional * f.tipocambio,1)) AS Descuento, SUM(round(((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional) * f.tipocambio,1)) AS SubTotal, SUM(round(Cab.IvaAdicional * f.tipocambio,1)) AS Iva, SUM(round((Cab.TotalAdicional + Cab.RedondeoAdicional) * f.tipocambio,1)) As Total FROM CatAlmacen Suc, MovimientosVentasCab Cab INNER JOIN (SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM " &
            '        "MovimientosVentasDet GROUP BY FolioVenta) Det ON Cab.FolioVenta = Det.FolioVenta inner join facturas f on cab.foliofactura = f.foliofactura WHERE Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Suc.CodAlmacen = Cab.CodSucursal AND Suc.TipoAlmacen = 'P' AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' AND Cab.Estatus <> 'C'  AND Cab.CodSucursal = " & CInt(Numerico(txtCodSucursal)) & " GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta1 FULL OUTER JOIN  (SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen " &
            '        "AS Varchar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,SUM(Det.CantidadAdicional) AS CantArticulos, SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe,SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal,SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total From (SELECT FolioAdicional,det.foliofactura,ISNULL(CodSucursalAdicional,0) AS CodSucursal,SUM(CantidadAdicional) AS " &
            '        "CantidadAdicional,SUM(round((PrecioListaSinIvaAdicional * CantidadAdicional) * f.tipocambio,1)) AS SubTotalAdicional,round(RedondeoAdicional * f.tipocambio,1) as redondeoadicional, SUM(round(((ImptePromocionesAdicional + ImpteDescuentosAdicional) * CantidadAdicional) * f.tipocambio,1)) AS DescuentoAdicional,SUM(round((IvaRealAdicional * CantidadAdicional) * f.tipocambio,1)) AS IvaAdicional,SUM(round((PrecioRealAdicional * CantidadAdicional) * f.tipocambio,1)) AS TotalAdicional, FechaVentaAdicional From MovimientosVentasDet det inner join facturas f on det.foliofactura = f.foliofactura WHERE FolioAdicional <> '' AND det.FolioFactura <> '' AND EstatusAdicional <> 'O' GROUP BY " &
            '        "FolioAdicional,det.foliofactura,CodSucursalAdicional,RedondeoAdicional,FechaVentaAdicional,f.tipocambio) Det INNER JOIN CatAlmacen Suc ON Det.CodSucursal = Suc.CodAlmacen WHERE Det.FechaVentaAdicional Between '" & FechaInicial & "' AND '" & FechaFinal & "' AND Det.CodSucursal = " & CInt(Numerico(txtCodSucursal)) & " AND Suc.TipoAlmacen = 'P' GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta2 ON Vta1.Sucursal = Vta2.Sucursal FULL OUTER JOIN (Select Right(Replicate('0',3) + Cast(CodSucursal AS " &
            '        "VarChar(3)),3) + ' ' + DescAlmacen AS Sucursal, sum(Cantidad) as CantArticulos, sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal, sum(Iva) as Iva, sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.CodSucursal, A.DescAlmacen, sum(F.Cantidad) as Cantidad, case when f.moneda = 'P' then F.SubTotal else round(f.subtotal * f.tipocambio,1) end as subtotal, case when f.moneda = 'P' then F.Descuento else round(f.descuento * f.tipocambio,1) end as descuento,case when f.moneda = 'P' then F.Iva else round(f.iva * f.tipocambio,1) end as iva,case when f.moneda = 'P' then F.Total else round(f.total * f.tipocambio,1) end as total,case when f.moneda = 'P' then F.Redondeo else round(f.redondeo * f.tipocambio,1) end as redondeo, sum(case when f.moneda = 'P' then F.Importe else round(f.importe * f.tipocambio,1) end) as Importe From Facturas F Inner Join CatAlmacen A On F.CodSucursal = " &
            '        "A.CodAlmacen Where F.CodSucursal = " & CInt(Numerico(txtCodSucursal)) & " And F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' Group By F.FolioFactura, F.CodSucursal, A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total, F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group by Right(Replicate('0',3) + Cast(CodSucursal AS VarChar(3)),3) + ' ' + DescAlmacen ) as FactE on Vta2.Sucursal = FactE.Sucursal "

            'sql = "SELECT CASE WHEN ISNULL(Vta1.Sucursal,'') = '' THEN Vta2.Sucursal ELSE Vta1.Sucursal END AS Sucursal,(ISNULL(Vta1.CantArticulos,0) + ISNULL(Vta2.CantArticulos,0)) AS CantArticulos,(ISNULL(Vta1.Importe,0) + ISNULL(Vta2.Importe,0)) AS Importe,(ISNULL(Vta1.Descuento,0) + ISNULL(Vta2.Descuento,0)) AS Descuento,(ISNULL(Vta1.SubTotal,0) + ISNULL(Vta2.SubTotal,0)) AS SubTotal,(ISNULL(Vta1.Iva,0) + ISNULL(Vta2.Iva,0)) AS Iva,(ISNULL(Vta1.Total,0) + ISNULL(Vta2.Total,0)) AS Total From " &
            '        "(SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,SUM(Det.Cantidad) AS CantArticulos,SUM(Cab.SubTotalAdicional + Cab.RedondeoAdicional) AS Importe,SUM(Cab.DescuentoAdicional) AS Descuento,SUM((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional) AS SubTotal,SUM(Cab.IvaAdicional) AS Iva,SUM(Cab.TotalAdicional + Cab.RedondeoAdicional) As Total FROM " &
            '        "CatAlmacen Suc,MovimientosVentasCab Cab INNER JOIN (SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM MovimientosVentasDet GROUP BY FolioVenta) Det ON Cab.FolioVenta = Det.FolioVenta WHERE Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Suc.CodAlmacen = Cab.CodSucursal AND Suc.TipoAlmacen = 'P' " &
            '        "AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' AND Cab.Estatus <> 'C' AND Cab.CodSucursal = " & Numerico(txtCodSucursal) & " GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta1 FULL OUTER JOIN " &
            '        "(SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS Varchar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,SUM(Det.CantidadAdicional) AS CantArticulos,SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe,SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal,SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total From " &
            '        "(SELECT FolioAdicional,ISNULL(CodSucursalAdicional,0) AS CodSucursal,SUM(CantidadAdicional) AS CantidadAdicional,SUM(PrecioListaSinIvaAdicional * CantidadAdicional) AS SubTotalAdicional,RedondeoAdicional,SUM((ImptePromocionesAdicional + ImpteDescuentosAdicional) * CantidadAdicional) AS DescuentoAdicional,SUM(IvaRealAdicional * CantidadAdicional) AS IvaAdicional,SUM(PrecioRealAdicional * CantidadAdicional) AS TotalAdicional,FechaVentaAdicional From MovimientosVentasDet " &
            '        "WHERE FolioAdicional <> '' AND FolioFactura <> '' AND EstatusAdicional <> 'O' GROUP BY FolioAdicional,CodSucursalAdicional,RedondeoAdicional,FechaVentaAdicional) Det INNER JOIN CatAlmacen Suc ON Det.CodSucursal = Suc.CodAlmacen WHERE Det.FechaVentaAdicional BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Det.CodSucursal = " & Numerico(txtCodSucursal) & " AND Suc.TipoAlmacen = 'P' GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + " &
            '        "Suc.DescAlmacen) Vta2 ON Vta1.Sucursal = Vta2.Sucursal"
        ElseIf chkTodaslasSucursales.CheckState = 0 And Moneda = "D" Then
            If CShort(Numerico(txtCodSucursal.Text)) = 0 Then
                MsgBox("Proporcione el Codigo de la Sucursal ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtCodSucursal.Focus()
                Exit Sub
            End If

            sql = "select case when rtrim(ltrim(isnull(norm.sucursal,''))) = '' and rtrim(ltrim(isnull(esp.sucursal,''))) <> '' then rtrim(ltrim(esp.sucursal)) when rtrim(ltrim(isnull(norm.sucursal,''))) <> '' and rtrim(ltrim(isnull(esp.sucursal,''))) = '' " & "then rtrim(ltrim(norm.sucursal)) when rtrim(ltrim(isnull(norm.sucursal,''))) <> '' and rtrim(ltrim(isnull(esp.sucursal,''))) <> '' then rtrim(ltrim(norm.sucursal)) end as sucursal,isnull(norm.cantidad,0) + isnull(esp.cantidad,0) as cantidad," & "isnull(norm.importe,0) + isnull(esp.importe,0) as importe,isnull(norm.descuento,0) + isnull(esp.descuento,0) as descuento,isnull(norm.subtotal,0) + isnull(esp.subtotal,0) as subtotal,isnull(norm.iva,0) + isnull(esp.iva,0) as iva,isnull(norm.total,0) + isnull(esp.total,0) as total " & "From " & "(select RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,suc.codalmacen,sum(det.cantidadadicional) as cantidad,sum(case when f.moneda = 'D' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) / f.tipocambio,2) end) as importe," & "sum(case when f.moneda = 'D' then round(f.descuento,2) else round(f.descuento / f.tipocambio,2) end) as descuento,sum(case when f.moneda = 'D' then round((f.subtotal + f.redondeo) - f.descuento,2) else round(((f.subtotal + f.redondeo) - f.descuento) / f.tipocambio,2) end) as subtotal,sum(case when f.moneda = 'D' then round(f.iva,2) else round(f.iva / f.tipocambio,2) end) as iva," & "sum(case when f.moneda = 'D' then  round(f.total + f.redondeo,2) else round((f.total + f.redondeo) / f.tipocambio,2) end) as total from (select foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,sum(cantidad) as cantidad,subtotal,redondeo,descuento,iva,total from facturas where tipofactura = 'N' group by foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,subtotal,redondeo,descuento,iva,total) f " & "inner join (select distinct foliofactura,codsucursal,estatus from movimientosventascab) cab on f.foliofactura = cab.foliofactura Inner Join (select foliofactura,sum(cantidadadicional) as cantidadadicional from movimientosventasdet group by foliofactura) det on f.foliofactura = det.foliofactura and cab.foliofactura = det.foliofactura inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen and cab.codsucursal = suc.codalmacen " & "where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' and cab.estatus <> 'C' AND suc.codalmacen = " & Numerico(txtCodSucursal.Text) & " group by suc.codalmacen,suc.descalmacen) Norm " & "full Join " & "(select RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,suc.codalmacen,sum(f.cantidad) as cantidad,sum(case when f.moneda = 'D' then round(f.subtotal + f.redondeo,2) else round((f.subtotal + f.redondeo) / f.tipocambio,2) end) as importe,sum(case when f.moneda = 'D' then round(f.descuento,2) else round(f.descuento / f.tipocambio,2) end) as descuento,sum(case when f.moneda = 'D' then round((f.subtotal + f.redondeo) - f.descuento,2) else round(((f.subtotal + f.redondeo) - f.descuento) / f.tipocambio,2) end) as subtotal," & "sum(case when f.moneda = 'D' then round(f.iva,2) else round(f.iva / f.tipocambio,2) end) as iva,sum(case when f.moneda = 'D' then  round(f.total + f.redondeo,2) else round((f.total + f.redondeo) / f.tipocambio,2) end) as total from (select foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,sum(cantidad) as cantidad,subtotal,redondeo,descuento,iva,total from facturas where tipofactura = 'E' group by foliofactura,codsucursal,fechafactura,tipofactura,estatus,moneda,tipocambio,subtotal,redondeo,descuento,iva,total) f " & "inner join (select * from catalmacen where tipoalmacen = 'P') suc on f.codsucursal = suc.codalmacen where f.fechafactura between '" & FechaInicial & "' AND '" & FechaFinal & "' and f.estatus <> 'C' and suc.codalmacen = " & Numerico(txtCodSucursal.Text) & " group by suc.codalmacen,suc.descalmacen) Esp on norm.codalmacen = esp.codalmacen"

            TextoMoneda = "Los importes estan expresados en dólares"
            'sql = "SELECT CASE WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  FactE.Sucursal  WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' THEN  Vta2.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,''))) <> '' " &
            '        "THEN  Vta1.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(FactE.Sucursal,'')))  = '' THEN  Vta1.Sucursal WHEN ltrim(rtrim(ISNULL(Vta1.Sucursal,''))) <> '' And ltrim(rtrim(ISNULL(Vta2.Sucursal,'')))  = '' And ltrim(rtrim(ISNULL(FactE.Sucursal,'')))  = '' THEN  Vta1.Sucursal END AS Sucursal, (ISNULL(Vta1.CantArticulos,0) + ISNULL(Vta2.CantArticulos,0) + ISNULL(FactE.CantArticulos,0)) AS CantArticulos, " &
            '        "(ISNULL(Vta1.Importe,0) +  ISNULL(Vta2.Importe,0) + ISNULL(FactE.Importe,0)) AS Importe, (ISNULL(Vta1.Descuento,0) + ISNULL(Vta2.Descuento,0) + ISNULL(FactE.Descuento,0)) AS Descuento, (ISNULL(Vta1.SubTotal,0) + ISNULL(Vta2.SubTotal,0) + ISNULL(FactE.SubTotal,0)) AS SubTotal, (ISNULL(Vta1.Iva,0) + ISNULL(Vta2.Iva,0) + ISNULL(FactE.Iva,0)) AS Iva, (ISNULL(Vta1.Total,0) + ISNULL(Vta2.Total,0) + ISNULL(FactE.Total,0)) AS Total From ( SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen " &
            '        "AS VarChar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal, SUM(Det.Cantidad) AS CantArticulos, SUM((Cab.SubTotalAdicional + Cab.RedondeoAdicional)) AS Importe, SUM(Cab.DescuentoAdicional) AS Descuento, SUM(((Cab.SubTotalAdicional + Cab.RedondeoAdicional) - Cab.DescuentoAdicional)) AS SubTotal, SUM(Cab.IvaAdicional) AS Iva, SUM((Cab.TotalAdicional + Cab.RedondeoAdicional)) As Total FROM CatAlmacen Suc, MovimientosVentasCab Cab INNER JOIN (SELECT FolioVenta, SUM(Cantidad) AS Cantidad FROM " &
            '        "MovimientosVentasDet GROUP BY FolioVenta) Det ON Cab.FolioVenta = Det.FolioVenta inner join facturas f on cab.foliofactura = f.foliofactura WHERE Cab.FechaVenta BETWEEN '" & FechaInicial & "' AND '" & FechaFinal & "' AND Suc.CodAlmacen = Cab.CodSucursal AND Suc.TipoAlmacen = 'P' AND Cab.FolioFactura <> '' AND Cab.EstatusAdicional <> 'O' AND Cab.Estatus <> 'C'  AND Cab.CodSucursal = " & CInt(Numerico(txtCodSucursal)) & " GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta1 FULL OUTER JOIN  (SELECT RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen " &
            '        "AS Varchar(3)),3) + ' ' + Suc.DescAlmacen AS Sucursal,SUM(Det.CantidadAdicional) AS CantArticulos, SUM(Det.SubTotalAdicional + Det.RedondeoAdicional) AS Importe,SUM(Det.DescuentoAdicional) AS Descuento,SUM((Det.SubTotalAdicional + Det.RedondeoAdicional) - Det.DescuentoAdicional) AS SubTotal,SUM(Det.IvaAdicional) AS Iva,SUM(Det.TotalAdicional + Det.RedondeoAdicional) As Total From (SELECT FolioAdicional,det.foliofactura,ISNULL(CodSucursalAdicional,0) AS CodSucursal,SUM(CantidadAdicional) AS " &
            '        "CantidadAdicional,SUM((PrecioListaSinIvaAdicional * CantidadAdicional)) AS SubTotalAdicional,(RedondeoAdicional) as redondeoadicional, SUM(((ImptePromocionesAdicional + ImpteDescuentosAdicional) * CantidadAdicional)) AS DescuentoAdicional,SUM((IvaRealAdicional * CantidadAdicional)) AS IvaAdicional,SUM((PrecioRealAdicional * CantidadAdicional)) AS TotalAdicional, FechaVentaAdicional From MovimientosVentasDet det inner join facturas f on det.foliofactura = f.foliofactura WHERE FolioAdicional <> '' AND det.FolioFactura <> '' AND EstatusAdicional <> 'O' GROUP BY " &
            '        "FolioAdicional,det.foliofactura,CodSucursalAdicional,RedondeoAdicional,FechaVentaAdicional,f.tipocambio) Det INNER JOIN CatAlmacen Suc ON Det.CodSucursal = Suc.CodAlmacen WHERE Det.FechaVentaAdicional Between '" & FechaInicial & "' AND '" & FechaFinal & "' AND Det.CodSucursal = " & CInt(Numerico(txtCodSucursal)) & " AND Suc.TipoAlmacen = 'P' GROUP BY  RIGHT(REPLICATE('0',3) + CAST(Suc.CodAlmacen AS VarChar(3)),3) + ' ' + Suc.DescAlmacen) Vta2 ON Vta1.Sucursal = Vta2.Sucursal FULL OUTER JOIN (Select Right(Replicate('0',3) + Cast(CodSucursal AS " &
            '        "VarChar(3)),3) + ' ' + DescAlmacen AS Sucursal, sum(Cantidad) as CantArticulos, sum(SubTotal+Redondeo) as Importe, sum(Descuento) as Descuento, sum((SubTotal+Redondeo)-Descuento) as SubTotal, sum(Iva) as Iva, sum(Total+Redondeo) as Total From (Select F.FolioFactura, F.CodSucursal, A.DescAlmacen, sum(F.Cantidad) as Cantidad, case when f.moneda = 'D' then F.SubTotal else round(f.subtotal / f.tipocambio,2) end as subtotal, case when f.moneda = 'D' then F.Descuento else round(f.descuento / f.tipocambio,2) end as descuento,case when f.moneda = 'D' then F.Iva else round(f.iva / f.tipocambio,2) end as iva,case when f.moneda = 'D' then F.Total else round(f.total / f.tipocambio,2) end as total,case when f.moneda = 'D' then F.Redondeo else round(f.redondeo / f.tipocambio,2) end as redondeo, sum(case when f.moneda = 'D' then F.Importe else round(f.importe / f.tipocambio,2) end) as Importe From Facturas F Inner Join CatAlmacen A On F.CodSucursal = " &
            '        "A.CodAlmacen Where F.CodSucursal = " & CInt(Numerico(txtCodSucursal)) & " And F.FechaFactura Between '" & FechaInicial & "' And '" & FechaFinal & "' And F.TipoFactura = 'E' Group By F.FolioFactura, F.CodSucursal, A.DescAlmacen, F.SubTotal, F.Descuento, F.Iva, F.Total, F.Redondeo,f.moneda,f.tipocambio) as FactEsp Group by Right(Replicate('0',3) + Cast(CodSucursal AS VarChar(3)),3) + ' ' + DescAlmacen ) as FactE on Vta2.Sucursal = FactE.Sucursal "
        End If
        BorraCmd()
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = sql
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No Existe Información En Este Rango de Fechas...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If
        'frmReportes.Report = RptFactFacturacionGlobalXSucursal
        RptFactFacturacionGlobalXSucursal.SetDataSource(frmReportes.rsReport)

        If chkV.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            With frmReportes.rsReport
                '.Text7.Suppress = True
                '.Field7.Suppress = True
                '.Field15.Suppress = True
            End With
        Else
            With frmReportes.rsReport
                '.Text7.Suppress = False
                '.Field7.Suppress = False
                '.Field15.Suppress = False
            End With
        End If

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'frmReportes.rsReport = rsReporte
        'frmReportes.aFormula_ = New Object() {"NombreEmpresa", "NombreReporte", "PeriodoReporte", "TextoAdicional", "Moneda"}
        'frmReportes.aValues_ = New Object() {NombreEmpresa, NombreReporte, PeriodoReporte, TextoAdicional, TextoMoneda}

        If (NombreEmpresa <> Nothing Or NombreEmpresa <> "") Then
            pdvNum.Value = NombreEmpresa : pvNum.Add(pdvNum)
            RptFactFacturacionGlobalXSucursal.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            RptFactFacturacionGlobalXSucursal.DataDefinition.ParameterFields("NombreEmpresa").ApplyCurrentValues(pvNum)
        End If

        If (NombreReporte <> Nothing Or NombreReporte <> "") Then
            pdvNum.Value = NombreReporte : pvNum.Add(pdvNum)
            RptFactFacturacionGlobalXSucursal.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            RptFactFacturacionGlobalXSucursal.DataDefinition.ParameterFields("NombreReporte").ApplyCurrentValues(pvNum)
        End If

        If (PeriodoReporte <> Nothing Or PeriodoReporte <> "") Then
            pdvNum.Value = PeriodoReporte : pvNum.Add(pdvNum)
            RptFactFacturacionGlobalXSucursal.DataDefinition.ParameterFields("PeriodoReporte").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            RptFactFacturacionGlobalXSucursal.DataDefinition.ParameterFields("PeriodoReporte").ApplyCurrentValues(pvNum)
        End If

        'If (txtTextoAdicional.Text <> Nothing Or TextoAdicional <> "") Then
        '    pdvNum.Value = txtTextoAdicional.Text : pvNum.Add(pdvNum)
        '    RptFactFacturacionGlobalXSucursal.DataDefinition.ParameterFields("TextoAdicional").ApplyCurrentValues(pvNum)
        'Else
        '    pdvNum.Value = "" : pvNum.Add(pdvNum)
        '    RptFactFacturacionGlobalXSucursal.DataDefinition.ParameterFields("TextoAdicional").ApplyCurrentValues(pvNum)
        'End If

        'If (TextoMoneda <> Nothing Or TextoMoneda <> "") Then
        '    pdvNum.Value = TextoMoneda : pvNum.Add(pdvNum)
        '    RptFactFacturacionGlobalXSucursal.DataDefinition.ParameterFields("Moneda").ApplyCurrentValues(pvNum)
        'Else
        '    pdvNum.Value = "" : pvNum.Add(pdvNum)
        '    RptFactFacturacionGlobalXSucursal.DataDefinition.ParameterFields("Moneda").ApplyCurrentValues(pvNum)
        'End If

        frmReportes.Text = "Facturación Global por Sucursal"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        frmReportes.reporteActual = RptFactFacturacionGlobalXSucursal
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

    Private Sub chkTodaslasSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.Enter
        Pon_Tool()
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

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%' AND TipoAlmacen = 'P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla, dbcSucursal)
        intCodSucursal = 0
        FueraChange = True
        txtCodSucursal.Text = Format(String.Concat(intCodSucursal, "000"))
        FueraChange = False
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursal.Name Then
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
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcSucursal_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        'If dbcSucursal.SelectedItem <> 0 Then
        '    dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
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

        txtCodSucursal.Text = (intCodSucursal)

        For i = 0 To 2 - txtCodSucursal.TextLength
            txtCodSucursal.Text = String.Concat("0" + txtCodSucursal.Text)
        Next i

        FueraChange = False
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = dbcSucursal.Text
        'If dbcSucursal.SelectedItem <> 0 Then
        'dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
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

    Private Sub frmFactReportesFacturacionGlobalXSucursal_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmFactReportesFacturacionGlobalXSucursal_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmFactReportesFacturacionGlobalXSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmFactReportesFacturacionGlobalXSucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmFactReportesFacturacionGlobalXSucursal_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmFactReportesFacturacionGlobalXSucursal_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmFactReportesFacturacionGlobalXSucursal_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
        'MDIMenuPrincipalCorpo.mnuFacturacionRptFactOpc(0).Enabled = True
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
            txtCodSucursal.Text = "000"
            Exit Sub
        End If
        FueraChange = True

        For i = 0 To 3 - txtCodSucursal.TextLength
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

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtTextoAdicional = New System.Windows.Forms.TextBox()
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.optPesos = New System.Windows.Forms.RadioButton()
        Me.optDolares = New System.Windows.Forms.RadioButton()
        Me.chkV = New System.Windows.Forms.CheckBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicial = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFinal = New System.Windows.Forms.DateTimePicker()
        Me._Label2_1 = New System.Windows.Forms.Label()
        Me._Label2_0 = New System.Windows.Forms.Label()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.chkTodaslasSucursales = New System.Windows.Forms.CheckBox()
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
        Me.txtTextoAdicional.Location = New System.Drawing.Point(9, 230)
        Me.txtTextoAdicional.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTextoAdicional.MaxLength = 120
        Me.txtTextoAdicional.Multiline = True
        Me.txtTextoAdicional.Name = "txtTextoAdicional"
        Me.txtTextoAdicional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTextoAdicional.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtTextoAdicional.Size = New System.Drawing.Size(359, 71)
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
        Me.txtCodSucursal.Location = New System.Drawing.Point(68, 37)
        Me.txtCodSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucursal.MaxLength = 3
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(67, 20)
        Me.txtCodSucursal.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCodSucursal, "Codigo de Sucursal.")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optPesos)
        Me.Frame2.Controls.Add(Me.optDolares)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(14, 137)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(354, 40)
        Me.Frame2.TabIndex = 14
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Moneda"
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
        'chkV
        '
        Me.chkV.BackColor = System.Drawing.SystemColors.Control
        Me.chkV.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkV.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkV.Location = New System.Drawing.Point(350, 195)
        Me.chkV.Margin = New System.Windows.Forms.Padding(2)
        Me.chkV.Name = "chkV"
        Me.chkV.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkV.Size = New System.Drawing.Size(18, 19)
        Me.chkV.TabIndex = 13
        Me.chkV.Text = "Check1"
        Me.chkV.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dtpFechaInicial)
        Me.Frame1.Controls.Add(Me.dtpFechaFinal)
        Me.Frame1.Controls.Add(Me._Label2_1)
        Me.Frame1.Controls.Add(Me._Label2_0)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(14, 72)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(354, 46)
        Me.Frame1.TabIndex = 8
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Periodo"
        '
        'dtpFechaInicial
        '
        Me.dtpFechaInicial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicial.Location = New System.Drawing.Point(82, 20)
        Me.dtpFechaInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaInicial.Name = "dtpFechaInicial"
        Me.dtpFechaInicial.Size = New System.Drawing.Size(95, 20)
        Me.dtpFechaInicial.TabIndex = 3
        '
        'dtpFechaFinal
        '
        Me.dtpFechaFinal.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFinal.Location = New System.Drawing.Point(241, 20)
        Me.dtpFechaFinal.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaFinal.Name = "dtpFechaFinal"
        Me.dtpFechaFinal.Size = New System.Drawing.Size(95, 20)
        Me.dtpFechaFinal.TabIndex = 4
        '
        '_Label2_1
        '
        Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_1.Location = New System.Drawing.Point(190, 20)
        Me._Label2_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_1.Size = New System.Drawing.Size(58, 17)
        Me._Label2_1.TabIndex = 10
        Me._Label2_1.Text = "Hasta el :"
        '
        '_Label2_0
        '
        Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_0.Location = New System.Drawing.Point(26, 20)
        Me._Label2_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_0.Size = New System.Drawing.Size(60, 17)
        Me._Label2_0.TabIndex = 9
        Me._Label2_0.Text = "Desde el :"
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(144, 36)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(170, 21)
        Me.dbcSucursal.TabIndex = 2
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
        Me.chkTodaslasSucursales.Size = New System.Drawing.Size(130, 17)
        Me.chkTodaslasSucursales.TabIndex = 0
        Me.chkTodaslasSucursales.Text = "Todas las Sucursales"
        Me.chkTodaslasSucursales.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(10, 214)
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
        Me.Label1.Location = New System.Drawing.Point(12, 37)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(64, 17)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Sucursal :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(122, 318)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 115
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(7, 318)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 114
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmFactReportesFacturacionGlobalXSucursal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(380, 366)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.chkV)
        Me.Controls.Add(Me.txtTextoAdicional)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me.txtCodSucursal)
        Me.Controls.Add(Me.chkTodaslasSucursales)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(384, 239)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmFactReportesFacturacionGlobalXSucursal"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Facturación Global por Sucursal"
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

End Class