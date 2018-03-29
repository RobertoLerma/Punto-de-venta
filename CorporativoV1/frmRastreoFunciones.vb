Option Explicit On
Option Strict Off
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmRastreoFunciones
    Inherits System.Windows.Forms.Form

    Public components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents prgRastreo As System.Windows.Forms.ProgressBar
    Public WithEvents btnRastreo As System.Windows.Forms.Button
    Public WithEvents fraRastreo As System.Windows.Forms.GroupBox

    ''' ***************************************************************************************************************************************************
    ''' SE AGREGO REPORTE DE VTAS Y EXIST X FAMILIA - MENU VENTAS - SALIDA DE MERCANCIA - REPORTES EJECUTIVOS
    ''' 27OCT2010 - MAVF Ver
    '''
    ''' Ver 1.0       Estatus: Aprobado
    ''' ***************************************************************************************************************************************************


    'Constante que indica la posición de las columnas en la Matriz
    Const nFORMA As Integer = 0
    Const nDESC As Integer = 1

    Dim mblnSalir As Boolean

    Public Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRastreoFunciones))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me.prgRastreo = New System.Windows.Forms.ProgressBar
        Me.btnRastreo = New System.Windows.Forms.Button
        Me.fraRastreo = New System.Windows.Forms.GroupBox
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Text = "Rastreo de Módulos y Funciones"
        Me.ClientSize = New System.Drawing.Size(401, 113)
        Me.Location = New System.Drawing.Point(408, 286)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.MinimizeBox = False
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ControlBox = True
        Me.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = True
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmRastreoFunciones"
        Me.prgRastreo.Size = New System.Drawing.Size(353, 20)
        Me.prgRastreo.Location = New System.Drawing.Point(24, 34)
        Me.prgRastreo.TabIndex = 1
        Me.prgRastreo.Name = "prgRastreo"
        Me.btnRastreo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnRastreo.Text = "Iniciar Rastreo"
        Me.btnRastreo.Size = New System.Drawing.Size(129, 25)
        Me.btnRastreo.Location = New System.Drawing.Point(264, 80)
        Me.btnRastreo.TabIndex = 2
        Me.btnRastreo.BackColor = System.Drawing.SystemColors.Control
        Me.btnRastreo.CausesValidation = True
        Me.btnRastreo.Enabled = True
        Me.btnRastreo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRastreo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnRastreo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnRastreo.TabStop = True
        Me.btnRastreo.Name = "btnRastreo"
        Me.fraRastreo.Text = "Progreso del Rastreo"
        Me.fraRastreo.ForeColor = System.Drawing.Color.FromArgb(0, 0, 128)
        Me.fraRastreo.Size = New System.Drawing.Size(385, 65)
        Me.fraRastreo.Location = New System.Drawing.Point(8, 8)
        Me.fraRastreo.TabIndex = 0
        Me.fraRastreo.BackColor = System.Drawing.SystemColors.Control
        Me.fraRastreo.Enabled = True
        Me.fraRastreo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraRastreo.Visible = True
        Me.fraRastreo.Name = "fraRastreo"
        Me.Controls.Add(prgRastreo)
        Me.Controls.Add(btnRastreo)
        Me.Controls.Add(fraRastreo)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub


    Public Sub Rastreo()
        On Error GoTo Merr
        'Paimí 21/Mayo/2003
        Dim I As Object
        Dim J As Integer
        Dim IntCodModulo As Double
        Dim blnTransaction As Boolean
        Dim MODULOS(12) As String

        Dim CATALOGOS(22, 2) As String
        Dim VENTAS(35, 2) As String '''27OCT2010 - MAVF
        Dim FACTURACION(9, 2) As String
        Dim BANCOS(23, 2) As String
        Dim INVENTARIOS(29, 2) As String
        Dim CONFIGURACION(8, 2) As String
        Dim SEGURIDAD(3, 2) As String
        Dim CXP(17, 2) As String

        Dim VENTASPV(29, 2) As String
        Dim INVENTARIOSPV(18, 2) As String
        Dim CONFIGURACIONPV(5, 2) As String
        Dim CATALOGOSPV(2, 2) As String

        'MODULOS DEL CORPORATIVO
        MODULOS(1) = "CATALOGOS"
        MODULOS(2) = "VENTAS"
        MODULOS(3) = "FACTURACION"
        MODULOS(4) = "BANCOS"
        MODULOS(5) = "INVENTARIOS"
        MODULOS(6) = "CONFIGURACION"
        MODULOS(7) = "SEGURIDAD"
        MODULOS(8) = "CXP"

        MODULOS(9) = "VENTASPV"
        MODULOS(10) = "INVENTARIOSPV"
        MODULOS(11) = "CONFIGURACIONPV"
        MODULOS(12) = "CATALOGOSPV"
        '----------------------------------------------------------------------------------------------------------------
        'MODULO CATÁLOGOS
        CATALOGOS(0, 0) = "FRMCORPOABCARTICULOS"
        CATALOGOS(1, 0) = "FRMCORPOABCCUENTASBANCARIAS"
        CATALOGOS(2, 0) = "FRMCORPOABCFAMILIAS"
        CATALOGOS(3, 0) = "FRMCORPOABCGRUPOS"
        CATALOGOS(4, 0) = "FRMCORPOABCLINEAS"
        CATALOGOS(5, 0) = "FRMCORPOABCMARCA"
        CATALOGOS(6, 0) = "FRMCORPOABCMODELOS"
        CATALOGOS(7, 0) = "FRMCORPOABCSUBLINEAS"
        CATALOGOS(8, 0) = "FRMCORPOABCCLIENTES"
        CATALOGOS(9, 0) = "FRMCORPOABCORIGENYAPLICACIONDERECURSOS"
        CATALOGOS(10, 0) = "FRMCORPOABCRUBROSDEAPLICACIONYORIGEN"
        CATALOGOS(11, 0) = "FRMCORPOABCBANCOS"
        CATALOGOS(12, 0) = "FRMCORPOABCFORMASDEPAGO"
        CATALOGOS(13, 0) = "FRMCORPOABCPROVACREED"
        CATALOGOS(14, 0) = "FRMCORPOABCSUCURSALES"
        CATALOGOS(15, 0) = "FRMCORPOABCTALLERES"
        CATALOGOS(16, 0) = "FRMCORPOABCTIPOSMATERIAL"
        CATALOGOS(17, 0) = "FRMCORPOABCVENDEDORES"
        CATALOGOS(18, 0) = "FRMCORPOABCDESCUENTOSVENDEXTERNOS"
        CATALOGOS(19, 0) = "FRMPROGRAMACIONPROMOCIONES"
        CATALOGOS(20, 0) = "FRMCORPOTARJETASCREDITOENPROMOCION"
        CATALOGOS(21, 0) = "FRMCORPOABCCOMISIONES"
        'DESCRIPCION DE LAS FUNCIONES DE CATÁLOGOS
        CATALOGOS(0, 1) = "ABC A ARTICULOS"
        CATALOGOS(1, 1) = "ABC A CUENTAS BANCARIAS"
        CATALOGOS(2, 1) = "ABC A FAMILIAS DE ARTICULOS"
        CATALOGOS(3, 1) = "ABC A GRUPOS DE ARICULOS"
        CATALOGOS(4, 1) = "ABC A LINEAS DE ARTICULOS"
        CATALOGOS(5, 1) = "ABC A MARCAS DE RELOJERIA"
        CATALOGOS(6, 1) = "ABC A MODELOS DE RELOJERIA"
        CATALOGOS(7, 1) = "ABC A SUBLINEAS DE JOYERIA"
        CATALOGOS(8, 1) = "ABC A CLIENTES"
        CATALOGOS(9, 1) = "ABC A ORIGEN Y APLICACION DE RECURSOS"
        CATALOGOS(10, 1) = "ABC A RUBROS DE ORIGEN Y APLICACION"
        CATALOGOS(11, 1) = "ABC A BANCOS"
        CATALOGOS(12, 1) = "ABC A FORMAS DE PAGO"
        CATALOGOS(13, 1) = "ABC A PROVEEDORES / ACREEDORES"
        CATALOGOS(14, 1) = "ABC A SUCURSALES"
        CATALOGOS(15, 1) = "ABC A TALLERES"
        CATALOGOS(16, 1) = "ABC A TIPOS DE MATERIAL"
        CATALOGOS(17, 1) = "ABC A VENDEDORES"
        CATALOGOS(18, 1) = "ABC A DESCUENTOS DE VENDEDORES EXTERNOS"
        CATALOGOS(19, 1) = "ABC A PROGRAMACION DE PROMOCIONES"
        CATALOGOS(20, 1) = "ABC A PROMOCIONES DE TARJETAS BANCARIAS"
        CATALOGOS(21, 1) = "ABC A COMISIONES POR VENTAS DE VENDEDORES"
        '----------------------------------------------------------------------------------------------------------------
        'VENTAS
        VENTAS(0, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIA"
        VENTAS(1, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIAPORPROV"
        VENTAS(2, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIACOMPARA"
        VENTAS(3, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIACLASIFARTIC"
        VENTAS(4, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIAUTILIDAD"
        VENTAS(5, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIARELOJERIA"
        VENTAS(6, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIARELOJMATERIAL"
        VENTAS(7, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIAFLUJOVENTA"
        VENTAS(8, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIAPORCLIENTE"
        VENTAS(9, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIAPORVENDEDOR"
        VENTAS(10, 0) = "FRMVTASRPTVENTASSALIDADEMERCANCIACOMISIONVENDEDOR"
        VENTAS(11, 0) = "FRMVTASRPTINGRESOSGENERALES" ''9/Oct/2003
        VENTAS(12, 0) = "FRMVTASRPTINGRESOSPORPERIODOYSUCURSAL"
        VENTAS(13, 0) = "FRMVTASRPTINGRESOSPORABONOS"
        VENTAS(14, 0) = "FRMVTASRPTINGRESOSPORREPARACIONES"
        VENTAS(15, 0) = "FRMVTASRPTINGRESOSPORCONCEPTODEPAGO"
        VENTAS(16, 0) = "FRMVTASVEENTRADADEMERCANCIA"
        VENTAS(17, 0) = "FRMVTASVESALIDADEMERCANCIA"
        VENTAS(18, 0) = "FRMVTASVEEXISTENCIAS"
        VENTAS(19, 0) = "FRMVTASVELIQUIDACIONVENDEDOREXTERNO"
        VENTAS(20, 0) = "FRMVTASVEINGRESOSSALIDADEMERCANCIAAVENDEXT"
        VENTAS(21, 0) = "FRMVTASVEDETALLADODEENTRADASSALIDAS"
        VENTAS(22, 0) = "FRMVTASREPORTEDEAPARTADOS"
        VENTAS(23, 0) = "FRMVTASREPORTEDEREPARACIONES"
        VENTAS(24, 0) = "FRMVTASREPORTEDECUENTASPORCOBRAR"
        VENTAS(25, 0) = "FRMVTASESTADODERESULTADOS"
        VENTAS(26, 0) = "FRMVTASRELACIONGASTOS"
        VENTAS(27, 0) = "FRMCORPOCONTROLREPARACIONES"
        VENTAS(28, 0) = "FRMVERIFICADORPRECIOS"
        VENTAS(29, 0) = "FRMVTASVENTASYEXISTENCIASPORPROVEEDOR"
        VENTAS(30, 0) = "FRMVTASVENTASPORRESURTIR"
        VENTAS(31, 0) = "FRMVTASVENTASYUTILIDAD"
        VENTAS(32, 0) = "FRMVENTASPORGRUPO"
        VENTAS(33, 0) = "FRMUTILIDADPORGRUPO"
        VENTAS(34, 0) = "FRMVTASVENTASYEXISTXFAM" '''27OCT2010 - MAVF

        VENTAS(0, 1) = "SAL MCIA. SALIDA DE MERCANCIA POR PERIODO Y SUCURSAL"
        VENTAS(1, 1) = "SAL MCIA. SALIDA DE MERCANCIA POR PROVEEDOR Y SUCURSAL"
        VENTAS(2, 1) = "SAL MCIA. COMPARATIVO DE VENTAS DIARIAS CON AÑO ANTERIOR"
        VENTAS(3, 1) = "SAL MCIA. SALIDA DE MERCANCÍA POR CLASIF. DE ARTICULOS"
        VENTAS(4, 1) = "SAL MCIA. SALIDA DE MERCANCÍA UTILIDAD POR LÍNEA"
        VENTAS(5, 1) = "SAL MCIA. SALIDA DE RELOJERÍA POR MARCA Y MODELO"
        VENTAS(6, 1) = "SAL MCIA. SALIDA DE RELOJERÍA POR MATERIAL DE FABRICACIÓN"
        VENTAS(7, 1) = "SAL MCIA. SALIDA DE MERCANCIA, FLUJO DE VENTA POR PROVEEDOR"
        VENTAS(8, 1) = "SAL MCIA. SALIDA DE MERCANCIA, VENTAS POR CLIENTE"
        VENTAS(9, 1) = "SAL MCIA. SALIDA DE MERCANCIA, POR VENDEDOR"
        VENTAS(10, 1) = "SAL MCIA. SALIDA DE MERCANCIA, COMISION POR VENDEDOR"
        VENTAS(11, 1) = "VTAS ING. GENERALES"
        VENTAS(12, 1) = "VTAS ING. PERÍODO Y SUCURSAL"
        VENTAS(13, 1) = "VTAS ING. ABONOS A DOCUMENTOS"
        VENTAS(14, 1) = "VTAS ING. REPARACIONES"
        VENTAS(15, 1) = "VTAS ING. CONCEPTO DE PAGO"
        VENTAS(16, 1) = "VEND EXT. ENTRADA DE MERCANCIA VENDEDOR EXTERNO"
        VENTAS(17, 1) = "VEND EXT. SALIDA DE MERCANCIA VENDEDOR EXTERNO"
        VENTAS(18, 1) = "VEND EXT. REPORTE DE EXISTENCIAS A PRECIO PUBLICO O AL COSTO"
        VENTAS(19, 1) = "VEND EXT. LIQUIDACION DE VENDEDOR EXTERNO"
        VENTAS(20, 1) = "VEND EXT. REPORTE DE INGRESOS POR VENTA A VEND. EXTERNO"
        VENTAS(21, 1) = "VEND EXT. REPORTE DE ENTRADAS/SALIDAS"
        VENTAS(22, 1) = "SAL MCIA. REPORTE DE APARTADOS"
        VENTAS(23, 1) = "SAL MCIA. REPORTE DE REPARACIONES"
        VENTAS(24, 1) = "SAL MCIA. CUENTAS POR COBRAR"
        VENTAS(25, 1) = "SAL MCIA. ESTADO DE RESULTADOS"
        VENTAS(26, 1) = "REPORTE DE RELACION DE GASTOS"
        VENTAS(27, 1) = "SAL MCIA. CONTROL DE REPARACIONES"
        VENTAS(28, 1) = "VERIFICADOR DE PRECIOS"
        VENTAS(29, 1) = "REPORTE DE VENTAS Y EXISTENCIAS POR PROVEEDOR"
        VENTAS(30, 1) = "REPORTE DE VENTAS POR RESURTIR"
        VENTAS(31, 1) = "REPORTE DE VENTAS Y UTILIDAD GLOBAL POR GRUPO"
        VENTAS(32, 1) = "REPORTE DE VENTAS POR GRUPO"
        VENTAS(33, 1) = "REPORTE DE UTILIDAD POR GRUPO"
        VENTAS(34, 1) = "REPORTE DE VENTAS Y EXISTENCIAS POR FAMILIA" '''27OCT2010 - MAVF

        '----------------------------------------------------------------------------------------------------------------
        'MODULO FACTURACION
        FACTURACION(0, 0) = "FRMFACTANALISISVENTAS"
        FACTURACION(1, 0) = "FRMFACTFACTURACIONESPECIAL"
        FACTURACION(2, 0) = "FRMFACTREPORTESFACTURACIONGLOBALXSUCURSAL"
        FACTURACION(3, 0) = "FRMFACTREPORTESFACTURACIONDETALLADAXSUCURSAL"
        FACTURACION(4, 0) = "FRMFACTREPORTESIMPRESIONTICKETS"
        FACTURACION(5, 0) = "FRMFACTREPORTESMEJORESCLIENTES"
        FACTURACION(6, 0) = "FRMREIMPRESIONCORTEFINAL"
        FACTURACION(7, 0) = "FRMPVDIARIOMOVTOS"
        FACTURACION(8, 0) = "FRMFACTREPORTEVTASTARJETACREDITOXSUCURSAL"
        FACTURACION(0, 1) = "ANÁLISIS DE LAS VENTAS"
        FACTURACION(1, 1) = "FACTURACIÓN ESPECIAL"
        FACTURACION(2, 1) = "REPORTES DE FACTURACIÓN GLOBAL POR SUCURSAL"
        FACTURACION(3, 1) = "REPORTES DE FACTURACIÓN DETALLADOS POR SUCURSAL"
        FACTURACION(4, 1) = "REIMPRESIÓN DE TICKETS"
        FACTURACION(5, 1) = "REPORTE DE LOS MEJORES CLIENTES"
        FACTURACION(6, 1) = "REIMPRESION DEL CORTE FINAL"
        FACTURACION(7, 1) = "DIARIO DE MOVIMIENTOS"
        FACTURACION(8, 1) = "REPORTE DE VENTAS CON TARJETA DE CREDITO"

        '----------------------------------------------------------------------------------------------------------------
        'MODULO COMPRAS Y CUENTAS POR PAGAR
        CXP(0, 0) = "FRMCXPORDENCOMPRA"
        CXP(1, 0) = "FRMCXPREGFACTCOMPRAS"
        CXP(2, 0) = "FRMCXPREGFACTGASTOS"
        CXP(3, 0) = "FRMCXPPROGPAGOS"
        CXP(4, 0) = "FRMCXPREGNOTASCREDITO"
        CXP(5, 0) = "FRMCXPRPTMEJORESPROV"
        CXP(6, 0) = "FRMCXPRPTCOMPRASPORPROVEEDOR"
        CXP(7, 0) = "FRMCXPRPTOC"
        CXP(8, 0) = "FRMCXPRPTARTICULOSPENDIENTES"
        CXP(9, 0) = "FRMCXPRPTNOTASCREDITO"
        CXP(10, 0) = "FRMCXPRPTFACTURAS"
        CXP(11, 0) = "FRMCXPEMISIONPAGOS"
        CXP(12, 0) = "FRMCXPCUENTASPORPAGAR"
        CXP(13, 0) = "FRMCXPREPORTESALDOXPROVEEDOR"
        CXP(14, 0) = "FRMCXPPRESUPUESTADO"
        CXP(15, 0) = "FRMCXPREPORTESALDOXPROVEEDORES"
        CXP(16, 0) = "FRMCXPREGFACTCOMPRASCARGAINICIAL"

        CXP(0, 1) = "REGISTRO DE ORDEN DE COMPRA"
        CXP(1, 1) = "REGISTRO DE FACTURAS DE COMPRAS"
        CXP(2, 1) = "REGISTRO DE FACTURAS DE GASTOS"
        CXP(3, 1) = "PROGRAMACION ESPECIAL DE PAGOS"
        CXP(4, 1) = "NOTAS DE CREDITO"
        CXP(5, 1) = "REPORTE DE LOS MEJORES PROVEEDORES"
        CXP(6, 1) = "ANALISIS ANUAL DE LAS COMPRAS"
        CXP(7, 1) = "REPORTE DE ORDENES DE COMPRA"
        CXP(8, 1) = "REPORTE DE ARTICULOS PENDIENTES"
        CXP(9, 1) = "REPORTE DE NOTAS DE CRÉDITO"
        CXP(10, 1) = "REPORTE DE FACTURAS DE CXP"
        CXP(11, 1) = "EMISION DE PAGOS"
        CXP(12, 1) = "CUENTAS POR PAGAR"
        CXP(13, 1) = "AUXILIAR DE PROVEEDORES"
        CXP(14, 1) = "CXP PRESUPUESTADO"
        CXP(15, 1) = "SALDOS POR PROVEEDOR"
        CXP(16, 1) = "CARGA INICIAL DE FACTURAS CXP"

        '----------------------------------------------------------------------------------------------------------------
        'MODULO BANCOS
        BANCOS(0, 0) = "FRMBANCOSPROCESODIARIOREGISTRODEPAGOS"
        BANCOS(1, 0) = "FRMBANCOSPROCESODIARIOREGISTRODEDEPOSITOS"
        BANCOS(2, 0) = "FRMBANCOSPROCESODIARIOCARGOSDIVERSOS"
        BANCOS(3, 0) = "FRMBANCOSPROCESODIARIOTRASPASOSBANCARIOS"
        BANCOS(4, 0) = "FRMBANCOSPROCESODIARIOANTICIPOPROVEEDORESACREED"
        BANCOS(5, 0) = "FRMBANCOSPROCESODIARIOREGISTRODEOTROSINGRESOS"
        BANCOS(6, 0) = "FRMBANCOSPROCESODIARIOCANCELACIONDEMOVIMIENTOSBANC"
        BANCOS(7, 0) = "FRMBANCOSPROCESODIARIOCONSULTADESALDOS"
        BANCOS(8, 0) = "FRMBANCOSPROCESODIARIOCIERREDIARIOBANCOS"
        BANCOS(9, 0) = "FRMBANCOSPROCESOMENSUALCONCILIACIONMENSUAL"
        BANCOS(10, 0) = "FRMBANCOSPROCESOMENSUALMOVIMIENTOSENCONCILIACION"
        BANCOS(11, 0) = "FRMBANCOSPROCESOMENSUALFLUJOCAJAGENERAL"
        BANCOS(12, 0) = "FRMBANCOSPROCESOMENSUALCONSULTAORIGENAPLICREC"
        BANCOS(13, 0) = "FRMBANCOSPROCESOMENSUALREPORTEORIGENYAPLICACION"
        BANCOS(14, 0) = "FRMBANCOSPROCESOMENSUALCIERRECONCILIACION"
        BANCOS(15, 0) = "FRMBANCOSPROCESOMENSUALDEPURACIONDEMOVIMIENTOSHIST"
        BANCOS(16, 0) = "FRMABCBANCOS"
        BANCOS(17, 0) = "FRMABCCUENTASBANCARIAS"
        BANCOS(18, 0) = "FRMABCORIGENYAPLICACIONDERECURSOS"
        BANCOS(19, 0) = "FRMABCRUBROSDEAPLICACIONYORIGEN"
        BANCOS(20, 0) = "FRMBANCOSREPORTEMOVBANCARIOSXTIPO"
        BANCOS(21, 0) = "FRMBANCOSREPORTEDEMOVIMIENTOSBANCARIOS"
        BANCOS(22, 0) = "FRMBANCOSREPORTEANALISISDIARIOBANCOS"

        BANCOS(0, 1) = "REGISTRO DE PAGOS"
        BANCOS(1, 1) = "REGISTRO DE DEPOSITOS"
        BANCOS(2, 1) = "REGISTRO DE CARGOS DIVERSOS"
        BANCOS(3, 1) = "TRASPASOS BANCARIOS"
        BANCOS(4, 1) = "ANTICIPO A PROVEEDORES / ACREEDORES"
        BANCOS(5, 1) = "REGISTRO DE OTROS INGRESOS"
        BANCOS(6, 1) = "CANCELACION DE MOVIMIENTOS"
        BANCOS(7, 1) = "CONSULTA DE SALDOS"
        BANCOS(8, 1) = "CIERRE DIARIO DE BANCOS"
        BANCOS(9, 1) = "CONCILIACION MENSUAL"
        BANCOS(10, 1) = "REPORTE DE MOVIMIENTOS EN CONCILIACION"
        BANCOS(11, 1) = "FLUJO DE LA CAJA GENERAL"
        BANCOS(12, 1) = "CONSULTA DE ORIGEN Y APLICACION DE RECURSOS"
        BANCOS(13, 1) = "REPORTE DE ORIGEN Y APLICACION DE RECURSOS"
        BANCOS(14, 1) = "CIERRE DE CONCILIACION MENSUAL"
        BANCOS(15, 1) = "DEPURACION DE MOVIMIENTOS HISTORICOS"
        BANCOS(16, 1) = "ABC BANCOS"
        BANCOS(17, 1) = "ABC CUENTAS BANCARIAS"
        BANCOS(18, 1) = "ABC ORIGEN Y APLICACION DE RECURSOS"
        BANCOS(19, 1) = "ABC RUBROS DE ORIGEN Y APLICACION DE RECURSOS"
        BANCOS(20, 1) = "MOVIMIENTOS BANCARIOS POR TIPO"
        BANCOS(21, 1) = "MOVIMIENTOS BANCARIOS"
        BANCOS(22, 1) = "ANALISIS DIARO DE BANCOS"

        '----------------------------------------------------------------------------------------------------------------
        'MODULO DE INVENTARIOS - CORPO
        INVENTARIOS(0, 0) = "frmInvSalidaPorVenta"
        INVENTARIOS(1, 0) = "frmInvSalidaPorTransferencia"
        INVENTARIOS(2, 0) = "frmInvSalidaPorDevolSobreCompra"
        INVENTARIOS(3, 0) = "frmInvSalidaAVendedoresExternos"
        INVENTARIOS(4, 0) = "frmInvSalidaPorVentaAVendedoresExternos"
        INVENTARIOS(5, 0) = "frmInvSalidaPorObsequio"
        INVENTARIOS(6, 0) = "frmInvSalidaPorPrestamo"
        INVENTARIOS(7, 0) = "frmInvSalidaporAjuste"
        INVENTARIOS(8, 0) = "frmInvEntradaPorCompra"
        INVENTARIOS(9, 0) = "frmInvEntradaPorTransferencia"
        INVENTARIOS(10, 0) = "frmInvEntradaPorDevolSobreVenta"
        INVENTARIOS(11, 0) = "frmInvEntradaPorDevoldeVendedoresExternos"
        INVENTARIOS(12, 0) = "frmInvEntradaPorDevolSobreVentaAVendedoresExternos"
        INVENTARIOS(13, 0) = "frmInvEntradaPorDevolSobreObsequio"
        INVENTARIOS(14, 0) = "frmInvEntradaPorDevolPorPrestamo"
        INVENTARIOS(15, 0) = "frmInvEntradaporAjuste"

        '    INVENTARIOS(16, 0) = "frmInvHojadeControl"
        '    INVENTARIOS(17, 0) = "frmInvCapturaInvFisico"
        '    INVENTARIOS(18, 0) = "frmInvAnalisisComparativo"
        '    INVENTARIOS(19, 0) = "frmInvGenerarAjustes"
        INVENTARIOS(20, 0) = "frmRptKardexArticulo"
        INVENTARIOS(21, 0) = "frmReportePrestamosPendientes"
        INVENTARIOS(22, 0) = "frmRptExistenciasyCostos"
        INVENTARIOS(23, 0) = "frmrptComparacionExistenciaStock"
        INVENTARIOS(24, 0) = "frmRptTransferenciasNoConciliadas"
        INVENTARIOS(25, 0) = "frmStockBasicoTienda"
        INVENTARIOS(26, 0) = "frmImpresionEtiquetas"
        INVENTARIOS(27, 0) = "frmInvHojadeControl"
        INVENTARIOS(28, 0) = "frmInvAnalisisComparativo"

        INVENTARIOS(0, 1) = "INV. SALIDA POR VENTA"
        INVENTARIOS(1, 1) = "INV. SALIDA POR TRANSFERENCIA"
        INVENTARIOS(2, 1) = "INV. SALIDA POR DEVOLUCION SOBRE COMPRA"
        INVENTARIOS(3, 1) = "INV. SALIDA A VENDEDORES EXTERNOS"
        INVENTARIOS(4, 1) = "INV. SALIDA POR VENTA A VENDEDORES EXTERNOS"
        INVENTARIOS(5, 1) = "INV. SALIDA POR OBSEQUIO"
        INVENTARIOS(6, 1) = "INV. SALIDA POR PRESTAMO"
        INVENTARIOS(7, 1) = "INV. SALIDA POR AJUSTE"
        INVENTARIOS(8, 1) = "INV. ENTRADA POR COMPRA"
        INVENTARIOS(9, 1) = "INV. ENTRADA POR TRANSFERENCIA"
        INVENTARIOS(10, 1) = "INV. ENTRADA POR DEVOLUCION SOBRE VTA"
        INVENTARIOS(11, 1) = "INV. ENTRADA POR DEVOLUCION DE VENDEDORES EXTERNOS"
        INVENTARIOS(12, 1) = "INV. ENTRADA POR DEVOLUCION SOBRE VENTA A VENDEDORES EXTERNOS"
        INVENTARIOS(13, 1) = "INV. ENTRADA POR DEVOLUCION SOBRE OBSEQUIO"
        INVENTARIOS(14, 1) = "INV. ENTRADA POR DEVOLUSION SOBRE PRESTAMO"
        INVENTARIOS(15, 1) = "INV. ENTRADA POR AJUSTE"

        '    INVENTARIOS(16, 1) = "INV PROC. HOJA DE CONTROL"
        '    INVENTARIOS(17, 1) = "INV PROC. INVENTARIO FISICO"
        '    INVENTARIOS(18, 1) = "INV PROC. ANALISIS COMPARATIVO"
        '    INVENTARIOS(19, 1) = "INV PROC. GENERACION DE AJUSTES"
        INVENTARIOS(20, 1) = "INV RPT. KARDEX POR ARTICULO"
        INVENTARIOS(21, 1) = "INV RPT. PRESTAMOS PENDIENTES"
        INVENTARIOS(22, 1) = "INV RPT. EXISTENCIAS Y COSTOS"
        INVENTARIOS(23, 1) = "INV RPT. COMPARACION EXISTENCIA-STOCK"
        INVENTARIOS(24, 1) = "INV RPT. TRANSFERENCIAS NO CONCILIADAS"
        INVENTARIOS(25, 1) = "INV. STOCK BASICO DE TIENDA"
        INVENTARIOS(26, 1) = "INV. IMPRESION DE ETIQUETAS"
        INVENTARIOS(27, 1) = "IMPRESION DE HOJA DE CONTROL"
        INVENTARIOS(28, 1) = "ANALISIS COMPARATIVO"

        '----------------------------------------------------------------------------------------------------------------
        'MODULO CONFIGURACION
        CONFIGURACION(0, 0) = "FRMCONFIGGRALCORPORATIVO"
        CONFIGURACION(1, 0) = "FRMPVCONFIGPUNTOVENTA"
        CONFIGURACION(2, 0) = "FRMPVCONFIGTICKETVENTA"
        CONFIGURACION(3, 0) = "FRMPVCONFIGFACTURACION"
        CONFIGURACION(4, 0) = "FRMPVCONFIGFOLIOS"
        CONFIGURACION(5, 0) = "FRMPVCONFIGCAJA"
        CONFIGURACION(6, 0) = "FRMCONFIGURACION"
        CONFIGURACION(7, 0) = "FRMIMPORTACIONIMAGENES"
        '''CONFIGURACION(6, 0) = "CAMBIOUSUARIO"

        CONFIGURACION(0, 1) = "CONFIGURACION GENERAL DEL CORPORATIVO"
        CONFIGURACION(1, 1) = "CONFIGURACION GENERAL DEL PUNTO DE VENTA"
        CONFIGURACION(2, 1) = "CONFIGURACION GENERAL TICKET PUNTO VENTA"
        CONFIGURACION(3, 1) = "CONFIGURACION GENERAL FACTURAS PUNTO VENTA"
        CONFIGURACION(4, 1) = "CONFIGURACION FOLIOS DEL PUNTO VENTA"
        CONFIGURACION(5, 1) = "CONFIGURACION CAJAS DEL PUNTO VENTA"
        CONFIGURACION(6, 1) = "CONFIGURACION DE IMPRESORAS"
        CONFIGURACION(7, 1) = "IMPORTACION DE IMAGENES"

        '''CONFIGURACION(6, 1) = "CAMBIO DE USUARIO"
        '----------------------------------------------------------------------------------------------------------------
        'MODULO SEGURIDAD
        SEGURIDAD(0, 0) = "FRMRASTREOFUNCIONES"
        SEGURIDAD(1, 0) = "FRMABCMODULOS"
        SEGURIDAD(2, 0) = "FRMABCUSUARIOS"

        SEGURIDAD(0, 1) = "RASTREO DE FUNCIONES"
        SEGURIDAD(1, 1) = "ABC DE MÓDULOS Y FUNCIONES DEL SISTEMA"
        SEGURIDAD(2, 1) = "ABC DE USUARIOS"
        '----------------------------------------------------------------------------------------------------------------
        '''MODULO DE VENTAS - PV
        '''    VENTASPV(0, 0) = "FRMPVCONFIGCAJA"
        '''    VENTASPV(1, 0) = "FRMPVCONFIGFACTURACION"
        '''    VENTASPV(2, 0) = "FRMPVCONFIGFOLIOS"
        '''    VENTASPV(3, 0) = "FRMPVCONFTICKETVTA"
        '''    VENTASPV(4, 0) = "FRMPVCONFIGPUNTOVENTA"
        VENTASPV(5, 0) = "FRMVENTASSALMERCANCIA"
        VENTASPV(6, 0) = "FRMREGISTROAPARTADOS"
        VENTASPV(7, 0) = "FRMABONOAPARTADOS"
        VENTASPV(8, 0) = "FRMREPORTEAPARTADOS"
        VENTASPV(9, 0) = "FRMESTADOCUENTAAPARTADOS"
        VENTASPV(10, 0) = "FRMOBSEQUIOS"
        VENTASPV(11, 0) = "FRMCONTROLREPARACIONES"
        VENTASPV(12, 0) = "FRMREPORTEREPARACIONES"
        VENTASPV(13, 0) = "FRMABONODOCUMENTOS"
        VENTASPV(14, 0) = "FRMDEVOLUCIONMERCANCIA"
        VENTASPV(15, 0) = "FRMFACTURACIONTICKETS"
        VENTASPV(16, 0) = "FRMRETVALADEPOSITOS"
        VENTASPV(17, 0) = "FRMRETVALENVIOCAJAGRAL"
        VENTASPV(18, 0) = "FRMRETVALPAGOSTERCEROS"
        VENTASPV(19, 0) = "FRMVERIFICADORPRECIOS"
        '''    VENTASPV(20, 0) = "FRMCONFIGURACION"
        VENTASPV(21, 0) = "FRMREPORTEOBSEQUIOS"
        VENTASPV(22, 0) = "FRMRPTDIARIOMOVIMIENTOS"
        VENTASPV(23, 0) = "FRMCORTEFINALDIA"
        VENTASPV(24, 0) = "FRMCORTEAUDITORIA"
        VENTASPV(25, 0) = "FRMREIMPRESIONCORTEFINAL"
        VENTASPV(26, 0) = "FRMAPARTADOSPORCATALOGO"
        VENTASPV(27, 0) = "FRMCORPOCONTROLREPARACIONES_CORPO"
        VENTASPV(28, 0) = "FRMVENTASFACTPENDXIMPRIMIR"

        '''    VENTASPV(0, 1) = "CONS. CONFIGURACION CAJA"
        '''    VENTASPV(1, 1) = "CONS. CONFIGURACION FACTURA"
        '''    VENTASPV(2, 1) = "CONS. CONFIGURACION FOLIOS"
        '''    VENTASPV(3, 1) = "CONS. CONFIGURACION TICKET VENTA"
        '''    VENTASPV(4, 1) = "CONS. CONFIGURACINO GRAL PTO VTA"
        VENTASPV(5, 1) = "VTAS. VENTAS CREDITO/CONTADO"
        VENTASPV(6, 1) = "VTAS. REGISTRO DE APARTADOS"
        VENTASPV(7, 1) = "VTAS. ABONOS DE APARTADOS"
        VENTASPV(8, 1) = "VTAS. REPORTE DE APARTADOS"
        VENTASPV(9, 1) = "VTAS. ESTADO DE CUENTA DE APARTADOS"
        VENTASPV(10, 1) = "VTAS. REGISTRO DE OBSEQUIOS"
        VENTASPV(11, 1) = "VTAS. CONTROL DE REPARACIONES"
        VENTASPV(12, 1) = "VTAS. REPORTE DE REPARACIONES PENDIENTES"
        VENTASPV(13, 1) = "VTAS. ABONOS A DOCUMENTOS"
        VENTASPV(14, 1) = "VTAS. DEVOLUCION DE MERCANCIA"
        VENTASPV(15, 1) = "VTAS. FACTURACION DE TICKETS"
        VENTASPV(16, 1) = "RETIRO DE VALORES A DEPOSITOS"
        VENTASPV(17, 1) = "RETIRO DE VALORES A CAJA GENERAL"
        VENTASPV(18, 1) = "RETIRO DE VALORES POR PAGO A TERCERSO"
        VENTASPV(19, 1) = "VERIFICADOR DE PRECIOS"
        '''    VENTASPV(20, 1) = "CONFIGURACION DE IMPRESORA"
        VENTASPV(21, 1) = "REPORTE DE OBSEQUIOS"
        VENTASPV(22, 1) = "REPORTE DE MOVIMIENTOS DIARIOS"
        VENTASPV(23, 1) = "CORTE. FINAL DEL DIA"
        VENTASPV(24, 1) = "CORTE. AUDITORIA"
        VENTASPV(25, 1) = "CORTE. REIMPRESION DE CORTE FINAL"
        VENTASPV(26, 1) = "VTAS. APARTADOS POR CATALOGO"
        VENTASPV(27, 1) = "ADMINISTRACION DE REPARACIONES"
        VENTASPV(28, 1) = "FACTURAS PENDIENTES POR IMPRIMIR"

        '----------------------------------------------------------------------------------------------------------------
        'MODULO DE INVENTARIOS - PUNTO DE VENTA
        INVENTARIOSPV(0, 0) = "frmInvSalidaPorTransferencia"
        INVENTARIOSPV(1, 0) = "frmInvSalidaPorVenta"
        INVENTARIOSPV(2, 0) = "frmInvSalidaPorObsequio"
        INVENTARIOSPV(3, 0) = "frmInvSalidaPorPrestamo"
        INVENTARIOSPV(4, 0) = "frmInvSalidaporAjuste"
        INVENTARIOSPV(5, 0) = "frmInvEntradaPorTransferencia"
        INVENTARIOSPV(6, 0) = "frmInvEntradaPorDevolSobreVenta"
        INVENTARIOSPV(7, 0) = "frmInvEntradaPorDevolSobreObsequio"
        INVENTARIOSPV(8, 0) = "frmInvEntradaPorDevolPorPrestamo"
        INVENTARIOSPV(9, 0) = "frmInvEntradaporAjuste"

        INVENTARIOSPV(10, 0) = "frmInvHojadeControl"
        INVENTARIOSPV(11, 0) = "frmInvCapturaInvFisico"
        INVENTARIOSPV(12, 0) = "frmInvAnalisisComparativo"
        INVENTARIOSPV(13, 0) = "frmInvGenerarAjustes"
        INVENTARIOSPV(14, 0) = "frmRptKardexArticulo"
        INVENTARIOSPV(15, 0) = "frmReportePrestamosPendientes"
        INVENTARIOSPV(16, 0) = "frmRptExistenciasyCostos"
        INVENTARIOSPV(17, 0) = "frminvelectronico"

        INVENTARIOSPV(0, 1) = "INV. SALIDA POR TRANSFERENCIA"
        INVENTARIOSPV(1, 1) = "INV. SALIDA POR VENTA"
        INVENTARIOSPV(2, 1) = "INV. SALIDA POR OBSEQUIO"
        INVENTARIOSPV(3, 1) = "INV. SALIDA POR PRESTAMO"
        INVENTARIOSPV(4, 1) = "INV. SALIDA POR AJUSTE"
        INVENTARIOSPV(5, 1) = "INV. ENTRADA POR TRANSFERENCIA"
        INVENTARIOSPV(6, 1) = "INV. ENTRADA POR DEVOLUCION SOBRE VTA"
        INVENTARIOSPV(7, 1) = "INV. ENTRADA POR DEVOLUCION SOBRE OBSEQUIO"
        INVENTARIOSPV(8, 1) = "INV. ENTRADA POR DEVOLUSION SOBRE PRESTAMO"
        INVENTARIOSPV(9, 1) = "INV. ENTRADA POR AJUSTE"

        INVENTARIOSPV(10, 1) = "INV PROC. HOJA DE CONTROL"
        INVENTARIOSPV(11, 1) = "INV PROC. INVENTARIO FISICO"
        INVENTARIOSPV(12, 1) = "INV PROC. ANALISIS COMPARATIVO"
        INVENTARIOSPV(13, 1) = "INV PROC. GENERACION DE AJUSTES"
        INVENTARIOSPV(14, 1) = "INV RPT. KARDEX POR ARTICULO"
        INVENTARIOSPV(15, 1) = "INV RPT. PRESTAMOS PENDIENTES"
        INVENTARIOSPV(16, 1) = "INV RPT. EXISTENCIAS Y COSTOS"
        INVENTARIOSPV(17, 1) = "INV PROC. INVENTARIO ELECTRONICO"
        '----------------------------------------------------------------------------------------------------------------
        '''CATALOGOS - PTOVTA
        CATALOGOSPV(0, 0) = "frmPVABCClientes"
        CATALOGOSPV(1, 0) = "frmPVABCRFC"

        CATALOGOSPV(0, 1) = "ABC A CLIENTES"
        CATALOGOSPV(1, 1) = "ABC A DATOS FISCALES"
        '----------------------------------------------------------------------------------------------------------------
        '''CONFIGURACION PV
        CONFIGURACIONPV(0, 0) = "FRMPVCONFIGPUNTOVENTA"
        CONFIGURACIONPV(1, 0) = "FRMPVCONFIGCAJA"
        CONFIGURACIONPV(2, 0) = "FRMPVCONFIGFOLIOS"
        CONFIGURACIONPV(3, 0) = "FRMCONFIGURACION"
        '''CONFIGURACIONPV(4, 0) = "CAMBIOUSUARIO"

        CONFIGURACIONPV(0, 1) = "CONF. CONFIGURACION PUNTO VENTA"
        CONFIGURACIONPV(1, 1) = "CONF. CONFIGURACION CAJA"
        CONFIGURACIONPV(2, 1) = "CONF. CONFIGURACION FOLIOS"
        CONFIGURACIONPV(3, 1) = "CONF. CONFIGURACION DE IMPRESORA"
        '''CONFIGURACIONPV(4, 1) = "CONF. CAMBIO DE USUARIO"
        '----------------------------------------------------------------------------------------------------------------

        Cnn.BeginTrans()
        blnTransaction = True
        IntCodModulo = 0
        'Cambiar la apariencia del Cursor del Mouse
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        ''*******************************************************************************
        ''* En esta parte inicia el proceso de rastreo de Modulos, usados en el sistema *
        ''* Lo que se hara es buscar cada modulo en la tabla de CatModulos, si no se    *
        ''* encuentra el modulo se insertará en la tabla                                *
        ''*******************************************************************************
        Me.prgRastreo.Maximum = 12
        Me.prgRastreo.Value = 0
        For I = 1 To 12
            Me.prgRastreo.Value = I
            ModEstandar.BorraCmd()
            gStrSql = "Select * From CatModulos Where DescModulo LIKE '" & MODULOS(I) & "'"
            Cmd.CommandText = "UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            IntCodModulo = 0
            If RsGral.RecordCount = 0 Then
                ModStoredProcedures.PR_IMECatModulos(Str(0), Trim(MODULOS(I)), C_INSERCION, CStr(0))
                Cmd.Execute()
                IntCodModulo = Cmd.Parameters("ID").Value
            Else
                IntCodModulo = RsGral.Fields("CodModulo").Value
            End If
            ModEstandar.BorraCmd()
            gStrSql = "Select * From CatModulos Where CodModulo =" & IntCodModulo
            Cmd.CommandText = "UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute

            If IntCodModulo > 0 Then
                '''*******************************************************************************
                ''* En esta parte inicia el proceso de rastro de Funciones, usadas en el sistema *
                ''* Lo que se hara es buscar cada Funcion en la tabla de CatFunciones, si no se  *
                ''* encuentra la Funcion se insertará en la tabla                                *
                ''********************************************************************************
                Select Case Trim(RsGral.Fields("DescModulo").Value)
                    Case "CATALOGOS"
                        For J = 1 To 22
                            If Trim(CATALOGOS(J - 1, nFORMA)) <> "" And Trim(CATALOGOS(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(CATALOGOS(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), CATALOGOS(J - 1, nDESC), CATALOGOS(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J
                        '''27OCT2010 - MAVF
                    Case "VENTAS"
                        For J = 1 To 35
                            If Trim(VENTAS(J - 1, nFORMA)) <> "" And Trim(VENTAS(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(VENTAS(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), VENTAS(J - 1, nDESC), VENTAS(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J
                    Case "CONFIGURACION"
                        For J = 1 To 8
                            If Trim(CONFIGURACION(J - 1, nFORMA)) <> "" And Trim(CONFIGURACION(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(CONFIGURACION(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), CONFIGURACION(J - 1, nDESC), CONFIGURACION(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J
                    Case "FACTURACION"
                        For J = 1 To 9
                            If Trim(FACTURACION(J - 1, nFORMA)) <> "" And Trim(FACTURACION(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(FACTURACION(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), FACTURACION(J - 1, nDESC), FACTURACION(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J
                    Case "SEGURIDAD"
                        For J = 1 To 4
                            If Trim(SEGURIDAD(J - 1, nFORMA)) <> "" And Trim(SEGURIDAD(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(SEGURIDAD(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), SEGURIDAD(J - 1, nDESC), SEGURIDAD(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J
                    Case "CXP"
                        For J = 1 To 17
                            If Trim(CXP(J - 1, nFORMA)) <> "" And Trim(CXP(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(CXP(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), CXP(J - 1, nDESC), CXP(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J
                    Case "BANCOS"
                        For J = 1 To 23
                            If Trim(BANCOS(J - 1, nFORMA)) <> "" And Trim(BANCOS(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(BANCOS(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), BANCOS(J - 1, nDESC), BANCOS(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J
                    Case "INVENTARIOS"
                        For J = 1 To 29
                            If Trim(INVENTARIOS(J - 1, nFORMA)) <> "" And Trim(INVENTARIOS(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(INVENTARIOS(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), INVENTARIOS(J - 1, nDESC), INVENTARIOS(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J

                        '''PUNTO DE VENTA
                    Case "VENTASPV"
                        For J = 1 To 29
                            If Trim(VENTASPV(J - 1, nFORMA)) <> "" And Trim(VENTASPV(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(VENTASPV(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), VENTASPV(J - 1, nDESC), VENTASPV(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J

                    Case "INVENTARIOSPV"
                        For J = 1 To 18
                            If Trim(INVENTARIOSPV(J - 1, nFORMA)) <> "" And Trim(INVENTARIOSPV(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(INVENTARIOSPV(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), INVENTARIOSPV(J - 1, nDESC), INVENTARIOSPV(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J

                    Case "CATALOGOSPV"
                        For J = 1 To 2
                            If Trim(CATALOGOSPV(J - 1, nFORMA)) <> "" And Trim(CATALOGOSPV(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(CATALOGOSPV(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), CATALOGOSPV(J - 1, nDESC), CATALOGOSPV(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J

                    Case "CONFIGURACIONPV"
                        For J = 1 To 5
                            If Trim(CONFIGURACIONPV(J - 1, nFORMA)) <> "" And Trim(CONFIGURACIONPV(J - 1, nDESC)) <> "" Then
                                ModEstandar.BorraCmd()
                                gStrSql = "Select * From CatFunciones Where codmodulo=" & IntCodModulo & " and Forma LIKE '" & Trim(CONFIGURACIONPV(J - 1, nFORMA)) & "'"
                                Cmd.CommandText = "UP_Select_Datos"
                                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                                RsGral = Cmd.Execute
                                If RsGral.RecordCount = 0 Then
                                    'Insertar la función
                                    ModStoredProcedures.PR_IMECatFunciones(Str(IntCodModulo), Str(0), CONFIGURACIONPV(J - 1, nDESC), CONFIGURACIONPV(J - 1, nFORMA), C_INSERCION, CStr(0))
                                    Cmd.Execute()
                                End If
                            End If
                        Next J
                End Select
            End If
        Next I
        If I = 13 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            MsgBox("Integración de funciones finalizada con éxito", MsgBoxStyle.Information, "Aviso")
            Me.prgRastreo.Value = 0
        End If
        Cnn.CommitTrans()
        blnTransaction = False

        System.Array.Clear(MODULOS, 0, MODULOS.Length)
        System.Array.Clear(CATALOGOS, 0, CATALOGOS.Length)
        System.Array.Clear(VENTAS, 0, VENTAS.Length)
        System.Array.Clear(FACTURACION, 0, FACTURACION.Length)
        System.Array.Clear(BANCOS, 0, BANCOS.Length)
        System.Array.Clear(INVENTARIOS, 0, INVENTARIOS.Length)
        System.Array.Clear(CONFIGURACION, 0, CONFIGURACION.Length)
        System.Array.Clear(SEGURIDAD, 0, SEGURIDAD.Length)
        System.Array.Clear(CXP, 0, CXP.Length)

        System.Array.Clear(VENTASPV, 0, VENTASPV.Length)
        System.Array.Clear(INVENTARIOSPV, 0, INVENTARIOSPV.Length)
        System.Array.Clear(CONFIGURACIONPV, 0, CONFIGURACIONPV.Length)
        System.Array.Clear(CATALOGOSPV, 0, CATALOGOSPV.Length)

Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub btnRastreo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnRastreo.Click
        Me.btnRastreo.Enabled = False
        Rastreo()
        Me.btnRastreo.Enabled = True
    End Sub

    Private Sub frmRastreoFunciones_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmRastreoFunciones_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmRastreoFunciones_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Escape
                mblnSalir = True
                Me.Close()
        End Select
    End Sub

    Private Sub frmRastreoFunciones_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
    End Sub

    Private Sub frmRastreoFunciones_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If mblnSalir Then
            mblnSalir = False
            Cancel = 0
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmRastreoFunciones_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub


End Class