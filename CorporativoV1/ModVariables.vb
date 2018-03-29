'**********************************************************************************************************************'
'*PROGRAMA: MODULO VARIABLES JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018    
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB
Imports System
Imports System.Windows.Forms
Imports System.Data
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6

Public Module ModVariables
    Public Printer As Printer
    Public rutaArchivoTxt = My.Application.Info.DirectoryPath & "\Sistema\CJoyeria.Txt"
    Public conexionLocal As String = ""
    Public conexionServidor As String = ""
    Public Servidor As String = ""
    Public Bd As String = ""
    Public ArchivoTxt, F As Object


    Public gIntContFallos As Integer 'Vatiable para llevar el control de los fallos al intentar igresar al sistema}
    Public gIntCodUsuario As Integer 'Variable para almacenar el Codigo deUsuario que se está conectando al sistema
    Public gStrNomUsuario As String 'contiene el nombre del Usuario que se está conectando al sistema

    Public Const C_FECHAINICIAL = #1/1/1900#
    Public Const C_FECHAFINAL = #12/31/2078#

    Public Const C_DESACTIVADO As Integer = 0
    Public Const C_ACTIVADO As Integer = 1
    Public Const C_SINCAMBIO As Integer = 2
    Public NombreMaquina As String
    Public NombreServidor As String
    Public NombreBaseDatos As String
    Public RsGral As New ADODB.Recordset
    Public Cnn As New ADODB.Connection
    Public Cmd As New ADODB.Command
    Public RsgralPV As New ADODB.Recordset

    'Public CnnPVenta As ADODB.Connection
    'Public CmdPVenta As New ADODB.Command
    Public gStrSql As String
    Public gstrNombCortoEmpresa As String
    'Public date As Date

    'Contantes Definidas
    Public Const C_msgFALTADATO As String = "Información incompleta... Proporcione: "
    Public Const C_msgSINDATOS As String = "No hay datos registrados"
    Public Const C_msgGUARDAR As String = "¿Desea GUARDAR los cambios?"
    Public Const C_msgBORRAR As String = "¿Esta seguro de BORRAR esta información?"
    Public Const C_msgACTUALIZADO As String = "Los cambios se han grabado exitosamente !!"
    Public Const C_msgSALIR As String = "¿Desea abandonar la captura?"
    Public Const C_msgCOORDNOVALIDA As String = "Coordenada no Valida"

    'Constantes para el manejo de los stored procedures
    Public Const C_INSERCION = "I"
    Public Const C_MODIFICACION = "M"
    Public Const C_ELIMINACION = "E"
    Public C_FECHADEFAULT As Date

    'Constantes para la edicion de detalles en un grid
    Public Const C_NUEVO As String = "N" 'Nuevo detalle
    Public Const C_MODIFICADO As String = "M" 'el detalle a sido modificado
    Public Const C_ELIMINADO As String = "E" 'El detalle a sido eliminado
    Public Const C_ACTIVO As String = "A" 'El detalle esta activo(No ha havido ningun cambio)
    Public Const C_SUSPENDIDO As String = "S" 'SUSPENDIDO

    'Public Const C_FORMATFECHAGUARDAR As String = "dd/MMM/yyyy"
    'Public Const C_FORMATFECHAMOSTRAR As String = "MM/dd/yyyy"

    Public Const C_FORMATFECHAGUARDAR As String = "yyyy/MM/dd HH:mm:ss"
    Public Const C_FORMATFECHAMOSTRAR As String = "yyyy/MM/dd HH:mm:ss"

    'Constantes para el tipo de Cuenta Bancaria
    Public Const C_NORMAL As String = "N"
    Public Const C_INVERSION As String = "I"

    'Constantes para el Tipo o Nivel de Acceso de los Usuarios
    Public Const C_TADMIN = "A"
    Public Const C_TSUPERVISOR = "S"
    Public Const C_TEMPLEADO = "E"

    'Constantes para el Tipo de Proveedor
    Public Const C_TPROVEEDOR = "P"
    Public Const C_TACREEDOR = "A"

    'Constantes y variables públicas para los Grupos de Arículos
    Public Const C_JOYERIA = "JOYERIA"
    Public Const C_RELOJERIA = "RELOJERIA"
    Public Const C_VARIOS = "VARIOS"

    'A las variables se les da valor en el MenuPrincipal
    Public gCODJOYERIA As Integer = 1
    Public gCODRELOJERIA As Integer = 2
    Public gCODVARIOS As Integer = 3

    'Se usa en el Formulario de Formas de Pago y en Form de Denominaciones
    'Rosaura Torres
    Public gblnMostrarDatosGrid As Boolean

    'Se usa en el Menu Principal y Contiene todas las Formas existentes en el Proyecto
    'Public VFormas(0, 0) As Object 'Vector que almacena todas las formas existentes en el Sistema,
    Public VFormas(2, 158) As Object
    'Public VFormas()


    'Public Const C_DESACTIVADO = 0
    'Public Const C_ACTIVADO = 1
    'Public Const C_SINCAMBIO = 2

    'Variables que son usadas durante todo el proyecto.
    'Contienen cada uno de los Parámetros de Configuracion General,
    'los cuales son necesarios en algunos de los procesos del sistema
    Public gintCodAlmacen As Integer 'Código del Almacen Configurado para el Punto de Venta
    Public gblnCapturarCantidadArts As Boolean 'Capturar Cantidad de Articulos
    Public gcurRedondeo As Decimal 'Redondeo de Montos en ($)
    Public gstrTipodeTrasfencia As String 'Tipo de Transferencia(Electrónica o Diskette)
    Public gbytPosicionesDecimal As Byte 'Cantidad de Posiones de Decimales
    Public gstrSimboloMonedaNacional As String 'Simbolo de Pesos para Totales
    Public gcurEfectivoMaximoCaja As Decimal 'Cantidad de Efectivo Maximo en Caj
    Public gstrRutaImagenLogotipo As String 'Ruta de la Ubicación del Logotipo
    Public gstrRutaArchivoInvElectronico As String 'Ruta de la Ubicacion del Archivo para Inv. Electronico
    Public gstrNomArchivoInvElectronico As String 'Nombre del Archivo para Inventario Electronico
    Public gstrSeparador As String
    Public gbytEspacios As Byte

    Public gblnPermitirVtassinExistencia As Boolean 'Permitir vnetas sin existencias
    Public gblnConsultarxDescripcion As Boolean 'Permitir Consultar por descripcion
    Public gblnAutCambiarCodCapturado As Boolean 'Autorizacion para cambiar código Capturado
    Public gblnAutSuprLineaCapturada As Boolean 'Autorizacion para suprimir linea capturada
    Public gblnAutAbandCapturaIniciada As Boolean 'Autorización para Abandonar Captura Iniciada
    Public gblnIndSiProdNoSoportaDescto As Boolean 'Indicar si el Prod no soporta Descuento
    Public gblnAutConsultaFoliosVta As Boolean 'Autorizacion para consultar Folios de Venta
    Public gblnAutModificarDesctos As Boolean 'autorizacion para modificar descuentos.
    Public gblnDescAumentar As Boolean 'Aumento en el Descuento
    Public gblnDescDisminuir As Boolean 'Disminucion del Descuento
    Public gcurUtilMinporOperacion As Decimal 'Utilidad Mínima por Operacion
    Public gstrMensajeFiscal As String
    Public gstrMensajeNormal As String
    Public gstrMensajeCredito As String
    Public gstrMensajeDevoluciones As String
    Public gstrImpresionTransferencias As String
    Public gblnAplicarIEPS As Boolean 'Aplicar IEPS
    Public gintCodCaja As Integer 'Número de Caja cargado en la configuración
    Public gcurDescuentoGral As Decimal 'Esta Variable Maneja el descuento general aplicable a todos los Articulos, siempre y cuanod dicho articulo no esté en promocion

    Public gintLonCliente As Integer
    Public gintLonDireccion As Integer
    Public gintLonColonia As Integer
    Public gintLonCiudad As Integer
    Public gintLonEstado As Integer
    Public gintLonLeyenda As Integer
    Public gintLonDescProducto As Integer

    'Las siguientes Variables corresponden a la configuracion general del Corporativo, pero también serán usadas en el Punto de Venta.
    'Variables públicas del corporativo
    Public gcurCorpoPORCUTILMINOPERACION As Decimal
    Public gstrCorpoRUTAIMAGENES As String
    Public gcurCorpoTASAIVA As Decimal
    Public gstrCorpoNOMBREEMPRESA As String
    Public gstrCorpoRFCEMPRESA As String
    Public gstrCorpoDOMICILIOEMPRESA As String
    Public gcurCorpoTIPOCAMBIODOLAR As Decimal
    Public gcurCorpoTIPOCAMBIOEURO As Decimal
    Public gintCodAlmacenGral As Integer
    Public gstrCodificacionImportes As String
    Public gintLapsoDifStock As Integer
    Public gstrCorpoTransferEntreSucursales As Boolean
    Public gstrCorpoDriveLocal As String

    '----------------------------------------------------------------------------------------------------------------
    'Constantes para el estatus de las Órdenes de Compra
    Public Const C_STVIGENTE = "V"
    Public Const C_STGENERADA = "G"
    Public Const C_STCANCELADA = "C"
    Public Const C_STREGISTRADA = "R"
    Public Const C_STAPLICADA = "A"

    'Constante para indicar la inicial del módulo de CxP. Se utiliza en el proceso de registro de pagos en bancos para indicar que el pago procede de CxP
    Public Const C_MODULOCXP = "C"

    Public Const C_CONCILIADO = "C"
    Public Const C_RESURTIDO = "R"
    Public Const C_CR = "CR"

    'Constantes para el Tipo de Factura y el Tipo de Gasto(frmCXPRegFactCompras, frmCXPRegFactGastos)
    Public Const C_TIPOFACTURAPROV = "P"
    Public Const C_TIPOFACTURAACRE = "A"
    Public Const C_TIPOGASTOPERSONAL = "P"
    Public Const C_TIPOGASTOJOYERIA = "J"

    'Constantes para el tipo de nota de crédito, que sería "B"-onificación y "D"-evolución
    Public Const C_TIPONOTADEVOLUCION = "D"
    Public Const C_TIPONOTABONIFICACION = "B"

    'Constantes para el tipo de monedas
    Public Const C_DOLAR = "D"
    Public Const C_PESO = "P"
    Public Const C_EURO = "E"

    Public Const C_DESCDOLARES = "DOLARES"
    Public Const C_DESCPESOS = "PESOS"
    Public Const C_DESCEUROS = "EUROS"

    'Constantes para el tipo de programación de pagos
    Public Const C_TIPOPAGONORMAL = ""
    Public Const C_TIPOPAGOAUTOMATICO = "A"
    Public Const C_DESCTIPOPAGONORMAL = "NORM"
    Public Const C_DESCTIPOPAGOAUTOMATICO = "AUTO"

    'Constantes y variables para manejar la programación de pagos y la programación de la frecuencia
    Public Const C_STPAGADO = "P"
    Public blnCambiaSerie As Boolean
    Public blnFrecuencia As Boolean
    Public Const C_DIASPAGO As Integer = 2000000 'Constante para el número de fechas que se generarán para los pagos
    Public aDiasPago(C_DIASPAGO) As Date
    Public fldFrecuencia As String
    Public fldTipoIntervalo As Integer
    Public fldRepeticiones As Integer
    Public fldFechaInicio As Date
    Public fldFechaFin As Date
    Public fldPeriodo As Integer
    Public fldDiaSemana As String
    Public fldDiaMes As Integer
    Public fldMes As Integer
    Public fldOpcion As Integer
    Public fldCual As Integer
    Public fldCuando As Integer

    'Constantes para Movimientos Bancarios
    Public Const C_MOVPAGO As String = "PA"
    Public Const C_MOVDEPOSITO As String = "DE"
    Public Const C_MOVTRASPASO As String = "TB"
    Public Const C_MOVCARGOS As String = "CD"
    Public Const C_MOVANTICIPOS As String = "AP"
    Public Const C_MOVCANCELACION As String = "CA"
    Public Const C_OTROSINGRESOS As String = "OI"

    Public Const C_TIPOMOVINGRESO As String = "I"
    Public Const C_TIPOMOVEGRESO As String = "E"
    Public Const C_TIPOMOVCANCELACION As String = "C"

    Public Const C_NATURALEZACOMERCIAL As String = "C"
    Public Const C_NATURALEZAINTERNA As String = "I"

    Public Const C_FORMAPAGOEFECTIVO As String = "F"
    Public Const C_FORMAPAGOCHEQUE As String = "Q"
    Public Const C_FORMAPAGOELECTRONICO As String = "E"

    Public Const C_TIPOPAGOJOYERIA As String = "J"
    Public Const C_TIPOPAGOPERSONAL As String = "P"

    Public Const C_MODULOBANCOS As String = "B"

    Public Const C_RETIROCAJAGENERAL As String = "C"


    '----------------------------------------------------------------------------------------------------------------

    'Parametros para Hacer la Conexión al Punto de Venta
    Public gstrBasedeDatosPV As String
    Public gstrServidorPV As String

    Public gstrFormatoCantidad As String = "" 'Esta Variable Contiene la Mascara para las cantidades
    Public gblnSalir As Boolean 'Esta Variable nos Sirve para validar si un Formulario es Descargado
    Public gstrMovimiento As String

    Public frmPagos As New frmBancosProcesoDiarioOrigenyAplicacion
    Public frmDepositos As New frmBancosProcesoDiarioOrigenyAplicacion
    Public frmCargos As New frmBancosProcesoDiarioOrigenyAplicacion
    Public frmAnticipos As New frmBancosProcesoDiarioOrigenyAplicacion
    Public frmOtrosIngresos As New frmBancosProcesoDiarioOrigenyAplicacion
    Public frmDepositosIntPes As New frmBancosProcesoDiarioOrigenyAplicacion
    Public frmDepositosIntDol As New frmBancosProcesoDiarioOrigenyAplicacion
    Public frmConsultaOrigenAplicacion As New frmBancosProcesoDiarioOrigenyAplicacion

    Public frmDesgloseDepositos As New frmBancosProcesoDiarioDesglosedeDepositos
    Public frmDesgloseCargosDiversos As New frmBancosProcesoDiarioDesglosedeDepositos
    Public frmDesgloseOtrosIngresos As New frmBancosProcesoDiarioDesglosedeDepositos

    'Iva para los Anticipos
    Public gcurIvaAnticipos As Decimal

    'Constantes para el Manejo de Inventarios en el Punto de Venta
    'Contiene los Códigos de Movimientos de Almacen
    Public Const C_PrefijoFoliosAlmacen As String = "I"
    Public Const C_ENTRADA As String = "E"
    Public Const C_SALIDA As String = "S"
    Public Const C_EntradaPorCompra As Integer = 1
    Public Const C_EntradaPorTransferencia As Integer = 2
    Public Const C_EntradaPorDevolucionSobreVenta As Integer = 3
    Public Const C_EntradaPorDevoluciondeVendedoresExternos As Integer = 4
    Public Const C_EntradaPorDevolucionSobreVentadeVendedoresExternos As Integer = 5
    Public Const C_EntradaPorDevolucionSobrePrestamo As Integer = 6
    Public Const C_EntradaPorDevoluciondeObsequio As Integer = 7
    Public Const C_EntradaPorAjustedeInventario As Integer = 8
    Public Const C_EntradaaAlmacendeVendedorExterno As Integer = 9

    Public Const C_SalidaPorVenta As Integer = 51
    Public Const C_SalidaPorTransferencia As Integer = 52
    Public Const C_SalidaPorDevolucionSobreCompra As Integer = 53
    Public Const C_SalidaAVendedoresExternos As Integer = 54
    Public Const C_SalidaPorVentadeVendedoresExternos As Integer = 55
    Public Const C_SalidaPorPrestamodeArticulos As Integer = 56
    Public Const C_SalidaPorObsequio As Integer = 57
    Public Const C_SalidaPorAjustedeInventario As Integer = 58
    Public Const C_SalidadeAlmacendeVendedorExterno As Integer = 59

    '--Rosaura
    'VAriable se usa en el Form de Notas Credito(Ventas Salida de Mercancia), y en frmAutorizacionConfig.
    'Para Saber si el usuario dado tiene autorizacion para modificar un dato o no.
    Public gblnAutorizacionAceptada As Boolean
    Public gblnSalioSinValidar As Boolean 'Esta variable se usa en el form. de AutorizacionCnfig, para saber si se salió del formulario sin validar los datos proporcionados.
    'Lo cual puede suceder si un usuario presiona el boton de cerrado en la parte superior del form.
    Public Const C_msgSINAUTORIZACION As String = "No posee Autorización para "

    Public gstrRutaImpresora As String '''Contiene la Ruta de la Impresora para Facturas
    Public gstrTicketPrinter As String '''Contiene la Ruta de la Impresora de tickets

    '''Forma auxiliar para abc de bancos en bancos
    'Public FrmAbcBancos As New frmCorpoAbcBancos
    'Public frmAbcCuentasBancarias As New frmCorpoABCCuentasBancarias
    'Public frmAbcOrigenyAplicaciondeRecursos As New frmCorpoABCOrigenyAplicaciondeRecursos
    'Public frmAbcRubrosdeAplicacionyOrigen As New frmCorpoABCRubrosdeAplicacionyOrigen

    Public gstrProcesoqueGeneraError As String
    Public gbytCantidadDecimales As Byte 'Esta Variable se usa para especificar el número  de decimales que tendrá una cantidad, este numero es 2, sólo que para no modificar todo el codigo, uso esta variable que contiene el valor de 2

    '----------------------------------------------------------------------------------------------------------------
    'Constantes que almacenan la NOTA de los reportes de Ventas Ingresos cuando se activa o desactiva
    'la opción de descontar Comisión Bancaria
    Public Const C_DESCUENTOPORCOMISIONES_SI As String = "** Los importes recibidos por pagos que implican una transacción bancaria tienen descontadas las comisiones correspondientes"
    Public Const C_DESCUENTOPORCOMISIONES_NO As String = "** Los importes recibidos por pagos que implican una transacción bancaria NO tienen descontadas las comisiones correspondientes"

    'Bandera para mostrar el código viejo
    Public gblnTransfCodigoViejo As Boolean

    'vARIBLES GLOBALES USADAS EN EL ANALISIS DE LAS VENTAS
    Public gintCodRFC As Integer
    Public gstrNombreCliente As String
    Public gstrRFCCliente As String

    Public gblnPagoVentasconTarjeta As Boolean '''Para saber si se pago con tarjeta de credito

    Public Const C_BDACCESS As String = "IMPETIQ"

    Public gstrNombreForma As String
    Public Const C_REDONDEO As Integer = 0

    Public Const C_GASTOPERSONAL As String = "P"
    Public Const C_GASTOJOYERIA As String = "J"

    'Public frmFactAnalisisVentasImpresionTickets As New frmFactAnalisisVentasImpresionTickets()

    Public Function AgregarHoraAFecha(ByVal FechaOrig As String) As String

        Dim Dia As String
        Dim Ano As String
        Dim Mes As String
        Dim Hrs As String
        Dim Min As String
        Dim Seg As String
        Dim H24 As String
        Dim lafecha As String = FechaOrig.Split("/")(2).Substring(0, 4) & "-" & FechaOrig.Split("/")(1) & "-" & FechaOrig.Split("/")(0)



        Dia = Mid(FechaOrig, 1, 2)
        Mes = Mid(FechaOrig, 4, 2)
        Ano = Mid(FechaOrig, 7, 4)
        Dim HoraActual As DateTime = DateTime.Now



        Hrs = Mid(HoraActual, 12, 2)
        Min = Mid(HoraActual, 15, 2)
        Seg = Mid(HoraActual, 18, 2)
        H24 = Mid(HoraActual, 20, 4)



        If H24.Trim.ToUpper.Substring(0, 1) = "P" Then
            If Hrs <> 12 Then
                Hrs = Val(Hrs) + 12
            Else
                Hrs = 12
            End If
        Else
            If Val(Hrs) = 12 Then
                Hrs = "00"
            Else
                Hrs = Val(Hrs)
            End If
        End If

        Dim lahora As String = Hrs & ":" & Min & ":00.000"
        Dim fechatotal = lafecha & " " & lahora
        Return fechatotal
    End Function

End Module