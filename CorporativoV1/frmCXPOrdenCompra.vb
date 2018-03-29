Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports VB = Microsoft.VisualBasic

Public Class frmCXPOrdenCompra
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtFolioApartado As System.Windows.Forms.TextBox
    Public WithEvents _lblOrden_2 As System.Windows.Forms.Label
    Public WithEvents fraApartado As System.Windows.Forms.Panel
    Public WithEvents btnAsignarCodigos As System.Windows.Forms.Button
    Public WithEvents btnProv As System.Windows.Forms.Button
    Public WithEvents txtDesctoFinanciero As System.Windows.Forms.TextBox
    Public WithEvents txtTipoCambioConciliado As System.Windows.Forms.TextBox
    Public WithEvents txtTipoCambioEuroConciliado As System.Windows.Forms.TextBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents _fraOrden_0 As System.Windows.Forms.GroupBox
    Public WithEvents txtPorcDescto As System.Windows.Forms.TextBox
    Public WithEvents txtTasaIva As System.Windows.Forms.TextBox
    Public WithEvents txtRemision As System.Windows.Forms.TextBox
    Public WithEvents txtPedido As System.Windows.Forms.TextBox
    Public WithEvents _fraOrden_3 As System.Windows.Forms.GroupBox
    Public WithEvents txtTipoCambioEuro As System.Windows.Forms.TextBox
    Public WithEvents txtTipoCambio As System.Windows.Forms.TextBox
    Public WithEvents _optMoneda_2 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblEuro As System.Windows.Forms.Label
    Public WithEvents lblDolar As System.Windows.Forms.Label
    Public WithEvents fraMoneda As System.Windows.Forms.GroupBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblOrden_3 As System.Windows.Forms.Label
    Public WithEvents fraFecha As System.Windows.Forms.Panel
    Public WithEvents _fraOrden_5 As System.Windows.Forms.GroupBox
    Public WithEvents txtTotal As System.Windows.Forms.TextBox
    Public WithEvents txtIVA As System.Windows.Forms.TextBox
    Public WithEvents txtDescuento As System.Windows.Forms.TextBox
    Public WithEvents txtSubTotal As System.Windows.Forms.TextBox
    Public WithEvents rtEntregaren As System.Windows.Forms.RichTextBox
    Public WithEvents fraEntregarEn As System.Windows.Forms.GroupBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents txtOtrosDatos As System.Windows.Forms.Label
    Public WithEvents fraOtrosDatos As System.Windows.Forms.GroupBox
    Public WithEvents txtCostosIndirectos As System.Windows.Forms.TextBox
    Public WithEvents txtCostoAdicional As System.Windows.Forms.TextBox
    Public WithEvents _lblOrden_8 As System.Windows.Forms.Label
    Public WithEvents _lblOrden_7 As System.Windows.Forms.Label
    Public WithEvents fraCostos As System.Windows.Forms.GroupBox
    Public WithEvents txtFolio As System.Windows.Forms.TextBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents dtpFechaEntrega As System.Windows.Forms.DateTimePicker
    Public WithEvents dbcOrigen As System.Windows.Forms.ComboBox
    Public WithEvents dbcGrupo As System.Windows.Forms.ComboBox
    Public WithEvents _lblOrden_6 As System.Windows.Forms.Label
    Public WithEvents _lblOrden_5 As System.Windows.Forms.Label
    Public WithEvents _lblOrden_4 As System.Windows.Forms.Label
    Public WithEvents fraEntrega As System.Windows.Forms.GroupBox
    Public WithEvents mshFlex As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents lblDescProv As System.Windows.Forms.Label
    Public WithEvents lblDesctoFinanciero As System.Windows.Forms.Label
    Public WithEvents lblPorcDescto As System.Windows.Forms.Label
    Public WithEvents lblTasaIva As System.Windows.Forms.Label
    Public WithEvents lblRemision As System.Windows.Forms.Label
    Public WithEvents lblPedido As System.Windows.Forms.Label
    Public WithEvents txtDescripcion As System.Windows.Forms.Label
    Public WithEvents lblCR As System.Windows.Forms.Label
    Public WithEvents _lblOrden_17 As System.Windows.Forms.Label
    Public WithEvents _lblOrden_16 As System.Windows.Forms.Label
    Public WithEvents lblResurtido As System.Windows.Forms.Label
    Public WithEvents _lblOrden_15 As System.Windows.Forms.Label
    Public WithEvents lblConciliado As System.Windows.Forms.Label
    Public WithEvents lblEstatus As System.Windows.Forms.Label
    Public WithEvents _lblOrden_14 As System.Windows.Forms.Label
    Public WithEvents _lblOrden_13 As System.Windows.Forms.Label
    Public WithEvents _lblOrden_12 As System.Windows.Forms.Label
    Public WithEvents _lblOrden_11 As System.Windows.Forms.Label
    Public WithEvents _lblOrden_1 As System.Windows.Forms.Label
    Public WithEvents _lblOrden_0 As System.Windows.Forms.Label
    Public WithEvents fraOrden As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblOrden As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray


    Public FolioAlmacen As String
    Const C_RENENCABEZADO As Integer = 0

    'Cuando no tenga Código Auxiliar, Se considerará una entrada nueva a la tabla OrdenesCompraPreCat
    Const C_COLCODIGO As Integer = 0
    Const C_COLDESCRIPCION As Integer = 1
    Const C_COLUNIDAD As Integer = 2
    Const C_COLCANTIDAD As Integer = 3
    Const C_COLPRECIOUNITARIO As Integer = 4
    Const C_COLCOSTO As Integer = 5
    Const C_COLDESCTO As Integer = 6
    Const C_COLDESCTOPORC As Integer = 45
    Const C_COLDESCTOPORCTAG As Integer = 46

    Const C_COLIVA As Integer = 7
    Const C_COLPORCIVA As Integer = 43

    Const C_COLCODAUX As Integer = 8
    Const C_ColSTATUS As Integer = 9
    Const C_COLSTATUSTAG As Integer = 44

    Const C_COLCOSTOFACTURA As Integer = 41
    Const C_COLPORCFACTURA As Integer = 42

    Const C_ColDESCRIPCIONTAG As Integer = 10
    Const C_COLUNIDADTAG As Integer = 11
    Const C_COLCANTIDADTAG As Integer = 12
    Const C_COLPRECIOUNITARIOTAG As Integer = 13
    Const C_COLCOSTOTAG As Integer = 14
    Const C_COLDESCTOTAG As Integer = 15
    Const C_COLIVATAG As Integer = 16

    'Constantes para los valores de las columnas de la tabla OrdenesCompraPreCat
    Const C_COLCOSTOADICIONAL As Integer = 17
    Const C_COLCOSTOINDIRECTOS As Integer = 18
    Const C_ColCODGRUPO As Integer = 19
    Const C_COLCODFAMILIA As Integer = 20
    Const C_COLCODLINEA As Integer = 21
    Const C_COLCODSUBLINEA As Integer = 22
    Const C_COLCODKILATES As Integer = 52
    Const C_COLCODMARCA As Integer = 23
    Const C_COLCODMODELO As Integer = 24
    Const C_COLCODTIPOMATERIAL As Integer = 25
    Const C_COLGENERO As Integer = 26
    Const C_COLMOVIMIENTO As Integer = 27
    Const C_COLCRONO As Integer = 54
    Const C_COLCODIGOARTICULOPROV As Integer = 28

    Const C_COLCOSTOADICIONALTAG As Integer = 29
    Const C_COLCOSTOINDIRECTOSTAG As Integer = 30
    Const C_COLCODGRUPOTAG As Integer = 31
    Const C_COLCODFAMILIATAG As Integer = 32
    Const C_COLCODLINEATAG As Integer = 33
    Const C_COLCODSUBLINEATAG As Integer = 34
    Const C_COLCODKILATESTAG As Integer = 53
    Const C_COLCODMARCATAG As Integer = 35
    Const C_COLCODMODELOTAG As Integer = 36
    Const C_COLCODTIPOMATERIALTAG As Integer = 37
    Const C_COLGENEROTAG As Integer = 38
    Const C_COLMOVIMIENTOTAG As Integer = 39
    Const C_COLCRONOTAG As Integer = 55
    Const C_COLCODIGOARTICULOPROVTAG As Integer = 40

    '47 en adelante, son para almacenar el valor currency en 4 dígitos decimales,
    'estos valores se cargarán en la tabla con sus correspondientes 4 dígitos decimales
    Const C_COLCOSTOCUR As Integer = 47
    Const C_COLDESCUENTOCUR As Integer = 48
    Const C_COLIVACUR As Integer = 49
    Const C_COLCOSTOADICIONALCUR As Integer = 50
    Const C_COLCOSTOINDIRECTOSCUR As Integer = 51
    Const C_COLPRECIOUNITARIO4DEC As Integer = 56
    Const C_ColIMPORTE As Integer = 57

    Const C_COLADICIONAL As Integer = 58
    Const C_COLPRECIOPUBDOLAR As Integer = 59
    Const C_COLMONEDAPP As Integer = 60
    Const C_COLORIGENANT As Integer = 61
    Const C_ColCODIGOANT As Integer = 62
    Const C_ColIMAGEN As Integer = 63

    Const C_ColCODGRUPOX As Integer = 64
    Const C_COLCODFAMILIAX As Integer = 65
    Const C_COLCODLINEAX As Integer = 66
    Const C_COLCODSUBLINEAX As Integer = 67
    Const C_COLCODKILATESX As Integer = 68
    Const C_COLCODMARCAX As Integer = 69
    Const C_COLCODMODELOX As Integer = 70
    Const C_COLCODTIPOMATERIALX As Integer = 71
    Const C_COLGENEROX As Integer = 72
    Const C_COLMOVIMIENTOX As Integer = 73
    Const C_COLCRONOX As Integer = 74
    Const C_COLADICIONALX As Integer = 75
    Const C_COLSTATUSX As Integer = 76

    Const C_COLADICIONALTAG As Integer = 77
    Const C_COLPRECIOPUBDOLARTAG As Integer = 78
    Const C_COLMONEDAPPTAG As Integer = 79
    Const C_COLORIGENANTTAG As Integer = 80
    Const C_ColCODIGOANTTAG As Integer = 81
    Const C_ColIMAGENTAG As Integer = 82

    Const C_ColMDSPESO As Integer = 83 '''27OCT2010 - MAVF
    Const C_ColMDSCOLOR As Integer = 84 '''27OCT2010 - MAVF
    Const C_ColMDSPUREZA As Integer = 85 '''27OCT2010 - MAVF
    Const C_ColMDSCERTIFICADO As Integer = 86 '''27OCT2010 - MAVF

    'Variables en las que se almacenarán los totales con 4 dígitos decimales
    Dim mcurSubTotal As Decimal
    Dim mcurDESCUENTO As Decimal
    Dim mcurIVA As Decimal
    Dim mcurTotal As Decimal

    'Variable para el estatus de la orden de compra
    Dim cESTATUSORDEN As String

    Dim mblnCambiosEnCodigo As Boolean
    Dim mblnNuevo As Boolean
    Dim mblnConciliar As Boolean

    Dim mblnSalir As Boolean
    Dim rsLocal As ADODB.Recordset
    Dim cMonedadeCompra As String
    Dim cMonedadeCompraTag As String
    Dim UltimaMoneda As String
    Dim mblnLoad As Boolean

    Dim mblnFueraChange As Boolean
    Dim tecla As Integer
    Public mintCodProveedor As Integer ' Sólo se modifica en este formulario, para los demás formularios es de sólo consulta
    Dim mintCodGrupo As Integer
    Public mintCodOrigen As Integer

    Public mintRenglonAnt As Integer
    Public mintRenglonAct As Integer
    Public mintRenglonSig As Integer
    Public mintTotalPartidasCapt As Integer
    Dim ResBusquedaArt As Integer
    Public WithEvents btnCancelar As Button
    Public WithEvents btnNuevo As Button
    Public WithEvents btnGuardar As Button
    Friend WithEvents btnBuscar As Button
    Dim CodAux As Integer

    Public Sub ActualizaCantidades()
        On Error GoTo Err_Renamed
        'Variables para calcular la columna de IVA
        Dim nIva As Decimal
        Dim nIVACal As Decimal
        Dim nIVAIMPORTECal As Decimal

        'Variable para el uso del For
        Dim I As Integer

        Dim nCostoAdicional As Decimal
        Dim nCostoIndirectos As Decimal

        'Variables para almacenar los valores de las columnas del renglón que se estará calculando
        Dim nCantidad As Decimal
        Dim nCOSTO As Decimal
        Dim nCOSTOUNITARIO As Decimal
        Dim nDESCTO As Decimal
        Dim nDesctoPorc As Decimal
        Dim nIMPIVA As Decimal

        Dim nCantidadCur As Decimal
        Dim nCOSTOCur As Decimal
        Dim nCOSTOUNITARIOCur As Decimal
        Dim nDESCTOCur As Decimal
        Dim nDesctoPorcCur As Decimal
        Dim nIMPIVACur As Decimal
        Dim nCostoAdicionalCur As Decimal
        Dim nCostoIndirectosCur As Decimal
        Dim nCostoParaProrrateoCur As Decimal
        Dim nPorcFacturaCur As Decimal
        Dim nCostoFacturaCur As Decimal


        Dim nCostoParaProrrateo As Decimal 'Esta variable la voy a utilizar para almacenar la sumatoria de (Cantidad*CostoUnitario)
        'de los productos conciliados y sacar el porcentaje que le corresponde a cada producto sobre el costo adicional
        'y el costo indirecto

        'Variables para calcular los valores de los totales
        Dim nSubTotal As Decimal 'Sumatoria de (Costo Unitario * Cantidad)
        Dim nDescuento As Decimal 'Sumatoria de ((CostoUnitario * Cantidad) - Descuento)
        Dim nTotalIVA As Decimal 'Sumatoria de la columna de IVA
        Dim nTotal As Decimal 'El (SubTotal - Descuento) + TotalIva

        mintTotalPartidasCapt = 0
        nIva = CDec(VB6.Format(CDec(Numerico((Me.txtTasaIva.Text))) / 100, "##0.00"))
        nIVACal = 0
        nIVAIMPORTECal = 0

        nCostoAdicional = CDec(VB6.Format(CDec(Numerico((Me.txtCostoAdicional.Text))), "###,###,##0.00"))
        nCostoIndirectos = CDec(VB6.Format(CDec(Numerico((Me.txtCostosIndirectos.Text))), "###,###,##0.00"))

        With mshFlex
            'Primero ve realiza las operaciones correspondientes con los Conciliados y Conciliados/Resurtidos
            mcurSubTotal = 0
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then Exit For

                'Si el estatus es "C" - Conciliado, o "CR" - Conciliado y Resurtido
                'Calcula el valor del Costo Factura por artículo (de los que se han conciliado)
                If Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CONCILIADO Or Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CR Then
                    nCantidadCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))), "###,###,##0.0000"))
                    nCOSTOUNITARIOCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))), "###,###,##0.0000"))
                    nCOSTOCur = CDec(VB6.Format(nCantidadCur * nCOSTOUNITARIOCur, "###,###,##0.0000"))
                    nDESCTOCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))
                    nDesctoPorcCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTOPORC))), "###,###,##0.0000"))

                    nCantidad = CDec(Numerico(.get_TextMatrix(I, C_COLCANTIDAD)))
                    nCOSTOUNITARIO = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))), "###,###,##0.0000"))
                    nCOSTO = CDec(VB6.Format(nCantidad * nCOSTOUNITARIO, "###,###,##0.0000"))
                    nDESCTO = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))
                    nDesctoPorc = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTOPORC))), "###,###,##0.0000"))
                    If nDesctoPorc = 0 Then
                        If nDESCTO = 0 Then
                            nDESCTOCur = 0
                            nDESCTO = 0
                        Else
                            nDESCTOCur = CDec(VB6.Format(nDESCTOCur, "###,###,##0.0000"))
                            nDESCTO = nDESCTO
                        End If
                    Else
                        'Calcular el importe de descuento en base al porcentaje dado
                        nDESCTOCur = CDec(VB6.Format((nCOSTOUNITARIOCur * nDesctoPorcCur) / 100, "###,###,##0.0000"))
                        nDESCTO = CDec(VB6.Format((nCOSTOUNITARIO * nDesctoPorc) / 100, "###,###,##0.0000"))
                    End If
                    'Calcular el Importe de IVA por artículo
                    nIMPIVACur = CDec(VB6.Format((nCOSTOUNITARIOCur - nDESCTOCur) * nIva, "###,###,##0.0000"))
                    nIMPIVA = CDec(VB6.Format((nCOSTOUNITARIO - nDESCTO) * nIva, "###,###,##0.0000"))

                    mcurSubTotal = mcurSubTotal + nCOSTOCur

                    nCostoParaProrrateoCur = CDec(VB6.Format(nCostoParaProrrateoCur + nCOSTOCur, "###,###,##0.0000"))
                    nCostoParaProrrateo = nCostoParaProrrateo + nCOSTO

                End If
            Next I
            'Calcular el porcentaje del costofactura de cada una de las filas y prorratear (distribuir) uniformemente
            'el Costo Adicional y los Costos Indirectos (debe ser por artículo)
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then
                    Exit For
                End If
                'Si el estatus es "C" - Conciliado, o "CR" - Conciliado y Resurtido
                If Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CONCILIADO Or Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CR Then
                    nCantidadCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))), "###,###,##0.0000"))
                    nCOSTOUNITARIOCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))), "###,###,##0.0000"))
                    nCOSTOCur = CDec(VB6.Format(nCantidadCur * nCOSTOUNITARIOCur, "###,###,##0.0000"))
                    nDESCTOCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))
                    nDesctoPorcCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTOPORC))), "###,###,##0.0000"))

                    nCantidad = CDec(Numerico(.get_TextMatrix(I, C_COLCANTIDAD)))
                    nCOSTOUNITARIO = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))), "###,###,##0.0000"))
                    nCOSTO = CDec(VB6.Format(nCantidad * nCOSTOUNITARIO, "###,###,##0.0000"))
                    nDESCTO = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))
                    nDesctoPorc = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTOPORC))), "###,###,##0.0000"))
                    If nDesctoPorc = 0 Then
                        If nDESCTO = 0 Then
                            nDESCTOCur = 0
                            nDESCTO = 0
                        Else
                            nDESCTOCur = CDec(VB6.Format(nDESCTOCur, "###,###,##0.0000"))
                            nDESCTO = nDESCTO
                        End If
                    Else
                        'Calcular el importe de descuento en base al porcentaje dado
                        nDESCTOCur = CDec(VB6.Format((nCOSTOUNITARIOCur * nDesctoPorcCur) / 100, "###,###,##0.0000"))
                        nDESCTO = CDec(VB6.Format((nCOSTOUNITARIO * nDesctoPorc) / 100, "###,###,##0.0000"))
                    End If
                    'Calcular el Importe de IVA por artículo
                    nIMPIVACur = CDec(VB6.Format((nCOSTOUNITARIOCur - nDESCTOCur) * nIva, "###,###,##0.0000"))
                    nIMPIVA = CDec(VB6.Format((nCOSTOUNITARIO - nDESCTO) * nIva, "###,###,##0.0000"))

                    'Calcular el porcentaje que ocupa la partida en la factura y guardar el resultado en C_COLPORCFACTURA
                    'Si nCostoParaProrrateo - 100%, nCOSTO - ?
                    nPorcFacturaCur = IIf(nCostoParaProrrateoCur <> 0, VB6.Format((nCOSTOCur * 100) / nCostoParaProrrateoCur, "###,###,##0.0000"), VB6.Format(0, "###,###,##0.0000"))
                    nCostoFacturaCur = CDec(VB6.Format((nPorcFacturaCur * nCostoParaProrrateoCur) / 100, "###,###,##0.0000"))
                    nCostoAdicionalCur = CDec(VB6.Format(IIf(mcurSubTotal <> 0, (nCOSTOUNITARIOCur / mcurSubTotal), 0) * nCostoAdicional, "###,###,##0.0000"))
                    nCostoIndirectosCur = CDec(VB6.Format(IIf(mcurSubTotal <> 0, (nCOSTOUNITARIOCur / mcurSubTotal), 0) * nCostoIndirectos, "###,###,##0.0000"))

                    .set_TextMatrix(I, C_COLCOSTOCUR, nCOSTOCur)
                    .set_TextMatrix(I, C_COLDESCUENTOCUR, nDESCTOCur)
                    .set_TextMatrix(I, C_COLIVACUR, nIMPIVACur)
                    .set_TextMatrix(I, C_COLCOSTOADICIONALCUR, nCostoAdicionalCur)
                    .set_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR, nCostoIndirectosCur)

                    .set_TextMatrix(I, C_COLPORCFACTURA, VB6.Format(IIf(nCostoParaProrrateo <> 0, (nCOSTO * 100) / nCostoParaProrrateo, 0), "###,###,##0.0000"))
                    .set_TextMatrix(I, C_COLCOSTOFACTURA, VB6.Format((CDec(Numerico(.get_TextMatrix(I, C_COLPORCFACTURA))) * nCostoParaProrrateo) / 100, "###,###,##0.0000"))
                    .set_TextMatrix(I, C_COLCOSTOADICIONAL, nCostoAdicionalCur)
                    .set_TextMatrix(I, C_COLCOSTOINDIRECTOS, nCostoIndirectosCur)

                    .set_TextMatrix(I, C_COLCOSTOADICIONALTAG, .get_TextMatrix(I, C_COLCOSTOADICIONAL))
                    .set_TextMatrix(I, C_COLCOSTOINDIRECTOSTAG, .get_TextMatrix(I, C_COLCOSTOINDIRECTOS))

                    'Pasar los valores de la variables a las columnas
                    .set_TextMatrix(I, C_COLCANTIDAD, nCantidad)
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO, nCOSTOUNITARIO)
                    .set_TextMatrix(I, C_COLDESCTO, nDESCTO)
                    .set_TextMatrix(I, C_COLIVA, nIMPIVA)

                End If
            Next I
        End With

        With mshFlex
            'A continuación realiza las operaciones con los que tienen estatus vacío, y los que son resurtidos
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then Exit For

                If Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_RESURTIDO Or Trim(.get_TextMatrix(I, C_ColSTATUS)) = "" Then
                    nCantidadCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))), "###,###,##0.0000"))
                    nCOSTOUNITARIOCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))), "###,###,##0.0000"))
                    nCOSTOCur = CDec(VB6.Format(nCantidadCur * nCOSTOUNITARIOCur, "###,###,##0.0000"))
                    nDESCTOCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))
                    nDesctoPorcCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTOPORC))), "###,###,##0.0000"))

                    nCantidad = CDec(Numerico(.get_TextMatrix(I, C_COLCANTIDAD)))
                    nCOSTOUNITARIO = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))), "###,###,##0.0000"))
                    nCOSTO = CDec(VB6.Format(nCantidad * nCOSTOUNITARIO, "###,###,##0.0000"))
                    nDESCTO = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))
                    nDesctoPorc = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTOPORC))), "###,###,##0.0000"))
                    If nDesctoPorc = 0 Then
                        If nDESCTO = 0 Then
                            nDESCTOCur = 0
                            nDESCTO = 0
                        Else
                            nDESCTOCur = CDec(VB6.Format(nDESCTOCur, "###,###,##0.0000"))
                            nDESCTO = nDESCTO
                        End If
                    Else
                        'Calcular el importe de descuento en base al porcentaje dado
                        nDESCTOCur = CDec(VB6.Format((nCOSTOUNITARIOCur * nDesctoPorcCur) / 100, "###,###,##0.0000"))
                        nDESCTO = CDec(VB6.Format((nCOSTOUNITARIO * nDesctoPorc) / 100, "###,###,##0.0000"))
                    End If
                    'Calcular el Importe de IVA por artículo
                    nIMPIVACur = CDec(VB6.Format((nCOSTOUNITARIOCur - nDESCTOCur) * nIva, "###,###,##0.0000"))
                    nIMPIVA = CDec(VB6.Format((nCOSTOUNITARIO - nDESCTO) * nIva, "###,###,##0.0000"))

                    'Poner cero en las columnas C_COLPORCFACTURA-C_COLCOSTOFACTURA-C_COLCOSTOADICIONAL-C_COLCOSTOINDIRECTOS
                    .set_TextMatrix(I, C_COLPORCFACTURA, 0)
                    .set_TextMatrix(I, C_COLCOSTOFACTURA, 0)
                    .set_TextMatrix(I, C_COLCOSTOADICIONAL, 0)
                    .set_TextMatrix(I, C_COLCOSTOINDIRECTOS, 0)

                    .set_TextMatrix(I, C_COLCOSTOADICIONALTAG, .get_TextMatrix(I, C_COLCOSTOADICIONAL))
                    .set_TextMatrix(I, C_COLCOSTOINDIRECTOSTAG, .get_TextMatrix(I, C_COLCOSTOINDIRECTOS))

                    'CostoCur
                    .set_TextMatrix(I, C_COLCOSTOCUR, nCOSTOCur)
                    .set_TextMatrix(I, C_COLDESCUENTOCUR, nDESCTOCur)
                    .set_TextMatrix(I, C_COLIVACUR, nIMPIVACur)
                    .set_TextMatrix(I, C_COLCOSTOADICIONALCUR, 0)
                    .set_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR, 0)

                    'Pasar los valores de la variables a las columnas
                    .set_TextMatrix(I, C_COLCANTIDAD, nCantidad)
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO, nCOSTOUNITARIO)
                    .set_TextMatrix(I, C_COLDESCTO, nDESCTO)
                    .set_TextMatrix(I, C_COLIVA, nIMPIVA)

                End If
            Next I
        End With

        'Antes de hacer las sumatorias (Totales), debemos asegurarnos de que, en caso estar generada, la orden sólo puede
        'mostrar los totales de los artículos conciliados
        'Pero si está cancelada o vigente, debe tomar en cuenta todos y cada uno de los registros del Grid para
        'calcular los Totales
        mcurSubTotal = 0
        mcurDESCUENTO = 0
        mcurIVA = 0
        mcurTotal = 0
        If Trim(cESTATUSORDEN) = C_STGENERADA Then
            With mshFlex
                For I = 1 To .Rows - 1
                    If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then
                        Exit For
                    End If
                    If Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CONCILIADO Or Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CR Then
                        nCantidadCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))), "###,###,##0.0000"))
                        nCOSTOUNITARIOCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))), "###,###,##0.0000"))
                        nCOSTOCur = CDec(VB6.Format(nCantidadCur * nCOSTOUNITARIOCur, "###,###,##0.0000"))
                        nDESCTOCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))
                        nDesctoPorcCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTOPORC))), "###,###,##0.0000"))

                        nCantidad = CDec(Numerico(.get_TextMatrix(I, C_COLCANTIDAD)))
                        nCOSTOUNITARIO = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))), "###,###,##0.0000"))
                        nCOSTO = CDec(VB6.Format(nCantidad * nCOSTOUNITARIO, "###,###,##0.0000"))
                        nDESCTO = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))
                        nDesctoPorc = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTOPORC))), "###,###,##0.0000"))
                        If nDesctoPorc = 0 Then
                            If nDESCTO = 0 Then
                                nDESCTOCur = 0
                                nDESCTO = 0
                            Else
                                nDESCTOCur = CDec(VB6.Format(nDESCTOCur, "###,###,##0.0000"))
                                nDESCTO = nDESCTO
                            End If
                        Else
                            'Calcular el importe de descuento en base al porcentaje dado
                            nDESCTOCur = CDec(VB6.Format((nCOSTOUNITARIOCur * nDesctoPorcCur) / 100, "###,###,##0.0000"))
                            nDESCTO = CDec(VB6.Format((nCOSTOUNITARIO * nDesctoPorc) / 100, "###,###,##0.0000"))
                        End If
                        'Calcular el Importe de IVA por artículo
                        nIMPIVACur = CDec(VB6.Format((nCOSTOUNITARIOCur - nDESCTOCur) * nIva, "###,###,##0.0000"))
                        nIMPIVA = CDec(VB6.Format((nCOSTOUNITARIO - nDESCTO) * nIva, "###,###,##0.0000"))

                        'Calcular totales de los artículos conciliados
                        nSubTotal = nSubTotal + nCOSTO
                        nDescuento = nDescuento + (nDESCTO * nCantidad)
                        nIVACal = nIVACal + (nIMPIVA * nCantidad)

                        mcurSubTotal = CDec(VB6.Format(mcurSubTotal + nCOSTOCur, "###,###,##0.0000"))
                        mcurDESCUENTO = CDec(VB6.Format(mcurDESCUENTO + (nDESCTOCur * nCantidadCur), "###,###,##0.0000"))
                        mcurIVA = CDec(VB6.Format(mcurIVA + (nIMPIVACur * nCantidadCur), "###,###,##0.0000"))
                    End If
                Next I
            End With
        ElseIf Trim(cESTATUSORDEN) <> C_STGENERADA Then
            With mshFlex
                For I = 1 To .Rows - 1
                    If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then
                        Exit For
                    End If
                    nCantidadCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))), "###,###,##0.0000"))
                    nCOSTOUNITARIOCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))), "###,###,##0.0000"))
                    nCOSTOCur = CDec(VB6.Format(nCantidadCur * nCOSTOUNITARIOCur, "###,###,##0.0000"))
                    nDESCTOCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))
                    nDesctoPorcCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTOPORC))), "###,###,##0.0000"))

                    nCantidad = CDec(Numerico(.get_TextMatrix(I, C_COLCANTIDAD)))
                    nCOSTOUNITARIO = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))), "###,###,##0.0000"))
                    nCOSTO = CDec(VB6.Format(nCantidad * nCOSTOUNITARIO, "###,###,##0.0000"))
                    nDESCTO = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))
                    nDesctoPorc = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTOPORC))), "###,###,##0.0000"))
                    If nDesctoPorc = 0 Then
                        If nDESCTO = 0 Then
                            nDESCTOCur = 0
                            nDESCTO = 0
                        Else
                            nDESCTOCur = CDec(VB6.Format(nDESCTOCur, "###,###,##0.0000"))
                            nDESCTO = nDESCTO
                        End If
                    Else
                        'Calcular el importe de descuento en base al porcentaje dado
                        nDESCTOCur = CDec(VB6.Format((nCOSTOUNITARIOCur * nDesctoPorcCur) / 100, "###,###,##0.0000"))
                        nDESCTO = CDec(VB6.Format((nCOSTOUNITARIO * nDesctoPorc) / 100, "###,###,##0.0000"))
                    End If
                    'Calcular el Importe de IVA por artículo
                    nIMPIVACur = CDec(VB6.Format((nCOSTOUNITARIOCur - nDESCTOCur) * nIva, "###,###,##0.0000"))
                    nIMPIVA = CDec(VB6.Format((nCOSTOUNITARIO - nDESCTO) * nIva, "###,###,##0.00"))

                    'Calcular totales de los artículos conciliados
                    nSubTotal = nSubTotal + nCOSTO
                    nDescuento = nDescuento + (nDESCTO * nCantidad)
                    nIVACal = nIVACal + (nIMPIVA * nCantidad)

                    mcurSubTotal = CDec(VB6.Format(mcurSubTotal + nCOSTOCur, "###,###,##0.0000"))
                    mcurDESCUENTO = CDec(VB6.Format(mcurDESCUENTO + (nDESCTOCur * nCantidadCur), "###,###,##0.0000"))
                    mcurIVA = CDec(VB6.Format(mcurIVA + (nIMPIVACur * nCantidadCur), "###,###,##0.0000"))
                Next I
            End With
        End If

        mcurTotal = CDec(VB6.Format((mcurSubTotal - mcurDESCUENTO) + mcurIVA, "###,###,##0.0000"))

        Me.txtSubTotal.Text = VB6.Format(nSubTotal, "###,###,##0.00")
        Me.txtDescuento.Text = VB6.Format(nDescuento, "###,###,##0.00")
        Me.txtIVA.Text = VB6.Format(nIVACal, "###,###,##0.00")
        nTotal = CDec(VB6.Format((CDbl(Numerico(txtSubTotal.Text)) - CDbl(Numerico(txtDescuento.Text))) + CDbl(Numerico(txtIVA.Text)), "###,##0.00"))
        Me.txtTotal.Text = VB6.Format(nTotal, "###,###,##0.00")

        'Formatear las cantidades en el grid
        With mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then
                    Exit For
                End If
                mintTotalPartidasCapt = mintTotalPartidasCapt + 1
                .set_TextMatrix(I, C_COLCANTIDAD, VB6.Format(Numerico(.get_TextMatrix(I, C_COLCANTIDAD)), "###,###,##0"))
                .set_TextMatrix(I, C_COLPRECIOUNITARIO, VB6.Format(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO)), "###,###,##0.00"))
                .set_TextMatrix(I, C_COLDESCTO, VB6.Format(Numerico(.get_TextMatrix(I, C_COLDESCTO)), "###,###,##0.00"))
                .set_TextMatrix(I, C_COLIVA, VB6.Format(Numerico(.get_TextMatrix(I, C_COLIVA)), "###,###,##0.00"))
                .set_TextMatrix(I, C_ColIMPORTE, VB6.Format(((CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDbl(Numerico(.get_TextMatrix(I, C_COLDESCTO)))) + CDbl(Numerico(.get_TextMatrix(I, C_COLIVA)))) * CDbl(Numerico(.get_TextMatrix(I, C_COLCANTIDAD))), "###,##0.00"))

                .set_TextMatrix(I, C_COLCANTIDADTAG, VB6.Format(Numerico(.get_TextMatrix(I, C_COLCANTIDADTAG)), "###,###,##0"))
                .set_TextMatrix(I, C_COLPRECIOUNITARIOTAG, VB6.Format(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIOTAG)), "###,###,##0.00"))
                .set_TextMatrix(I, C_COLDESCTOTAG, VB6.Format(Numerico(.get_TextMatrix(I, C_COLDESCTOTAG)), "###,###,##0.00"))
                .set_TextMatrix(I, C_COLIVATAG, VB6.Format(Numerico(.get_TextMatrix(I, C_COLIVATAG)), "###,###,##0.00"))
            Next I
        End With
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub Conciliar()
        On Error GoTo Merr
        Dim I As Integer
        Dim blnTransaction As Boolean
        Dim nCodArticulo As Integer
        Dim nCostoFactura As Decimal
        Dim nCostoFacturaIva As Decimal
        Dim nCostoReal As Decimal
        Dim nCostoAdicional As Decimal
        Dim nCostoIndirectos As Decimal
        Dim nTipoCambioEuro As Decimal
        Dim nDescuento As Decimal
        Dim nConciliados As Integer

        Dim nCostoFacturaCur As Decimal
        Dim nCostoFacturaCurIva As Decimal
        Dim nCostoRealCur As Decimal
        Dim nCostoAdicionalCur As Decimal
        Dim nCostoIndirectosCur As Decimal
        Dim nTipoCambioEuroCur As Decimal

        Dim nCostoFacturaPesosCur As Decimal
        Dim nCostoAdicionalPesosCur As Decimal
        Dim nCostoIndirectosPesosCur As Decimal

        Dim CodMovtoAlm As Integer
        Dim NumPartida As Integer
        Dim PrecioVenta As Decimal
        Dim Estatus As String
        Dim TipoAlmacen As String
        Dim UltimoCostoMN As String
        Dim mblnEncabezadoAlmacen As Boolean
        Dim lTCConcil As Decimal

        Dim nPrecioPP As Decimal
        Dim nMonedaPP As String
        Dim nOrigenAnt As Integer
        Dim nCodigoAnt As Integer
        Dim lRuta As String
        Dim lNomArch As String
        Dim lExtension As String
        Dim lImagen As String
        Dim lNvaRuta As String

        NumPartida = 0
        CodMovtoAlm = C_EntradaPorCompra
        If txtFolio.Text = "" Then
            MsgBox("Seleccione la Orden de compra que desea Conciliar", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Me.txtFolio.Focus()
            ModEstandar.SelTextoTxt((Me.txtFolio))
            Exit Sub
        Else
            'Busca la Orden de compra y, si la encuentra, pregunta cuál es su estatus
            gStrSql = "SELECT FolioOrdenCompra FROM OrdenesCompra WHERE FolioOrdenCompra = '" & Trim(txtFolio.Text) & "'"

            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            rsLocal = Cmd.Execute
            If rsLocal.RecordCount <= 0 Then
                MsgBox("La Orden de Compra que desea Conciliar no existe, verifique por favor", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.txtFolio.Focus()
                ModEstandar.SelTextoTxt((Me.txtFolio))
                Exit Sub
            End If
        End If

        'Verifica que haya uno ó más registros marcados para conciliar, si no los hay, muestra un mensaje
        nConciliados = 0
        With mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then
                    Exit For
                End If
                If Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CONCILIADO Or Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CR Then
                    nConciliados = nConciliados + 1
                End If
            Next I
        End With
        If nConciliados = 0 Then
            MsgBox("No existe ningún registro a conciliar, verifique los datos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If

        'Si el valor de los tipos de Cambio es igual a Cero, no procede, tiene que salir del procedimiento
        If CDec(Numerico((txtTipoCambioConciliado.Text))) = 0 Then
            MsgBox("Indique el equivalente en pesos, del Dólar", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If CDec(Numerico((txtTipoCambioEuroConciliado.Text))) = 0 Then
            MsgBox("Indique el equivalente en pesos, del Euro", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If

        Cnn.BeginTrans()

        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        blnTransaction = True

        'Poner la orden de compra como GENERADA (Conciliada)
        ModStoredProcedures.PR_IMEOrdenesCompra(Trim(Me.txtFolio.Text), CStr(mintCodProveedor), VB6.Format(Me.dtpFecha.Value, C_FORMATFECHAGUARDAR), VB6.Format(Me.dtpFechaEntrega.Value, C_FORMATFECHAGUARDAR), Trim(Me.txtRemision.Text), Trim(Me.txtPedido.Text), CStr(mintCodOrigen), CStr(mintCodGrupo), CStr(ModEstandar.Numerico((Me.txtCostoAdicional.Text))), CStr(ModEstandar.Numerico((Me.txtCostosIndirectos.Text))), Trim(Me.rtEntregaren.Text), CStr(mcurSubTotal), CStr(mcurDESCUENTO), CStr(mcurIVA), CStr(mcurTotal), Trim(cMonedadeCompra), C_STGENERADA, CStr(CDate(#1/1/1900#)), Trim(Me.txtTasaIva.Text), Trim(Me.txtPorcDescto.Text), Trim(Me.txtTipoCambio.Text), Trim(Me.txtTipoCambioEuro.Text), Trim(Me.txtTipoCambioConciliado.Text), Trim(Me.txtTipoCambioEuroConciliado.Text), Trim(Me.txtDesctoFinanciero.Text), VB6.Format(Today, C_FORMATFECHAGUARDAR), "", C_MODIFICACION, CStr(0))
        Cmd.Execute()

        'Guardar los datos del Grid en la tabla OrdenesCompraPreCat
        With mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then Exit For

                'Es una modificación del registro
                'Convertir los costos a dólares
                '''Se agregó al costofactura el iva, por el manejo de la empresa
                '''pero el costo contable permanece sin iva
                Select Case cMonedadeCompra
                    Case C_DOLAR
                        nCostoAdicionalCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOADICIONALCUR))), "###,###,##0.0000"))
                        nCostoIndirectosCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR))), "###,###,##0.0000"))
                        nCostoFacturaCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCUENTOCUR))), "###,###,##0.0000"))
                        nCostoFacturaCurIva = CDec(VB6.Format((CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCUENTOCUR)))) + CDec(Numerico(.get_TextMatrix(I, C_COLIVACUR))), "###,###,##0.0000"))
                        '''nCostoRealCur = Format(nCostoFacturaCur + nCostoAdicionalCur + nCostoIndirectosCur, "###,###,##0.0000")
                        nCostoRealCur = CDec(VB6.Format(nCostoFacturaCurIva + nCostoAdicionalCur + nCostoIndirectosCur, "###,###,##0.0000"))
                        nDescuento = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.0000"))

                        nCostoAdicional = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOADICIONAL))), "###,###,##0.00"))
                        nCostoIndirectos = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOINDIRECTOS))), "###,###,##0.00"))
                        nCostoFactura = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))), "###,###,##0.00"))
                        nCostoFacturaIva = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))) + CDec(Numerico(.get_TextMatrix(I, C_COLIVA))), "###,###,##0.00"))
                        '''nCostoReal = Format(nCostoFactura + nCostoAdicional + nCostoIndirectos, "###,###,##0.00")
                        nCostoReal = CDec(VB6.Format(nCostoFacturaIva + nCostoAdicional + nCostoIndirectos, "###,###,##0.00"))
                    Case C_PESO
                        nCostoAdicionalCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOADICIONALCUR))) / CDec(Numerico((Me.txtTipoCambioConciliado.Text))), "###,###,##0.0000"))
                        nCostoIndirectosCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR))) / CDec(Numerico((Me.txtTipoCambioConciliado.Text))), "###,###,##0.0000"))
                        nCostoFacturaCur = CDec(VB6.Format((CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCUENTOCUR)))) / CDec(Numerico((Me.txtTipoCambioConciliado.Text))), "###,###,##0.0000"))
                        nCostoFacturaCurIva = CDec(VB6.Format((CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCUENTOCUR))) + CDec(Numerico(.get_TextMatrix(I, C_COLIVACUR)))) / CDec(Numerico((Me.txtTipoCambioConciliado.Text))), "###,###,##0.0000"))
                        '''nCostoRealCur = Format(nCostoFacturaCur + nCostoAdicionalCur + nCostoIndirectosCur, "###,###,##0.0000")
                        nCostoRealCur = CDec(VB6.Format(nCostoFacturaCurIva + nCostoAdicionalCur + nCostoIndirectosCur, "###,###,##0.0000"))
                        nDescuento = CDec(VB6.Format(CDec(CDbl(Numerico(.get_TextMatrix(I, C_COLDESCTO))) / CDec(Numerico((Me.txtTipoCambioConciliado.Text)))), "###,###,##0.0000"))

                        nCostoAdicional = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOADICIONAL))) / CDec(Numerico((Me.txtTipoCambioConciliado.Text))), "###,###,##0.00"))
                        nCostoIndirectos = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOINDIRECTOS))) / CDec(Numerico((Me.txtTipoCambioConciliado.Text))), "###,###,##0.00"))
                        nCostoFactura = CDec(VB6.Format((CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO)))) / CDec(Numerico((Me.txtTipoCambioConciliado.Text))), "###,###,##0.00"))
                        nCostoFacturaIva = CDec(VB6.Format((CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))) + CDec(Numerico(.get_TextMatrix(I, C_COLIVA)))) / CDec(Numerico((Me.txtTipoCambioConciliado.Text))), "###,###,##0.00"))
                        '''nCostoReal = Format(nCostoFactura + nCostoAdicional + nCostoIndirectos, "###,###,##0.00")
                        nCostoReal = CDec(VB6.Format(nCostoFacturaIva + nCostoAdicional + nCostoIndirectos, "###,###,##0.00"))
                    Case C_EURO
                        nTipoCambioEuro = CDec(VB6.Format(CDec(Numerico((Me.txtTipoCambioEuroConciliado.Text))) / CDec(Numerico((Me.txtTipoCambioConciliado.Text))), "###,###,##0.0000")) 'Su equivalencia en dólares

                        nCostoAdicionalCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOADICIONALCUR))) * nTipoCambioEuro, "###,###,##0.0000"))
                        nCostoIndirectosCur = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR))) * nTipoCambioEuro, "###,###,##0.0000"))
                        nCostoFacturaCur = CDec(VB6.Format((CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCUENTOCUR)))) * nTipoCambioEuro, "###,###,##0.0000"))
                        nCostoFacturaCurIva = CDec(VB6.Format((CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCUENTOCUR))) + CDec(Numerico(.get_TextMatrix(I, C_COLIVACUR)))) * nTipoCambioEuro, "###,###,##0.0000"))
                        '''nCostoRealCur = Format(nCostoFacturaCur + nCostoAdicionalCur + nCostoIndirectosCur, "###,###,##0.0000")
                        nCostoRealCur = CDec(VB6.Format(nCostoFacturaCurIva + nCostoAdicionalCur + nCostoIndirectosCur, "###,###,##0.0000"))
                        nDescuento = CDec(VB6.Format(CDec(CDbl(Numerico(.get_TextMatrix(I, C_COLDESCTO))) * nTipoCambioEuro), "###,###,##0.0000"))

                        nCostoAdicional = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOADICIONAL))) * nTipoCambioEuro, "###,###,##0.00"))
                        nCostoIndirectos = CDec(VB6.Format(CDec(Numerico(.get_TextMatrix(I, C_COLCOSTOINDIRECTOS))) * nTipoCambioEuro, "###,###,##0.00"))
                        nCostoFactura = CDec(VB6.Format((CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO))) + CDec(Numerico(.get_TextMatrix(I, C_COLIVA)))) * nTipoCambioEuro, "###,###,##0.00"))
                        nCostoFacturaIva = CDec(VB6.Format((CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO))) - CDec(Numerico(.get_TextMatrix(I, C_COLDESCTO)))) * nTipoCambioEuro, "###,###,##0.00"))
                        '''nCostoReal = Format(nCostoFactura + nCostoAdicional + nCostoIndirectos, "###,###,##0.00")
                        nCostoReal = CDec(VB6.Format(nCostoFacturaIva + nCostoAdicional + nCostoIndirectos, "###,###,##0.00"))
                End Select
                '''            'Calcular el Costo Factura, CostoAdicional y Costos Indirectos en Pesos.
                '''            El TC a considerar para el paso a inventarios es el de concilicación, porque este puede ser modificado
                '''            según lo pactado con el cliente.  Por default el TC Concil es = al TC de la OC.
                '''            nCostoAdicionalPesosCur = nCostoAdicionalCur * gcurCorpoTIPOCAMBIODOLAR
                '''            nCostoFacturaPesosCur = nCostoFacturaCur * gcurCorpoTIPOCAMBIODOLAR
                '''            nCostoIndirectosPesosCur = nCostoIndirectosCur * gcurCorpoTIPOCAMBIODOLAR

                'Calcular el Costo Factura, CostoAdicional y Costos Indirectos en Pesos.
                lTCConcil = CDec(Numerico((txtTipoCambioConciliado.Text)))
                nCostoAdicionalPesosCur = nCostoAdicionalCur * lTCConcil
                nCostoFacturaPesosCur = nCostoFacturaCur * lTCConcil
                nCostoIndirectosPesosCur = nCostoIndirectosCur * lTCConcil

                nPrecioPP = CDec(Numerico(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR)))
                nMonedaPP = IIf(Trim(.get_TextMatrix(I, C_COLMONEDAPP)) = "D", "False", "True")
                nOrigenAnt = CInt(Numerico(.get_TextMatrix(I, C_COLORIGENANT)))
                nCodigoAnt = CInt(Numerico(.get_TextMatrix(I, C_ColCODIGOANT)))

                If Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CONCILIADO Then
                    'Cuando el producto se ha conciliado

                    'Si es un artículo que no tiene código debe darse de alta en CatArticulos
                    If CInt(Numerico(.get_TextMatrix(I, C_COLCODIGO))) = 0 Then
                        ''' 27OCT2010 - MAVF --> SE AGREGARON 4 CAMPOS PARA MANEJO DE DIAMANTE SUELTO
                        '''Se cambio CostoFacturaCur --> CostoFacturaCurIva
                        ModStoredProcedures.PR_IMECatArticulos(CStr(0), Trim(.get_TextMatrix(I, C_COLDESCRIPCION)), CStr(Numerico(.get_TextMatrix(I, C_ColCODGRUPO))), CStr(Numerico(.get_TextMatrix(I, C_COLCODFAMILIA))), CStr(Numerico(.get_TextMatrix(I, C_COLCODLINEA))), CStr(Numerico(.get_TextMatrix(I, C_COLCODSUBLINEA))), CStr(Numerico(CStr(Val(.get_TextMatrix(I, C_COLCODKILATES))))), CStr(Numerico(.get_TextMatrix(I, C_COLCODMARCA))), CStr(Numerico(.get_TextMatrix(I, C_COLCODMODELO))), CStr(Numerico(.get_TextMatrix(I, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(I, C_COLGENERO)), Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)), CStr(.get_TextMatrix(I, C_COLCRONO)), CStr(BuscaCodUnidad(.get_TextMatrix(I, C_COLUNIDAD))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)), Trim(cMonedadeCompra), Trim(CStr(nPrecioPP)), CStr(nCostoFacturaCurIva), CStr(nCostoAdicionalCur), CStr(nCostoIndirectosCur), CStr(nCostoRealCur), CStr(nCostoFacturaPesosCur), CStr(nCostoAdicionalPesosCur), CStr(nCostoIndirectosPesosCur), nMonedaPP, CStr(nOrigenAnt), CStr(nCodigoAnt), Trim(.get_TextMatrix(I, C_COLADICIONAL)), CStr(Numerico(.get_TextMatrix(I, C_ColMDSPESO))), Trim(.get_TextMatrix(I, C_ColMDSCOLOR)), Trim(.get_TextMatrix(I, C_ColMDSPUREZA)), Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)), C_INSERCION, CStr(0))
                        Cmd.Execute()
                        'Obtiene el código del artículo para guardar el código en el PreCatálogo
                        nCodArticulo = Cmd.Parameters("ID").Value
                        .set_TextMatrix(I, C_COLCODIGO, nCodArticulo)
                        '''Graba la imagen con el no. del articulo asociado
                    Else
                        nCodArticulo = CInt(Numerico(.get_TextMatrix(I, C_COLCODIGO)))
                        ''' 27OCT2010 - MAVF --> SE AGREGARON 4 CAMPOS PARA MANEJO DE DIAMANTE SUELTO
                        ModStoredProcedures.PR_IMECatArticulos(Trim(.get_TextMatrix(I, C_COLCODIGO)), Trim(.get_TextMatrix(I, C_COLDESCRIPCION)), CStr(Numerico(.get_TextMatrix(I, C_ColCODGRUPO))), CStr(Numerico(.get_TextMatrix(I, C_COLCODFAMILIA))), CStr(Numerico(.get_TextMatrix(I, C_COLCODLINEA))), CStr(Numerico(.get_TextMatrix(I, C_COLCODSUBLINEA))), CStr(Numerico(CStr(Val(.get_TextMatrix(I, C_COLCODKILATES))))), CStr(Numerico(.get_TextMatrix(I, C_COLCODMARCA))), CStr(Numerico(.get_TextMatrix(I, C_COLCODMODELO))), CStr(Numerico(.get_TextMatrix(I, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(I, C_COLGENERO)), Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)), CStr(.get_TextMatrix(I, C_COLCRONO)), CStr(BuscaCodUnidad(.get_TextMatrix(I, C_COLUNIDAD))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)), Trim(cMonedadeCompra), CStr(nPrecioPP), CStr(nCostoFacturaCurIva), CStr(nCostoAdicionalCur), CStr(nCostoIndirectosCur), CStr(nCostoRealCur), CStr(nCostoFacturaPesosCur), CStr(nCostoAdicionalPesosCur), CStr(nCostoIndirectosPesosCur), nMonedaPP, CStr(nOrigenAnt), CStr(nCodigoAnt), Trim(.get_TextMatrix(I, C_COLADICIONAL)), CStr(Numerico(.get_TextMatrix(I, C_ColMDSPESO))), Trim(.get_TextMatrix(I, C_ColMDSCOLOR)), Trim(.get_TextMatrix(I, C_ColMDSPUREZA)), Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)), C_MODIFICACION, CStr(1))
                        Cmd.Execute()
                    End If

                    '''Imagen
                    If Trim(.get_TextMatrix(I, C_ColIMAGEN)) <> "" Then
                        lRuta = Trim(.get_TextMatrix(I, C_ColIMAGEN))
                        lNomArch = Mid(lRuta, InStrRev(lRuta, "\") + 1, InStrRev(lRuta, ".") - (InStrRev(lRuta, "\") + 1))
                        lExtension = Mid(lRuta, InStrRev(lRuta, ".") + 1, Len(lRuta) - InStrRev(lRuta, "."))

                        If Trim(lRuta) <> "" Then
                            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                            lImagen = Dir(My.Application.Info.DirectoryPath & "\Sistema\Imagenes\" & Trim(CStr(nCodArticulo)) & "." & lExtension)
                            If lImagen <> "" Then
                                Kill(My.Application.Info.DirectoryPath & "\Sistema\Imagenes\" & Trim(CStr(nCodArticulo)) & "." & lExtension)
                            End If
                            lNvaRuta = My.Application.Info.DirectoryPath & "\Sistema\Imagenes\" & Trim(CStr(nCodArticulo)) & "." & lExtension
                            FileCopy(lRuta, lNvaRuta)
                            '''Elimina el termporal de PreCat
                            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                            If Dir(lRuta) <> "" Then
                                Kill(My.Application.Info.DirectoryPath & "\Sistema\PreCat\" & Trim(lNomArch) & "." & lExtension)
                            End If
                        End If
                    End If

                    ModStoredProcedures.PR_IMEOrdenesCompraPreCatAux(Trim(Me.txtFolio.Text), Trim(.get_TextMatrix(I, C_COLCODAUX)), Trim(.get_TextMatrix(I, C_COLCODIGO)), Trim(.get_TextMatrix(I, C_COLDESCRIPCION)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLPRECIOUNITARIO)), Trim(.get_TextMatrix(I, C_COLCOSTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCUENTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCTOPORC)), Trim(.get_TextMatrix(I, C_COLIVACUR)), Trim(.get_TextMatrix(I, C_COLCOSTOADICIONALCUR)), Trim(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR)), CStr(mintCodGrupo), Trim(.get_TextMatrix(I, C_COLCODFAMILIA)), Trim(.get_TextMatrix(I, C_COLCODLINEA)), Trim(.get_TextMatrix(I, C_COLCODSUBLINEA)), Trim(.get_TextMatrix(I, C_COLCODKILATES)), Trim(.get_TextMatrix(I, C_COLCODMARCA)), Trim(.get_TextMatrix(I, C_COLCODMODELO)), Trim(.get_TextMatrix(I, C_COLCODTIPOMATERIAL)), Trim(.get_TextMatrix(I, C_COLGENERO)), Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(I, C_COLCRONO)), CStr(BuscaCodUnidad(Trim(.get_TextMatrix(I, C_COLUNIDAD)))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)), C_CONCILIADO, .get_TextMatrix(I, C_COLADICIONAL), .get_TextMatrix(I, C_COLPRECIOPUBDOLAR), .get_TextMatrix(I, C_COLMONEDAPP), .get_TextMatrix(I, C_COLORIGENANT), .get_TextMatrix(I, C_ColCODIGOANT), .get_TextMatrix(I, C_ColIMAGEN), CStr(Numerico(.get_TextMatrix(I, C_ColMDSPESO))), Trim(.get_TextMatrix(I, C_ColMDSCOLOR)), Trim(.get_TextMatrix(I, C_ColMDSPUREZA)), Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)), C_MODIFICACION, CStr(1))
                    Cmd.Execute()
                    NumPartida = NumPartida + 1
                    PrecioVenta = 0
                    Estatus = "V"

                    If Not mblnEncabezadoAlmacen Then
                        'Aquí se debe relizar el Movimiento de Almacén. La Entrada de Artículos al Almacén.
                        'EL Registro es para el Cabecero de Movimientos de Almacen.
                        If RealizarMovimientosDeAlmacenCAB(CStr(Today), gintCodAlmacenGral, CInt(CStr(mintCodOrigen)), Trim(Me.txtFolio.Text), CInt(CStr(mintCodProveedor)), "", CodMovtoAlm, "E", "", "") = False Then
                            Err.Raise(0,  , "Error al Registrar la Entrada de Artículos al Inventario Por Concepto de Compra(CAB) (Form. frmCXPOrdenCompra)")
                        End If
                        mblnEncabezadoAlmacen = True
                    End If

                    'Aquí se debe relizar el Movimiento de Almacén. La Entrada de Artículos al Almacén.
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(FolioAlmacen, CStr(NumPartida), VB6.Format(Today, C_FORMATFECHAGUARDAR), CStr(nCodArticulo), CStr(mintCodOrigen), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), CStr(nCostoRealCur), CStr(PrecioVenta), CStr(0), Estatus, "01/01/1900", CStr(0), C_INSERCION, CStr(0))
                    Cmd.Execute()

                    'Se Realiza la Entrada al Inventario
                    TipoAlmacen = ObtenerTipoAlmacen(gintCodAlmacenGral)
                    UltimoCostoMN = CStr(nCostoRealCur * CDbl(Numerico(txtTipoCambio.Text)))
                    ModStoredProcedures.PR_IE_Inventario(CStr(gintCodAlmacenGral), IIf((TipoAlmacen = "P"), 1, 0), CStr(nCodArticulo), CStr(mintCodOrigen), CStr(0), CStr(0), CStr(0), CStr(UltimoCostoMN), CStr(nCostoRealCur), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), CStr(0), CStr(0), CStr(CodMovtoAlm), VB6.Format(Today, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()

                ElseIf Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CR Then
                    'Cuando el producto está conciliado y es resurtido
                    'Afecta el almacén y Actualiza los costos del Artículo en el catálogo de Artículos
                    '''Se cambio CostoFacturaCur --> CostoFacturaCurIva
                    '''ModStoredProcedures.PR_IMECatArticulos Trim(.TextMatrix(I, C_COLCODIGO)), _
                    ''''Trim(.TextMatrix(I, C_ColDESCRIPCION)), CStr(Numerico(.TextMatrix(I, C_ColCODGRUPO))), CStr(Numerico(.TextMatrix(I, C_COLCODFAMILIA))), CStr(Numerico(.TextMatrix(I, C_COLCODLINEA))), CStr(Numerico(.TextMatrix(I, C_COLCODSUBLINEA))), CStr(Numerico(Val(.TextMatrix(I, C_COLCODKILATES)))), CStr(Numerico(.TextMatrix(I, C_COLCODMARCA))), CStr(Numerico(.TextMatrix(I, C_COLCODMODELO))), CStr(Numerico(.TextMatrix(I, C_COLCODTIPOMATERIAL))), Trim(.TextMatrix(I, C_COLGENERO)), Trim(.TextMatrix(I, C_COLMOVIMIENTO)), CStr(.TextMatrix(I, C_COLCRONO)), CStr(BuscaCodUnidad(.TextMatrix(I, C_ColUNIDAD))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.TextMatrix(I, C_COLCODIGOARTICULOPROV)), Trim(cMonedadeCompra), CStr(0), CStr(nCostoFacturaCurIva), CStr(nCostoAdicionalCur), CStr(nCostoIndirectosCur), CStr(nCostoRealCur), CStr(nCostoFacturaPesosCur), CStr(nCostoAdicionalPesosCur), CStr(nCostoIndirectosPesosCur), 0, 0, 0, "", C_MODIFICACION, 1

                    If Trim(.get_TextMatrix(I, C_ColIMAGEN)) <> "" Then

                        lRuta = Trim(.get_TextMatrix(I, C_ColIMAGEN))
                        lNomArch = Mid(lRuta, InStrRev(lRuta, "\") + 1, InStrRev(lRuta, ".") - (InStrRev(lRuta, "\") + 1))
                        lExtension = Mid(lRuta, InStrRev(lRuta, ".") + 1, Len(lRuta) - InStrRev(lRuta, "."))
                        If Trim(lRuta) <> "" Then
                            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                            lImagen = Dir(My.Application.Info.DirectoryPath & "\Sistema\Imagenes\" & Trim(.get_TextMatrix(I, C_COLCODIGO)) & "." & lExtension)
                            If lImagen <> "" Then
                                Kill(My.Application.Info.DirectoryPath & "\Sistema\Imagenes\" & Trim(.get_TextMatrix(I, C_COLCODIGO)) & "." & lExtension)
                            End If
                            lNvaRuta = My.Application.Info.DirectoryPath & "\Sistema\Imagenes\" & Trim(.get_TextMatrix(I, C_COLCODIGO)) & "." & lExtension
                            FileCopy(lRuta, lNvaRuta)
                            '''Elimina el termporal de PreCat
                            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                            If Dir(lRuta) <> "" Then
                                Kill(My.Application.Info.DirectoryPath & "\Sistema\PreCat\" & Trim(lNomArch) & "." & lExtension)
                            End If
                        End If
                    End If

                    ''' 27OCT2010 - MAVF --> SE AGREGARON 4 CAMPOS PARA MANEJO DE DIAMANTE SUELTO
                    ModStoredProcedures.PR_IMECatArticulos(Trim(.get_TextMatrix(I, C_COLCODIGO)), Trim(.get_TextMatrix(I, C_COLDESCRIPCION)), CStr(Numerico(.get_TextMatrix(I, C_ColCODGRUPO))), CStr(Numerico(.get_TextMatrix(I, C_COLCODFAMILIA))), CStr(Numerico(.get_TextMatrix(I, C_COLCODLINEA))), CStr(Numerico(.get_TextMatrix(I, C_COLCODSUBLINEA))), CStr(Numerico(CStr(Val(.get_TextMatrix(I, C_COLCODKILATES))))), CStr(Numerico(.get_TextMatrix(I, C_COLCODMARCA))), CStr(Numerico(.get_TextMatrix(I, C_COLCODMODELO))), CStr(Numerico(.get_TextMatrix(I, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(I, C_COLGENERO)), Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)), CStr(.get_TextMatrix(I, C_COLCRONO)), CStr(BuscaCodUnidad(.get_TextMatrix(I, C_COLUNIDAD))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)), Trim(cMonedadeCompra), CStr(nPrecioPP), CStr(nCostoFacturaCurIva), CStr(nCostoAdicionalCur), CStr(nCostoIndirectosCur), CStr(nCostoRealCur), CStr(nCostoFacturaPesosCur), CStr(nCostoAdicionalPesosCur), CStr(nCostoIndirectosPesosCur), nMonedaPP, CStr(nOrigenAnt), CStr(nCodigoAnt), Trim(.get_TextMatrix(I, C_COLADICIONAL)), CStr(Numerico(.get_TextMatrix(I, C_ColMDSPESO))), Trim(.get_TextMatrix(I, C_ColMDSCOLOR)), Trim(.get_TextMatrix(I, C_ColMDSPUREZA)), Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)), C_MODIFICACION, CStr(1))
                    Cmd.Execute()

                    nCodArticulo = CInt(Trim(.get_TextMatrix(I, C_COLCODIGO)))
                    NumPartida = NumPartida + 1
                    PrecioVenta = 0
                    Estatus = "V"

                    If Not mblnEncabezadoAlmacen Then
                        'Aquí se debe relizar el Movimiento de Almacén. La Entrada de Artículos al Almacén.
                        'EL Registro es para el Cabecero de Movimientos de Almacen.
                        If RealizarMovimientosDeAlmacenCAB(CStr(Today), gintCodAlmacenGral, CInt(CStr(mintCodOrigen)), Trim(Me.txtFolio.Text), CInt(CStr(mintCodProveedor)), "", CodMovtoAlm, "E", "", "") = False Then
                            Err.Raise(0,  , "Error al Registrar la Entrada de Artículos al Inventario Por Concepto de Compra(CAB) (Form. frmCXPOrdenCompra)")
                        End If
                        mblnEncabezadoAlmacen = True
                    End If

                    'Aquí se debe relizar el Movimiento de Almacén. La Entrada de Artículos al Almacén.
                    ModStoredProcedures.PR_IE_MovtosAlmacenDet(FolioAlmacen, CStr(NumPartida), VB6.Format(Today, C_FORMATFECHAGUARDAR), CStr(nCodArticulo), CStr(mintCodOrigen), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), CStr(nCostoRealCur), CStr(PrecioVenta), CStr(0), Estatus, "01/01/1900", CStr(0), C_INSERCION, CStr(0))
                    Cmd.Execute()

                    'Se Realiza la Entrada al Inventario
                    TipoAlmacen = ObtenerTipoAlmacen(gintCodAlmacenGral)
                    UltimoCostoMN = CStr(nCostoRealCur * CDbl(txtTipoCambio.Text))
                    ModStoredProcedures.PR_IE_Inventario(CStr(gintCodAlmacenGral), IIf((TipoAlmacen = "P"), 1, 0), CStr(nCodArticulo), CStr(mintCodOrigen), CStr(0), CStr(0), CStr(0), CStr(UltimoCostoMN), CStr(nCostoRealCur), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), CStr(0), CStr(0), CStr(CodMovtoAlm), VB6.Format(Today, C_FORMATFECHAGUARDAR), C_INSERCION, CStr(0))
                    Cmd.Execute()

                Else 'Si no está conciliado, lo borra del Pre - Catálogo y del grid

                    '''Elimina el archivo del artículo que no llegó
                    If Trim(.get_TextMatrix(I, C_ColIMAGEN)) <> "" Then
                        lRuta = Trim(.get_TextMatrix(I, C_ColIMAGEN))
                        lNomArch = Mid(lRuta, InStrRev(lRuta, "\") + 1, InStrRev(lRuta, ".") - (InStrRev(lRuta, "\") + 1))
                        lExtension = Mid(lRuta, InStrRev(lRuta, ".") + 1, Len(lRuta) - InStrRev(lRuta, "."))

                        If Trim(lRuta) <> "" Then
                            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                            lImagen = Dir(My.Application.Info.DirectoryPath & "\Sistema\PreCat\" & Trim(lNomArch) & "." & lExtension)
                            If lImagen <> "" Then '''Existe
                                Kill(My.Application.Info.DirectoryPath & "\Sistema\PreCat\" & Trim(lNomArch) & "." & lExtension)
                            End If
                        End If
                    End If

                    If Trim(.get_TextMatrix(I, C_COLCODAUX)) <> "" Then
                        ModStoredProcedures.PR_IMEOrdenesCompraPreCatAux(Trim(Me.txtFolio.Text), Trim(.get_TextMatrix(I, C_COLCODAUX)), Trim(.get_TextMatrix(I, C_COLCODIGO)), Trim(.get_TextMatrix(I, C_COLDESCRIPCION)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLPRECIOUNITARIO)), Trim(.get_TextMatrix(I, C_COLCOSTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCUENTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCTOPORC)), Trim(.get_TextMatrix(I, C_COLIVACUR)), Trim(.get_TextMatrix(I, C_COLCOSTOADICIONALCUR)), Trim(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR)), CStr(mintCodGrupo), Trim(.get_TextMatrix(I, C_COLCODFAMILIA)), Trim(.get_TextMatrix(I, C_COLCODLINEA)), Trim(.get_TextMatrix(I, C_COLCODSUBLINEA)), Trim(.get_TextMatrix(I, C_COLCODKILATES)), Trim(.get_TextMatrix(I, C_COLCODMARCA)), Trim(.get_TextMatrix(I, C_COLCODMODELO)), Trim(.get_TextMatrix(I, C_COLCODTIPOMATERIAL)), Trim(.get_TextMatrix(I, C_COLGENERO)), Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(I, C_COLCRONO)), CStr(BuscaCodUnidad(Trim(.get_TextMatrix(I, C_COLUNIDAD)))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)), "", .get_TextMatrix(I, C_COLADICIONAL), .get_TextMatrix(I, C_COLPRECIOPUBDOLAR), .get_TextMatrix(I, C_COLMONEDAPP), .get_TextMatrix(I, C_COLORIGENANT), .get_TextMatrix(I, C_ColCODIGOANT), .get_TextMatrix(I, C_ColIMAGEN), CStr(Numerico(.get_TextMatrix(I, C_ColMDSPESO))), Trim(.get_TextMatrix(I, C_ColMDSCOLOR)), Trim(.get_TextMatrix(I, C_ColMDSPUREZA)), Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)), C_ELIMINACION, CStr(0))
                        Cmd.Execute()
                    End If
                    Me.mshFlex.RemoveItem((I))
                    I = I - 1
                End If
            Next I
        End With

        ActualizaCantidades()

        'Modificar los importes de la orden de compra
        ModStoredProcedures.PR_IMEOrdenesCompra(Trim(Me.txtFolio.Text), CStr(mintCodProveedor), VB6.Format(Me.dtpFecha.Value, C_FORMATFECHAGUARDAR), VB6.Format(Me.dtpFechaEntrega.Value, C_FORMATFECHAGUARDAR), Trim(Me.txtRemision.Text), Trim(Me.txtPedido.Text), CStr(mintCodOrigen), CStr(mintCodGrupo), CStr(ModEstandar.Numerico((Me.txtCostoAdicional.Text))), CStr(ModEstandar.Numerico((Me.txtCostosIndirectos.Text))), Trim(Me.rtEntregaren.Text), CStr(mcurSubTotal), CStr(mcurDESCUENTO), CStr(mcurIVA), CStr(mcurTotal), Trim(cMonedadeCompra), C_STGENERADA, VB6.Format(#1/1/1900#, C_FORMATFECHAGUARDAR), Trim(Me.txtTasaIva.Text), Trim(Me.txtPorcDescto.Text), Trim(Me.txtTipoCambio.Text), Trim(Me.txtTipoCambioEuro.Text), Trim(Me.txtTipoCambioConciliado.Text), Trim(Me.txtTipoCambioEuroConciliado.Text), Trim(Me.txtDesctoFinanciero.Text), VB6.Format(Today, C_FORMATFECHAGUARDAR), "", C_MODIFICACION, CStr(0))
        Cmd.Execute()

        Cnn.CommitTrans()
        blnTransaction = False
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Los artículos han sido registrados en el inventario correctamente", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)

        Select Case MsgBox("¿Desea Imprimir la Orden de Compra?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, gstrNombCortoEmpresa)
            Case MsgBoxResult.Yes
                Imprime()
        End Select
        Nuevo()
        Limpiar()

Merr:
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Sub ConvertirCantidades(ByRef MonedaAct As String, ByRef MonedaaConvertir As String)
        On Error GoTo Err_Renamed
        Dim I As Integer
        With mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then
                    Exit For
                End If
                If MonedaAct = "D" And MonedaaConvertir = "P" Then
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO4DEC, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC))) * CDbl(Numerico(txtTipoCambio.Text)), "###,##0.0000"))
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO, VB6.Format(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC)), "#####0.00"))
                ElseIf MonedaAct = "D" And MonedaaConvertir = "E" Then
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO4DEC, VB6.Format(CDbl(VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC))) * CDbl(Numerico(txtTipoCambio.Text)), "#####0.0000")) / CDbl(Numerico(txtTipoCambioEuro.Text)), "###,##0.0000"))
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO, VB6.Format(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC)), "#####0.00"))
                ElseIf MonedaAct = "P" And MonedaaConvertir = "D" Then
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO4DEC, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC))) / CDbl(Numerico(txtTipoCambio.Text)), "#####0.0000"))
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO, VB6.Format(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC)), "#####0.00"))
                ElseIf MonedaAct = "P" And MonedaaConvertir = "E" Then
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO4DEC, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC))) / CDbl(Numerico(txtTipoCambioEuro.Text)), "#####0.0000"))
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO, VB6.Format(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC)), "#####0.00"))
                ElseIf MonedaAct = "E" And MonedaaConvertir = "D" Then
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO4DEC, VB6.Format(CDbl(VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC))) * CDbl(Numerico(txtTipoCambioEuro.Text)), "#####0.0000")) / CDbl(Numerico(txtTipoCambio.Text)), "#####0.0000"))
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO, VB6.Format(VB6.Format(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC)), "#####0.0"), "###,##0.00"))
                ElseIf MonedaAct = "E" And MonedaaConvertir = "P" Then
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO4DEC, VB6.Format(CDbl(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC))) * CDbl(Numerico(txtTipoCambioEuro.Text)), "#####0.0000"))
                    .set_TextMatrix(I, C_COLPRECIOUNITARIO, VB6.Format(VB6.Format(Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO4DEC)), "#####0.0"), "###,##0.00"))
                End If
            Next
            ActualizaCantidades()
        End With
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function RealizarMovimientosDeAlmacenCAB(ByRef FechaAlmacen As String, ByRef CodAlmacen As Integer, ByRef CodAlmacenOrigen As Integer, ByRef FolioOC As String, ByRef CodProvAcreed As Integer, ByRef Factura As String, ByRef CodMovtoAlm As Integer, ByRef TipoMovimiento As String, ByRef Transporta As String, ByRef Entrega As String) As Boolean
        'TipoMovimiento = Tipo de Movimiento de Almacen (Entrada o Salida)
        On Error GoTo Merr
        '    Dim FolioAlmacen As String
        Dim Consecutivo As Integer
        Dim Estatus As String
        Dim NickUsuario As String
        Dim NumPartida As Integer
        Dim CodArticulo As Integer
        Dim Cantidad As Integer
        Dim CostoUnitario As Decimal
        Dim UltimoCostoMN As Decimal
        Dim PrecioVenta As Decimal
        Dim Descuento As Decimal
        Dim TipoCambio As Decimal
        Dim TipoAlmacen As String
        Dim Concepto As String
        Dim PorcIva As Decimal
        RealizarMovimientosDeAlmacenCAB = True

        Estatus = "V"
        TipoCambio = CDec(Numerico(txtTipoCambio.Text))
        NickUsuario = Trim(gStrNomUsuario)
        Concepto = "Entrada Por Compra " & FolioOC

        ModStoredProcedures.PR_I_FoliosAlmacen(CStr(gintCodAlmacenGral), CStr(0), C_INSERCION, CStr(0))
        RsGral = Cmd.Execute
        Consecutivo = Cmd.Parameters("Consecutivo").Value
        FolioAlmacen = C_PrefijoFoliosAlmacen & VB6.Format(gintCodAlmacenGral, "00") & CStr(Year(CDate(FechaAlmacen))) & VB6.Format(CStr(Month(CDate(FechaAlmacen))), "00") & VB6.Format(CStr(VB.Day(CDate(FechaAlmacen))), "00") & VB6.Format(Consecutivo, "000000")

        ModStoredProcedures.PR_IE_MovtosAlmacenCab(FolioAlmacen, VB6.Format(CDate(FechaAlmacen), C_FORMATFECHAGUARDAR), CStr(gintCodAlmacenGral), CStr(CodAlmacenOrigen), FolioOC, CStr(CodProvAcreed), Factura, CStr(0), CStr(CodMovtoAlm), TipoMovimiento, Transporta, Entrega, "", Concepto, Estatus, NickUsuario, "01/01/1900", "", "", "01/01/1900", CStr(0), "", "01/01/1900", CStr(TipoCambio), "", C_INSERCION, CStr(0))
        Cmd.Execute()
        Exit Function
Merr:
        RealizarMovimientosDeAlmacenCAB = False
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function


    'Obtener el Prefijo del folio de la orden de compra
    'Esta función es llamada después de haber añadido una orden de compra Nueva
    Public Function BuscaFolio() As String
        On Error GoTo Merr
        Dim Anio As String
        Dim Mes As String
        Dim Dia As String
        gStrSql = "SELECT Prefijo, Consecutivo FROM FoliosCorporativo WHERE CodFolio = 2 "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            'Obtiene el año, mes y día en formato string
            Anio = VB6.Format(Year(Today), "0000")
            Mes = VB6.Format(Month(Today), "00")
            Dia = VB6.Format(VB.Day(Today), "00")
            BuscaFolio = Trim(rsLocal.Fields("Prefijo").Value) & Anio & Mes & Dia & VB6.Format(rsLocal.Fields("Consecutivo").Value, "0000000000")
        Else
            BuscaFolio = CStr(0)
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Sub PonerColor()
        Dim I As Integer
        Dim Ctl As System.Windows.Forms.Control
        Dim nCol As Integer
        With mshFlex
            nCol = .Col
            Select Case Trim(.get_TextMatrix(.Row, C_ColSTATUS))
                Case C_CONCILIADO
                    Ctl = lblConciliado
                Case C_RESURTIDO
                    Ctl = lblResurtido
                Case C_CR
                    Ctl = lblCR
                Case ""
                    Ctl = mshFlex
            End Select
            For I = 0 To 7
                .Col = I
                'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.BackColor. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .CellBackColor = Ctl.BackColor
            Next
            .Col = C_ColIMPORTE
            'UPGRADE_WARNING: Couldn't resolve default property of object Ctl.BackColor. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .CellBackColor = Ctl.BackColor
            .Col = nCol
        End With
        Call ActualizaCantidades()
    End Sub

    Public Sub ScrollGrid()
        'Procedimiento que pone el enfoque en el primer renglón vacío del Grid
        Dim I As Integer
        Dim nCont As Integer 'Cuenta los renglones que están ocupados (que no están vacíos)
        Dim nRen As Integer
        'Aparecen 6 renglones disponibles en el Grid
        'Si son menos de seis registros ocupados, no se utiliza el .TopRow
        'Pero, si son 6 ó más registros, el .TopRow manda el enfoque al primer renglón vacío
        'después de los renglones ocupados
        nRen = 6 'El máximo de renglones que aparece en el grid (Además del encabezado)
        nCont = 0
        With Me.mshFlex
            For I = 1 To .Rows
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) <> "" Then
                    nCont = nCont + 1
                Else
                    Exit For
                End If
            Next I
            If nCont < 7 Then
                'Hay menos de 7 registros
                .Row = nCont + 1
                .Col = C_COLCODIGO
            Else
                'Hay 6 ó más registros, hay que recorrer el grid
                .TopRow = (nCont - nRen) + 2
                .Row = nCont + 1
                .Col = C_COLCODIGO
            End If
        End With
    End Sub

    Public Function BuscaOrigen(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codAlmacenOrigen, DescAlmacenOrigen FROM CatOrigen WHERE CodAlmacenOrigen = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaOrigen = Trim(rsLocal.Fields("DescAlmacenorigen").Value)
        Else
            BuscaOrigen = ""
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaUnidad(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codUnidad, DescUnidad FROM CatUnidades WHERE CodUnidad = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaUnidad = Trim(rsLocal.Fields("DescUnidad").Value)
        Else
            BuscaUnidad = ""
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaCodUnidad(ByRef Descripcion As String) As Integer
        On Error GoTo Merr
        gStrSql = "SELECT codUnidad, DescUnidad FROM CatUnidades WHERE LTrim(RTrim(DescUnidad)) = '" & Trim(Descripcion) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaCodUnidad = rsLocal.Fields("CodUnidad").Value
        Else
            BuscaCodUnidad = 0
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Function BuscaGrupo(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codGrupo, DescGrupo FROM CatGrupos WHERE CodGrupo = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaGrupo = Trim(rsLocal.Fields("DescGrupo").Value)
        Else
            BuscaGrupo = ""
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaNombreProveedor(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT DescProvAcreed FROM CatProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' and codProvAcreed = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaNombreProveedor = Trim(rsLocal.Fields("DescProvACreed").Value)
        Else
            BuscaNombreProveedor = ""
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaDesctoProveedor(ByRef Codigo As Integer) As Decimal
        On Error GoTo Merr
        gStrSql = "SELECT codProvAcreed, DesctoVolumen FROM CatProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' and codProvAcreed = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaDesctoProveedor = rsLocal.Fields("DesctoVolumen").Value
        Else
            BuscaDesctoProveedor = 0
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaDatosProveedor(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        Dim cTaxID As String
        gStrSql = "SELECT codProvAcreed, Nacional, DescProvAcreed, Domicilio, Ciudad, CP, Pais, Telefono, RFC, TaxId FROM CatProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' and codProvAcreed = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            If rsLocal.Fields("Nacional").Value Then
                cTaxID = Trim(rsLocal.Fields("Rfc").Value)
            Else
                cTaxID = Trim(rsLocal.Fields("TaxId").Value)
            End If
            BuscaDatosProveedor = Trim(rsLocal.Fields("Domicilio").Value) & vbNewLine & Trim(rsLocal.Fields("Ciudad").Value) & ", " & Trim(rsLocal.Fields("Pais").Value) & ", C.P." & Trim(rsLocal.Fields("CP").Value) & vbNewLine & "RFC: " & Trim(cTaxID) & vbNewLine & "Tel. " & Trim(rsLocal.Fields("Telefono").Value)
        Else
            BuscaDatosProveedor = ""
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaDesctoFinanciero(ByRef CodProveedor As Integer) As Integer
        On Error GoTo Merr
        gStrSql = "SELECT codProvAcreed, DesctoFinanciero FROM CatProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' and codProvAcreed = " & CodProveedor
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaDesctoFinanciero = CInt(VB6.Format(rsLocal.Fields("DesctoFinanciero").Value, "##0.00"))
        Else
            BuscaDesctoFinanciero = 0
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub LlenaLineaGrid()
        Dim I As Integer
        Dim Codigo As Integer
        Dim rsLineaGrid As ADODB.Recordset
        With Me.mshFlex
            If .Col = C_COLCODIGO Or .Col = C_COLDESCRIPCION Then
                If Trim(.get_TextMatrix(.Row - 1, C_COLDESCRIPCION)) <> "" And Trim(.get_TextMatrix(.Row, C_COLDESCRIPCION)) = "" Then
                    .Rows = .Rows + 1
                End If
            End If
        End With
        Codigo = CInt(Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_COLCODIGO))
        I = Me.mshFlex.Row
        gStrSql = "select * from CatArticulos where CodArticulo =" & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            With Me.mshFlex
                .set_TextMatrix(I, C_COLCODIGO, RsGral.Fields("CodArticulo").Value)
                .set_TextMatrix(I, C_COLDESCRIPCION, Trim(RsGral.Fields("DescArticulo").Value))
                'Hacer una función para Unidad [BuscaUnidad()]
                .set_TextMatrix(I, C_COLUNIDAD, Trim(BuscaUnidad(RsGral.Fields("CodUnidad").Value)))
                .set_TextMatrix(I, C_COLPORCIVA, gcurCorpoTASAIVA)
                .set_TextMatrix(I, C_COLCODAUX, "")
                .set_TextMatrix(I, C_ColCODGRUPO, RsGral.Fields("CodGrupo").Value)
                .set_TextMatrix(.Row, C_ColIMPORTE, "0.00")
                Select Case RsGral.Fields("CodGrupo").Value
                    Case gCODJOYERIA
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_COLCODFAMILIA, IIf(IsDBNull(RsGral.Fields("CodFamilia").Value), 0, RsGral.Fields("CodFamilia").Value))
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_COLCODLINEA, IIf(IsDBNull(RsGral.Fields("COdLinea").Value), 0, RsGral.Fields("COdLinea").Value))
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_COLCODSUBLINEA, IIf(IsDBNull(RsGral.Fields("CodSubLinea").Value), 0, RsGral.Fields("CodSubLinea").Value))
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_COLCODKILATES, IIf(IsDBNull(RsGral.Fields("CodKilates").Value), 0, RsGral.Fields("CodKilates").Value))
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_ColMDSPESO, IIf(IsDBNull(RsGral.Fields("mdsPeso").Value), 0, RsGral.Fields("mdsPeso").Value)) '''27OCT2010 - MAVF
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_ColMDSCOLOR, IIf(IsDBNull(RsGral.Fields("mdsColor").Value), 0, RsGral.Fields("mdsColor").Value)) '''27OCT2010 - MAVF
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_ColMDSPUREZA, IIf(IsDBNull(RsGral.Fields("mdsPureza").Value), 0, RsGral.Fields("mdsPureza").Value)) '''27OCT2010 - MAVF
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_ColMDSCERTIFICADO, IIf(IsDBNull(RsGral.Fields("mdsCertificado").Value), 0, RsGral.Fields("mdsCertificado").Value)) '''27OCT2010 - MAVF
                        .set_TextMatrix(I, C_COLCODMARCA, 0)
                        .set_TextMatrix(I, C_COLCODMODELO, 0)
                        .set_TextMatrix(I, C_COLGENERO, "")
                        .set_TextMatrix(I, C_COLMOVIMIENTO, "")
                        .set_TextMatrix(I, C_COLCRONO, False)
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        If IsDBNull(RsGral.Fields("CodFamilia").Value) Then
                            MsgBox("El código de la familia para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario" & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        ElseIf IsDBNull(RsGral.Fields("CodKilates").Value) Then
                            MsgBox("El código de kilates para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario" & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        ElseIf IsDBNull(RsGral.Fields("CodTipoMaterial").Value) Then
                            MsgBox("El código de Tipo de Material para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario" & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        ElseIf IsDBNull(RsGral.Fields("CodUnidad").Value) Then
                            MsgBox("El código de Unidad para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario" & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                        End If
                    Case gCODRELOJERIA
                        .set_TextMatrix(I, C_COLCODFAMILIA, 0)
                        .set_TextMatrix(I, C_COLCODLINEA, 0)
                        .set_TextMatrix(I, C_COLCODSUBLINEA, 0)
                        .set_TextMatrix(I, C_COLCODKILATES, 0)
                        .set_TextMatrix(I, C_ColMDSPESO, 0) '''27OCT2010 - MAVF
                        .set_TextMatrix(I, C_ColMDSCOLOR, "") '''27OCT2010 - MAVF
                        .set_TextMatrix(I, C_ColMDSPUREZA, "") '''27OCT2010 - MAVF
                        .set_TextMatrix(I, C_ColMDSCERTIFICADO, "") '''27OCT2010 - MAVF
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_COLCODMARCA, IIf(IsDBNull(RsGral.Fields("CodMArca").Value), 0, RsGral.Fields("CodMArca").Value))
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_COLCODMODELO, IIf(IsDBNull(RsGral.Fields("CodModelo").Value), 0, RsGral.Fields("CodModelo").Value))
                        .set_TextMatrix(I, C_COLGENERO, Trim(RsGral.Fields("Genero").Value))
                        .set_TextMatrix(I, C_COLMOVIMIENTO, Trim(RsGral.Fields("Movimiento").Value))
                        .set_TextMatrix(I, C_COLCRONO, RsGral.Fields("Crono").Value)
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        If IsDBNull(RsGral.Fields("CodMArca").Value) Then
                            MsgBox("El código de Marca para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario," & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                        ElseIf Trim(RsGral.Fields("Genero").Value) = "" Then
                            MsgBox("El campo Genero para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario," & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                        ElseIf Trim(RsGral.Fields("Movimiento").Value) = "" Then
                            MsgBox("El campo Movimiento para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario," & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        ElseIf IsDBNull(RsGral.Fields("CodTipoMaterial").Value) Then
                            MsgBox("El código de Tipo de Material para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario," & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        ElseIf IsDBNull(RsGral.Fields("CodUnidad").Value) Then
                            MsgBox("El código de Unidad para este artículo no existe" & vbNewLine & "debe clasificarse en el catálogo de Artículos de lo contario," & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                        End If
                    Case gCODVARIOS
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_COLCODFAMILIA, IIf(IsDBNull(RsGral.Fields("CodFamilia").Value), 0, RsGral.Fields("CodFamilia").Value))
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_COLCODLINEA, IIf(IsDBNull(RsGral.Fields("COdLinea").Value), 0, RsGral.Fields("COdLinea").Value))
                        .set_TextMatrix(I, C_COLCODSUBLINEA, 0)
                        .set_TextMatrix(I, C_COLCODKILATES, 0)
                        .set_TextMatrix(I, C_ColMDSPESO, 0) '''27OCT2010 - MAVF
                        .set_TextMatrix(I, C_ColMDSCOLOR, "") '''27OCT2010 - MAVF
                        .set_TextMatrix(I, C_ColMDSPUREZA, "") '''27OCT2010 - MAVF
                        .set_TextMatrix(I, C_ColMDSCERTIFICADO, "") '''27OCT2010 - MAVF
                        .set_TextMatrix(I, C_COLCODMARCA, 0)
                        .set_TextMatrix(I, C_COLCODMODELO, 0)
                        .set_TextMatrix(I, C_COLGENERO, "")
                        .set_TextMatrix(I, C_COLMOVIMIENTO, "")
                        .set_TextMatrix(I, C_COLCRONO, False)
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        If IsDBNull(RsGral.Fields("CodFamilia").Value) Then
                            MsgBox("El código de Familia para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario" & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        ElseIf IsDBNull(RsGral.Fields("CodTipoMaterial").Value) Then
                            MsgBox("El código de Tipo de Material para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario" & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        ElseIf IsDBNull(RsGral.Fields("CodUnidad").Value) Then
                            MsgBox("El código de Unidad para este artículo no existe" & vbNewLine & "Debe clasificarse en el catálogo de Artículos de lo contario" & vbNewLine & "no podrá generarse correctamente la Orden de Compra" & vbNewLine & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                        End If
                End Select
                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                .set_TextMatrix(I, C_COLCODTIPOMATERIAL, IIf(IsDBNull(RsGral.Fields("CodTipoMaterial").Value), 0, RsGral.Fields("CodTipoMaterial").Value))
                .set_TextMatrix(I, C_COLCODIGOARTICULOPROV, Trim(RsGral.Fields("CodigoArticuloProv").Value))
                .set_TextMatrix(I, C_COLDESCTO, 0)
                .set_TextMatrix(I, C_COLDESCTOPORC, Numerico((txtPorcDescto.Text)))
                .set_TextMatrix(I, C_COLDESCTOPORCTAG, Numerico((txtPorcDescto.Text)))
                .set_TextMatrix(I, C_COLADICIONAL, Trim(RsGral.Fields("Adicional").Value))

                .set_TextMatrix(I, C_COLPRECIOPUBDOLAR, Trim(RsGral.Fields("PrecioPubDolar").Value))
                .set_TextMatrix(I, C_COLMONEDAPP, IIf(RsGral.Fields("PesosFijos").Value, "P", "D"))
                .set_TextMatrix(I, C_COLORIGENANT, Trim(RsGral.Fields("OrigenAnt").Value))
                .set_TextMatrix(I, C_ColCODIGOANT, Trim(RsGral.Fields("CodigoAnt").Value))
                .set_TextMatrix(I, C_ColIMAGEN, "")

                .Col = C_COLCANTIDAD
                'Para indicar si es, o no, un resurtido necesito ver si hay alguna existencia en el Almacén
                gStrSql = "select * from Inventario where CodArticulo =" & Codigo
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.UP_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                rsLineaGrid = Cmd.Execute
                If rsLineaGrid.RecordCount > 0 Then
                    .set_TextMatrix(I, C_ColSTATUS, C_RESURTIDO)
                    'Obtenemos el Codigo Anterior
                    With mshFlex
                        gStrSql = "SELECT OCP.FolioOrdenCompra,OCP.CostoUnitario,OC.Moneda,OC.TipoCambio,OC.TipoCambioEuro FROM OrdenesCompraPreCat OCP INNER JOIN OrdenesCompra OC ON OCP.FolioOrdenCompra = OC.FolioOrdenCompra WHERE OCP.CodArticulo = " & Numerico(.get_TextMatrix(.Row, C_COLCODIGO)) & " AND OCP.CodAlmacenOrigen = " & mintCodOrigen & " AND OCP.CodGrupo = " & mintCodGrupo & " AND OCP.CodProveedor = " & mintCodProveedor & " AND OC.Estatus <> 'C' ORDER BY OCP.FolioOrdenCompra DESC"
                        ModEstandar.BorraCmd()
                        Cmd.CommandText = "dbo.UP_Select_Datos"
                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                        RsGral = Cmd.Execute
                        If RsGral.RecordCount > 0 Then
                            If optMoneda(0).Checked And RsGral.Fields("Moneda").Value = "D" Then
                                .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, VB6.Format(RsGral.Fields("CostoUnitario").Value, "###,##0.00"))
                            ElseIf optMoneda(0).Checked And RsGral.Fields("Moneda").Value = "P" Then
                                .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, VB6.Format(VB6.Format(RsGral.Fields("CostoUnitario").Value / CDbl(Numerico(txtTipoCambio.Text)), "#####0.0"), "###,##0.00"))
                            ElseIf optMoneda(0).Checked And RsGral.Fields("Moneda").Value = "E" Then
                                .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, VB6.Format(VB6.Format(CDbl(VB6.Format(RsGral.Fields("CostoUnitario").Value * CDbl(Numerico(txtTipoCambioEuro.Text)), "#####0.00")) / CDbl(Numerico(txtTipoCambio.Text)), "#####0.0"), "###,##0.00"))
                            End If
                            If optMoneda(1).Checked And RsGral.Fields("Moneda").Value = "P" Then
                                .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, VB6.Format(RsGral.Fields("CostoUnitario").Value, "###,##0.00"))
                            ElseIf optMoneda(1).Checked And RsGral.Fields("Moneda").Value = "D" Then
                                .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, VB6.Format(RsGral.Fields("CostoUnitario").Value * CDbl(Numerico(txtTipoCambio.Text)), "###,##0.00"))
                            ElseIf optMoneda(1).Checked And RsGral.Fields("Moneda").Value = "E" Then
                                .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, VB6.Format(VB6.Format(RsGral.Fields("CostoUnitario").Value * CDbl(Numerico(txtTipoCambioEuro.Text)), "#####0.0"), "###,##0.00"))
                            End If
                            If optMoneda(2).Checked And RsGral.Fields("Moneda").Value = "E" Then
                                .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, VB6.Format(RsGral.Fields("CostoUnitario").Value, "###,##0.00"))
                            ElseIf optMoneda(2).Checked And RsGral.Fields("Moneda").Value = "D" Then
                                .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, VB6.Format(CDbl(VB6.Format(RsGral.Fields("CostoUnitario").Value * CDbl(Numerico(txtTipoCambio.Text)), "#####0.00")) / CDbl(Numerico(txtTipoCambioEuro.Text)), "###,##0.00"))
                            ElseIf optMoneda(2).Checked And RsGral.Fields("Moneda").Value = "P" Then
                                .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, VB6.Format(RsGral.Fields("CostoUnitario").Value / CDbl(Numerico(txtTipoCambioEuro.Text)), "#####0.00"))
                            End If
                        Else
                            .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, "0.00")
                        End If
                        .set_TextMatrix(.Row, C_COLPRECIOUNITARIO4DEC, .get_TextMatrix(.Row, C_COLPRECIOUNITARIO))
                    End With
                Else
                    .set_TextMatrix(I, C_ColSTATUS, "")
                End If
            End With
        End If
    End Sub

    Public Sub LlenaDatos()
        On Error GoTo Merr
        Dim I As Integer
        If Len(Trim(Me.txtFolio.Text)) = 0 Then
            Nuevo()
            dbcProveedor.Focus()
            Exit Sub
        End If
        gStrSql = "select * from OrdenesCompra where FolioOrdenCompra ='" & Trim(Me.txtFolio.Text) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            'Verificar el estado en el que se encuentra la orden de compra
            If RsGral.Fields("Estatus").Value = C_STVIGENTE Then
                Me.dbcProveedor.Text = True
                Me.optMoneda(0).Enabled = True
                Me.optMoneda(1).Enabled = True
                Me.optMoneda(2).Enabled = True
                Me.txtTipoCambio.ReadOnly = True
                Me.txtTipoCambioEuro.ReadOnly = True
                Me.txtTasaIva.ReadOnly = False
                Me.txtPorcDescto.ReadOnly = False
                Me.txtDesctoFinanciero.ReadOnly = False
                Me.txtRemision.ReadOnly = False
                Me.txtPedido.ReadOnly = False
                Me.dtpFecha.Enabled = True
                Me.dtpFechaEntrega.Enabled = True
                Me.dbcOrigen.Text = True
                Me.dbcGrupo.Text = True
                Me.txtCostoAdicional.ReadOnly = False
                Me.txtCostosIndirectos.ReadOnly = False
                Me.rtEntregaren.ReadOnly = True
                Me.txtTipoCambioConciliado.Enabled = True
                Me.txtTipoCambioEuroConciliado.Enabled = True
                Me.btnAsignarCodigos.Enabled = True
                Me.btnProv.Enabled = True
            ElseIf RsGral.Fields("Estatus").Value = C_STGENERADA Or RsGral.Fields("Estatus").Value = C_STCANCELADA Or RsGral.Fields("Estatus").Value = C_STREGISTRADA Then
                Me.dbcProveedor.Text = True
                Me.optMoneda(0).Enabled = False
                Me.optMoneda(1).Enabled = False
                Me.optMoneda(2).Enabled = False
                Me.txtTipoCambio.ReadOnly = True
                Me.txtTipoCambioEuro.ReadOnly = True
                Me.txtTasaIva.ReadOnly = True
                Me.txtPorcDescto.ReadOnly = True
                Me.txtDesctoFinanciero.ReadOnly = True
                Me.txtRemision.ReadOnly = True
                Me.txtPedido.ReadOnly = True
                Me.dtpFecha.Enabled = False
                Me.dtpFechaEntrega.Enabled = False
                Me.dbcOrigen.Text = True
                Me.dbcGrupo.Text = True
                Me.txtCostoAdicional.ReadOnly = True
                Me.txtCostosIndirectos.ReadOnly = True
                Me.rtEntregaren.ReadOnly = True
                Me.txtTipoCambioConciliado.Enabled = False
                Me.txtTipoCambioEuroConciliado.Enabled = False
                Me.btnAsignarCodigos.Enabled = False
                Me.btnProv.Enabled = False
            End If

            If Trim(RsGral.Fields("FolioApartado").Value) <> "" Then
                Me.txtFolioApartado.Text = Trim(RsGral.Fields("FolioApartado").Value)
                fraApartado.Visible = True
            End If
            Me.txtTasaIva.Text = VB6.Format(RsGral.Fields("PorcIva").Value, "##0.00")
            Me.txtTasaIva.Tag = Me.txtTasaIva.Text
            Me.txtTipoCambio.Text = VB6.Format(RsGral.Fields("TipoCambio").Value, "###,###,##0.00")
            Me.txtTipoCambio.Tag = Me.txtTipoCambio.Text
            Me.txtTipoCambioEuro.Text = VB6.Format(RsGral.Fields("TipoCambioEuro").Value, "###,###,##0.00")
            Me.txtTipoCambioEuro.Tag = Me.txtTipoCambioEuro.Text
            Me.txtTipoCambioConciliado.Text = VB6.Format(RsGral.Fields("TipoCambioC").Value, "###,###,##0.00")
            Me.txtTipoCambioConciliado.Tag = Me.txtTipoCambioConciliado.Text
            Me.txtTipoCambioEuroConciliado.Text = VB6.Format(RsGral.Fields("TipoCambioEuroC").Value, "###,###,##0.00")
            Me.txtTipoCambioEuroConciliado.Tag = Me.txtTipoCambioEuroConciliado.Text
            Me.txtPorcDescto.Text = VB6.Format(RsGral.Fields("PorcDescto").Value, "##0.00")
            Me.txtPorcDescto.Tag = Me.txtPorcDescto.Text
            Me.txtDesctoFinanciero.Text = VB6.Format(RsGral.Fields("PorcDesctoFinanciero").Value, "##0.00")
            Me.txtDesctoFinanciero.Tag = Me.txtDesctoFinanciero.Text

            mblnFueraChange = True
            mintCodProveedor = RsGral.Fields("CodProvAcreed").Value
            'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
            Me.dbcProveedor.Text = Trim(BuscaNombreProveedor(mintCodProveedor))
            'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
            Me.dbcProveedor.Tag = Me.dbcProveedor.Text
            mblnFueraChange = False
            Me.txtOtrosDatos.Text = Trim(BuscaDatosProveedor(mintCodProveedor))
            Me.txtOtrosDatos.Tag = Me.txtOtrosDatos.Text
            Me.dtpFecha.Value = VB6.Format(RsGral.Fields("FechaOrdenCompra").Value, "dd/MMM/yyyy")
            Me.dtpFecha.Tag = Me.dtpFecha.Value
            Me.dtpFechaEntrega.Value = VB6.Format(RsGral.Fields("FechaEntrega").Value, "dd/MMM/yyyy")
            Me.dtpFechaEntrega.Tag = Me.dtpFechaEntrega.Value
            Me.txtRemision.Text = Trim(RsGral.Fields("Remision").Value)
            Me.txtRemision.Tag = Me.txtRemision.Text
            Me.txtPedido.Text = Trim(RsGral.Fields("Pedido").Value)
            Me.txtPedido.Tag = Me.txtPedido.Text
            'Hacer una función que me devuelva el Origen
            mblnFueraChange = True
            mintCodOrigen = RsGral.Fields("Origen").Value
            'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
            Me.dbcOrigen.Text = Trim(BuscaOrigen(mintCodOrigen))
            'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
            Me.dbcOrigen.Tag = Me.dbcOrigen.Text
            'Hacer una función que me devuelva el Grupo
            mintCodGrupo = RsGral.Fields("CodGrupo").Value
            'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
            Me.dbcGrupo.Text = Trim(BuscaGrupo(mintCodGrupo))
            'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
            Me.dbcGrupo.Tag = Me.dbcGrupo.Text
            mblnFueraChange = False

            Me.txtCostoAdicional.Text = VB6.Format(RsGral.Fields("CostoAdicional").Value, "###,###,##0.00")
            Me.txtCostoAdicional.Tag = Me.txtCostoAdicional.Text
            Me.txtCostosIndirectos.Text = VB6.Format(RsGral.Fields("CostoIndirectos").Value, "###,###,##0.00")
            Me.txtCostosIndirectos.Tag = Me.txtCostosIndirectos.Text
            Me.rtEntregaren.Text = Trim(RsGral.Fields("Entregar").Value)
            Me.rtEntregaren.Tag = Me.rtEntregaren.Text

            Me.txtSubTotal.Text = VB6.Format(RsGral.Fields("SubTotal").Value, "###,###,##0.00")
            Me.txtSubTotal.Tag = Me.txtSubTotal.Text
            Me.txtDescuento.Text = VB6.Format(RsGral.Fields("Descuento").Value, "###,###,##0.00")
            Me.txtDescuento.Tag = Me.txtDescuento.Text
            Me.txtIVA.Text = VB6.Format(RsGral.Fields("Iva").Value, "###,###,##0.00")
            Me.txtIVA.Tag = Me.txtIVA.Text
            Me.txtTotal.Text = VB6.Format(RsGral.Fields("Total").Value, "###,###,##0.00")
            Me.txtTotal.Tag = Me.txtTotal.Text

            mcurSubTotal = CDec(VB6.Format(RsGral.Fields("SubTotal").Value, "###,###,##0.0000"))
            mcurDESCUENTO = CDec(VB6.Format(RsGral.Fields("Descuento").Value, "###,###,##0.0000"))
            mcurIVA = CDec(VB6.Format(RsGral.Fields("Iva").Value, "###,###,##0.0000"))
            mcurTotal = CDec(VB6.Format(RsGral.Fields("Total").Value, "###,###,##0.0000"))

            If RsGral.Fields("Moneda").Value = C_DOLAR Then
                Me.optMoneda(0).Checked = True
                Me.optMoneda(1).Checked = False
                Me.optMoneda(2).Checked = False
            ElseIf RsGral.Fields("Moneda").Value = C_PESO Then
                Me.optMoneda(0).Checked = False
                Me.optMoneda(1).Checked = True
                Me.optMoneda(2).Checked = False
            Else
                Me.optMoneda(0).Checked = False
                Me.optMoneda(1).Checked = False
                Me.optMoneda(2).Checked = True
            End If
            cMonedadeCompra = Trim(RsGral.Fields("Moneda").Value)
            cMonedadeCompraTag = cMonedadeCompra

            'Estatus de la Orden de Compra
            cESTATUSORDEN = UCase(Trim(RsGral.Fields("Estatus").Value))
            Me.lblEstatus.Visible = True
            If Trim(cESTATUSORDEN) = C_STVIGENTE Then
                Me.lblEstatus.Text = "VIGENTE"
            ElseIf Trim(cESTATUSORDEN) = C_STGENERADA Then
                Me.lblEstatus.Text = "GENERADA" 'Conciliada
            ElseIf Trim(cESTATUSORDEN) = C_STCANCELADA Then
                Me.lblEstatus.Text = "CANCELADA"
            ElseIf Trim(cESTATUSORDEN) = C_STREGISTRADA Then
                Me.lblEstatus.Text = "REGISTRADA"
            Else
                Me.lblEstatus.Visible = False
            End If

            'Llenar el Grid
            '''ojo quitar aux ya que este terminado
            '''gStrSql = "Select * From OrdenesCompraPreCat_Aux Where FolioOrdenCompra ='" & Trim(txtFolio.text) & "'"
            gStrSql = "Select * From OrdenesCompraPreCat     Where FolioOrdenCompra ='" & Trim(txtFolio.Text) & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            Encabezado()

            If RsGral.RecordCount > 0 Then
                RsGral.MoveFirst()
                With Me.mshFlex
                    If RsGral.RecordCount < 11 Then
                        .Rows = 11
                    Else
                        .Rows = RsGral.RecordCount + 2
                    End If
                    For I = 1 To RsGral.RecordCount
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        If IsDBNull(RsGral.Fields("CodArticulo").Value) Then
                            .set_TextMatrix(I, C_COLCODIGO, "")
                        Else
                            .set_TextMatrix(I, C_COLCODIGO, RsGral.Fields("CodArticulo").Value)
                        End If
                        .set_TextMatrix(I, C_COLDESCRIPCION, Trim(RsGral.Fields("DescArticulo").Value))
                        .set_TextMatrix(I, C_ColDESCRIPCIONTAG, Trim(RsGral.Fields("DescArticulo").Value))
                        'Hacer una función para Unidad [BuscaUnidad()]
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_COLUNIDAD, Trim(BuscaUnidad(IIf(IsDBNull(RsGral.Fields("CodUnidad").Value), 0, RsGral.Fields("CodUnidad").Value))))
                        .set_TextMatrix(I, C_COLUNIDADTAG, .get_TextMatrix(I, C_COLUNIDAD))
                        .set_TextMatrix(I, C_COLCANTIDAD, Trim(RsGral.Fields("CantidadRecepcion").Value))

                        .set_TextMatrix(I, C_COLCANTIDADTAG, .get_TextMatrix(I, C_COLCANTIDAD))
                        .set_TextMatrix(I, C_COLPRECIOUNITARIO, VB6.Format(RsGral.Fields("CostoUnitario").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLPRECIOUNITARIOTAG, .get_TextMatrix(I, C_COLPRECIOUNITARIO))
                        .set_TextMatrix(I, C_COLPRECIOUNITARIO4DEC, .get_TextMatrix(I, C_COLPRECIOUNITARIO))
                        .set_TextMatrix(I, C_COLCOSTO, VB6.Format(RsGral.Fields("Costo").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLCOSTOTAG, .get_TextMatrix(I, C_COLCOSTO))
                        .set_TextMatrix(I, C_COLDESCTO, VB6.Format(RsGral.Fields("Descuento").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLDESCTOTAG, .get_TextMatrix(I, C_COLDESCTO))
                        .set_TextMatrix(I, C_COLDESCTOPORC, VB6.Format(RsGral.Fields("PorcDescuento").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLDESCTOPORCTAG, .get_TextMatrix(I, C_COLDESCTOPORC))
                        .set_TextMatrix(I, C_COLIVA, VB6.Format(RsGral.Fields("Iva").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLIVATAG, .get_TextMatrix(I, C_COLIVA))

                        'Calcular el valor de la columna C_COLPORCIVA
                        If (RsGral.Fields("Costo").Value - RsGral.Fields("Descuento").Value) > 0 Then
                            .set_TextMatrix(I, C_COLPORCIVA, (RsGral.Fields("Iva").Value * 100) / ((RsGral.Fields("Costo").Value + RsGral.Fields("CostoAdicional").Value + RsGral.Fields("CostoIndirectos").Value) - RsGral.Fields("Descuento").Value))
                        Else 'El descuento no es válido
                            .set_TextMatrix(I, C_COLPORCIVA, 0)
                        End If
                        .set_TextMatrix(I, C_COLCODAUX, RsGral.Fields("NumPartida").Value)

                        .set_TextMatrix(I, C_COLCOSTOADICIONAL, VB6.Format(RsGral.Fields("CostoAdicional").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLCOSTOADICIONALTAG, .get_TextMatrix(I, C_COLCOSTOADICIONAL))
                        .set_TextMatrix(I, C_COLCOSTOINDIRECTOS, VB6.Format(RsGral.Fields("CostoIndirectos").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLCOSTOINDIRECTOSTAG, .get_TextMatrix(I, C_COLCOSTOINDIRECTOS))

                        'CostoCur
                        .set_TextMatrix(I, C_COLCOSTOCUR, VB6.Format(RsGral.Fields("Costo").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLDESCUENTOCUR, VB6.Format(RsGral.Fields("Descuento").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLIVACUR, VB6.Format(RsGral.Fields("Iva").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLCOSTOADICIONALCUR, VB6.Format(RsGral.Fields("CostoAdicional").Value, "###,###,##0.0000"))
                        .set_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR, VB6.Format(RsGral.Fields("CostoIndirectos").Value, "###,###,##0.0000"))

                        .set_TextMatrix(I, C_ColCODGRUPO, RsGral.Fields("CodGrupo").Value)
                        .set_TextMatrix(I, C_COLCODGRUPOTAG, .get_TextMatrix(I, C_ColCODGRUPO))
                        Select Case RsGral.Fields("CodGrupo").Value
                            Case gCODJOYERIA
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODFAMILIA, IIf(IsDBNull(RsGral.Fields("CodFamilia").Value), 0, RsGral.Fields("CodFamilia").Value))
                                .set_TextMatrix(I, C_COLCODFAMILIATAG, .get_TextMatrix(I, C_COLCODFAMILIA))
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODLINEA, IIf(IsDBNull(RsGral.Fields("COdLinea").Value), 0, RsGral.Fields("COdLinea").Value))
                                .set_TextMatrix(I, C_COLCODLINEATAG, .get_TextMatrix(I, C_COLCODLINEA))
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODSUBLINEA, IIf(IsDBNull(RsGral.Fields("CodSubLinea").Value), 0, RsGral.Fields("CodSubLinea").Value))
                                .set_TextMatrix(I, C_COLCODSUBLINEATAG, .get_TextMatrix(I, C_COLCODSUBLINEA))
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODKILATES, IIf(IsDBNull(RsGral.Fields("CodKilates").Value), 0, RsGral.Fields("CodKilates").Value))
                                .set_TextMatrix(I, C_COLCODKILATESTAG, .get_TextMatrix(I, C_COLCODKILATES))

                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_ColMDSPESO, IIf(IsDBNull(RsGral.Fields("mdsPeso").Value), 0, RsGral.Fields("mdsPeso").Value)) '''27OCT2010 - MAVF
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_ColMDSCOLOR, IIf(IsDBNull(RsGral.Fields("mdsColor").Value), "", RsGral.Fields("mdsColor").Value)) '''27OCT2010 - MAVF
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_ColMDSPUREZA, IIf(IsDBNull(RsGral.Fields("mdsPureza").Value), "", RsGral.Fields("mdsPureza").Value)) '''27OCT2010 - MAVF
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_ColMDSCERTIFICADO, IIf(IsDBNull(RsGral.Fields("mdsCertificado").Value), "", RsGral.Fields("mdsCertificado").Value)) '''27OCT2010 - MAVF

                                .set_TextMatrix(I, C_COLCODMARCA, 0)
                                .set_TextMatrix(I, C_COLCODMARCATAG, 0)
                                .set_TextMatrix(I, C_COLCODMODELO, 0)
                                .set_TextMatrix(I, C_COLCODMODELOTAG, 0)
                                .set_TextMatrix(I, C_COLGENERO, "")
                                .set_TextMatrix(I, C_COLGENEROTAG, "")
                                .set_TextMatrix(I, C_COLMOVIMIENTO, "")
                                .set_TextMatrix(I, C_COLMOVIMIENTOTAG, "")
                                .set_TextMatrix(I, C_COLCRONO, False)
                                .set_TextMatrix(I, C_COLCRONOTAG, False)

                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODFAMILIAX, IIf(IsDBNull(RsGral.Fields("CodFamilia").Value), 0, RsGral.Fields("CodFamilia").Value))
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODLINEAX, IIf(IsDBNull(RsGral.Fields("COdLinea").Value), 0, RsGral.Fields("COdLinea").Value))
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODSUBLINEAX, IIf(IsDBNull(RsGral.Fields("CodSubLinea").Value), 0, RsGral.Fields("CodSubLinea").Value))
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODKILATESX, IIf(IsDBNull(RsGral.Fields("CodKilates").Value), 0, RsGral.Fields("CodKilates").Value))
                                .set_TextMatrix(I, C_COLCODMARCAX, 0)
                                .set_TextMatrix(I, C_COLCODMODELOX, 0)
                                .set_TextMatrix(I, C_COLGENEROX, "")
                                .set_TextMatrix(I, C_COLMOVIMIENTOX, "")
                                .set_TextMatrix(I, C_COLCRONOX, False)
                                .set_TextMatrix(I, C_COLADICIONALX, Trim(RsGral.Fields("Adicional").Value))

                            Case gCODRELOJERIA
                                .set_TextMatrix(I, C_COLCODFAMILIA, 0)
                                .set_TextMatrix(I, C_COLCODFAMILIATAG, 0)
                                .set_TextMatrix(I, C_COLCODLINEA, 0)
                                .set_TextMatrix(I, C_COLCODLINEATAG, 0)
                                .set_TextMatrix(I, C_COLCODSUBLINEA, 0)
                                .set_TextMatrix(I, C_COLCODSUBLINEATAG, 0)
                                .set_TextMatrix(I, C_COLCODKILATES, 0)
                                .set_TextMatrix(I, C_COLCODKILATESTAG, 0)

                                .set_TextMatrix(I, C_ColMDSPESO, "0.00") '''27OCT2010 - MAVF
                                .set_TextMatrix(I, C_ColMDSCOLOR, "") '''27OCT2010 - MAVF
                                .set_TextMatrix(I, C_ColMDSPUREZA, "") '''27OCT2010 - MAVF
                                .set_TextMatrix(I, C_ColMDSCERTIFICADO, "") '''27OCT2010 - MAVF

                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODMARCA, IIf(IsDBNull(RsGral.Fields("CodMArca").Value), 0, RsGral.Fields("CodMArca").Value))
                                .set_TextMatrix(I, C_COLCODMARCATAG, .get_TextMatrix(I, C_COLCODMARCA))
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODMODELO, IIf(IsDBNull(RsGral.Fields("CodModelo").Value), 0, RsGral.Fields("CodModelo").Value))
                                .set_TextMatrix(I, C_COLCODMODELOTAG, .get_TextMatrix(I, C_COLCODMODELO))
                                .set_TextMatrix(I, C_COLGENERO, Trim(RsGral.Fields("Genero").Value))
                                .set_TextMatrix(I, C_COLGENEROTAG, .get_TextMatrix(I, C_COLGENERO))
                                .set_TextMatrix(I, C_COLMOVIMIENTO, Trim(RsGral.Fields("Movimiento").Value))
                                .set_TextMatrix(I, C_COLMOVIMIENTOTAG, .get_TextMatrix(I, C_COLMOVIMIENTO))
                                .set_TextMatrix(I, C_COLCRONO, RsGral.Fields("Crono").Value)
                                .set_TextMatrix(I, C_COLCRONOTAG, .get_TextMatrix(I, C_COLCRONO))

                                .set_TextMatrix(I, C_COLCODFAMILIAX, 0)
                                .set_TextMatrix(I, C_COLCODLINEAX, 0)
                                .set_TextMatrix(I, C_COLCODSUBLINEAX, 0)
                                .set_TextMatrix(I, C_COLCODKILATESX, 0)
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODMARCAX, IIf(IsDBNull(RsGral.Fields("CodMArca").Value), 0, RsGral.Fields("CodMArca").Value))
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODMODELOX, IIf(IsDBNull(RsGral.Fields("CodModelo").Value), 0, RsGral.Fields("CodModelo").Value))
                                .set_TextMatrix(I, C_COLGENEROX, Trim(RsGral.Fields("Genero").Value))
                                .set_TextMatrix(I, C_COLMOVIMIENTOX, Trim(RsGral.Fields("Movimiento").Value))
                                .set_TextMatrix(I, C_COLCRONOX, RsGral.Fields("Crono").Value)
                                .set_TextMatrix(I, C_COLADICIONALX, Trim(RsGral.Fields("Adicional").Value))

                            Case gCODVARIOS
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODFAMILIA, IIf(IsDBNull(RsGral.Fields("CodFamilia").Value), 0, RsGral.Fields("CodFamilia").Value))
                                .set_TextMatrix(I, C_COLCODFAMILIATAG, .get_TextMatrix(I, C_COLCODFAMILIA))
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODLINEA, IIf(IsDBNull(RsGral.Fields("COdLinea").Value), 0, RsGral.Fields("COdLinea").Value))
                                .set_TextMatrix(I, C_COLCODLINEATAG, .get_TextMatrix(I, C_COLCODLINEA))
                                .set_TextMatrix(I, C_COLCODSUBLINEA, 0)
                                .set_TextMatrix(I, C_COLCODSUBLINEATAG, 0)
                                .set_TextMatrix(I, C_COLCODKILATES, 0)
                                .set_TextMatrix(I, C_COLCODKILATESTAG, 0)

                                .set_TextMatrix(I, C_ColMDSPESO, "0.00") '''27OCT2010 - MAVF
                                .set_TextMatrix(I, C_ColMDSCOLOR, "") '''27OCT2010 - MAVF
                                .set_TextMatrix(I, C_ColMDSPUREZA, "") '''27OCT2010 - MAVF
                                .set_TextMatrix(I, C_ColMDSCERTIFICADO, "") '''27OCT2010 - MAVF

                                .set_TextMatrix(I, C_COLCODMARCA, 0)
                                .set_TextMatrix(I, C_COLCODMARCATAG, 0)
                                .set_TextMatrix(I, C_COLCODMODELO, 0)
                                .set_TextMatrix(I, C_COLCODMODELOTAG, 0)
                                .set_TextMatrix(I, C_COLGENERO, "")
                                .set_TextMatrix(I, C_COLGENEROTAG, "")
                                .set_TextMatrix(I, C_COLMOVIMIENTO, "")
                                .set_TextMatrix(I, C_COLMOVIMIENTOTAG, "")
                                .set_TextMatrix(I, C_COLCRONO, False)
                                .set_TextMatrix(I, C_COLCRONOTAG, False)

                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODFAMILIAX, IIf(IsDBNull(RsGral.Fields("CodFamilia").Value), 0, RsGral.Fields("CodFamilia").Value))
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                .set_TextMatrix(I, C_COLCODLINEAX, IIf(IsDBNull(RsGral.Fields("COdLinea").Value), 0, RsGral.Fields("COdLinea").Value))
                                .set_TextMatrix(I, C_COLCODSUBLINEAX, 0)
                                .set_TextMatrix(I, C_COLCODKILATESX, 0)
                                .set_TextMatrix(I, C_COLCODMARCAX, 0)
                                .set_TextMatrix(I, C_COLCODMODELOX, 0)
                                .set_TextMatrix(I, C_COLGENEROX, "")
                                .set_TextMatrix(I, C_COLMOVIMIENTOX, "")
                                .set_TextMatrix(I, C_COLCRONOX, False)
                                .set_TextMatrix(I, C_COLADICIONALX, Trim(RsGral.Fields("Adicional").Value))

                        End Select
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        .set_TextMatrix(I, C_COLCODTIPOMATERIAL, IIf(IsDBNull(RsGral.Fields("CodTipoMaterial").Value), 0, RsGral.Fields("CodTipoMaterial").Value))
                        .set_TextMatrix(I, C_COLCODTIPOMATERIALTAG, .get_TextMatrix(I, C_COLCODTIPOMATERIAL))
                        .set_TextMatrix(I, C_COLCODIGOARTICULOPROV, Trim(RsGral.Fields("CodigoArticuloProv").Value))
                        .set_TextMatrix(I, C_COLCODIGOARTICULOPROVTAG, .get_TextMatrix(I, C_COLCODIGOARTICULOPROV))
                        .set_TextMatrix(I, C_ColSTATUS, Trim(RsGral.Fields("Estatus").Value))
                        .set_TextMatrix(I, C_COLSTATUSTAG, .get_TextMatrix(I, C_ColSTATUS))

                        .set_TextMatrix(I, C_COLADICIONAL, Trim(RsGral.Fields("Adicional").Value))
                        .set_TextMatrix(I, C_COLPRECIOPUBDOLAR, Trim(RsGral.Fields("PrecioPubDolar").Value))
                        .set_TextMatrix(I, C_COLMONEDAPP, Trim(RsGral.Fields("MonedaPP").Value))
                        .set_TextMatrix(I, C_COLORIGENANT, Trim(RsGral.Fields("OrigenAnt").Value))
                        .set_TextMatrix(I, C_ColCODIGOANT, Trim(RsGral.Fields("CodigoAnt").Value))
                        .set_TextMatrix(I, C_ColIMAGEN, Trim(RsGral.Fields("Imagen").Value))

                        .set_TextMatrix(I, C_COLADICIONALTAG, Trim(RsGral.Fields("Adicional").Value))
                        .set_TextMatrix(I, C_COLPRECIOPUBDOLARTAG, Trim(RsGral.Fields("PrecioPubDolar").Value))
                        .set_TextMatrix(I, C_COLMONEDAPPTAG, Trim(RsGral.Fields("MonedaPP").Value))
                        .set_TextMatrix(I, C_COLORIGENANTTAG, Trim(RsGral.Fields("OrigenAnt").Value))
                        .set_TextMatrix(I, C_ColCODIGOANTTAG, Trim(RsGral.Fields("CodigoAnt").Value))
                        .set_TextMatrix(I, C_ColIMAGENTAG, Trim(RsGral.Fields("Imagen").Value))

                        .set_TextMatrix(I, C_COLSTATUSX, "")
                        RsGral.MoveNext()
                    Next I
                    '--------------------------------------------------------------------------------
                    For I = 1 To RsGral.RecordCount
                        .Row = I
                        .Col = 1
                        Call PonerColor()
                    Next I
                    '--------------------------------------------------------------------------------
                End With
            End If
            mblnCambiosEnCodigo = False
            mblnNuevo = False
            'Activamos el Boton de Imprimir
            ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Else
            MsjNoExiste("el Folio de la orden de compra", gstrNombCortoEmpresa)
            Me.mshFlex.Rows = 11
            Limpiar()
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub Buscar()
        On Error GoTo Merr
        Dim nCol As Integer
        Dim nRow As Integer
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendrá el string del tag que se le mandara al fromulario de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim strControlActual As String 'Nombre del control actual
        Dim nColumnaActual As Integer

        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        Select Case UCase(Me.ActiveControl.Name)
            Case "TXTFOLIO"
                'frmCXPConsultaOrden.cFORM = "FRMCXPORDENCOMPRA"
                'frmCXPConsultaOrden.ShowDialog()
                Exit Sub
            Case "MSHFLEX"
                Exit Sub
        End Select

        If Trim(cESTATUSORDEN) <> "" Then
            If Trim(cESTATUSORDEN) <> Trim(C_STVIGENTE) Then
                Exit Sub
            End If
        End If
        'Sólo se puede consultar si el renglón está listo para recibir un registro
        With Me.mshFlex
            If mintCodGrupo = 0 Then
                MsgBox("Debe seleccionar un Grupo de Artículos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.dbcGrupo.Focus()
                ModEstandar.SelTxt()
                Exit Sub
            End If
            nCol = .Col
            nRow = .Row
            '''en esta parte se validará si es el rengón, columna que le
            '''corresponde editarse
            If (.Row > 1) Then
                '''de tal modo que si el renglón es mayor que 1
                '''y si un renglón antes del renglón actual está vacío,
                '''el renglón actual no se editará
                If Trim(.get_TextMatrix(.Row - 1, C_COLDESCRIPCION)) = "" Then
                    .Focus()
                    Exit Sub
                End If
            End If
            If .Col = C_COLCODIGO Or .Col = C_COLDESCRIPCION Then
                If Trim(.get_TextMatrix(.Row - 1, C_COLDESCRIPCION)) <> "" And Trim(.get_TextMatrix(.Row, C_COLDESCRIPCION)) = "" Then
                    .Rows = .Rows + 1
                End If
            End If
        End With

        nColumnaActual = mshFlex.Col

        If nColumnaActual <> C_COLCODIGO And nColumnaActual <> C_COLDESCRIPCION Then
            Exit Sub
        End If

        If strControlActual = "MSHFLEX" Then
            '''        Select Case mintCodGrupo
            '''            Case gCODJOYERIA
            '''                strCaptionForm = "Consulta de Joyería"
            '''                If nColumnaActual = C_COLCODIGO Then
            '''                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, b.DescTipoMaterial AS MATERIAL, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a inner join CatTipoMaterial b on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODJOYERIA & " and a.CodProveedor = " & mintCodProveedor & " and a.CodAlmacenOrigen = " & mintCodOrigen & " ORDER BY a.CodArticulo"
            '''                ElseIf nColumnaActual = C_ColDESCRIPCION Then
            '''                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, b.DescTipoMaterial AS MATERIAL, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a inner join CatTipoMaterial b on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODJOYERIA & " and a.CodProveedor = " & mintCodProveedor & " and a.CodAlmacenOrigen = " & mintCodOrigen & " ORDER BY a.DescArticulo"
            '''                End If
            '''            Case gCODRELOJERIA
            '''                strCaptionForm = "Consulta de Relojería"
            '''                If nColumnaActual = C_COLCODIGO Then
            '''                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a inner join CatTipoMaterial b on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODRELOJERIA & " and a.CodProveedor = " & mintCodProveedor & " and a.CodAlmacenOrigen = " & mintCodOrigen & " ORDER BY a.CodArticulo"
            '''                ElseIf nColumnaActual = C_ColDESCRIPCION Then
            '''                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a inner join CatTipoMaterial b on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODRELOJERIA & " and a.CodProveedor = " & mintCodProveedor & " and a.CodAlmacenOrigen = " & mintCodOrigen & " ORDER BY a.DescArticulo"
            '''                End If
            '''            Case gCODVARIOS
            '''                strCaptionForm = "Consulta de Artículos Varios"
            '''                If nColumnaActual = C_COLCODIGO Then
            '''                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, b.DescTipoMaterial AS MATERIAL, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a inner join CatTipoMaterial b on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODVARIOS & " and a.CodProveedor = " & mintCodProveedor & " and a.CodAlmacenOrigen = " & mintCodOrigen & " ORDER BY a.CodArticulo"
            '''                ElseIf nColumnaActual = C_ColDESCRIPCION Then
            '''                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, b.DescTipoMaterial AS MATERIAL, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a inner join CatTipoMaterial b on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODVARIOS & " and a.CodProveedor = " & mintCodProveedor & " and a.CodAlmacenOrigen = " & mintCodOrigen & " ORDER BY a.DescArticulo"
            '''                End If
            '''            Case Else
            '''                'Sale de este sub para que no ejecute ninguna opción
            '''                Exit Sub
            '''        End Select
        ElseIf strControlActual = "TXTFLEX" Then
            Select Case mintCodGrupo
                Case gCODJOYERIA
                    strCaptionForm = "Consulta de Joyería"
                    If nColumnaActual = C_COLCODIGO Then
                        gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, b.DescTipoMaterial AS MATERIAL, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a inner join CatTipoMaterial b on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODJOYERIA & " and a.CodProveedor = " & mintCodProveedor & " and a.CodArticulo >= " & CInt(Numerico((Me.txtFlex.Text))) & " and a.CodAlmacenOrigen = " & mintCodOrigen & " ORDER BY a.CodArticulo"
                    Else
                        Exit Sub
                    End If
                Case gCODRELOJERIA
                    strCaptionForm = "Consulta de Relojería"
                    If nColumnaActual = C_COLCODIGO Then
                        gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a inner join CatTipoMaterial b on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODRELOJERIA & " and a.CodProveedor = " & mintCodProveedor & " and a.CodArticulo >= " & CInt(Numerico((Me.txtFlex.Text))) & " and a.CodAlmacenOrigen = " & mintCodOrigen & " ORDER BY a.CodArticulo"
                    Else
                        Exit Sub
                    End If
                Case gCODVARIOS
                    strCaptionForm = "Consulta de Artículos Varios"
                    If nColumnaActual = C_COLCODIGO Then
                        gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, b.DescTipoMaterial AS MATERIAL, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a inner join CatTipoMaterial b on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODVARIOS & " and a.CodProveedor = " & mintCodProveedor & " and a.CodArticulo >= " & CInt(Numerico((Me.txtFlex.Text))) & " and a.CodAlmacenOrigen = " & mintCodOrigen & " ORDER BY a.CodArticulo"
                    Else
                        Exit Sub
                    End If
                Case Else
                    'Sale de este sub para que no ejecute ninguna opción
                    Exit Sub
            End Select
        End If

        strSQL = gStrSql 'Se hace uso de una variable temporal para el query

        'Si hubo cambios y es una modificacion entonces preguntará si desea grabar los cambios
        '    If Cambios() And Not mblnNuevo Then
        '        Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
        '            Case vbYes: 'Guardar el registro
        '                If Not Guardar() Then
        '                    Exit Sub
        '                End If
        '            Case vbNo: 'No hace nada y permite que se cargue la consulta
        '            Case vbCancel: 'Cancela la consulta
        '                Exit Sub
        '        End Select
        '    End If

        gStrSql = strSQL 'Se regresa el valor de la variable temporal a la variable original

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            If strControlActual <> "TXTFLEX" Then
                'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
                MsgBox("No existen artículos registrados que provengan de este Proveedor ( " & Trim(Me.dbcProveedor.Text) & " ) en el grupo de " & Trim(Me.dbcGrupo.Text), MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Else
                MsgBox("No existen artículos cuyo código sea mayor o igual al código introducido", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            End If
            Me.mshFlex.Focus()
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta
        'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        'Load(FrmConsultas)
        If mintCodGrupo = gCODRELOJERIA Then
            Call ConfiguraConsultas(FrmConsultas, 8790, RsGral, strTag, strCaptionForm)
        Else
            Call ConfiguraConsultas(FrmConsultas, 10845, RsGral, strTag, strCaptionForm)
        End If

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "MSHFLEX"
                    If mintCodGrupo <> gCODRELOJERIA Then
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColWidth(1, 0, 6000) 'Columna de la Descripción
                        .set_ColWidth(2, 0, 2055) 'Columna de Tipo de Material
                        .set_ColWidth(3, 0, 1890) 'Columna del Código del Proveedor
                        .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                        .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    Else
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColWidth(1, 0, 6000) 'Columna de la Descripción
                        .set_ColWidth(2, 0, 1890) 'Columna del Código del Proveedor
                        .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                        .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    End If
                Case "TXTFLEX"
                    If mintCodGrupo <> gCODRELOJERIA Then
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColWidth(1, 0, 6000) 'Columna de la Descripción
                        .set_ColWidth(2, 0, 2055) 'Columna de Tipo de Material
                        .set_ColWidth(3, 0, 1890) 'Columna del Código del Proveedor
                        .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                        .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    Else
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColWidth(1, 0, 6000) 'Columna de la Descripción
                        .set_ColWidth(2, 0, 1890) 'Columna del Código del Proveedor
                        .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                        .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    End If
            End Select
        End With

        FrmConsultas.ShowDialog()
        Select Case strControlActual
            Case "MSHFLEX"
                Me.mshFlex.Focus()
            Case "TXTFLEX"
                With Me.mshFlex
                    If Trim(.get_TextMatrix(.Row, C_COLCODIGO)) <> "" Then
                        'Quiere decir que escogió un artículo
                        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left 'Alinear a la izquierda
                        txtFlex.Text = ""
                        txtFlex.Visible = False
                        .Focus()
                    Else
                        If Me.txtFlex.Visible Then
                            Me.txtFlex.Focus()
                        Else
                            .Focus()
                        End If
                    End If
                End With
        End Select
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub Cancelar()
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        If Trim(Me.txtFolio.Text) = "" Then
            MsgBox("Elija un Folio para cancelar", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        'Buscar el resistro en la tabla y verificar su estatus
        'Si está cancelada o generada, no hace nada
        gStrSql = "SELECT FolioOrdenCompra, Estatus FROM OrdenesCompra WHERE ltrim(rtrim(FolioOrdenCompra)) = '" & Trim(Me.txtFolio.Text) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            'Sólo puede eliminar aquellas que NO han sido generadas o canceladas
            If RsGral.Fields("Estatus").Value = C_STGENERADA Then
                MsgBox("No puede cancelar esta orden de compra debido a que ya ha sido GENERADA", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Sub
            End If
            If RsGral.Fields("Estatus").Value = C_STCANCELADA Then
                MsgBox("No puede cancelar esta orden de compra debido a que ya ha sido CANCELADA", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Sub
            End If
            If RsGral.Fields("Estatus").Value = C_STREGISTRADA Then
                MsgBox("No puede cancelar esta orden de compra debido a que ya ha sido REGISTRADA en una Factura", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Sub
            End If
            'Preguntar si desea borrar el registro
            If MsgBox("¿Está seguro de CANCELAR esta Orden?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
                Exit Sub
            End If
            Cnn.BeginTrans()
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            blnTransaction = True
            ModStoredProcedures.PR_IMEOrdenesCompra(Trim(Me.txtFolio.Text), CStr(mintCodProveedor), VB6.Format(Me.dtpFecha.Value, C_FORMATFECHAGUARDAR), VB6.Format(Me.dtpFechaEntrega.Value, C_FORMATFECHAGUARDAR), Trim(Me.txtRemision.Text), Trim(Me.txtPedido.Text), CStr(mintCodOrigen), CStr(mintCodGrupo), CStr(ModEstandar.Numerico((Me.txtCostoAdicional.Text))), CStr(ModEstandar.Numerico((Me.txtCostosIndirectos.Text))), Trim(Me.rtEntregaren.Text), CStr(mcurSubTotal), CStr(mcurDESCUENTO), CStr(mcurIVA), CStr(mcurTotal), Trim(cMonedadeCompra), C_STCANCELADA, VB6.Format(Today, C_FORMATFECHAGUARDAR), Trim(Me.txtTasaIva.Text), Trim(Me.txtPorcDescto.Text), Trim(Me.txtTipoCambio.Text), Trim(Me.txtTipoCambioEuro.Text), Trim(Me.txtTipoCambioConciliado.Text), Trim(Me.txtTipoCambioEuroConciliado.Text), Trim(Me.txtDesctoFinanciero.Text), VB6.Format(Today, C_FORMATFECHAGUARDAR), "", C_MODIFICACION, CStr(1)) 'Sólo cancela la orden
            Cmd.Execute()
            Cnn.CommitTrans()
            blnTransaction = False
            Limpiar()
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Else
            If Trim(Me.txtFolio.Text) = "" Then
                MsgBox("Debe especificar un Folio válido para Cancelar", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Exit Sub
            End If
        End If
Merr:
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function Guardar() As Boolean
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        Dim I As Integer
        Dim lNomArch As String
        Dim lExtension As String
        Dim lRuta As String
        Dim lImagen As String
        Dim lNvaRuta As String
        Dim lNombre As String

        lNomArch = ""
        lExtension = ""
        lExtension = ""
        lRuta = ""
        lImagen = ""
        lNombre = ""

        'Valida si todos los datos han sido llenados correctamnte para poder ser guardados
        If Not ValidaDatos() Then
            mblnNuevo = True
            Exit Function
        End If
        If Not ValidaFlex() Then Exit Function
        If Not Cambios() Then
            Limpiar()
            Exit Function
        End If

        Cnn.BeginTrans()
        blnTransaction = True
        If mblnNuevo Then
            'Dar de alta la nueva orden de compra
            ModStoredProcedures.PR_IMEOrdenesCompra(Trim(Me.txtFolio.Text), CStr(mintCodProveedor), VB6.Format(Me.dtpFecha.Value, C_FORMATFECHAGUARDAR), VB6.Format(Me.dtpFechaEntrega.Value, C_FORMATFECHAGUARDAR), Trim(Me.txtRemision.Text), Trim(Me.txtPedido.Text), CStr(mintCodOrigen), CStr(mintCodGrupo), CStr(ModEstandar.Numerico((Me.txtCostoAdicional.Text))), CStr(ModEstandar.Numerico((Me.txtCostosIndirectos.Text))), Trim(Me.rtEntregaren.Text), CStr(mcurSubTotal), CStr(mcurDESCUENTO), CStr(mcurIVA), CStr(mcurTotal), Trim(cMonedadeCompra), C_STVIGENTE, VB6.Format(#1/1/1900#, C_FORMATFECHAGUARDAR), Trim(Me.txtTasaIva.Text), Trim(Me.txtPorcDescto.Text), Trim(Me.txtTipoCambio.Text), Trim(Me.txtTipoCambioEuro.Text), Trim(Me.txtTipoCambioConciliado.Text), Trim(Me.txtTipoCambioEuroConciliado.Text), Trim(Me.txtDesctoFinanciero.Text), VB6.Format(#1/1/1900#, C_FORMATFECHAGUARDAR), "", C_INSERCION, CStr(0))
            Cmd.Execute()
            Me.txtFolio.Text = BuscaFolio()
        Else
            'Modificar los datos generales de la orden de compra pero sin modificar el estatus
            ModStoredProcedures.PR_IMEOrdenesCompra(Trim(Me.txtFolio.Text), CStr(mintCodProveedor), VB6.Format(Me.dtpFecha.Value, C_FORMATFECHAGUARDAR), VB6.Format(Me.dtpFechaEntrega.Value, C_FORMATFECHAGUARDAR), Trim(Me.txtRemision.Text), Trim(Me.txtPedido.Text), CStr(mintCodOrigen), CStr(mintCodGrupo), CStr(ModEstandar.Numerico((Me.txtCostoAdicional.Text))), CStr(ModEstandar.Numerico((Me.txtCostosIndirectos.Text))), Trim(Me.rtEntregaren.Text), CStr(mcurSubTotal), CStr(mcurDESCUENTO), CStr(mcurIVA), CStr(mcurTotal), Trim(cMonedadeCompra), "", VB6.Format(#1/1/1900#, C_FORMATFECHAGUARDAR), Trim(Me.txtTasaIva.Text), Trim(Me.txtPorcDescto.Text), Trim(Me.txtTipoCambio.Text), Trim(Me.txtTipoCambioEuro.Text), Trim(Me.txtTipoCambioConciliado.Text), Trim(Me.txtTipoCambioEuroConciliado.Text), Trim(Me.txtDesctoFinanciero.Text), VB6.Format(#1/1/1900#, C_FORMATFECHAGUARDAR), "", C_MODIFICACION, CStr(2))
            Cmd.Execute()
        End If
        'Guardar los datos del Grid en la tabla OrdenesCompraPreCat
        With mshFlex
            For I = 1 To .Rows - 1
                lNvaRuta = ""
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then Exit For
                If Trim(.get_TextMatrix(I, C_COLCRONO)) = "" Then .set_TextMatrix(I, C_COLCRONO, CStr(False))

                If Trim(.get_TextMatrix(I, C_ColIMAGEN)) <> "" Then
                    lRuta = Trim(.get_TextMatrix(I, C_ColIMAGEN))
                    lNomArch = Mid(lRuta, InStrRev(lRuta, "\") + 1, InStrRev(lRuta, ".") - (InStrRev(lRuta, "\") + 1))
                    lExtension = Mid(lRuta, InStrRev(lRuta, ".") + 1, Len(lRuta) - InStrRev(lRuta, "."))
                    lNombre = NombreArchArticulo((txtFolio.Text), CInt(Trim(CStr(I)))) '''I  -->  es el numero de partida cuando es nuevo

                    If Trim(Mid(lRuta, 1, 1) & ":") = gstrCorpoDriveLocal Then
                        If Trim(lRuta) <> "" Then
                            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                            lImagen = Dir(My.Application.Info.DirectoryPath & "\Sistema\PreCat\" & Trim(lNombre) & "." & lExtension)
                            If lImagen <> "" Then '''Si Existe la elimina
                                Kill(lImagen)
                            End If
                            lNvaRuta = My.Application.Info.DirectoryPath & "\Sistema\PreCat\" & Trim(lNombre) & "." & lExtension
                            FileCopy(lRuta, lNvaRuta)
                        End If
                    Else '''este archivo ya existe en la carpeta del PreCatalogo, no es necesario hacerle nada
                        lNvaRuta = lRuta
                    End If

                Else
                    lNombre = NombreArchArticulo((txtFolio.Text), CInt(Trim(CStr(I)))) '''I  -->  es el numero de partida cuando es nuevo
                    lExtension = "*"

                    'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                    lImagen = Dir(My.Application.Info.DirectoryPath & "\Sistema\PreCat\" & Trim(lNombre) & "." & lExtension)
                    If lImagen <> "" Then '''no existe
                        Kill(My.Application.Info.DirectoryPath & "\Sistema\PreCat\" & Trim(lNombre) & "." & lExtension)
                        lNvaRuta = ""
                    End If
                End If

                If Trim(.get_TextMatrix(I, C_COLCODAUX)) = "" Then
                    'Es un registro Nuevo
                    ModStoredProcedures.PR_IMEOrdenesCompraPreCatAux(Trim(txtFolio.Text), CStr(0), Trim(.get_TextMatrix(I, C_COLCODIGO)), Trim(.get_TextMatrix(I, C_COLDESCRIPCION)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLPRECIOUNITARIO)), Trim(.get_TextMatrix(I, C_COLCOSTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCUENTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCTOPORC)), Trim(.get_TextMatrix(I, C_COLIVACUR)), Trim(.get_TextMatrix(I, C_COLCOSTOADICIONALCUR)), Trim(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR)), CStr(mintCodGrupo), Trim(.get_TextMatrix(I, C_COLCODFAMILIA)), Trim(.get_TextMatrix(I, C_COLCODLINEA)), Trim(.get_TextMatrix(I, C_COLCODSUBLINEA)), Trim(.get_TextMatrix(I, C_COLCODKILATES)), Trim(.get_TextMatrix(I, C_COLCODMARCA)), Trim(.get_TextMatrix(I, C_COLCODMODELO)), Trim(.get_TextMatrix(I, C_COLCODTIPOMATERIAL)), Trim(.get_TextMatrix(I, C_COLGENERO)), Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(I, C_COLCRONO)), CStr(BuscaCodUnidad(Trim(.get_TextMatrix(I, C_COLUNIDAD)))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(I, C_ColSTATUS)), Trim(.get_TextMatrix(I, C_COLADICIONAL)), Trim(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR)), Trim(.get_TextMatrix(I, C_COLMONEDAPP)), Trim(.get_TextMatrix(I, C_COLORIGENANT)), Trim(.get_TextMatrix(I, C_ColCODIGOANT)), Trim(lNvaRuta), CStr(Numerico(.get_TextMatrix(I, C_ColMDSPESO))), Trim(.get_TextMatrix(I, C_ColMDSCOLOR)), Trim(.get_TextMatrix(I, C_ColMDSPUREZA)), Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    .set_TextMatrix(I, C_COLCODAUX, Cmd.Parameters("ID").Value)
                Else
                    'Es una modificación del registro
                    If Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CONCILIADO Then
                        'Cuando el producto se ha conciliado
                        ModStoredProcedures.PR_IMEOrdenesCompraPreCatAux(Trim(Me.txtFolio.Text), Trim(.get_TextMatrix(I, C_COLCODAUX)), Trim(.get_TextMatrix(I, C_COLCODIGO)), Trim(.get_TextMatrix(I, C_COLDESCRIPCION)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLPRECIOUNITARIO)), Trim(.get_TextMatrix(I, C_COLCOSTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCUENTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCTOPORC)), Trim(.get_TextMatrix(I, C_COLIVACUR)), Trim(.get_TextMatrix(I, C_COLCOSTOADICIONALCUR)), Trim(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR)), CStr(mintCodGrupo), Trim(.get_TextMatrix(I, C_COLCODFAMILIA)), Trim(.get_TextMatrix(I, C_COLCODLINEA)), Trim(.get_TextMatrix(I, C_COLCODSUBLINEA)), Trim(.get_TextMatrix(I, C_COLCODKILATES)), Trim(.get_TextMatrix(I, C_COLCODMARCA)), Trim(.get_TextMatrix(I, C_COLCODMODELO)), Trim(.get_TextMatrix(I, C_COLCODTIPOMATERIAL)), Trim(.get_TextMatrix(I, C_COLGENERO)), Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)), CStr(CBool(.get_TextMatrix(I, C_COLCRONO))), CStr(BuscaCodUnidad(Trim(.get_TextMatrix(I, C_COLUNIDAD)))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)), C_CONCILIADO, .get_TextMatrix(I, C_COLADICIONAL), .get_TextMatrix(I, C_COLPRECIOPUBDOLAR), .get_TextMatrix(I, C_COLMONEDAPP), .get_TextMatrix(I, C_COLORIGENANT), .get_TextMatrix(I, C_ColCODIGOANT), lNvaRuta, CStr(Numerico(.get_TextMatrix(I, C_ColMDSPESO))), Trim(.get_TextMatrix(I, C_ColMDSCOLOR)), Trim(.get_TextMatrix(I, C_ColMDSPUREZA)), Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)), C_MODIFICACION, CStr(1))
                        Cmd.Execute()
                    ElseIf Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_CR Then
                        'Cuando el producto está conciliado y es resurtido
                        ModStoredProcedures.PR_IMEOrdenesCompraPreCatAux(Trim(Me.txtFolio.Text), Trim(.get_TextMatrix(I, C_COLCODAUX)), Trim(.get_TextMatrix(I, C_COLCODIGO)), Trim(.get_TextMatrix(I, C_COLDESCRIPCION)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLPRECIOUNITARIO)), Trim(.get_TextMatrix(I, C_COLCOSTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCUENTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCTOPORC)), Trim(.get_TextMatrix(I, C_COLIVACUR)), Trim(.get_TextMatrix(I, C_COLCOSTOADICIONALCUR)), Trim(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR)), CStr(mintCodGrupo), Trim(.get_TextMatrix(I, C_COLCODFAMILIA)), Trim(.get_TextMatrix(I, C_COLCODLINEA)), Trim(.get_TextMatrix(I, C_COLCODSUBLINEA)), Trim(.get_TextMatrix(I, C_COLCODKILATES)), Trim(.get_TextMatrix(I, C_COLCODMARCA)), Trim(.get_TextMatrix(I, C_COLCODMODELO)), Trim(.get_TextMatrix(I, C_COLCODTIPOMATERIAL)), Trim(.get_TextMatrix(I, C_COLGENERO)), Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(I, C_COLCRONO)), CStr(BuscaCodUnidad(Trim(.get_TextMatrix(I, C_COLUNIDAD)))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)), C_CR, .get_TextMatrix(I, C_COLADICIONAL), .get_TextMatrix(I, C_COLPRECIOPUBDOLAR), .get_TextMatrix(I, C_COLMONEDAPP), .get_TextMatrix(I, C_COLORIGENANT), .get_TextMatrix(I, C_ColCODIGOANT), lNvaRuta, CStr(Numerico(.get_TextMatrix(I, C_ColMDSPESO))), Trim(.get_TextMatrix(I, C_ColMDSCOLOR)), Trim(.get_TextMatrix(I, C_ColMDSPUREZA)), Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)), C_MODIFICACION, CStr(1))
                        Cmd.Execute()
                    ElseIf Trim(.get_TextMatrix(I, C_ColSTATUS)) = C_RESURTIDO Then
                        'Cuando el producto es un resurtido
                        ModStoredProcedures.PR_IMEOrdenesCompraPreCatAux(Trim(Me.txtFolio.Text), Trim(.get_TextMatrix(I, C_COLCODAUX)), Trim(.get_TextMatrix(I, C_COLCODIGO)), Trim(.get_TextMatrix(I, C_COLDESCRIPCION)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLPRECIOUNITARIO)), Trim(.get_TextMatrix(I, C_COLCOSTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCUENTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCTOPORC)), Trim(.get_TextMatrix(I, C_COLIVACUR)), Trim(.get_TextMatrix(I, C_COLCOSTOADICIONALCUR)), Trim(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR)), CStr(mintCodGrupo), Trim(.get_TextMatrix(I, C_COLCODFAMILIA)), Trim(.get_TextMatrix(I, C_COLCODLINEA)), Trim(.get_TextMatrix(I, C_COLCODSUBLINEA)), Trim(.get_TextMatrix(I, C_COLCODKILATES)), Trim(.get_TextMatrix(I, C_COLCODMARCA)), Trim(.get_TextMatrix(I, C_COLCODMODELO)), Trim(.get_TextMatrix(I, C_COLCODTIPOMATERIAL)), Trim(.get_TextMatrix(I, C_COLGENERO)), Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(I, C_COLCRONO)), CStr(BuscaCodUnidad(Trim(.get_TextMatrix(I, C_COLUNIDAD)))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)), C_RESURTIDO, .get_TextMatrix(I, C_COLADICIONAL), .get_TextMatrix(I, C_COLPRECIOPUBDOLAR), .get_TextMatrix(I, C_COLMONEDAPP), .get_TextMatrix(I, C_COLORIGENANT), .get_TextMatrix(I, C_ColCODIGOANT), lNvaRuta, CStr(Numerico(.get_TextMatrix(I, C_ColMDSPESO))), Trim(.get_TextMatrix(I, C_ColMDSCOLOR)), Trim(.get_TextMatrix(I, C_ColMDSPUREZA)), Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)), C_MODIFICACION, CStr(1))
                        Cmd.Execute()
                    Else
                        ModStoredProcedures.PR_IMEOrdenesCompraPreCatAux(Trim(Me.txtFolio.Text), Trim(.get_TextMatrix(I, C_COLCODAUX)), Trim(.get_TextMatrix(I, C_COLCODIGO)), Trim(.get_TextMatrix(I, C_COLDESCRIPCION)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLCANTIDAD)), Trim(.get_TextMatrix(I, C_COLPRECIOUNITARIO)), Trim(.get_TextMatrix(I, C_COLCOSTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCUENTOCUR)), Trim(.get_TextMatrix(I, C_COLDESCTOPORC)), Trim(.get_TextMatrix(I, C_COLIVACUR)), Trim(.get_TextMatrix(I, C_COLCOSTOADICIONALCUR)), Trim(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSCUR)), CStr(mintCodGrupo), Trim(.get_TextMatrix(I, C_COLCODFAMILIA)), Trim(.get_TextMatrix(I, C_COLCODLINEA)), Trim(.get_TextMatrix(I, C_COLCODSUBLINEA)), Trim(.get_TextMatrix(I, C_COLCODKILATES)), Trim(.get_TextMatrix(I, C_COLCODMARCA)), Trim(.get_TextMatrix(I, C_COLCODMODELO)), Trim(.get_TextMatrix(I, C_COLCODTIPOMATERIAL)), Trim(.get_TextMatrix(I, C_COLGENERO)), Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(I, C_COLCRONO)), CStr(BuscaCodUnidad(Trim(.get_TextMatrix(I, C_COLUNIDAD)))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)), "", .get_TextMatrix(I, C_COLADICIONAL), .get_TextMatrix(I, C_COLPRECIOPUBDOLAR), .get_TextMatrix(I, C_COLMONEDAPP), .get_TextMatrix(I, C_COLORIGENANT), .get_TextMatrix(I, C_ColCODIGOANT), lNvaRuta, CStr(Numerico(.get_TextMatrix(I, C_ColMDSPESO))), Trim(.get_TextMatrix(I, C_ColMDSCOLOR)), Trim(.get_TextMatrix(I, C_ColMDSPUREZA)), Trim(.get_TextMatrix(I, C_ColMDSCERTIFICADO)), C_MODIFICACION, CStr(1))
                        Cmd.Execute()
                    End If
                End If


            Next I
        End With

        Cnn.CommitTrans()
        blnTransaction = False
        If Not mblnConciliar Then
            If mblnNuevo Then
                MsgBox("La Orden de Compra ha sido GUARDADA correctamente", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Else
                MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            End If
            Nuevo()
            Guardar = True
            Limpiar()
        Else
            mblnCambiosEnCodigo = False
            mblnNuevo = False
            Guardar = True
        End If
Merr:
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Function

    Sub Imprime()
        Dim rptCXPOrdenCompra As Object
        On Error GoTo Err_Renamed
        Dim rsReporte As ADODB.Recordset
        Dim NombreEmpresa As String
        Dim NombreReporte As String
        Dim Moneda As String
        Dim lStrSql As String
        lStrSql = "SELECT OC.FolioOrdenCompra,OC.FechaOrdenCompra,OC.FechaCompraEI,OC.FechaCancel," & "CASE WHEN OC.Estatus = 'V' THEN 'VIGENTE' WHEN OC.Estatus = 'R' THEN 'REGISTRADA' WHEN OC.Estatus = 'G' THEN 'GENERADA' WHEN OC.Estatus = 'C' THEN 'CANCELADA' END AS Estatus," & "OC.CodProvAcreed,CP.DescProvAcreed,CP.Domicilio,RTRIM(LTRIM(CP.Ciudad)) + ', ' + RTRIM(LTRIM(CP.Pais)) + ' C.P. ' + CP.CP AS Pais," & "'RFC ' + CP.RFC AS RFC,'Tel ' + CP.Telefono AS Telefono,OC.FolioApartado," & "CASE WHEN OC.Moneda = 'P' THEN 'PESOS' WHEN OC.Moneda = 'D' THEN 'DOLARES' WHEN OC.Moneda = 'E' THEN 'EUROS' END AS Moneda," & "OC.TipoCambioC,OC.TipoCambioEuroC,CAST(OC.PorcIva AS VarChar) + '%' AS PorcIva,CAST(OC.PorcDescto AS VarChar) + '%' AS PorcDescto," & "CAST(OC.PorcDesctoFinanciero AS VarChar) + '%' AS PorcDesctoFinanciero,OC.Remision,OC.Pedido,OC.Origen,CO.DescAlmacenOrigen," & "OC.CodGrupo,CG.DescGrupo,OC.CostoAdicional,OC.CostoIndirectos,OC.SubTotal,OC.Descuento,OC.Iva,OC.Total," & "ISNULL(OCPC.CodArticulo,0) AS CodArticulo,OCPC.NumPartida,ISNULL(OCPC.DescArticulo,'') AS DescArticulo," & "OCPC.CodigoArticuloProv,OCPC.CantidadRecepcion,ISNULL(OCPC.CodUnidad,0) AS CodUnidad,ISNULL(CU.DescUnidad,'') AS DescUnidad," & "OCPC.CostoUnitario,OCPC.Descuento,OCPC.Iva AS IvaUnit,((OCPC.CostoUnitario - OCPC.Descuento) + OCPC.Iva) * OCPC.CantidadRecepcion AS TotalPartida " & "FROM OrdenesCompra OC INNER JOIN OrdenesCompraPreCat OCPC ON OC.FolioOrdenCompra = OCPC.FolioOrdenCompra " & "INNER JOIN CatProvAcreed CP ON OC.CodProvAcreed = CP.CodProvAcreed " & "INNER JOIN CatOrigen CO ON OC.Origen = CO.CodAlmacenOrigen " & "INNER JOIN CatGrupos CG ON OC.CodGrupo = CG.CodGrupo " & "LEFT OUTER JOIN CatUnidades CU ON OCPC.CodUnidad = CU.CodUnidad " & "WHERE OC.FolioOrdenCompra = '" & Trim(txtFolio.Text) & "' " & "ORDER BY OCPC.NumPartida"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, lStrSql))
        rsReporte = Cmd.Execute
        If rsReporte.RecordCount = 0 Then
            MsgBox("Folio de Orden de Compra no Existe...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            frmReportes.Report = rptCXPOrdenCompra
            NombreEmpresa = UCase(gstrNombCortoEmpresa)
            NombreReporte = "ORDEN DE COMPRA"
            If optMoneda(0).Checked = True Then
                Moneda = "*** Importes expresados en Dolares"
            ElseIf optMoneda(1).Checked = True Then
                Moneda = "*** Importes expresados en Pesos"
            ElseIf optMoneda(2).Checked = True Then
                Moneda = "*** Importes expresados en Euros"
            End If
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            frmReportes.rsReport = rsReporte
            'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'UPGRADE_WARNING: Couldn't resolve default property of object frmReportes.aFormula_. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            frmReportes.aFormula_ = New Object() {"Moneda", "NombreEmpresa", "NombreReporte"}
            'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'UPGRADE_WARNING: Couldn't resolve default property of object frmReportes.aValues_. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            frmReportes.aValues_ = New Object() {Moneda, NombreEmpresa, NombreReporte}
            frmReportes.Text = "Orden de Compra"
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            frmReportes.ShowDialog()
        End If
Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function Cambios() As Boolean
        Dim I As Integer
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Select Case True
            Case cMonedadeCompra <> cMonedadeCompraTag
                Cambios = True
                Exit Function
            Case Numerico((Me.txtTipoCambio.Text)) <> Numerico((Me.txtTipoCambio.Tag))
                Cambios = True
                Exit Function
            Case Numerico((Me.txtTipoCambioEuro.Text)) <> Numerico((Me.txtTipoCambioEuro.Tag))
                Cambios = True
                Exit Function
            Case Numerico((Me.txtTasaIva.Text)) <> Numerico((Me.txtTasaIva.Tag))
                Cambios = True
                Exit Function
            Case Numerico((Me.txtPorcDescto.Text)) <> Numerico((Me.txtPorcDescto.Tag))
                Cambios = True
                Exit Function
            Case Numerico((Me.txtDesctoFinanciero.Text)) <> Numerico((Me.txtDesctoFinanciero.Tag))
                Cambios = True
                Exit Function
            Case Trim(Me.txtRemision.Text) <> Trim(Me.txtRemision.Tag)
                Cambios = True
                Exit Function
            Case Trim(Me.txtPedido.Text) <> Trim(Me.txtPedido.Tag)
                Cambios = True
                Exit Function
            Case Trim(Me.txtOtrosDatos.Text) <> Trim(Me.txtOtrosDatos.Tag)
                Cambios = True
                Exit Function
            Case VB6.Format(Me.dtpFecha.Value, "dd/MMM/yyyy") <> VB6.Format(Me.dtpFecha.Tag, "dd/MMM/yyyy")
                Cambios = True
                Exit Function
            Case VB6.Format(Me.dtpFechaEntrega.Value, "dd/MMM/yyyy") <> VB6.Format(Me.dtpFechaEntrega.Tag, "dd/MMM/yyyy")
                Cambios = True
                Exit Function
            Case Trim(Me.dbcOrigen.Text) <> Trim(Me.dbcOrigen.Tag)
                Cambios = True
                Exit Function
            Case Trim(Me.dbcGrupo.Text) <> Trim(Me.dbcGrupo.Tag)
                Cambios = True
                Exit Function
            Case ModEstandar.Numerico((Me.txtCostoAdicional.Text)) <> ModEstandar.Numerico((Me.txtCostoAdicional.Tag))
                Cambios = True
                Exit Function
            Case ModEstandar.Numerico((Me.txtCostosIndirectos.Text)) <> ModEstandar.Numerico((Me.txtCostosIndirectos.Tag))
                Cambios = True
                Exit Function
            Case Trim(Me.rtEntregaren.Text) <> Trim(Me.rtEntregaren.Tag)
                Cambios = True
                Exit Function
            Case Numerico((Me.txtTipoCambioConciliado.Text)) <> Numerico((Me.txtTipoCambioConciliado.Tag))
                Cambios = True
                Exit Function
            Case Numerico((Me.txtTipoCambioEuroConciliado.Text)) <> Numerico((Me.txtTipoCambioEuroConciliado.Tag))
                Cambios = True
                Exit Function
                '        Case ModEstandar.Numerico(Me.txtSubTotal.text) <> ModEstandar.Numerico(Me.txtSubTotal.Tag)
                '            Cambios = True
                '            Exit Function
                '        Case ModEstandar.Numerico(Me.txtDescuento.text) <> ModEstandar.Numerico(Me.txtDescuento.Tag)
                '            Cambios = True
                '            Exit Function
                '        Case ModEstandar.Numerico(Me.txtIVA.text) <> ModEstandar.Numerico(Me.txtIVA.Tag)
                '            Cambios = True
                '            Exit Function
                '        Case ModEstandar.Numerico(Me.txtTotal.text) <> ModEstandar.Numerico(Me.txtTotal.Tag)
                '            Cambios = True
                '            Exit Function
        End Select
        'Ver si hay cambios en el grid
        With Me.mshFlex
            For I = 1 To .Rows - 1
                If I = 11 Then
                    I = 11
                End If
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then
                    Exit For
                End If
                Select Case True
                    Case Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) <> Trim(.get_TextMatrix(I, C_ColDESCRIPCIONTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_COLUNIDAD)) <> Trim(.get_TextMatrix(I, C_COLUNIDADTAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLCANTIDAD)) <> Numerico(.get_TextMatrix(I, C_COLCANTIDADTAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIO)) <> Numerico(.get_TextMatrix(I, C_COLPRECIOUNITARIOTAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLDESCTO)) <> Numerico(.get_TextMatrix(I, C_COLDESCTOTAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLIVA)) <> Numerico(.get_TextMatrix(I, C_COLIVATAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLCOSTOADICIONAL)) <> Numerico(.get_TextMatrix(I, C_COLCOSTOADICIONALTAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLCOSTOINDIRECTOS)) <> Numerico(.get_TextMatrix(I, C_COLCOSTOINDIRECTOSTAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_ColCODGRUPO)) <> Numerico(.get_TextMatrix(I, C_COLCODGRUPOTAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLCODFAMILIA)) <> Numerico(.get_TextMatrix(I, C_COLCODFAMILIATAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLCODLINEA)) <> Numerico(.get_TextMatrix(I, C_COLCODLINEATAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLCODSUBLINEA)) <> Numerico(.get_TextMatrix(I, C_COLCODSUBLINEATAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLCODKILATES)) <> Numerico(.get_TextMatrix(I, C_COLCODKILATESTAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLCODMARCA)) <> Numerico(.get_TextMatrix(I, C_COLCODMARCATAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLCODMODELO)) <> Numerico(.get_TextMatrix(I, C_COLCODMODELOTAG))
                        Cambios = True
                        Exit Function
                    Case Numerico(.get_TextMatrix(I, C_COLCODTIPOMATERIAL)) <> Numerico(.get_TextMatrix(I, C_COLCODTIPOMATERIALTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_COLGENERO)) <> Trim(.get_TextMatrix(I, C_COLGENEROTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_COLMOVIMIENTO)) <> Trim(.get_TextMatrix(I, C_COLMOVIMIENTOTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_COLCRONO)) <> Trim(.get_TextMatrix(I, C_COLCRONOTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROV)) <> Trim(.get_TextMatrix(I, C_COLCODIGOARTICULOPROVTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_ColSTATUS)) <> Trim(.get_TextMatrix(I, C_COLSTATUSTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_COLDESCTOPORC)) <> Trim(.get_TextMatrix(I, C_COLDESCTOPORCTAG))
                        Cambios = True
                        Exit Function


                    Case Trim(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR)) <> Trim(.get_TextMatrix(I, C_COLPRECIOPUBDOLAR))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_COLADICIONAL)) <> Trim(.get_TextMatrix(I, C_COLADICIONALTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_COLMONEDAPP)) <> Trim(.get_TextMatrix(I, C_COLMONEDAPPTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_COLORIGENANT)) <> Trim(.get_TextMatrix(I, C_COLORIGENANTTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_ColCODIGOANT)) <> Trim(.get_TextMatrix(I, C_ColCODIGOANTTAG))
                        Cambios = True
                        Exit Function
                    Case Trim(.get_TextMatrix(I, C_ColIMAGEN)) <> Trim(.get_TextMatrix(I, C_ColIMAGENTAG))
                        Cambios = True
                        Exit Function

                End Select
            Next I
        End With
        Cambios = False
    End Function

    Public Sub Eliminar()
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        gStrSql = "select FolioOrdenCompra, Estatus from OrdenesCompra where FolioOrdenCompra = '" & Trim(Me.txtFolio.Text) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            MsgBox("Proporcione un Folio válido para eliminar la Orden de Compra", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If
        'Sólo puede eliminar aquellas que NO han sido generadas o canceladas
        If RsGral.Fields("Estatus").Value = C_STGENERADA Then
            MsgBox("No puede eliminar esta orden de compra debido a que ya ha sido GENERADA", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If RsGral.Fields("Estatus").Value = C_STCANCELADA Then
            MsgBox("No puede eliminar esta orden de compra debido a que ya ha sido CANCELADA", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If RsGral.Fields("Estatus").Value = C_STREGISTRADA Then
            MsgBox("No puede cancelar esta orden de compra debido a que ya ha sido REGISTRADA en una Factura", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        'Preguntar si desea borrar el registro
        If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
            Exit Sub
        End If
        Cnn.BeginTrans()
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        blnTransaction = True
        ModStoredProcedures.PR_IMEOrdenesCompra(Trim(Me.txtFolio.Text), "", VB6.Format(Me.dtpFecha.Value, C_FORMATFECHAGUARDAR), VB6.Format(Me.dtpFechaEntrega.Value, C_FORMATFECHAGUARDAR), "", "", "", "", "0.00", "0.00", "", "0.00", "0.00", "0.00", "0.00", Trim(cMonedadeCompra), "", VB6.Format(Today, C_FORMATFECHAGUARDAR), Trim(Me.txtTasaIva.Text), Trim(Me.txtPorcDescto.Text), Trim(Me.txtTipoCambio.Text), Trim(Me.txtTipoCambioEuro.Text), Trim(Me.txtTipoCambioConciliado.Text), Trim(Me.txtTipoCambioEuroConciliado.Text), Trim(Me.txtDesctoFinanciero.Text), VB6.Format(Today, C_FORMATFECHAGUARDAR), "", C_ELIMINACION, CStr(0))
        Cmd.Execute()
        Cnn.CommitTrans()
        blnTransaction = False
        Limpiar()
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
Merr:
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Sub EliminarFlex()
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        Dim TopRowAnterior As Object
        Dim RowAnterior As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object TopRowAnterior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TopRowAnterior = Me.mshFlex.TopRow
        RowAnterior = Me.mshFlex.Row

        If Trim(cESTATUSORDEN) <> "" And Trim(mshFlex.get_TextMatrix(mshFlex.Row, C_ColSTATUS)) <> "" Then
            MsgBox("No puede eliminar artículos de la Orden de Compra" & vbNewLine & "ya que ésta ha sido registrada en el inventario...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            mshFlex.Focus()
            Exit Sub
        End If

        If Trim(Me.mshFlex.get_TextMatrix(mshFlex.Row, C_COLCODAUX)) = "" Then
            If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
                mshFlex.Focus()
                Exit Sub
            End If
            mshFlex.RemoveItem((mshFlex.Row))
            ActualizaCantidades()
            mshFlex_EnterCell(mshFlex, New System.EventArgs())
            mshFlex.Focus()
            Exit Sub
        End If
        If mshFlex.get_TextMatrix(mshFlex.Row, C_ColSTATUS) <> C_CONCILIADO And Me.mshFlex.get_TextMatrix(mshFlex.Row, C_ColSTATUS) <> C_CR Then
            If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
                Exit Sub
            End If
            Cnn.BeginTrans()
            blnTransaction = True
            With Me.mshFlex
                ModStoredProcedures.PR_IMEOrdenesCompraPreCatAux(Trim(Me.txtFolio.Text), Trim(.get_TextMatrix(.Row, C_COLCODAUX)), Trim(.get_TextMatrix(.Row, C_COLCODIGO)), Trim(.get_TextMatrix(.Row, C_COLDESCRIPCION)), Trim(.get_TextMatrix(.Row, C_COLCANTIDAD)), Trim(.get_TextMatrix(.Row, C_COLCANTIDAD)), Trim(.get_TextMatrix(.Row, C_COLPRECIOUNITARIO)), Trim(.get_TextMatrix(.Row, C_COLCOSTOCUR)), Trim(.get_TextMatrix(.Row, C_COLDESCUENTOCUR)), Trim(.get_TextMatrix(.Row, C_COLDESCTOPORC)), Trim(.get_TextMatrix(.Row, C_COLIVACUR)), Trim(.get_TextMatrix(.Row, C_COLCOSTOADICIONALCUR)), Trim(.get_TextMatrix(.Row, C_COLCOSTOINDIRECTOSCUR)), CStr(mintCodGrupo), Trim(.get_TextMatrix(.Row, C_COLCODFAMILIA)), Trim(.get_TextMatrix(.Row, C_COLCODLINEA)), Trim(.get_TextMatrix(.Row, C_COLCODSUBLINEA)), Trim(.get_TextMatrix(.Row, C_COLCODKILATES)), Trim(.get_TextMatrix(.Row, C_COLCODMARCA)), Trim(.get_TextMatrix(.Row, C_COLCODMODELO)), Trim(.get_TextMatrix(.Row, C_COLCODTIPOMATERIAL)), Trim(.get_TextMatrix(.Row, C_COLGENERO)), Trim(.get_TextMatrix(.Row, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(.Row, C_COLCRONO)), CStr(BuscaCodUnidad(Trim(.get_TextMatrix(.Row, C_COLUNIDAD)))), CStr(mintCodOrigen), CStr(mintCodProveedor), Trim(.get_TextMatrix(.Row, C_COLCODIGOARTICULOPROV)), "", .get_TextMatrix(.Row, C_COLADICIONAL), .get_TextMatrix(.Row, C_COLPRECIOPUBDOLAR), .get_TextMatrix(.Row, C_COLMONEDAPP), .get_TextMatrix(.Row, C_COLORIGENANT), .get_TextMatrix(.Row, C_ColCODIGOANT), .get_TextMatrix(.Row, C_ColIMAGEN), CStr(0), "", "", "", C_ELIMINACION, CStr(0))
                Cmd.Execute()
            End With
            Cnn.CommitTrans()
            blnTransaction = False
            mshFlex.RemoveItem((mshFlex.Row))
        Else
            'Si ya se ha conciliado, no debe borrarse
            MsgBox("No puede borrar este Artículo debido a que ya se ha conciliado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        ActualizaCantidades()
        'UPGRADE_WARNING: Couldn't resolve default property of object TopRowAnterior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mshFlex.TopRow = TopRowAnterior
        mshFlex.Row = RowAnterior

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
        If blnTransaction Then Cnn.RollbackTrans()
    End Sub

    Public Sub Limpiar()
        On Error Resume Next
        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                Case MsgBoxResult.Cancel 'Cancela la acción de limpiar pantalla
                    Exit Sub
            End Select
        End If
        Me.txtFolio.Text = ""
        Nuevo()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        Me.txtFolio.Focus()
    End Sub

    Public Sub Nuevo()
        On Error Resume Next
        If mblnNuevo Then
            Me.txtFolio.Text = ""
            Me.txtFolio.Tag = ""
        End If
        Me.txtTasaIva.Text = VB6.Format(gcurCorpoTASAIVA, "##0.00")
        Me.txtTasaIva.Tag = Me.txtTasaIva.Text
        Me.txtTipoCambio.Text = VB6.Format(gcurCorpoTIPOCAMBIODOLAR, "###,###,##0.00")
        Me.txtTipoCambio.Tag = Me.txtTipoCambio.Text
        Me.txtTipoCambioEuro.Text = VB6.Format(gcurCorpoTIPOCAMBIOEURO, "###,###,##0.00")
        Me.txtTipoCambioEuro.Tag = Me.txtTipoCambioEuro.Text
        Me.txtPorcDescto.Text = "0.00"
        Me.txtPorcDescto.Tag = Me.txtPorcDescto.Text
        Me.txtDesctoFinanciero.Text = "0.00"
        Me.txtDesctoFinanciero.Tag = Me.txtDesctoFinanciero.Text
        Me.txtOtrosDatos.Text = ""
        Me.txtOtrosDatos.Tag = ""
        Me.optMoneda(0).Checked = True
        Me.optMoneda(1).Checked = False
        Me.optMoneda(2).Checked = False

        Me.txtTipoCambioConciliado.Text = VB6.Format(gcurCorpoTIPOCAMBIODOLAR, "###,###,##0.00")
        Me.txtTipoCambioConciliado.Tag = Me.txtTipoCambioConciliado.Text
        Me.txtTipoCambioEuroConciliado.Text = VB6.Format(gcurCorpoTIPOCAMBIOEURO, "###,###,##0.00")
        Me.txtTipoCambioEuroConciliado.Tag = Me.txtTipoCambioEuroConciliado.Text
        mblnFueraChange = True
        mintCodProveedor = 0
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Me.dbcProveedor.Text = ""
        Me.dbcProveedor.Tag = ""
        mblnFueraChange = False

        cMonedadeCompra = C_DOLAR
        cMonedadeCompraTag = C_DOLAR
        Me.txtRemision.Text = ""
        Me.txtRemision.Tag = ""
        Me.txtPedido.Text = ""
        Me.txtPedido.Tag = ""
        Me.txtOtrosDatos.Text = ""
        Me.txtOtrosDatos.Tag = ""
        Me.dtpFecha.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.dtpFecha.Tag = VB6.Format(Today, "dd/MMM/yyyy")
        Me.dtpFechaEntrega.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.dtpFechaEntrega.Tag = VB6.Format(Today, "dd/MMM/yyyy")

        mblnFueraChange = True
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Me.dbcOrigen.Text = ""
        Me.dbcOrigen.Tag = ""
        mintCodOrigen = 0
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Me.dbcGrupo.Text = ""
        Me.dbcGrupo.Tag = ""
        mintCodGrupo = 0
        mblnFueraChange = False

        Me.txtCostoAdicional.Text = "0.00"
        Me.txtCostoAdicional.Tag = "0.00"
        Me.txtCostosIndirectos.Text = "0.00"
        Me.txtCostosIndirectos.Tag = "0.00"
        Encabezado()

        Me.mshFlex.TopRow = 1
        Me.mshFlex.Col = C_COLCODIGO
        Me.mshFlex.Row = 1

        Me.txtDescripcion.Text = ""
        Me.lblDescProv.Text = ""
        Me.txtDescripcion.Tag = ""
        Me.rtEntregaren.Text = ""
        Me.rtEntregaren.Tag = ""
        Me.txtSubTotal.Text = "0.00"
        Me.txtSubTotal.Tag = "0.00"
        Me.txtDescuento.Text = "0.00"
        Me.txtDescuento.Tag = "0.00"
        Me.txtIVA.Text = "0.00"
        Me.txtIVA.Tag = "0.00"
        Me.txtTotal.Text = "0.00"
        Me.txtTotal.Tag = "0.00"

        mcurSubTotal = 0
        mcurDESCUENTO = 0
        mcurIVA = 0
        mcurTotal = 0

        'El estatus de la Orden
        cESTATUSORDEN = ""
        Me.lblEstatus.Text = ""
        Me.lblEstatus.Visible = False
        fraApartado.Visible = False

        'Habilitar y deshabilitar controles al limpiar
        Me.dbcProveedor.Text = False
        Me.optMoneda(0).Enabled = True
        Me.optMoneda(1).Enabled = True
        Me.optMoneda(2).Enabled = True
        Me.txtTipoCambio.ReadOnly = False
        Me.txtTipoCambioEuro.ReadOnly = False
        Me.txtTipoCambioConciliado.Enabled = True
        Me.txtTipoCambioEuroConciliado.Enabled = True
        Me.txtTasaIva.ReadOnly = False
        Me.txtPorcDescto.ReadOnly = False
        Me.txtDesctoFinanciero.ReadOnly = False
        Me.txtRemision.ReadOnly = False
        Me.txtPedido.ReadOnly = False
        Me.dtpFecha.Enabled = True
        Me.dtpFechaEntrega.Enabled = True
        Me.dbcOrigen.Text = False
        Me.dbcGrupo.Text = False
        Me.txtCostoAdicional.ReadOnly = False
        Me.txtCostosIndirectos.ReadOnly = False
        Me.rtEntregaren.ReadOnly = False
        Me.btnAsignarCodigos.Enabled = True
        Me.btnProv.Enabled = True
        UltimaMoneda = "D"
        gstrNombreForma = "FRMCXPORDENCOMPRA"
        mintRenglonAnt = 0
        mintRenglonAct = 0
        mintRenglonSig = 0
        mintTotalPartidasCapt = 0
        ResBusquedaArt = 0
        'Desactivamos el boton de imprimir
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Public Function ValidaDatos() As Boolean
        Dim I As Integer
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Select Case True
            Case mintCodProveedor = 0
                MsgBox("Debe Indicar el proveedor", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dbcProveedor.Focus()
                ModEstandar.SelTxt()
                Exit Function
            Case Trim(Me.dbcOrigen.Text) = ""
                MsgBox("Debe indicar el almacén de origen", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dbcOrigen.Focus()
                ModEstandar.SelTxt()
                Exit Function
            Case mintCodGrupo = 0
                MsgBox("Debe indicar el Grupo de Artículos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dbcGrupo.Focus()
                ModEstandar.SelTxt()
                Exit Function
        End Select
        ValidaDatos = True
    End Function

    Public Function ValidaFlex() As Boolean
        Dim I As Integer
        'Ver si los datos en el Grid son válidos
        With Me.mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLDESCRIPCION)) = "" Then
                    If I = 1 Then
                        MsgBox("Para Guardar una Nueva Orden de Compra, ésta debe contener, por lo menos, un artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                        ValidaFlex = False
                        .Row = I
                        .Col = C_COLCANTIDAD
                        .Focus()
                        Exit Function
                    End If
                    Exit For
                End If
                Select Case True
                    'Case Numerico(.TextMatrix(i, C_COLCANTIDAD)) = 0
                    '    MsgBox "Introduzca la cantidad de artículos a pedir", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                    '    ValidaFlex = False
                    '    .Row = i
                    '    .Col = C_COLCANTIDAD
                    '    .SetFocus
                    '    Exit Function
                    '                Case Numerico(.TextMatrix(I, C_COLPRECIOUNITARIO)) = 0
                    '                    MsgBox "Introduzca el precio de cada artículo de esta categoría", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                    '                    ValidaFlex = False
                    '                    .Row = I
                    '                    .Col = C_COLPRECIOUNITARIO
                    '                    .SetFocus
                    '                    Exit Function
                End Select
            Next I
        End With
        ValidaFlex = True
    End Function

    Sub Encabezado()
        Dim LnContador As Integer

        With mshFlex
            If Not mblnLoad Then
                .Rows = 2
                .Rows = 12
                .set_Cols(0, 87) '''27OCT2010 - MAVF - antes 83
                .RemoveItem((1))
                Exit Sub
            End If
            .Height = VB6.TwipsToPixelsY(2050)
            .set_Cols(0, 87) '''27OCT2010 - MAVF - antes 83
            .Clear()

            For LnContador = 0 To (.get_Cols() - 1) Step 1
                .Col = LnContador
                .set_ColWidth(.Col, 0, 0)
            Next LnContador

            .set_ColWidth(C_COLCODIGO, 0, 1140)
            .set_ColWidth(C_COLDESCRIPCION, 0, 4185)
            .set_ColAlignment(C_COLDESCRIPCION, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColWidth(C_COLUNIDAD, 0, 690)
            .set_ColWidth(C_COLCANTIDAD, 0, 1354)
            .set_ColWidth(C_COLPRECIOUNITARIO, 0, 1354)
            .set_ColWidth(C_COLCOSTO, 0, 0) '1354
            .set_ColWidth(C_COLDESCTO, 0, 1354)
            .set_ColWidth(C_COLIVA, 0, 1354)
            .set_ColWidth(C_COLCODAUX, 0, 0)
            .set_ColWidth(C_ColSTATUS, 0, 0)

            .set_ColWidth(C_ColDESCRIPCIONTAG, 0, 0)
            .set_ColWidth(C_COLUNIDADTAG, 0, 0)
            .set_ColWidth(C_COLCANTIDADTAG, 0, 0)
            .set_ColWidth(C_COLPRECIOUNITARIOTAG, 0, 0)
            .set_ColWidth(C_COLCOSTOTAG, 0, 0)
            .set_ColWidth(C_COLDESCTOTAG, 0, 0)
            .set_ColWidth(C_COLIVATAG, 0, 0)

            .set_ColWidth(C_COLCOSTOADICIONAL, 0, 0)
            .set_ColWidth(C_COLCOSTOINDIRECTOS, 0, 0)
            .set_ColWidth(C_ColCODGRUPO, 0, 0)
            .set_ColWidth(C_COLCODFAMILIA, 0, 0)
            .set_ColWidth(C_COLCODLINEA, 0, 0)
            .set_ColWidth(C_COLCODSUBLINEA, 0, 0)
            .set_ColWidth(C_COLCODKILATES, 0, 0)
            .set_ColWidth(C_COLCODMARCA, 0, 0)
            .set_ColWidth(C_COLCODMODELO, 0, 0)
            .set_ColWidth(C_COLCODTIPOMATERIAL, 0, 0)
            .set_ColWidth(C_COLGENERO, 0, 0)
            .set_ColWidth(C_COLMOVIMIENTO, 0, 0)
            .set_ColWidth(C_COLCRONO, 0, 0)
            .set_ColWidth(C_COLCODIGOARTICULOPROV, 0, 0)

            .set_ColWidth(C_COLCOSTOADICIONALTAG, 0, 0)
            .set_ColWidth(C_COLCOSTOINDIRECTOSTAG, 0, 0)
            .set_ColWidth(C_COLCODGRUPOTAG, 0, 0)
            .set_ColWidth(C_COLCODFAMILIATAG, 0, 0)
            .set_ColWidth(C_COLCODLINEATAG, 0, 0)
            .set_ColWidth(C_COLCODSUBLINEATAG, 0, 0)
            .set_ColWidth(C_COLCODKILATESTAG, 0, 0)
            .set_ColWidth(C_COLCODMARCATAG, 0, 0)
            .set_ColWidth(C_COLCODMODELOTAG, 0, 0)
            .set_ColWidth(C_COLCODTIPOMATERIALTAG, 0, 0)
            .set_ColWidth(C_COLGENEROTAG, 0, 0)
            .set_ColWidth(C_COLMOVIMIENTOTAG, 0, 0)
            .set_ColWidth(C_COLCRONOTAG, 0, 0)
            .set_ColWidth(C_COLCODIGOARTICULOPROVTAG, 0, 0)

            .set_ColWidth(C_COLCOSTOFACTURA, 0, 0)
            .set_ColWidth(C_COLPORCFACTURA, 0, 0)
            .set_ColWidth(C_COLPORCIVA, 0, 0)
            .set_ColWidth(C_COLDESCTOPORC, 0, 0)
            .set_ColWidth(C_COLDESCTOPORCTAG, 0, 0)

            .set_ColWidth(C_COLCOSTOCUR, 0, 0)
            .set_ColWidth(C_COLDESCUENTOCUR, 0, 0)
            .set_ColWidth(C_COLIVACUR, 0, 0)
            .set_ColWidth(C_COLCOSTOADICIONALCUR, 0, 0)
            .set_ColWidth(C_COLCOSTOINDIRECTOSCUR, 0, 0)

            .set_ColWidth(C_COLSTATUSTAG, 0, 0)
            .set_ColWidth(C_COLPRECIOUNITARIO4DEC, 0, 0)
            .set_ColWidth(C_ColIMPORTE, 0, 1400)
            .set_ColWidth(C_COLADICIONAL, 0, 0)

            .set_ColWidth(C_COLPRECIOPUBDOLAR, 0, 0)
            .set_ColWidth(C_COLMONEDAPP, 0, 0)
            .set_ColWidth(C_COLORIGENANT, 0, 0)
            .set_ColWidth(C_ColCODIGOANT, 0, 0)
            .set_ColWidth(C_ColIMAGEN, 0, 0)
            .set_ColWidth(C_COLSTATUSX, 0, 0)

            .set_ColWidth(C_ColMDSPESO, 0, 0) '''27OCT2010 - MAVF
            .set_ColWidth(C_ColMDSCOLOR, 0, 0) '''27OCT2010 - MAVF
            .set_ColWidth(C_ColMDSPUREZA, 0, 0) '''27OCT2010 - MAVF
            .set_ColWidth(C_ColMDSCERTIFICADO, 0, 0) '''27OCT2010 - MAVF

            .set_TextMatrix(C_RENENCABEZADO, C_COLCODIGO, "Código")
            .set_TextMatrix(C_RENENCABEZADO, C_COLDESCRIPCION, "Descripción")
            .set_TextMatrix(C_RENENCABEZADO, C_COLUNIDAD, "Unidad")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCANTIDAD, "Cantidad")
            .set_TextMatrix(C_RENENCABEZADO, C_COLPRECIOUNITARIO, "Costo Unitario")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCOSTO, "Costo")
            .set_TextMatrix(C_RENENCABEZADO, C_COLDESCTO, "Descuento")
            .set_TextMatrix(C_RENENCABEZADO, C_COLIVA, "IVA")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODAUX, "Código Auxiliar")
            .set_TextMatrix(C_RENENCABEZADO, C_ColSTATUS, "STATUS")

            .set_TextMatrix(C_RENENCABEZADO, C_ColDESCRIPCIONTAG, "Descripción Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLUNIDADTAG, "Unidad Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCANTIDADTAG, "Cantidad Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCOSTOTAG, "Costo Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLDESCTOTAG, "Descuento Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLIVATAG, "IVA Tag")

            .set_TextMatrix(C_RENENCABEZADO, C_COLCOSTOADICIONAL, "Costo Adicional")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCOSTOINDIRECTOS, "Costos Indirectos")
            .set_TextMatrix(C_RENENCABEZADO, C_ColCODGRUPO, "Grupo")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODFAMILIA, "Familia")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODLINEA, "Linea")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODSUBLINEA, "SubLinea")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODMARCA, "Marca")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODMODELO, "Modelo")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODTIPOMATERIAL, "Tipo Material")
            .set_TextMatrix(C_RENENCABEZADO, C_COLGENERO, "Género")
            .set_TextMatrix(C_RENENCABEZADO, C_COLMOVIMIENTO, "Movimiento")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODIGOARTICULOPROV, "CodArticuloProv")

            .set_TextMatrix(C_RENENCABEZADO, C_COLCOSTOADICIONALTAG, "Costo Adicional Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCOSTOINDIRECTOSTAG, "Costos Indirectos Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODGRUPOTAG, "Grupo Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODFAMILIATAG, "Familia Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODLINEATAG, "Linea Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODSUBLINEATAG, "SubLinea Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODMARCATAG, "Marca Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODMODELOTAG, "Modelo Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODTIPOMATERIALTAG, "Tipo Material Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLGENEROTAG, "Género Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLMOVIMIENTOTAG, "Movimiento Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODIGOARTICULOPROVTAG, "CodArticuloProv Tag")

            .set_TextMatrix(C_RENENCABEZADO, C_COLCOSTOFACTURA, "Costo Factura")
            .set_TextMatrix(C_RENENCABEZADO, C_COLPORCFACTURA, "Porcentaje de Costo Factura")
            .set_TextMatrix(C_RENENCABEZADO, C_COLPORCIVA, "Porcentaje de IVA")
            .set_TextMatrix(C_RENENCABEZADO, C_COLDESCTOPORC, "Porcentaje de descuento")
            .set_TextMatrix(C_RENENCABEZADO, C_COLDESCTOPORCTAG, "Porcentaje de descuento Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLSTATUSTAG, "Status Tag")
            .set_TextMatrix(C_RENENCABEZADO, C_ColIMPORTE, "Importe")

            .set_TextMatrix(C_RENENCABEZADO, C_COLADICIONAL, "Adicional")
            .set_TextMatrix(C_RENENCABEZADO, C_COLPRECIOPUBDOLAR, "PrecioPubDolar")
            .set_TextMatrix(C_RENENCABEZADO, C_COLMONEDAPP, "MonedaPP")
            .set_TextMatrix(C_RENENCABEZADO, C_COLORIGENANT, "OrigenAnt")
            .set_TextMatrix(C_RENENCABEZADO, C_ColCODIGOANT, "CodigoAnt")
            .set_TextMatrix(C_RENENCABEZADO, C_ColIMAGEN, "Imagen")
            .set_TextMatrix(C_RENENCABEZADO, C_COLSTATUSX, "STATUSX")

            'Colocar los textos de los encabezados centrados
            .Row = C_RENENCABEZADO
            For LnContador = 0 To (.get_Cols() - 1) Step 1
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = False
            Next LnContador
        End With
    End Sub

    Private Sub btnAsignarCodigos_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAsignarCodigos.Click
        If Cambios() Then
            mblnConciliar = True
            If Guardar() Then
                mblnConciliar = False
            Else
                mblnConciliar = False
                Exit Sub
            End If
        End If
        Conciliar()
        mblnConciliar = False
    End Sub

    Private Sub btnAsignarCodigos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAsignarCodigos.Enter
        Pon_Tool()
    End Sub

    Private Sub btnProv_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnProv.Click
        frmCorpoAbcProvAcreed.Show()
    End Sub

    Private Sub dbcGrupo_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcGrupo.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        lStrSql = "SELECT codGrupo, rtrim(ltrim(descGrupo)) as descGrupo FROM catGrupos Where descGrupo LIKE '" & Trim(Me.dbcGrupo.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcGrupo)

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcGrupo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcGrupo.Enter
        Pon_Tool()
        gStrSql = "SELECT codGrupo, rtrim(ltrim(descGrupo)) as descGrupo FROM catGrupos ORDER BY DescGrupo"
        ModDCombo.DCGotFocus(gStrSql, dbcGrupo)
    End Sub

    Private Sub dbcGrupo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcGrupo.KeyDown
        Dim Aux As String
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Me.dbcOrigen.Focus()
                eventSender.KeyCode = 0
                '        Case vbKeyReturn
                '            Aux = Trim(Me.dbcGrupo.text)
                '            If dbcGrupo.SelectedItem <> 0 Then
                '                dbcGrupo_LostFocus
                '            End If
                '            'Me.dbcGrupo.text = Aux
                '            Exit Sub
                '        Case vbKeyTab
                '            Aux = Trim(Me.dbcGrupo.text)
                '            If dbcGrupo.SelectedItem <> 0 Then
                '                Me.dbcGrupo.text = Me.dbcGrupo.SelectedItem
                '                dbcGrupo_LostFocus
                '            End If
                '            'Me.dbcGrupo.text = Aux
                '            Exit Sub
        End Select
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcGrupo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcGrupo.Leave
        Dim I As Integer
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        gStrSql = "SELECT codGrupo, rtrim(ltrim(descGrupo)) as descGrupo FROM catGrupos Where descGrupo LIKE '" & Trim(Me.dbcGrupo.Text) & "%'"
        Aux = mintCodGrupo
        mintCodGrupo = 0
        ModDCombo.DCLostFocus((Me.dbcGrupo), gStrSql, mintCodGrupo)

        If Aux <> mintCodGrupo And Trim(Me.mshFlex.get_TextMatrix(1, C_COLDESCRIPCION)) <> "" Then
            If MsgBox("Si cambia de Grupo, se borrarán todos los artículos que haya registrado." & vbNewLine & "¿Desea continuar?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, "Advertencia ...") = MsgBoxResult.Yes Then
                Encabezado()
                Me.dbcGrupo.Focus()
                ModEstandar.SelTxt()
            Else
                mintCodGrupo = Aux
                mblnFueraChange = True
                'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
                Me.dbcGrupo.Text = BuscaGrupo(mintCodGrupo)
                mblnFueraChange = False
                Me.dbcGrupo.Focus()
                ModEstandar.SelTxt()
            End If
        End If
    End Sub

    Private Sub dbcGrupo_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcGrupo.MouseUp
        Dim Aux As String
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Aux = Trim(Me.dbcGrupo.Text)
        If dbcGrupo.SelectedItem <> 0 Then
            dbcGrupo_Leave(dbcGrupo, New System.EventArgs())
        End If
        'Me.dbcGrupo.text = Aux
    End Sub

    Private Sub dbcOrigen_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        lStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me.dbcOrigen.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcOrigen))

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcOrigen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen ORDER BY CodAlmacenOrigen"
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcOrigen))
    End Sub

    Private Sub dbcOrigen_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcOrigen.KeyDown
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                If Me.dtpFechaEntrega.Enabled Then
                    Me.dtpFechaEntrega.Focus()
                Else
                    Me.txtPedido.Focus()
                End If
                eventSender.KeyCode = 0
        End Select
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcOrigen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen.Leave
        Dim I As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me.dbcOrigen.Text) & "%'"
        mintCodOrigen = 0
        ModDCombo.DCLostFocus((Me.dbcOrigen), gStrSql, mintCodOrigen)
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String
        If mblnFueraChange Then Exit Sub
        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcProveedor.Name Then Exit Sub
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcProveedor)
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        If Trim(Me.dbcProveedor.Text) = "" Then
            dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        End If
        If dbcProveedor.SelectedItem <> "" Then
            Call dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcProveedor.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' ORDER BY descProvAcreed"
        ModDCombo.DCGotFocus(gStrSql, dbcProveedor)
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcProveedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.txtFolio.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
        Dim I As Integer
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        Aux = mintCodProveedor
        mintCodProveedor = 0
        ModDCombo.DCLostFocus(dbcProveedor, gStrSql, mintCodProveedor)
        Me.txtOtrosDatos.Text = BuscaDatosProveedor(mintCodProveedor)
        If Aux <> mintCodProveedor Then
            Me.txtPorcDescto.Text = VB6.Format(BuscaDesctoProveedor(mintCodProveedor), "##0.00")
            Me.txtDesctoFinanciero.Text = VB6.Format(Me.BuscaDesctoFinanciero(mintCodProveedor), "##0.00")
        End If
    End Sub

    Private Sub dtpFecha_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFecha.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpFechaEntrega_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaEntrega.Enter
        Pon_Tool()
    End Sub

    'UPGRADE_WARNING: Form event frmCXPOrdenCompra.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmCXPOrdenCompra_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        If mblnNuevo Then
            ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Else
            ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        End If
        'UPGRADE_WARNING: Form method frmCXPOrdenCompra.ZOrder has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        Me.BringToFront()
    End Sub

    'UPGRADE_WARNING: Form event frmCXPOrdenCompra.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmCXPOrdenCompra_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPOrdenCompra_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Or UCase(Me.ActiveControl.Name) = "TXTFLEX" Then
                    'mshFlex_KeyPress vbKeyReturn
                Else
                    ModEstandar.AvanzarTab(Me)
                End If
            Case System.Windows.Forms.Keys.Escape
                'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    Me.txtCostosIndirectos.Focus()
                    'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                ElseIf UCase(Me.ActiveControl.Name) = "TXTFOLIO" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
            Case System.Windows.Forms.Keys.Delete
                'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    If Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_COLDESCRIPCION) <> "" Then
                        Call EliminarFlex()
                    End If
                End If
        End Select
    End Sub

    Private Sub frmCXPOrdenCompra_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.gp_CampoMayusculas(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPOrdenCompra_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        mblnLoad = True
        Encabezado()
        Me.mshFlex.Rows = 11
        Nuevo()
        mblnLoad = False
        mblnCambiosEnCodigo = False
        mblnNuevo = True
    End Sub

    Private Sub frmCXPOrdenCompra_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If Not mblnSalir Then
            'Si desea cerrar la forma y ésta se encuentra minimizada, se debe restaurar
            ModEstandar.RestaurarForma(Me, False)
            If Cambios() Then ' And Not mblnNuevo
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        If Not Guardar() Then 'Si falla el guardar, no cierra la forma
                            Cancel = 1
                        Else
                            mblnNuevo = True
                            Cancel = 0
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se cierre el formulario
                        mblnNuevo = True
                        Cancel = 0
                    Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
                        Cancel = 1
                End Select
            End If
        Else 'Se quiere salir con escape
            mblnSalir = False
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Sale del Formulario
                    Cancel = 0
                Case MsgBoxResult.No 'No sale del formulario
                    Me.txtFolio.Focus()
                    ModEstandar.SelTextoTxt((Me.txtFolio))
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCXPOrdenCompra_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        gstrNombreForma = ""
        'UPGRADE_NOTE: Object frmCXPOrdenCompra may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub mshFlex_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.DblClick
        'Poner el color correspondiente
        If Trim(cESTATUSORDEN) <> "" Then
            If Trim(cESTATUSORDEN) <> Trim(C_STVIGENTE) Then Exit Sub
        End If

        With mshFlex
            If Trim(.get_TextMatrix(.Row, C_COLDESCRIPCION)) <> "" And (.Col = C_COLCODIGO Or .Col = C_COLDESCRIPCION) Then
                If Trim(.get_TextMatrix(.Row, C_COLUNIDAD)) = "" Then
                    MsgBox("Proporcione la Unidad Para Poder Conciliar la Orden de Compra.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Sub
                ElseIf CDec(Numerico(.get_TextMatrix(.Row, C_COLCANTIDAD))) <= 0 Then
                    MsgBox("Proporcione la Cantidad de Articulos para Poder Conciliar la Orden de Compra.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Sub
                ElseIf CDec(Numerico(.get_TextMatrix(.Row, C_COLPRECIOUNITARIO))) <= 0 Then
                    MsgBox("Proporcione el Precio Unitario Para Poder Conciliar la Orden de Compra.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Sub
                Else
                    Select Case Trim(.get_TextMatrix(.Row, C_ColSTATUS))
                        Case C_CONCILIADO
                            'Si es conciliado, se quita la marca de conciliado y se pone como vacío (quitar color)
                            .set_TextMatrix(.Row, C_ColSTATUS, "")
                        Case C_RESURTIDO
                            'Si es Resurtido, se pone como Conciliado/Resurtido y fija el color
                            .set_TextMatrix(.Row, C_ColSTATUS, C_CR)
                        Case C_CR
                            'Si es Conciliado y resurtido, se pone la marca de sólo resurtido y se pone el color
                            .set_TextMatrix(.Row, C_ColSTATUS, C_RESURTIDO)
                        Case ""
                            'Se pone sólo como conciliado
                            .set_TextMatrix(.Row, C_ColSTATUS, C_CONCILIADO)
                    End Select
                    PonerColor()
                End If
            Else
                mshFlex_KeyPressEvent(mshFlex, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
            End If
        End With
        'mshFlex_KeyPress (vbKeyReturn)
    End Sub

    Private Sub mshFlex_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.EnterCell
        txtDescripcion.Text = mshFlex.get_TextMatrix(Me.mshFlex.Row, C_COLDESCRIPCION)
        lblDescProv.Text = mshFlex.get_TextMatrix(Me.mshFlex.Row, C_COLCODIGOARTICULOPROV)
        PonerColor()
    End Sub

    Private Sub mshFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.Enter
        Pon_Tool()
        mshFlex_EnterCell(mshFlex, New System.EventArgs())
        PonerColor()
    End Sub

    Private Sub mshFlex_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles mshFlex.KeyDownEvent
        If mintCodGrupo = 0 Then
            eventArgs.keyCode = 0
        End If
        With Me.mshFlex
            Select Case eventArgs.keyCode
                Case System.Windows.Forms.Keys.Left
                    '.Col = C_COLDESCRIPCION
                Case System.Windows.Forms.Keys.Right
                    '.Col = C_COLDESCRIPCION
                Case System.Windows.Forms.Keys.Down
                    '.Col = C_COLDESCRIPCION
            End Select
        End With
    End Sub

    Private Sub mshFlex_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles mshFlex.KeyPressEvent
        Dim nCol As Integer
        Dim nRow As Integer
        Dim lRenglon As Integer

        If Trim(cESTATUSORDEN) <> "" Then
            If Trim(cESTATUSORDEN) <> Trim(C_STVIGENTE) Then
                Exit Sub
            End If
        End If
        If mintCodProveedor = 0 Then
            Exit Sub
        End If
        'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        If Trim(dbcOrigen.Text) = "" Then
            Exit Sub
        End If
        If mintCodGrupo = 0 Then
            Exit Sub
        End If
        With mshFlex
            '''si ya se capturo algo entonces se edita el grid
            '''ya sea con numeros, letras o enter
            If eventArgs.keyAscii = 13 Then
                If mintCodGrupo = 0 Then
                    MsgBox("Debe seleccionar un Grupo de Artículos antes de Guardar a la Orden de Compra", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    eventArgs.keyAscii = 0
                    dbcGrupo.Focus()
                    ModEstandar.SelTxt()
                    Exit Sub
                End If
                nCol = .Col
                nRow = .Row
                '''en esta parte se validará si es el rengón, columna que le
                '''corresponde editarse
                If (.Row > 1) Then
                    '''de tal modo que si el renglón es mayor que 1
                    '''y si un renglón antes del renglón actual está vacío,
                    '''el renglón actual no se editará
                    If Trim(.get_TextMatrix(.Row - 1, C_COLDESCRIPCION)) = "" Then
                        .Focus()
                        Exit Sub
                    End If
                End If
                If .Col = C_COLCODIGO Or .Col = C_COLDESCRIPCION Then
                    If Trim(.get_TextMatrix(.Row - 1, C_COLDESCRIPCION)) <> "" And Trim(.get_TextMatrix(.Row, C_COLDESCRIPCION)) = "" Then
                        .Rows = .Rows + 1
                    End If
                End If
                Select Case .Col
                    Case C_COLCODIGO
                        If .get_TextMatrix(.Row, C_COLDESCRIPCION) = "" Then
                            txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            txtFlex.BackColor = .CellBackColor
                            ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, eventArgs.keyAscii)
                            ModEstandar.SelTextoTxt(txtFlex)
                        Else
                            Exit Sub
                        End If
                    Case C_COLDESCRIPCION
                        '''vars para mostrar el renglon anterior de clasificacion
                        If .Row = 1 Then mintRenglonAnt = .Row Else mintRenglonAnt = .Row - 1
                        mintRenglonAct = .Row
                        '''Renglon
                        If Trim(.get_TextMatrix(mintRenglonAct, C_COLDESCRIPCION)) = "" Then
                            lRenglon = mintRenglonAnt
                        Else
                            lRenglon = mintRenglonAct
                        End If

                        'Mandar llamar la pantalla de captura de Artículos
                        Select Case mintCodGrupo
                            Case gCODJOYERIA
                                If txtFolioApartado.Text <> "" And Trim(.get_TextMatrix(mintRenglonAct, C_COLDESCRIPCION)) <> "" Then '''apartados x catalogo - 1ra vez
                                    If CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODFAMILIAX))) = 0 And CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODLINEAX))) = 0 And CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODSUBLINEAX))) = 0 And CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODKILATESX))) = 0 And CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODTIPOMATERIALX))) = 0 And Trim(.get_TextMatrix(mintRenglonAct, C_COLADICIONALX)) = "" Then
                                        If (Trim(.get_TextMatrix(mintRenglonAct, C_COLCODIGO)) = "" Or Trim(.get_TextMatrix(mintRenglonAct, C_COLCODAUX)) = "") Then
                                            '''27OCT2010 - MAVF
                                            frmCXPJoyeria.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(Trim(.get_TextMatrix(lRenglon, C_COLUNIDAD))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODFAMILIA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODLINEA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODSUBLINEA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODKILATES))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)), CDec(Trim(CStr(ModEstandar.Numerico(.get_TextMatrix(lRenglon, C_ColMDSPESO))))), Trim(.get_TextMatrix(lRenglon, C_ColMDSCOLOR)), Trim(.get_TextMatrix(lRenglon, C_ColMDSPUREZA)), Trim(.get_TextMatrix(lRenglon, C_ColMDSCERTIFICADO)))
                                        Else '''Apartados x Cat -> RESURTIDOS
                                            If Me.mshFlex.get_TextMatrix(mintRenglonAct, C_COLSTATUSX) = "M" Then
                                                '''27OCT2010 - MAVF
                                                frmCXPJoyeria.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(Trim(.get_TextMatrix(lRenglon, C_COLUNIDAD))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODFAMILIA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODLINEA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODSUBLINEA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODKILATES))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)), CDec(Trim(CStr(ModEstandar.Numerico(.get_TextMatrix(lRenglon, C_ColMDSPESO))))), Trim(.get_TextMatrix(lRenglon, C_ColMDSCOLOR)), Trim(.get_TextMatrix(lRenglon, C_ColMDSPUREZA)), Trim(.get_TextMatrix(lRenglon, C_ColMDSCERTIFICADO)))
                                            Else '''entra la primera vez que trata de clasificar correctamente el articulo originado por el apartado x cat
                                                '''27OCT2010 - MAVF
                                                frmCXPJoyeria.LlenaDatos(Trim(txtFolio.Text), .get_TextMatrix(.Row, C_COLCODAUX), nCol, lRenglon, mintRenglonAct, CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)), "1C", CDec(Trim(CStr(ModEstandar.Numerico(.get_TextMatrix(lRenglon, C_ColMDSPESO))))), Trim(.get_TextMatrix(lRenglon, C_ColMDSCOLOR)), Trim(.get_TextMatrix(lRenglon, C_ColMDSPUREZA)), Trim(.get_TextMatrix(lRenglon, C_ColMDSCERTIFICADO)))
                                            End If
                                        End If
                                    Else '''2da vez o +   Una vez que ya se ha clasificado correctamente el articulo del apartado x catalogo (en la OC aunque no haya sido grabado)
                                        If Me.mshFlex.get_TextMatrix(mintRenglonAct, C_COLSTATUSX) = "M" Then
                                            '''27OCT2010 - MAVF
                                            frmCXPJoyeria.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(Trim(.get_TextMatrix(lRenglon, C_COLUNIDAD))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODFAMILIA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODLINEA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODSUBLINEA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODKILATES))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)), CDec(Trim(CStr(ModEstandar.Numerico(.get_TextMatrix(lRenglon, C_ColMDSPESO))))), Trim(.get_TextMatrix(lRenglon, C_ColMDSCOLOR)), Trim(.get_TextMatrix(lRenglon, C_ColMDSPUREZA)), Trim(.get_TextMatrix(lRenglon, C_ColMDSCERTIFICADO)))
                                        Else
                                            '''27OCT2010 - MAVF
                                            frmCXPJoyeria.LlenaDatos(Trim(txtFolio.Text), .get_TextMatrix(.Row, C_COLCODAUX), nCol, lRenglon, mintRenglonAct, CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)), "2C", CDec(Trim(CStr(ModEstandar.Numerico(.get_TextMatrix(lRenglon, C_ColMDSPESO))))), Trim(.get_TextMatrix(lRenglon, C_ColMDSCOLOR)), Trim(.get_TextMatrix(lRenglon, C_ColMDSPUREZA)), Trim(.get_TextMatrix(lRenglon, C_ColMDSCERTIFICADO)))
                                        End If
                                    End If
                                Else
                                    '''27OCT2010 - MAVF
                                    '''siempre que es un articulo nuevo ( que no esta registrado en el ABC )
                                    frmCXPJoyeria.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(Trim(.get_TextMatrix(lRenglon, C_COLUNIDAD))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODFAMILIA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODLINEA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODSUBLINEA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODKILATES))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)), CDec(Trim(CStr(ModEstandar.Numerico(.get_TextMatrix(lRenglon, C_ColMDSPESO))))), Trim(.get_TextMatrix(lRenglon, C_ColMDSCOLOR)), Trim(.get_TextMatrix(lRenglon, C_ColMDSPUREZA)), Trim(.get_TextMatrix(lRenglon, C_ColMDSCERTIFICADO)))
                                End If
                                Enabled = False
                                .Row = mintRenglonAct
                                frmCXPJoyeria.Show()
                            Case gCODRELOJERIA
                                If Trim(.get_TextMatrix(mintRenglonAct, C_COLCRONO)) = "" Then .set_TextMatrix(mintRenglonAct, C_COLCRONO, False)

                                If txtFolioApartado.Text <> "" And Trim(.get_TextMatrix(mintRenglonAct, C_COLDESCRIPCION)) <> "" Then '''apartados x catalogo - 1ra vez
                                    If CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODFAMILIAX))) = 0 And CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODLINEAX))) = 0 And CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODSUBLINEAX))) = 0 And CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODKILATESX))) = 0 And CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODTIPOMATERIALX))) = 0 And Trim(.get_TextMatrix(mintRenglonAct, C_COLADICIONALX)) = "" Then
                                        If (Trim(.get_TextMatrix(mintRenglonAct, C_COLCODIGO)) = "" Or Trim(.get_TextMatrix(mintRenglonAct, C_COLCODAUX)) = "") Then
                                            frmCXPRelojeria.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(Trim(.get_TextMatrix(lRenglon, C_COLUNIDAD))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODMARCA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODMODELO))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), Trim(.get_TextMatrix(lRenglon, C_COLGENERO)), Trim(.get_TextMatrix(lRenglon, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(lRenglon, C_COLCRONO)), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)))
                                        Else '''Apartados x Cat -> RESURTIDOS
                                            '''FUE MODIFICADO DENTRO DE LA ORDEN DE COMPRA, POR LO QUE SE DEBERA MOSTRAR EL CAMBIO
                                            If Me.mshFlex.get_TextMatrix(mintRenglonAct, C_COLSTATUSX) = "M" Then
                                                frmCXPRelojeria.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(Trim(.get_TextMatrix(lRenglon, C_COLUNIDAD))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODMARCA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODMODELO))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), Trim(.get_TextMatrix(lRenglon, C_COLGENERO)), Trim(.get_TextMatrix(lRenglon, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(lRenglon, C_COLCRONO)), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)))
                                            Else
                                                frmCXPRelojeria.LlenaDatos(Trim(txtFolio.Text), .get_TextMatrix(.Row, C_COLCODAUX), nCol, lRenglon, mintRenglonAct, CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)), "1C")
                                            End If
                                        End If
                                    Else '''2da vez o +
                                        '''FUE MODIFICADO DENTRO DE LA ORDEN DE COMPRA, POR LO QUE SE DEBERA MOSTRAR EL CAMBIO
                                        If Me.mshFlex.get_TextMatrix(mintRenglonAct, C_COLSTATUSX) = "M" Then
                                            frmCXPRelojeria.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(Trim(.get_TextMatrix(lRenglon, C_COLUNIDAD))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODMARCA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODMODELO))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), Trim(.get_TextMatrix(lRenglon, C_COLGENERO)), Trim(.get_TextMatrix(lRenglon, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(lRenglon, C_COLCRONO)), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)))
                                        Else
                                            frmCXPRelojeria.LlenaDatos(Trim(txtFolio.Text), .get_TextMatrix(.Row, C_COLCODAUX), nCol, lRenglon, mintRenglonAct, CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)), "2C")
                                        End If
                                    End If
                                Else
                                    '''If (Trim(.TextMatrix(mintRenglonAct, C_COLCODIGO)) = "") Or (Trim(.TextMatrix(mintRenglonAct, C_COLCODIGO)) = "" And Trim(.TextMatrix(mintRenglonAct, C_COLCODAUX)) = "") Then
                                    frmCXPRelojeria.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(Trim(.get_TextMatrix(lRenglon, C_COLUNIDAD))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODMARCA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODMODELO))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), Trim(.get_TextMatrix(lRenglon, C_COLGENERO)), Trim(.get_TextMatrix(lRenglon, C_COLMOVIMIENTO)), Trim(.get_TextMatrix(lRenglon, C_COLCRONO)), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)))
                                    '''Else    '''Apartados x Cat / RESURITIDOS
                                    '''If frmCXPOrdenCompra.mshFlex.TextMatrix(mintRenglonAct, C_COLSTATUSX) = "M" Then
                                    '''   frmCXPRelojeria.LLenaForma CLng(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(Trim(.TextMatrix(lRenglon, C_ColUNIDAD))), CLng(Numerico(.TextMatrix(lRenglon, C_COLCODMARCA))), CLng(Numerico(.TextMatrix(lRenglon, C_COLCODMODELO))), CLng(Numerico(.TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(txtDescripcion.Caption), Trim(.TextMatrix(lRenglon, C_COLGENERO)), Trim(.TextMatrix(lRenglon, C_COLMOVIMIENTO)), Trim(.TextMatrix(lRenglon, C_COLCRONO)), Trim(.TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.TextMatrix(lRenglon, C_COLADICIONAL)), CCur(Numerico(.TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.TextMatrix(lRenglon, C_COLORIGENANT))), CLng(Numerico(.TextMatrix(lRenglon, C_ColCODIGOANT))), _
                                    'CInt(Numerico(.TextMatrix(lRenglon, C_ColCANTIDAD))), CCur(Numerico(.TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.TextMatrix(lRenglon, C_ColIMAGEN))
                                    '''Else
                                    '''frmCXPRelojeria.LlenaDatos Trim(txtFolio.text), .TextMatrix(.Row, C_COLCODAUX), nCol, lRenglon, mintRenglonAct, CInt(Numerico(.TextMatrix(lRenglon, C_ColCANTIDAD))), CCur(Numerico(.TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.TextMatrix(lRenglon, C_ColIMAGEN)), "2C"
                                    '''End If
                                    '''End If
                                End If

                                Enabled = False
                                .Row = mintRenglonAct
                                frmCXPRelojeria.Show()
                            Case gCODVARIOS
                                'If txtFolioApartado.Text <> "" And Trim(.get_TextMatrix(mintRenglonAct, C_COLDESCRIPCION)) <> "" Then '''apartados x catalogo - 1ra vez
                                '    If CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODFAMILIAX))) = 0 And CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODLINEAX))) = 0 And CInt(Numerico(.get_TextMatrix(mintRenglonAct, C_COLCODTIPOMATERIALX))) = 0 And Trim(.get_TextMatrix(mintRenglonAct, C_COLADICIONALX)) = "" Then
                                '        If (Trim(.get_TextMatrix(mintRenglonAct, C_COLCODIGO)) = "" Or Trim(.get_TextMatrix(mintRenglonAct, C_COLCODAUX)) = "") Then
                                '            frmCXPVarios.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(.get_TextMatrix(lRenglon, C_COLUNIDAD)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODFAMILIA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODLINEA))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)))
                                '        Else '''Apartados x Cat -> RESURTIDOS
                                '            If Me.mshFlex.get_TextMatrix(mintRenglonAct, C_COLSTATUSX) = "M" Then
                                '                frmCXPVarios.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(.get_TextMatrix(lRenglon, C_COLUNIDAD)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODFAMILIA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODLINEA))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)))
                                '            Else
                                '                frmCXPVarios.LlenaDatos(Trim(txtFolio.Text), .get_TextMatrix(.Row, C_COLCODAUX), nCol, lRenglon, mintRenglonAct, CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)), "1C")
                                '            End If
                                '        End If
                                '    Else '''2da vez o +
                                '        If Me.mshFlex.get_TextMatrix(mintRenglonAct, C_COLSTATUSX) = "M" Then
                                '            frmCXPVarios.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(.get_TextMatrix(lRenglon, C_COLUNIDAD)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODFAMILIA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODLINEA))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)))
                                '        Else
                                '            frmCXPVarios.LlenaDatos(Trim(txtFolio.Text), .get_TextMatrix(.Row, C_COLCODAUX), nCol, lRenglon, mintRenglonAct, CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)), "2C")
                                '        End If
                                '    End If
                                'Else
                                '    '''If (Trim(.TextMatrix(mintRenglonAct, C_COLCODIGO)) = "") Or (Trim(.TextMatrix(mintRenglonAct, C_COLCODIGO)) = "" And Trim(.TextMatrix(mintRenglonAct, C_COLCODAUX)) = "") Then
                                '    frmCXPVarios.LLenaForma(CInt(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(.get_TextMatrix(lRenglon, C_COLUNIDAD)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODFAMILIA))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODLINEA))), Trim(.get_TextMatrix(lRenglon, C_COLDESCRIPCION)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.get_TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.get_TextMatrix(lRenglon, C_COLADICIONAL)), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.get_TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLORIGENANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.get_TextMatrix(lRenglon, C_COLCANTIDAD))), CDec(Numerico(.get_TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.get_TextMatrix(lRenglon, C_ColIMAGEN)))
                                '    '''Else    '''Apartados x Cat / RESURTIDOS
                                '    '''    If frmCXPOrdenCompra.mshFlex.TextMatrix(mintRenglonAct, C_COLSTATUSX) = "M" Then
                                '    '''       frmCXPVarios.LLenaForma CLng(.Col), lRenglon, mintRenglonAct, BuscaCodUnidad(.TextMatrix(lRenglon, C_ColUNIDAD)), CLng(Numerico(.TextMatrix(lRenglon, C_COLCODFAMILIA))), CLng(Numerico(.TextMatrix(lRenglon, C_COLCODLINEA))), Trim(txtDescripcion.Caption), CLng(Numerico(.TextMatrix(lRenglon, C_COLCODTIPOMATERIAL))), Trim(.TextMatrix(lRenglon, C_COLCODIGOARTICULOPROV)), Trim(.TextMatrix(lRenglon, C_COLADICIONAL)), CCur(Numerico(.TextMatrix(lRenglon, C_COLPRECIOPUBDOLAR))), Trim(.TextMatrix(lRenglon, C_COLMONEDAPP)), CInt(Numerico(.TextMatrix(lRenglon, C_COLORIGENANT))), CLng(Numerico(.TextMatrix(lRenglon, C_ColCODIGOANT))), CInt(Numerico(.TextMatrix(lRenglon, C_ColCANTIDAD))), CCur(Numerico(.TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.TextMatrix(lRenglon, C_ColIMAGEN))
                                '    '''    Else
                                '    '''       frmCXPVarios.LlenaDatos Trim(txtFolio.text), .TextMatrix(.Row, C_COLCODAUX), nCol, lRenglon, mintRenglonAct, CInt(Numerico(.TextMatrix(lRenglon, C_ColCANTIDAD))), CCur(Numerico(.TextMatrix(lRenglon, C_COLPRECIOUNITARIO))), Trim(.TextMatrix(lRenglon, C_ColIMAGEN)), "2C"
                                '    '''    End If
                                '    '''End If
                                'End If

                                'Enabled = False
                                '.Row = mintRenglonAct
                                'frmCXPVarios.Show()
                        End Select

                    Case C_COLUNIDAD
                        .Focus()
                    Case C_COLCANTIDAD
                        If .get_TextMatrix(.Row, C_COLDESCRIPCION) <> "" Then
                            txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            txtFlex.BackColor = .CellBackColor
                            ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, eventArgs.keyAscii)
                            ModEstandar.SelTextoTxt(txtFlex)
                        Else
                            Exit Sub
                        End If
                    Case C_COLPRECIOUNITARIO
                        If .get_TextMatrix(.Row, C_COLDESCRIPCION) <> "" Then
                            txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            txtFlex.BackColor = .CellBackColor
                            ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, eventArgs.keyAscii)
                            ModEstandar.SelTextoTxt(txtFlex)
                        Else
                            Exit Sub
                        End If
                    Case C_COLDESCTO
                        If .get_TextMatrix(.Row, C_COLDESCRIPCION) = "" Then
                            Exit Sub
                        Else
                            If CDec(Numerico(.get_TextMatrix(.Row, C_COLDESCTOPORC))) > 0 Then
                                '        frmCXPDescuento.optDescuento(0).Checked = True
                                '    Else
                                '        If CDec(Numerico(.get_TextMatrix(.Row, C_COLDESCTO))) = 0 Then
                                '            frmCXPDescuento.optDescuento(0).Checked = True
                                '        Else
                                '            frmCXPDescuento.optDescuento(1).Checked = True
                                '        End If
                                '    End If
                                '    If frmCXPDescuento.optDescuento(1).Checked Then
                                '        frmCXPDescuento.txtDescuento(0).Enabled = False
                                '        frmCXPDescuento.txtDescuento(0).Text = "0.00"
                                '        frmCXPDescuento.txtDescuento(1).Enabled = True
                                '        frmCXPDescuento.txtDescuento(1).Text = VB6.Format(.get_TextMatrix(.Row, C_COLDESCTO), "###,###,##0.00")
                                '    Else
                                '        frmCXPDescuento.txtDescuento(1).Enabled = False
                                '        frmCXPDescuento.txtDescuento(1).Text = "0.00"
                                '        frmCXPDescuento.txtDescuento(0).Enabled = True
                                '        frmCXPDescuento.txtDescuento(0).Text = VB6.Format(.get_TextMatrix(.Row, C_COLDESCTOPORC), "##0.00")
                            End If
                            '    frmCXPDescuento.Tag = UCase(Name)
                            '    frmCXPDescuento.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Top) + VB6.PixelsToTwipsY(mshFlex.Top))
                            '    frmCXPDescuento.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Left) + mshFlex.get_ColWidth(C_COLCODIGO) + mshFlex.get_ColWidth(C_COLDESCRIPCION) + mshFlex.get_ColWidth(C_COLUNIDAD) + mshFlex.get_ColWidth(C_COLCANTIDAD))
                            '    frmCXPDescuento.ShowDialog()
                        End If
                            Case C_COLIVA
                        ScrollGrid()
                End Select
            ElseIf eventArgs.keyAscii = 27 Then
            ElseIf eventArgs.keyAscii = 32 Then
                If mintCodGrupo = 0 Then
                    MsgBox("Debe seleccionar un Grupo de Artículos antes de consultar o añadir artículos a la Orden de Compra", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    eventArgs.keyAscii = 0
                    dbcGrupo.Focus()
                    ModEstandar.SelTxt()
                    Exit Sub
                End If
                nCol = .Col
                nRow = .Row
                '''en esta parte se validará si es el rengón, columna que le
                '''corresponde editarse
                If (.Row > 1) Then
                    '''de tal modo que si el renglón es mayor que 1
                    '''y si un renglón antes del renglón actual está vacío,
                    '''el renglón actual no se editará
                    If Trim(.get_TextMatrix(.Row - 1, C_COLDESCRIPCION)) = "" Then
                        .Focus()
                        Exit Sub
                    End If
                End If
                If .Col = C_COLCODIGO Then
                    If .get_TextMatrix(.Row, C_COLDESCRIPCION) = "" Then
                        txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                        txtFlex.BackColor = .CellBackColor
                        ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, eventArgs.keyAscii)
                        ModEstandar.SelTextoTxt(txtFlex)
                    Else
                        Exit Sub
                    End If
                End If
            Else
                Select Case .Col
                    Case C_COLCODIGO
                        If Trim(.get_TextMatrix(.Row, C_COLDESCRIPCION)) = "" And Trim(.get_TextMatrix(.Row - 1, C_COLDESCRIPCION)) <> "" Then
                            txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            txtFlex.BackColor = .CellBackColor

                            'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            eventArgs.keyAscii = ModEstandar.MskCantidad((txtFlex.Text), eventArgs.keyAscii, 9, 0, (txtFlex.SelectionStart))

                            ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, eventArgs.keyAscii)
                            If Len(txtFlex.Text) <> 1 Then
                                ModEstandar.SelTextoTxt(txtFlex)
                            End If
                        Else
                            Exit Sub
                        End If
                    Case C_COLCANTIDAD
                        If .get_TextMatrix(.Row, C_COLDESCRIPCION) <> "" Then
                            txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            txtFlex.BackColor = .CellBackColor

                            'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            eventArgs.keyAscii = ModEstandar.MskCantidad((txtFlex.Text), eventArgs.keyAscii, 9, 0, (txtFlex.SelectionStart))

                            ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, eventArgs.keyAscii)
                            If Len(txtFlex.Text) <> 1 Then
                                ModEstandar.SelTextoTxt(txtFlex)
                            End If
                        Else
                            Exit Sub
                        End If
                    Case C_COLPRECIOUNITARIO
                        If .get_TextMatrix(.Row, C_COLDESCRIPCION) <> "" Then
                            txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            txtFlex.BackColor = .CellBackColor

                            'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            eventArgs.keyAscii = ModEstandar.MskCantidad((txtFlex.Text), eventArgs.keyAscii, 9, 2, (txtFlex.SelectionStart))

                            ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, eventArgs.keyAscii)
                            If Len(txtFlex.Text) <> 1 Then
                                ModEstandar.SelTextoTxt(txtFlex)
                            End If
                        Else
                            Exit Sub
                        End If
                    Case C_COLIVA
                        ScrollGrid()
                End Select
            End If
        End With
    End Sub

    'UPGRADE_WARNING: Event optMoneda.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optMoneda_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoneda.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = optMoneda.GetIndex(eventSender)
            If Me.optMoneda(0).Checked Then
                cMonedadeCompra = C_DOLAR
                If UltimaMoneda = "P" Then
                    ConvertirCantidades("P", "D")
                ElseIf UltimaMoneda = "E" Then
                    ConvertirCantidades("E", "D")
                End If
                UltimaMoneda = "D"
            ElseIf Me.optMoneda(1).Checked Then
                cMonedadeCompra = C_PESO
                If UltimaMoneda = "D" Then
                    ConvertirCantidades("D", "P")
                ElseIf UltimaMoneda = "E" Then
                    ConvertirCantidades("E", "P")
                End If
                UltimaMoneda = "P"
            Else
                cMonedadeCompra = C_EURO
                If UltimaMoneda = "D" Then
                    ConvertirCantidades("D", "E")
                ElseIf UltimaMoneda = "P" Then
                    ConvertirCantidades("P", "E")
                End If
                UltimaMoneda = "E"
            End If
        End If
    End Sub

    Private Sub optMoneda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoneda.Enter
        Dim Index As Integer = optMoneda.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    'UPGRADE_WARNING: Event txtCostoAdicional.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtCostoAdicional_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoAdicional.TextChanged
        Call Me.ActualizaCantidades()
    End Sub

    Private Sub txtCostoAdicional_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoAdicional.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtCostoAdicional)
    End Sub

    Private Sub txtCostoAdicional_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCostoAdicional.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtCostoAdicional.Text = VB6.Format(Numerico((Me.txtCostoAdicional.Text)), "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtCostoAdicional.Text), KeyAscii, 9, 2, (Me.txtCostoAdicional.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCostoAdicional_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoAdicional.Leave
        Me.txtCostoAdicional.Text = VB6.Format(Numerico((Me.txtCostoAdicional.Text)), "###,###,##0.00")
    End Sub

    'UPGRADE_WARNING: Event txtCostosIndirectos.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtCostosIndirectos_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostosIndirectos.TextChanged
        Call Me.ActualizaCantidades()
    End Sub

    Private Sub txtCostosIndirectos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostosIndirectos.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtCostosIndirectos)
    End Sub

    Private Sub txtCostosIndirectos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCostosIndirectos.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode <> System.Windows.Forms.Keys.Escape Then
            Me.mshFlex.Row = 1
            Me.mshFlex.Col = 0
            Me.mshFlex.TopRow = 1
        End If
    End Sub

    Private Sub txtCostosIndirectos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCostosIndirectos.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtCostosIndirectos.Text = VB6.Format(Numerico((Me.txtCostosIndirectos.Text)), "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtCostosIndirectos.Text), KeyAscii, 9, 2, (Me.txtCostosIndirectos.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCostosIndirectos_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostosIndirectos.Leave
        Me.txtCostosIndirectos.Text = VB6.Format(Numerico((Me.txtCostosIndirectos.Text)), "###,###,##0.00")
    End Sub

    Private Sub txtDesctoFinanciero_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesctoFinanciero.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt((Me.txtDesctoFinanciero))
    End Sub

    Private Sub txtDesctoFinanciero_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDesctoFinanciero.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtDesctoFinanciero.Text = VB6.Format(Numerico((Me.txtDesctoFinanciero.Text)), "##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtDesctoFinanciero.Text), KeyAscii, 3, 2, (Me.txtDesctoFinanciero.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDesctoFinanciero_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesctoFinanciero.Leave
        Me.txtDesctoFinanciero.Text = VB6.Format(Numerico((Me.txtDesctoFinanciero.Text)), "##0.00")
    End Sub

    Private Sub txtDescuento_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescuento.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtDescuento)
    End Sub

    Private Sub txtEntregarEn_GotFocus()
        Pon_Tool()
        'ModEstandar.SelTextoTxt(rtEntregaren.Text)
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Dim nCol As Object
        Dim nRen As Integer
        Dim nIva As Object
        Dim nIVAImporte As Object

        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        With mshFlex
            'UPGRADE_WARNING: Couldn't resolve default property of object nCol. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nCol = .Col
            nRen = .Row
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
                        Exit Sub
                    End If
                    Call ActualizaCantidades()
                    'UPGRADE_WARNING: Couldn't resolve default property of object nCol. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If nCol = C_COLCODIGO And CInt(Numerico((Me.txtFlex.Text))) = 0 Then
                        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left 'Alinear a la izquierda
                        txtFlex.Text = ""
                        txtFlex.Visible = False
                        'UPGRADE_WARNING: Couldn't resolve default property of object nCol. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ElseIf nCol = C_COLCODIGO And CInt(Numerico((Me.txtFlex.Text))) <> 0 Then
                        'No esconde nada
                        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left 'Alinear a la izquierda
                        txtFlex.Text = ""
                        txtFlex.Visible = False
                        'UPGRADE_WARNING: Couldn't resolve default property of object nCol. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ElseIf nCol <> C_COLCODIGO Then
                        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left 'Alinear a la izquierda
                        txtFlex.Text = ""
                        txtFlex.Visible = False
                    End If
                Case System.Windows.Forms.Keys.Return
                    Select Case .Col
                        Case C_COLCODIGO
                            'Llenar los datos del artículo solicitado
                            If CInt(Numerico((txtFlex.Text))) <> 0 Then

                                If Trim(txtFlex.Text) <> "" Then
                                    ''' busqueda dual - 26 May 2004
                                    ResBusquedaArt = BuscarCodigoArticulo_Cxp(Trim(txtFlex.Text))
                                    If ResBusquedaArt > 0 Or ResBusquedaArt = -1 Then


                                        'Llenar los datos del artículo solicitado, si es que existe para este proveedor y grupo de artículos
                                        '''ojo origen - 8 Nov 2003
                                        '''gStrSql = "select * from CatArticulos(Nolock) Where CodProveedor = " & mintCodProveedor & " and CodGrupo = " & mintCodGrupo & " and CodAlmacenOrigen = " & mintCodOrigen & " And CodArticulo = " & ResBusquedaArt
                                        gStrSql = "select * From CatArticulos(Nolock) Where CodArticulo = " & ResBusquedaArt
                                        ModEstandar.BorraCmd()
                                        Cmd.CommandText = "dbo.UP_Select_Datos"
                                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                                        RsGral = Cmd.Execute
                                        If RsGral.RecordCount > 0 Then

                                            If ArticuloRepetidoenGrid(ResBusquedaArt) Then
                                                MsgBox("Este artículo ya está registrado en la captura" & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                                                txtFlex.Text = ""
                                                txtFlex.Visible = False
                                                mshFlex_EnterCell(mshFlex, New System.EventArgs())
                                                .Focus()
                                                Exit Sub
                                            End If

                                            .set_TextMatrix(.Row, C_COLCODIGO, VB6.Format(ResBusquedaArt, "###,###,##0"))
                                            txtFlex.Text = ""
                                            txtFlex.Visible = False
                                            LlenaLineaGrid()
                                            mshFlex_EnterCell(mshFlex, New System.EventArgs())
                                            .Focus()
                                            Exit Sub
                                        Else
                                            'UPGRADE_NOTE: Text was upgraded to Text. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
                                            MsgBox("No se encontró ningún artículo que pertenezca al grupo " & Trim(Me.dbcGrupo.Text) & vbNewLine & "del proveedor " & Trim(Me.dbcProveedor.Text) & " con el código : " & Trim(Me.txtFlex.Text), MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                                            .set_TextMatrix(.Row, C_COLCODIGO, "")
                                            ModEstandar.SelTxt()
                                            Exit Sub
                                        End If
                                    ElseIf ResBusquedaArt = -2 Then
                                        CodAux = CInt(txtFlex.Text)
                                        txtFlex.Text = ""
                                        BuscarArticulosCxP(True, VB.Right(New String("0", 6) & Trim(CStr(CodAux)), 6))
                                    End If
                                End If
                            Else
                                .set_TextMatrix(.Row, C_COLCODIGO, "")
                            End If

                        Case C_COLCANTIDAD
                            .set_TextMatrix(.Row, C_COLCANTIDAD, VB6.Format(Numerico((txtFlex.Text)), "###,###,##0"))
                            .Col = C_COLPRECIOUNITARIO
                        Case C_COLPRECIOUNITARIO
                            .set_TextMatrix(.Row, C_COLPRECIOUNITARIO, VB6.Format(Numerico((txtFlex.Text)), "###,###,##0.00"))
                            .set_TextMatrix(.Row, C_COLPRECIOUNITARIO4DEC, .get_TextMatrix(.Row, C_COLPRECIOUNITARIO))
                            .Col = C_COLIVA
                    End Select
                    ActualizaCantidades()
                    .Focus()
                    txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Left 'Alinear a la izquierda
                    txtFlex.Text = ""
                    txtFlex.Visible = False
            End Select
        End With
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Select Case Me.mshFlex.Col
            Case C_COLCANTIDAD
                If KeyAscii = 13 Then
                    Me.txtFlex.Text = VB6.Format(Numerico((Me.txtFlex.Text)), "###,###,##0")
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                KeyAscii = ModEstandar.MskCantidad((Me.txtFlex.Text), KeyAscii, 9, 0, (Me.txtFlex.SelectionStart))
            Case C_COLPRECIOUNITARIO
                If KeyAscii = 13 Then
                    Me.txtFlex.Text = VB6.Format(Numerico((Me.txtFlex.Text)), "###,###,##0.00")
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                KeyAscii = ModEstandar.MskCantidad((Me.txtFlex.Text), KeyAscii, 9, 2, (Me.txtFlex.SelectionStart))
            Case C_COLCODIGO
                If KeyAscii = 13 Then
                    Me.txtFlex.Text = VB6.Format(Numerico((Me.txtFlex.Text)), "########0")
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                KeyAscii = ModEstandar.MskCantidad((Me.txtFlex.Text), KeyAscii, 9, 0, (Me.txtFlex.SelectionStart))
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    'UPGRADE_WARNING: Event txtFolio.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtFolio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtFolio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.Enter
        System.Windows.Forms.Application.DoEvents()
        SelTextoTxt(txtFolio)
        Pon_Tool()
    End Sub

    Private Sub txtFolio_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFolio.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Pregunta sólo en caso de que existan cambios en la clave (esto es, cuando se teclea una clave diferente a la actual)
        If Cambios() And KeyCode = System.Windows.Forms.Keys.Delete Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        KeyCode = 0
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se borre el contenido del text
                Case MsgBoxResult.Cancel
                    KeyCode = 0
                    Me.txtFolio.Focus()
            End Select
        End If
    End Sub

    Private Sub txtFolio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back And KeyAscii <> System.Windows.Forms.Keys.C Then
            KeyAscii = 0
        Else
            'Pregunta sólo si ha habido cambios
            If Cambios() And Not mblnNuevo Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        If Not Guardar() Then
                            KeyAscii = 0
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        KeyAscii = 0
                        Me.txtFolio.Focus()
                End Select
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolio.Leave
        If System.Windows.Forms.Form.ActiveForm.Text = Me.Text Then
            If mblnCambiosEnCodigo = True Then 'Si hubo cambios en el código hace la consulta
                LlenaDatos()
            End If
        End If
    End Sub

    Private Sub txtIVA_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtIVA.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtIVA)
    End Sub

    Private Sub txtPedido_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPedido.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtPedido)
    End Sub

    'UPGRADE_WARNING: Event txtPorcDescto.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtPorcDescto_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcDescto.TextChanged
        Dim I As Integer
        Dim nPorcDescto As Decimal
        If mintCodProveedor = 0 Then
            Exit Sub
        End If
        nPorcDescto = BuscaDesctoProveedor(mintCodProveedor)
        With mshFlex
            For I = 1 To .Rows
                If .get_TextMatrix(I, C_COLDESCRIPCION) = "" Then
                    Exit For
                End If
                .set_TextMatrix(I, C_COLDESCTOPORC, Me.txtPorcDescto.Text)
            Next I
        End With
        Call ActualizaCantidades()
    End Sub

    Private Sub txtPorcDescto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcDescto.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt((Me.txtPorcDescto))
    End Sub

    Private Sub txtPorcDescto_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPorcDescto.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtPorcDescto.Text = VB6.Format(Me.txtPorcDescto.Text, "##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtPorcDescto.Text), KeyAscii, 3, 2, (Me.txtPorcDescto.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPorcDescto_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcDescto.Leave
        Me.txtPorcDescto.Text = VB6.Format(Me.txtPorcDescto.Text, "##0.00")
    End Sub

    Private Sub txtRemision_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRemision.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtRemision)
    End Sub

    Private Sub txtSubTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSubTotal.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtSubTotal)
    End Sub

    'UPGRADE_WARNING: Event txtTasaIVA.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtTasaIVA_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTasaIva.TextChanged
        Call Me.ActualizaCantidades()
    End Sub

    Private Sub txtTasaIVA_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTasaIva.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt((Me.txtTasaIva))
    End Sub

    Private Sub txtTasaIVA_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTasaIva.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtTasaIva.Text = VB6.Format(Me.txtTasaIva.Text, "##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtTasaIva.Text), KeyAscii, 3, 2, (Me.txtTasaIva.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTasaIVA_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTasaIva.Leave
        Me.txtTasaIva.Text = VB6.Format(Me.txtTasaIva.Text, "##0.00")
    End Sub

    'UPGRADE_WARNING: Event txtTipoCambio.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtTipoCambio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.TextChanged
        Me.txtTipoCambioConciliado.Text = Me.txtTipoCambio.Text
    End Sub

    Private Sub txtTipoCambio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtTipoCambio)
    End Sub

    Private Sub txtTipoCambio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtTipoCambio.Text = VB6.Format(Me.txtTipoCambio.Text, "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtTipoCambio.Text), KeyAscii, 9, 2, (Me.txtTipoCambio.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambio.Leave
        Me.txtTipoCambio.Text = VB6.Format(Me.txtTipoCambio.Text, "###,###,##0.00")
    End Sub

    Private Sub txtTipoCambioConciliado_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioConciliado.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt((Me.txtTipoCambioConciliado))
    End Sub

    Private Sub txtTipoCambioConciliado_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambioConciliado.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtTipoCambioConciliado.Text = VB6.Format(Me.txtTipoCambioConciliado.Text, "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtTipoCambioConciliado.Text), KeyAscii, 9, 2, (Me.txtTipoCambioConciliado.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambioConciliado_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioConciliado.Leave
        Me.txtTipoCambioConciliado.Text = VB6.Format(Me.txtTipoCambioConciliado.Text, "###,###,##0.00")
    End Sub

    'UPGRADE_WARNING: Event txtTipoCambioEuro.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtTipoCambioEuro_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.TextChanged
        Me.txtTipoCambioEuroConciliado.Text = Me.txtTipoCambioEuro.Text
    End Sub

    Private Sub txtTipoCambioEuro_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtTipoCambioEuro)
    End Sub

    Private Sub txtTipoCambioEuro_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambioEuro.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtTipoCambioEuro.Text = VB6.Format(Me.txtTipoCambioEuro.Text, "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtTipoCambioEuro.Text), KeyAscii, 9, 2, (Me.txtTipoCambioEuro.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambioEuro_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.Leave
        Me.txtTipoCambioEuro.Text = VB6.Format(Me.txtTipoCambioEuro.Text, "###,###,##0.00")
    End Sub

    Private Sub txtTipoCambioEuroConciliado_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuroConciliado.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt((Me.txtTipoCambioEuroConciliado))
    End Sub

    Private Sub txtTipoCambioEuroConciliado_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambioEuroConciliado.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtTipoCambioEuroConciliado.Text = VB6.Format(Me.txtTipoCambioEuroConciliado.Text, "###,###,##0.00")
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.MskCantidad((Me.txtTipoCambioEuroConciliado.Text), KeyAscii, 9, 2, (Me.txtTipoCambioEuroConciliado.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambioEuroConciliado_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuroConciliado.Leave
        Me.txtTipoCambioEuroConciliado.Text = VB6.Format(Me.txtTipoCambioEuroConciliado.Text, "###,###,##0.00")
    End Sub

    Private Sub txtTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTotal.Enter
        Pon_Tool()
        ModEstandar.SelTextoTxt(txtTotal)
    End Sub

    Public Sub MuestraClasificacion()
        mshFlex_KeyPressEvent(mshFlex, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Public Function NombreArchArticulo(ByRef lFolioOc As String, ByRef lNumP As Integer) As String
        Dim lConsec As String
        Dim lNumPartida As String

        lConsec = VB.Right(lFolioOc, 6)
        lNumPartida = VB.Right("00" & Trim(CStr(lNumP)), 2)
        NombreArchArticulo = lConsec & lNumPartida
    End Function

    Sub BuscarArticulosCxP(ByRef BusquedaEspecial As Boolean, ByRef CodArticulo As String)
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim strControlActual As String 'Nombre del control actual
        Dim Columna As Integer

        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        strControlActual = IIf((BusquedaEspecial), "TXTCODARTICULO", strControlActual)
        Select Case strControlActual
            Case "TXTCODARTICULO"
                strCaptionForm = "Consulta de Artículos"
                If BusquedaEspecial Then
                    strSQL = "SELECT     CodArticulo AS CODIGO, RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION, " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR], " & "dbo.FormatCantidad(A.PrecioPubDolar)  AS [PRECIO PÚBLICO] , " & "case PesosFijos WHEN 0 THEN 'DÓLARES' WHEN 1 THEN 'PESOS' END AS [MONEDA] " & "From CatArticulos A cross Join Configuraciongeneral c WHERE ((CodArticulo = " & CInt(CodArticulo) & ") " & "OR   (OrigenAnt = " & CInt(VB.Left(CodArticulo, 1)) & ") AND (CodigoAnt = " & CInt(VB.Right(CodArticulo, 5)) & ")) And CodGrupo = " & mintCodGrupo & " And CodProveedor = " & mintCodProveedor & " And CodAlmacenOrigen = " & mintCodOrigen
                Else
                    strSQL = "SELECT     A.CodArticulo AS CODIGO, LTRIM(RTRIM(A.DescArticulo)) AS DESCRIPCION, M.DescTipoMaterial AS MATERIAL, LTrim(Rtrim(A.CodigoArticuloProv)) AS [ARTICULO PROV],  CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+ '-'+ RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]   " & "FROM dbo.CatArticulos A INNER JOIN dbo.CatTipoMaterial M ON A.CodTipoMaterial = M.CodTipoMaterial  " '''+ DevuelveQuery
                End If
            Case "TXTDESCARTICULO"
                strCaptionForm = "Consulta de Artículos"
                '''strSQL = "SELECT      LTRIM(RTRIM(A.DescArticulo))  AS DESCRIPCION,A.CodArticulo AS CODIGO, M.DescTipoMaterial AS MATERIAL, LTrim(Rtrim(A.CodigoArticuloProv)) AS [ARTICULO PROV],  CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+ '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]   " & _
                '"FROM dbo.CatArticulos A INNER JOIN dbo.CatTipoMaterial M ON A.CodTipoMaterial = M.CodTipoMaterial  " & _
                'DevuelveQuery + " AND  DescArticulo Like '" & Trim(txtDescArticulo) & "%'"
                strSQL = "SELECT      LTRIM(RTRIM(A.DescArticulo))  AS DESCRIPCION,A.CodArticulo AS CODIGO, M.DescTipoMaterial AS MATERIAL, LTrim(Rtrim(A.CodigoArticuloProv)) AS [ARTICULO PROV],  CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+ '-' + RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR]   " & "FROM dbo.CatArticulos A INNER JOIN dbo.CatTipoMaterial M ON A.CodTipoMaterial = M.CodTipoMaterial  " & " AND  DescArticulo Like '" & Trim(txtFlex.Text) & "%'"
            Case Else
                'Sale de este sub para ke no ejecute ninguna opcion
                Exit Sub
        End Select

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, strSQL))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            If BusquedaEspecial = True Then
                MsgBox("El Artículo no existe." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                RsGral.Close()
                Exit Sub
            Else
                MsgBox(C_msgSINDATOS & vbNewLine & "Verifique por favor....", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                RsGral.Close()
                Exit Sub
            End If
        End If

        'Carga el formulario de consulta
        'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        'Load(FrmConsultas)
        Call ConfiguraConsultas(FrmConsultas, 11050, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODARTICULO"
                    .set_ColWidth(0, 0, 900)
                    .set_ColWidth(1, 0, 4800)
                    .set_ColWidth(2, 0, 1900)
                    .set_ColWidth(3, 0, 1620)
                    .set_ColWidth(4, 0, 1800)

                    .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)

                Case "TXTDESCARTICULO"
                    .set_ColWidth(0, 0, 4800)
                    .set_ColWidth(1, 0, 900)
                    .set_ColWidth(2, 0, 1900)
                    .set_ColWidth(3, 0, 1620)
                    .set_ColWidth(4, 0, 1800)
                    .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            End Select
            .WordWrap = False
        End With
        mblnFueraChange = True
        CentrarForma(FrmConsultas)
        FrmConsultas.ShowDialog()
        mblnFueraChange = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function BuscarCodigoArticulo_Cxp(ByRef Codigo As String) As Integer
        On Error GoTo Merr
        Dim CodigoString As String
        Dim CodAnterior As Integer
        Dim CodOrigen As Integer
        BuscarCodigoArticulo_Cxp = 0
        'Esta función recibe como parámetro el código de artículo que se desea buscar.
        'Se buscará en la tabla de articulos, en codigo de Articulo y Código de Articulo anterior.
        'Es posible que se presenten tres situaciones:
        'el código buscado está en el campo código de Articulo de la Tabla
        '       --En este caso el codigo del articulo a buscar no cambia.
        'el código buscado está en el campo código anterior
        '       --En este caso, el codigo a buscar ahora será el que corresponda en el campo Codigo articulo del mismo registro.
        'El codigo a buscar está en los dos campos anteriores
        '       --En este caso, se mostrará una pantalla de ayuda para mostarle al usuario, los dos articulos encontrados. De los cuales debe seleccionar uno.

        ''Esta función regresa:
        '    -1 : Si el articulo no es encontró
        '    -2 : Si se encontró más de un Artículo
        If Len(CStr(Codigo)) = 6 Then
            CodigoString = VB.Right(New String("0", 6) & CStr(Codigo), 6)
            CodOrigen = CDbl(VB.Left(CodigoString, 1))
            CodAnterior = CInt(VB.Right(CodigoString, 5))
            gStrSql = "SELECT  * From CatArticulos WHERE ((CodArticulo = " & Codigo & ") " & "OR   (OrigenAnt = " & CodOrigen & " AND CodigoAnt = " & CodAnterior & ")) And CodGrupo = " & mintCodGrupo & " And CodProveedor = " & mintCodProveedor & " And CodAlmacenOrigen = " & mintCodOrigen
        Else
            gStrSql = "SELECT  * From CatArticulos WHERE CodArticulo = " & Codigo & " And CodGrupo = " & mintCodGrupo & " And CodProveedor = " & mintCodProveedor & " And CodAlmacenOrigen = " & mintCodOrigen
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        Select Case RsGral.RecordCount
            Case Is <= 0
                'No se encontró el código de articulo
                BuscarCodigoArticulo_Cxp = -1
            Case 1
                'encontro solo 1
                BuscarCodigoArticulo_Cxp = CInt(RsGral.Fields("CodArticulo").Value)
            Case Else
                'encontro mas de 1 - muestra busqueda
                BuscarCodigoArticulo_Cxp = -2
        End Select
        Exit Function

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ArticuloRepetidoenGrid(ByRef CodArticulo As Integer) As Boolean
        Dim I As Integer

        ArticuloRepetidoenGrid = False
        With mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLCODIGO)) = "" Then Exit Function

                If CInt(Numerico(.get_TextMatrix(I, C_COLCODIGO))) = CodArticulo Then
                    ArticuloRepetidoenGrid = True
                    Exit For
                End If
            Next
        End With
        Exit Function
    End Function

    Public Sub ArticuloRepetido(ByRef Articulo As Integer)
        With mshFlex
            If ArticuloRepetidoenGrid(Articulo) Then
                MsgBox("Este artículo ya está registrado en la captura" & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                txtFlex.Text = ""
                txtFlex.Visible = False
                mshFlex_EnterCell(mshFlex, New System.EventArgs())
                Exit Sub
            End If

            .set_TextMatrix(.Row, C_COLCODIGO, VB6.Format(Articulo, "###,###,##0"))
            txtFlex.Text = ""
            txtFlex.Visible = False
            LlenaLineaGrid()
            mshFlex_EnterCell(mshFlex, New System.EventArgs())
        End With
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCXPOrdenCompra))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnAsignarCodigos = New System.Windows.Forms.Button()
        Me.btnProv = New System.Windows.Forms.Button()
        Me.txtDesctoFinanciero = New System.Windows.Forms.TextBox()
        Me.txtTipoCambioConciliado = New System.Windows.Forms.TextBox()
        Me.txtTipoCambioEuroConciliado = New System.Windows.Forms.TextBox()
        Me.txtPorcDescto = New System.Windows.Forms.TextBox()
        Me.txtTasaIva = New System.Windows.Forms.TextBox()
        Me.txtRemision = New System.Windows.Forms.TextBox()
        Me.txtPedido = New System.Windows.Forms.TextBox()
        Me.txtTipoCambioEuro = New System.Windows.Forms.TextBox()
        Me.txtTipoCambio = New System.Windows.Forms.TextBox()
        Me._optMoneda_2 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me.txtTotal = New System.Windows.Forms.TextBox()
        Me.txtIVA = New System.Windows.Forms.TextBox()
        Me.txtDescuento = New System.Windows.Forms.TextBox()
        Me.txtSubTotal = New System.Windows.Forms.TextBox()
        Me.txtCostosIndirectos = New System.Windows.Forms.TextBox()
        Me.txtCostoAdicional = New System.Windows.Forms.TextBox()
        Me.txtFolio = New System.Windows.Forms.TextBox()
        Me.fraApartado = New System.Windows.Forms.Panel()
        Me.txtFolioApartado = New System.Windows.Forms.TextBox()
        Me._lblOrden_2 = New System.Windows.Forms.Label()
        Me._fraOrden_0 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me._fraOrden_3 = New System.Windows.Forms.GroupBox()
        Me.fraMoneda = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblEuro = New System.Windows.Forms.Label()
        Me.lblDolar = New System.Windows.Forms.Label()
        Me.fraFecha = New System.Windows.Forms.Panel()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me._lblOrden_3 = New System.Windows.Forms.Label()
        Me._fraOrden_5 = New System.Windows.Forms.GroupBox()
        Me.fraEntregarEn = New System.Windows.Forms.GroupBox()
        Me.rtEntregaren = New System.Windows.Forms.RichTextBox()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.fraOtrosDatos = New System.Windows.Forms.GroupBox()
        Me.txtOtrosDatos = New System.Windows.Forms.Label()
        Me.fraCostos = New System.Windows.Forms.GroupBox()
        Me._lblOrden_8 = New System.Windows.Forms.Label()
        Me._lblOrden_7 = New System.Windows.Forms.Label()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me.fraEntrega = New System.Windows.Forms.GroupBox()
        Me.dtpFechaEntrega = New System.Windows.Forms.DateTimePicker()
        Me.dbcOrigen = New System.Windows.Forms.ComboBox()
        Me.dbcGrupo = New System.Windows.Forms.ComboBox()
        Me._lblOrden_6 = New System.Windows.Forms.Label()
        Me._lblOrden_5 = New System.Windows.Forms.Label()
        Me._lblOrden_4 = New System.Windows.Forms.Label()
        Me.mshFlex = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.lblDescProv = New System.Windows.Forms.Label()
        Me.lblDesctoFinanciero = New System.Windows.Forms.Label()
        Me.lblPorcDescto = New System.Windows.Forms.Label()
        Me.lblTasaIva = New System.Windows.Forms.Label()
        Me.lblRemision = New System.Windows.Forms.Label()
        Me.lblPedido = New System.Windows.Forms.Label()
        Me.txtDescripcion = New System.Windows.Forms.Label()
        Me.lblCR = New System.Windows.Forms.Label()
        Me._lblOrden_17 = New System.Windows.Forms.Label()
        Me._lblOrden_16 = New System.Windows.Forms.Label()
        Me.lblResurtido = New System.Windows.Forms.Label()
        Me._lblOrden_15 = New System.Windows.Forms.Label()
        Me.lblConciliado = New System.Windows.Forms.Label()
        Me.lblEstatus = New System.Windows.Forms.Label()
        Me._lblOrden_14 = New System.Windows.Forms.Label()
        Me._lblOrden_13 = New System.Windows.Forms.Label()
        Me._lblOrden_12 = New System.Windows.Forms.Label()
        Me._lblOrden_11 = New System.Windows.Forms.Label()
        Me._lblOrden_1 = New System.Windows.Forms.Label()
        Me._lblOrden_0 = New System.Windows.Forms.Label()
        Me.fraOrden = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblOrden = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnCancelar = New System.Windows.Forms.Button()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.fraApartado.SuspendLayout()
        Me._fraOrden_0.SuspendLayout()
        Me.fraMoneda.SuspendLayout()
        Me.fraFecha.SuspendLayout()
        Me.fraEntregarEn.SuspendLayout()
        Me.fraOtrosDatos.SuspendLayout()
        Me.fraCostos.SuspendLayout()
        Me.fraEntrega.SuspendLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraOrden, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblOrden, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnAsignarCodigos
        '
        Me.btnAsignarCodigos.BackColor = System.Drawing.SystemColors.Control
        Me.btnAsignarCodigos.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAsignarCodigos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAsignarCodigos.Location = New System.Drawing.Point(280, 552)
        Me.btnAsignarCodigos.Name = "btnAsignarCodigos"
        Me.btnAsignarCodigos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAsignarCodigos.Size = New System.Drawing.Size(137, 25)
        Me.btnAsignarCodigos.TabIndex = 49
        Me.btnAsignarCodigos.Text = "Asignar Códigos"
        Me.ToolTip1.SetToolTip(Me.btnAsignarCodigos, "Dar de alta los artículos de la Orden en Inventario y Existencias")
        Me.btnAsignarCodigos.UseVisualStyleBackColor = False
        '
        'btnProv
        '
        Me.btnProv.BackColor = System.Drawing.SystemColors.Control
        Me.btnProv.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnProv.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnProv.Location = New System.Drawing.Point(490, 552)
        Me.btnProv.Name = "btnProv"
        Me.btnProv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnProv.Size = New System.Drawing.Size(137, 25)
        Me.btnProv.TabIndex = 56
        Me.btnProv.Text = "ABC de Proveedores"
        Me.ToolTip1.SetToolTip(Me.btnProv, "Dar de alta los artículos de la Orden en Inventario y Existencias")
        Me.btnProv.UseVisualStyleBackColor = False
        '
        'txtDesctoFinanciero
        '
        Me.txtDesctoFinanciero.AcceptsReturn = True
        Me.txtDesctoFinanciero.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtDesctoFinanciero.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDesctoFinanciero.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDesctoFinanciero.Location = New System.Drawing.Point(834, 88)
        Me.txtDesctoFinanciero.MaxLength = 0
        Me.txtDesctoFinanciero.Name = "txtDesctoFinanciero"
        Me.txtDesctoFinanciero.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDesctoFinanciero.Size = New System.Drawing.Size(49, 20)
        Me.txtDesctoFinanciero.TabIndex = 19
        Me.txtDesctoFinanciero.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDesctoFinanciero, "Porcentaje Adicional a la Factura")
        '
        'txtTipoCambioConciliado
        '
        Me.txtTipoCambioConciliado.AcceptsReturn = True
        Me.txtTipoCambioConciliado.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipoCambioConciliado.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambioConciliado.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambioConciliado.Location = New System.Drawing.Point(50, 19)
        Me.txtTipoCambioConciliado.MaxLength = 0
        Me.txtTipoCambioConciliado.Name = "txtTipoCambioConciliado"
        Me.txtTipoCambioConciliado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambioConciliado.Size = New System.Drawing.Size(57, 20)
        Me.txtTipoCambioConciliado.TabIndex = 46
        Me.txtTipoCambioConciliado.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambioConciliado, "Tipo de Cambio al Conciliar (de Dólares a Pesos)")
        '
        'txtTipoCambioEuroConciliado
        '
        Me.txtTipoCambioEuroConciliado.AcceptsReturn = True
        Me.txtTipoCambioEuroConciliado.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipoCambioEuroConciliado.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambioEuroConciliado.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambioEuroConciliado.Location = New System.Drawing.Point(50, 50)
        Me.txtTipoCambioEuroConciliado.MaxLength = 0
        Me.txtTipoCambioEuroConciliado.Name = "txtTipoCambioEuroConciliado"
        Me.txtTipoCambioEuroConciliado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambioEuroConciliado.Size = New System.Drawing.Size(57, 20)
        Me.txtTipoCambioEuroConciliado.TabIndex = 48
        Me.txtTipoCambioEuroConciliado.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambioEuroConciliado, "Tipo de Cambio al Conciliar (de Euros a Pesos)")
        '
        'txtPorcDescto
        '
        Me.txtPorcDescto.AcceptsReturn = True
        Me.txtPorcDescto.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtPorcDescto.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPorcDescto.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPorcDescto.Location = New System.Drawing.Point(618, 88)
        Me.txtPorcDescto.MaxLength = 0
        Me.txtPorcDescto.Name = "txtPorcDescto"
        Me.txtPorcDescto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPorcDescto.Size = New System.Drawing.Size(49, 20)
        Me.txtPorcDescto.TabIndex = 17
        Me.txtPorcDescto.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtPorcDescto, "Porcentaje de Descuento para la Orden de Compra")
        '
        'txtTasaIva
        '
        Me.txtTasaIva.AcceptsReturn = True
        Me.txtTasaIva.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtTasaIva.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTasaIva.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTasaIva.Location = New System.Drawing.Point(392, 88)
        Me.txtTasaIva.MaxLength = 0
        Me.txtTasaIva.Name = "txtTasaIva"
        Me.txtTasaIva.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTasaIva.Size = New System.Drawing.Size(49, 20)
        Me.txtTasaIva.TabIndex = 15
        Me.txtTasaIva.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTasaIva, "Porcentaje de IVA")
        '
        'txtRemision
        '
        Me.txtRemision.AcceptsReturn = True
        Me.txtRemision.BackColor = System.Drawing.SystemColors.Window
        Me.txtRemision.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRemision.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRemision.Location = New System.Drawing.Point(494, 128)
        Me.txtRemision.MaxLength = 10
        Me.txtRemision.Name = "txtRemision"
        Me.txtRemision.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRemision.Size = New System.Drawing.Size(121, 20)
        Me.txtRemision.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.txtRemision, "Revisión")
        '
        'txtPedido
        '
        Me.txtPedido.AcceptsReturn = True
        Me.txtPedido.BackColor = System.Drawing.SystemColors.Window
        Me.txtPedido.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPedido.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPedido.Location = New System.Drawing.Point(735, 128)
        Me.txtPedido.MaxLength = 10
        Me.txtPedido.Name = "txtPedido"
        Me.txtPedido.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPedido.Size = New System.Drawing.Size(121, 20)
        Me.txtPedido.TabIndex = 24
        Me.ToolTip1.SetToolTip(Me.txtPedido, "Pedido")
        '
        'txtTipoCambioEuro
        '
        Me.txtTipoCambioEuro.AcceptsReturn = True
        Me.txtTipoCambioEuro.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtTipoCambioEuro.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambioEuro.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambioEuro.Location = New System.Drawing.Point(258, 40)
        Me.txtTipoCambioEuro.MaxLength = 0
        Me.txtTipoCambioEuro.Name = "txtTipoCambioEuro"
        Me.txtTipoCambioEuro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambioEuro.Size = New System.Drawing.Size(57, 20)
        Me.txtTipoCambioEuro.TabIndex = 13
        Me.txtTipoCambioEuro.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambioEuro, "Tipo de Cambio (de Euros a Pesos)")
        '
        'txtTipoCambio
        '
        Me.txtTipoCambio.AcceptsReturn = True
        Me.txtTipoCambio.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtTipoCambio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambio.Location = New System.Drawing.Point(152, 40)
        Me.txtTipoCambio.MaxLength = 0
        Me.txtTipoCambio.Name = "txtTipoCambio"
        Me.txtTipoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambio.Size = New System.Drawing.Size(57, 20)
        Me.txtTipoCambio.TabIndex = 11
        Me.txtTipoCambio.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambio, "Tipo de Cambio (de Dólares a Pesos)")
        '
        '_optMoneda_2
        '
        Me._optMoneda_2.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_2, CType(2, Short))
        Me._optMoneda_2.Location = New System.Drawing.Point(232, 16)
        Me._optMoneda_2.Name = "_optMoneda_2"
        Me._optMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_2.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_2.TabIndex = 8
        Me._optMoneda_2.TabStop = True
        Me._optMoneda_2.Text = "Euros"
        Me.ToolTip1.SetToolTip(Me._optMoneda_2, "Modeda de Compra (Euros)")
        Me._optMoneda_2.UseVisualStyleBackColor = False
        '
        '_optMoneda_1
        '
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_1, CType(1, Short))
        Me._optMoneda_1.Location = New System.Drawing.Point(136, 16)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_1.TabIndex = 7
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me._optMoneda_1, "Modeda de Compra (Pesos)")
        Me._optMoneda_1.UseVisualStyleBackColor = False
        '
        '_optMoneda_0
        '
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_0, CType(0, Short))
        Me._optMoneda_0.Location = New System.Drawing.Point(32, 16)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_0.TabIndex = 6
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me._optMoneda_0, "Moneda de Compra (Dólares)")
        Me._optMoneda_0.UseVisualStyleBackColor = False
        '
        'txtTotal
        '
        Me.txtTotal.AcceptsReturn = True
        Me.txtTotal.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.txtTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTotal.Location = New System.Drawing.Point(756, 552)
        Me.txtTotal.MaxLength = 0
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.ReadOnly = True
        Me.txtTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotal.Size = New System.Drawing.Size(129, 20)
        Me.txtTotal.TabIndex = 65
        Me.txtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTotal, "Total Neto")
        '
        'txtIVA
        '
        Me.txtIVA.AcceptsReturn = True
        Me.txtIVA.BackColor = System.Drawing.SystemColors.Info
        Me.txtIVA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtIVA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtIVA.Location = New System.Drawing.Point(756, 520)
        Me.txtIVA.MaxLength = 0
        Me.txtIVA.Name = "txtIVA"
        Me.txtIVA.ReadOnly = True
        Me.txtIVA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtIVA.Size = New System.Drawing.Size(129, 20)
        Me.txtIVA.TabIndex = 63
        Me.txtIVA.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtIVA, "Total de los Impuestos")
        '
        'txtDescuento
        '
        Me.txtDescuento.AcceptsReturn = True
        Me.txtDescuento.BackColor = System.Drawing.SystemColors.Info
        Me.txtDescuento.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescuento.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescuento.Location = New System.Drawing.Point(756, 488)
        Me.txtDescuento.MaxLength = 0
        Me.txtDescuento.Name = "txtDescuento"
        Me.txtDescuento.ReadOnly = True
        Me.txtDescuento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescuento.Size = New System.Drawing.Size(129, 20)
        Me.txtDescuento.TabIndex = 61
        Me.txtDescuento.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDescuento, "Total del Descuento")
        '
        'txtSubTotal
        '
        Me.txtSubTotal.AcceptsReturn = True
        Me.txtSubTotal.BackColor = System.Drawing.SystemColors.Info
        Me.txtSubTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSubTotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubTotal.Location = New System.Drawing.Point(756, 456)
        Me.txtSubTotal.MaxLength = 0
        Me.txtSubTotal.Name = "txtSubTotal"
        Me.txtSubTotal.ReadOnly = True
        Me.txtSubTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSubTotal.Size = New System.Drawing.Size(129, 20)
        Me.txtSubTotal.TabIndex = 59
        Me.txtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtSubTotal, "Total del Importe sin Impuestos")
        '
        'txtCostosIndirectos
        '
        Me.txtCostosIndirectos.AcceptsReturn = True
        Me.txtCostosIndirectos.BackColor = System.Drawing.SystemColors.Window
        Me.txtCostosIndirectos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCostosIndirectos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCostosIndirectos.Location = New System.Drawing.Point(16, 88)
        Me.txtCostosIndirectos.MaxLength = 0
        Me.txtCostosIndirectos.Name = "txtCostosIndirectos"
        Me.txtCostosIndirectos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCostosIndirectos.Size = New System.Drawing.Size(137, 20)
        Me.txtCostosIndirectos.TabIndex = 38
        Me.txtCostosIndirectos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCostosIndirectos, "Costos o Gastos Indirectos")
        '
        'txtCostoAdicional
        '
        Me.txtCostoAdicional.AcceptsReturn = True
        Me.txtCostoAdicional.BackColor = System.Drawing.SystemColors.Window
        Me.txtCostoAdicional.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCostoAdicional.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCostoAdicional.Location = New System.Drawing.Point(16, 40)
        Me.txtCostoAdicional.MaxLength = 0
        Me.txtCostoAdicional.Name = "txtCostoAdicional"
        Me.txtCostoAdicional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCostoAdicional.Size = New System.Drawing.Size(137, 20)
        Me.txtCostoAdicional.TabIndex = 36
        Me.txtCostoAdicional.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCostoAdicional, "Costo Adicional a la factura de compra")
        '
        'txtFolio
        '
        Me.txtFolio.AcceptsReturn = True
        Me.txtFolio.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolio.Location = New System.Drawing.Point(80, 12)
        Me.txtFolio.MaxLength = 19
        Me.txtFolio.Name = "txtFolio"
        Me.txtFolio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolio.Size = New System.Drawing.Size(161, 20)
        Me.txtFolio.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtFolio, "Folio de la Orden de Compra")
        '
        'fraApartado
        '
        Me.fraApartado.BackColor = System.Drawing.SystemColors.Control
        Me.fraApartado.Controls.Add(Me.txtFolioApartado)
        Me.fraApartado.Controls.Add(Me._lblOrden_2)
        Me.fraApartado.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraApartado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraApartado.Location = New System.Drawing.Point(374, 40)
        Me.fraApartado.Name = "fraApartado"
        Me.fraApartado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraApartado.Size = New System.Drawing.Size(281, 33)
        Me.fraApartado.TabIndex = 69
        Me.fraApartado.Visible = False
        '
        'txtFolioApartado
        '
        Me.txtFolioApartado.AcceptsReturn = True
        Me.txtFolioApartado.BackColor = System.Drawing.SystemColors.Info
        Me.txtFolioApartado.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioApartado.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioApartado.Location = New System.Drawing.Point(88, 9)
        Me.txtFolioApartado.MaxLength = 0
        Me.txtFolioApartado.Name = "txtFolioApartado"
        Me.txtFolioApartado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioApartado.Size = New System.Drawing.Size(129, 20)
        Me.txtFolioApartado.TabIndex = 71
        '
        '_lblOrden_2
        '
        Me._lblOrden_2.AutoSize = True
        Me._lblOrden_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_2, CType(2, Short))
        Me._lblOrden_2.Location = New System.Drawing.Point(8, 13)
        Me._lblOrden_2.Name = "_lblOrden_2"
        Me._lblOrden_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_2.Size = New System.Drawing.Size(81, 13)
        Me._lblOrden_2.TabIndex = 70
        Me._lblOrden_2.Text = "Folio Apartado :"
        '
        '_fraOrden_0
        '
        Me._fraOrden_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraOrden_0.Controls.Add(Me.txtTipoCambioConciliado)
        Me._fraOrden_0.Controls.Add(Me.txtTipoCambioEuroConciliado)
        Me._fraOrden_0.Controls.Add(Me.Label4)
        Me._fraOrden_0.Controls.Add(Me.Label3)
        Me._fraOrden_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraOrden.SetIndex(Me._fraOrden_0, CType(0, Short))
        Me._fraOrden_0.Location = New System.Drawing.Point(280, 464)
        Me._fraOrden_0.Name = "_fraOrden_0"
        Me._fraOrden_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraOrden_0.Size = New System.Drawing.Size(177, 81)
        Me._fraOrden_0.TabIndex = 44
        Me._fraOrden_0.TabStop = False
        Me._fraOrden_0.Text = "Tipo de Cambio al Conciliar"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(18, 23)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(32, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Dólar"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(18, 54)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(29, 13)
        Me.Label3.TabIndex = 47
        Me.Label3.Text = "Euro"
        '
        '_fraOrden_3
        '
        Me._fraOrden_3.BackColor = System.Drawing.SystemColors.Control
        Me._fraOrden_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraOrden.SetIndex(Me._fraOrden_3, CType(3, Short))
        Me._fraOrden_3.Location = New System.Drawing.Point(344, 117)
        Me._fraOrden_3.Name = "_fraOrden_3"
        Me._fraOrden_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraOrden_3.Size = New System.Drawing.Size(541, 2)
        Me._fraOrden_3.TabIndex = 20
        Me._fraOrden_3.TabStop = False
        '
        'fraMoneda
        '
        Me.fraMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.fraMoneda.Controls.Add(Me.txtTipoCambioEuro)
        Me.fraMoneda.Controls.Add(Me.txtTipoCambio)
        Me.fraMoneda.Controls.Add(Me._optMoneda_2)
        Me.fraMoneda.Controls.Add(Me._optMoneda_1)
        Me.fraMoneda.Controls.Add(Me._optMoneda_0)
        Me.fraMoneda.Controls.Add(Me.Label1)
        Me.fraMoneda.Controls.Add(Me.lblEuro)
        Me.fraMoneda.Controls.Add(Me.lblDolar)
        Me.fraMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraMoneda.Location = New System.Drawing.Point(8, 77)
        Me.fraMoneda.Name = "fraMoneda"
        Me.fraMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMoneda.Size = New System.Drawing.Size(329, 73)
        Me.fraMoneda.TabIndex = 5
        Me.fraMoneda.TabStop = False
        Me.fraMoneda.Text = "Moneda ..."
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Tipo de Cambio"
        '
        'lblEuro
        '
        Me.lblEuro.AutoSize = True
        Me.lblEuro.BackColor = System.Drawing.SystemColors.Control
        Me.lblEuro.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEuro.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEuro.Location = New System.Drawing.Point(226, 44)
        Me.lblEuro.Name = "lblEuro"
        Me.lblEuro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEuro.Size = New System.Drawing.Size(29, 13)
        Me.lblEuro.TabIndex = 12
        Me.lblEuro.Text = "Euro"
        '
        'lblDolar
        '
        Me.lblDolar.AutoSize = True
        Me.lblDolar.BackColor = System.Drawing.SystemColors.Control
        Me.lblDolar.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDolar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDolar.Location = New System.Drawing.Point(118, 44)
        Me.lblDolar.Name = "lblDolar"
        Me.lblDolar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDolar.Size = New System.Drawing.Size(32, 13)
        Me.lblDolar.TabIndex = 10
        Me.lblDolar.Text = "Dólar"
        '
        'fraFecha
        '
        Me.fraFecha.BackColor = System.Drawing.SystemColors.Control
        Me.fraFecha.Controls.Add(Me.dtpFecha)
        Me.fraFecha.Controls.Add(Me._lblOrden_3)
        Me.fraFecha.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraFecha.Enabled = False
        Me.fraFecha.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraFecha.Location = New System.Drawing.Point(686, 16)
        Me.fraFecha.Name = "fraFecha"
        Me.fraFecha.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraFecha.Size = New System.Drawing.Size(210, 25)
        Me.fraFecha.TabIndex = 66
        '
        'dtpFecha
        '
        Me.dtpFecha.Location = New System.Drawing.Point(87, 0)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(113, 20)
        Me.dtpFecha.TabIndex = 68
        '
        '_lblOrden_3
        '
        Me._lblOrden_3.AutoSize = True
        Me._lblOrden_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_3, CType(3, Short))
        Me._lblOrden_3.Location = New System.Drawing.Point(39, 4)
        Me._lblOrden_3.Name = "_lblOrden_3"
        Me._lblOrden_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_3.Size = New System.Drawing.Size(37, 13)
        Me._lblOrden_3.TabIndex = 67
        Me._lblOrden_3.Text = "Fecha"
        '
        '_fraOrden_5
        '
        Me._fraOrden_5.BackColor = System.Drawing.SystemColors.Control
        Me._fraOrden_5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraOrden.SetIndex(Me._fraOrden_5, CType(5, Short))
        Me._fraOrden_5.Location = New System.Drawing.Point(655, 440)
        Me._fraOrden_5.Name = "_fraOrden_5"
        Me._fraOrden_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraOrden_5.Size = New System.Drawing.Size(2, 137)
        Me._fraOrden_5.TabIndex = 57
        Me._fraOrden_5.TabStop = False
        '
        'fraEntregarEn
        '
        Me.fraEntregarEn.BackColor = System.Drawing.SystemColors.Control
        Me.fraEntregarEn.Controls.Add(Me.rtEntregaren)
        Me.fraEntregarEn.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraEntregarEn.Location = New System.Drawing.Point(8, 464)
        Me.fraEntregarEn.Name = "fraEntregarEn"
        Me.fraEntregarEn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraEntregarEn.Size = New System.Drawing.Size(233, 113)
        Me.fraEntregarEn.TabIndex = 42
        Me.fraEntregarEn.TabStop = False
        Me.fraEntregarEn.Text = "Entregar en ..."
        '
        'rtEntregaren
        '
        Me.rtEntregaren.Location = New System.Drawing.Point(16, 24)
        Me.rtEntregaren.Name = "rtEntregaren"
        Me.rtEntregaren.Size = New System.Drawing.Size(201, 73)
        Me.rtEntregaren.TabIndex = 43
        Me.rtEntregaren.Text = ""
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(31, 352)
        Me.txtFlex.MaxLength = 50
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(81, 20)
        Me.txtFlex.TabIndex = 40
        Me.txtFlex.Visible = False
        '
        'fraOtrosDatos
        '
        Me.fraOtrosDatos.BackColor = System.Drawing.SystemColors.Control
        Me.fraOtrosDatos.Controls.Add(Me.txtOtrosDatos)
        Me.fraOtrosDatos.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraOtrosDatos.Location = New System.Drawing.Point(8, 159)
        Me.fraOtrosDatos.Name = "fraOtrosDatos"
        Me.fraOtrosDatos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraOtrosDatos.Size = New System.Drawing.Size(329, 121)
        Me.fraOtrosDatos.TabIndex = 25
        Me.fraOtrosDatos.TabStop = False
        Me.fraOtrosDatos.Text = "Datos Adicionales"
        '
        'txtOtrosDatos
        '
        Me.txtOtrosDatos.BackColor = System.Drawing.SystemColors.Info
        Me.txtOtrosDatos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtOtrosDatos.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtOtrosDatos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtOtrosDatos.Location = New System.Drawing.Point(16, 24)
        Me.txtOtrosDatos.Name = "txtOtrosDatos"
        Me.txtOtrosDatos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOtrosDatos.Size = New System.Drawing.Size(297, 81)
        Me.txtOtrosDatos.TabIndex = 26
        Me.txtOtrosDatos.Text = "OtrosDatos"
        '
        'fraCostos
        '
        Me.fraCostos.BackColor = System.Drawing.SystemColors.Control
        Me.fraCostos.Controls.Add(Me.txtCostosIndirectos)
        Me.fraCostos.Controls.Add(Me.txtCostoAdicional)
        Me.fraCostos.Controls.Add(Me._lblOrden_8)
        Me.fraCostos.Controls.Add(Me._lblOrden_7)
        Me.fraCostos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraCostos.Location = New System.Drawing.Point(713, 159)
        Me.fraCostos.Name = "fraCostos"
        Me.fraCostos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCostos.Size = New System.Drawing.Size(169, 121)
        Me.fraCostos.TabIndex = 34
        Me.fraCostos.TabStop = False
        '
        '_lblOrden_8
        '
        Me._lblOrden_8.AutoSize = True
        Me._lblOrden_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_8, CType(8, Short))
        Me._lblOrden_8.Location = New System.Drawing.Point(16, 72)
        Me._lblOrden_8.Name = "_lblOrden_8"
        Me._lblOrden_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_8.Size = New System.Drawing.Size(88, 13)
        Me._lblOrden_8.TabIndex = 37
        Me._lblOrden_8.Text = "Costos Indirectos"
        '
        '_lblOrden_7
        '
        Me._lblOrden_7.AutoSize = True
        Me._lblOrden_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_7, CType(7, Short))
        Me._lblOrden_7.Location = New System.Drawing.Point(16, 24)
        Me._lblOrden_7.Name = "_lblOrden_7"
        Me._lblOrden_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_7.Size = New System.Drawing.Size(80, 13)
        Me._lblOrden_7.TabIndex = 35
        Me._lblOrden_7.Text = "Costo Adicional"
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(80, 48)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(249, 21)
        Me.dbcProveedor.TabIndex = 4
        '
        'fraEntrega
        '
        Me.fraEntrega.BackColor = System.Drawing.SystemColors.Control
        Me.fraEntrega.Controls.Add(Me.dtpFechaEntrega)
        Me.fraEntrega.Controls.Add(Me.dbcOrigen)
        Me.fraEntrega.Controls.Add(Me.dbcGrupo)
        Me.fraEntrega.Controls.Add(Me._lblOrden_6)
        Me.fraEntrega.Controls.Add(Me._lblOrden_5)
        Me.fraEntrega.Controls.Add(Me._lblOrden_4)
        Me.fraEntrega.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraEntrega.Location = New System.Drawing.Point(389, 159)
        Me.fraEntrega.Name = "fraEntrega"
        Me.fraEntrega.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraEntrega.Size = New System.Drawing.Size(273, 121)
        Me.fraEntrega.TabIndex = 27
        Me.fraEntrega.TabStop = False
        '
        'dtpFechaEntrega
        '
        Me.dtpFechaEntrega.Location = New System.Drawing.Point(152, 16)
        Me.dtpFechaEntrega.Name = "dtpFechaEntrega"
        Me.dtpFechaEntrega.Size = New System.Drawing.Size(113, 20)
        Me.dtpFechaEntrega.TabIndex = 29
        '
        'dbcOrigen
        '
        Me.dbcOrigen.Location = New System.Drawing.Point(56, 52)
        Me.dbcOrigen.Name = "dbcOrigen"
        Me.dbcOrigen.Size = New System.Drawing.Size(209, 21)
        Me.dbcOrigen.TabIndex = 31
        '
        'dbcGrupo
        '
        Me.dbcGrupo.Location = New System.Drawing.Point(56, 88)
        Me.dbcGrupo.Name = "dbcGrupo"
        Me.dbcGrupo.Size = New System.Drawing.Size(209, 21)
        Me.dbcGrupo.TabIndex = 33
        '
        '_lblOrden_6
        '
        Me._lblOrden_6.AutoSize = True
        Me._lblOrden_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_6, CType(6, Short))
        Me._lblOrden_6.Location = New System.Drawing.Point(8, 92)
        Me._lblOrden_6.Name = "_lblOrden_6"
        Me._lblOrden_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_6.Size = New System.Drawing.Size(36, 13)
        Me._lblOrden_6.TabIndex = 32
        Me._lblOrden_6.Text = "Grupo"
        '
        '_lblOrden_5
        '
        Me._lblOrden_5.AutoSize = True
        Me._lblOrden_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_5, CType(5, Short))
        Me._lblOrden_5.Location = New System.Drawing.Point(8, 56)
        Me._lblOrden_5.Name = "_lblOrden_5"
        Me._lblOrden_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_5.Size = New System.Drawing.Size(38, 13)
        Me._lblOrden_5.TabIndex = 30
        Me._lblOrden_5.Text = "Origen"
        '
        '_lblOrden_4
        '
        Me._lblOrden_4.AutoSize = True
        Me._lblOrden_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_4, CType(4, Short))
        Me._lblOrden_4.Location = New System.Drawing.Point(56, 16)
        Me._lblOrden_4.Name = "_lblOrden_4"
        Me._lblOrden_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_4.Size = New System.Drawing.Size(92, 13)
        Me._lblOrden_4.TabIndex = 28
        Me._lblOrden_4.Text = "Fecha de Entrega"
        '
        'mshFlex
        '
        Me.mshFlex.DataSource = Nothing
        Me.mshFlex.Location = New System.Drawing.Point(8, 288)
        Me.mshFlex.Name = "mshFlex"
        Me.mshFlex.OcxState = CType(resources.GetObject("mshFlex.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mshFlex.Size = New System.Drawing.Size(875, 137)
        Me.mshFlex.TabIndex = 39
        '
        'lblDescProv
        '
        Me.lblDescProv.BackColor = System.Drawing.SystemColors.Info
        Me.lblDescProv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDescProv.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDescProv.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblDescProv.Location = New System.Drawing.Point(461, 433)
        Me.lblDescProv.Name = "lblDescProv"
        Me.lblDescProv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDescProv.Size = New System.Drawing.Size(168, 21)
        Me.lblDescProv.TabIndex = 72
        Me.lblDescProv.Text = "Descripcion"
        '
        'lblDesctoFinanciero
        '
        Me.lblDesctoFinanciero.AutoSize = True
        Me.lblDesctoFinanciero.BackColor = System.Drawing.SystemColors.Control
        Me.lblDesctoFinanciero.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDesctoFinanciero.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDesctoFinanciero.Location = New System.Drawing.Point(722, 92)
        Me.lblDesctoFinanciero.Name = "lblDesctoFinanciero"
        Me.lblDesctoFinanciero.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDesctoFinanciero.Size = New System.Drawing.Size(113, 13)
        Me.lblDesctoFinanciero.TabIndex = 18
        Me.lblDesctoFinanciero.Text = "Descto. Financiero (%)"
        '
        'lblPorcDescto
        '
        Me.lblPorcDescto.AutoSize = True
        Me.lblPorcDescto.BackColor = System.Drawing.SystemColors.Control
        Me.lblPorcDescto.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPorcDescto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPorcDescto.Location = New System.Drawing.Point(498, 92)
        Me.lblPorcDescto.Name = "lblPorcDescto"
        Me.lblPorcDescto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPorcDescto.Size = New System.Drawing.Size(123, 13)
        Me.lblPorcDescto.TabIndex = 16
        Me.lblPorcDescto.Text = "Descto. por Volumen (%)"
        '
        'lblTasaIva
        '
        Me.lblTasaIva.AutoSize = True
        Me.lblTasaIva.BackColor = System.Drawing.SystemColors.Control
        Me.lblTasaIva.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTasaIva.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTasaIva.Location = New System.Drawing.Point(352, 92)
        Me.lblTasaIva.Name = "lblTasaIva"
        Me.lblTasaIva.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTasaIva.Size = New System.Drawing.Size(41, 13)
        Me.lblTasaIva.TabIndex = 14
        Me.lblTasaIva.Text = "IVA (%)"
        '
        'lblRemision
        '
        Me.lblRemision.AutoSize = True
        Me.lblRemision.BackColor = System.Drawing.SystemColors.Control
        Me.lblRemision.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRemision.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRemision.Location = New System.Drawing.Point(438, 132)
        Me.lblRemision.Name = "lblRemision"
        Me.lblRemision.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRemision.Size = New System.Drawing.Size(50, 13)
        Me.lblRemision.TabIndex = 21
        Me.lblRemision.Text = "Remisión"
        '
        'lblPedido
        '
        Me.lblPedido.AutoSize = True
        Me.lblPedido.BackColor = System.Drawing.SystemColors.Control
        Me.lblPedido.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPedido.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPedido.Location = New System.Drawing.Point(695, 132)
        Me.lblPedido.Name = "lblPedido"
        Me.lblPedido.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPedido.Size = New System.Drawing.Size(40, 13)
        Me.lblPedido.TabIndex = 23
        Me.lblPedido.Text = "Pedido"
        '
        'txtDescripcion
        '
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Info
        Me.txtDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDescripcion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtDescripcion.Location = New System.Drawing.Point(8, 433)
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(449, 21)
        Me.txtDescripcion.TabIndex = 41
        Me.txtDescripcion.Text = "Descripcion"
        '
        'lblCR
        '
        Me.lblCR.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(200, Byte), Integer), CType(CType(145, Byte), Integer))
        Me.lblCR.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCR.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCR.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCR.Location = New System.Drawing.Point(490, 520)
        Me.lblCR.Name = "lblCR"
        Me.lblCR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCR.Size = New System.Drawing.Size(17, 17)
        Me.lblCR.TabIndex = 54
        '
        '_lblOrden_17
        '
        Me._lblOrden_17.AutoSize = True
        Me._lblOrden_17.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_17, CType(17, Short))
        Me._lblOrden_17.Location = New System.Drawing.Point(514, 520)
        Me._lblOrden_17.Name = "_lblOrden_17"
        Me._lblOrden_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_17.Size = New System.Drawing.Size(116, 13)
        Me._lblOrden_17.TabIndex = 55
        Me._lblOrden_17.Text = "Conciliados/Resurtidos"
        '
        '_lblOrden_16
        '
        Me._lblOrden_16.AutoSize = True
        Me._lblOrden_16.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_16, CType(16, Short))
        Me._lblOrden_16.Location = New System.Drawing.Point(514, 496)
        Me._lblOrden_16.Name = "_lblOrden_16"
        Me._lblOrden_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_16.Size = New System.Drawing.Size(57, 13)
        Me._lblOrden_16.TabIndex = 53
        Me._lblOrden_16.Text = "Resurtidos"
        '
        'lblResurtido
        '
        Me.lblResurtido.BackColor = System.Drawing.Color.FromArgb(CType(CType(173, Byte), Integer), CType(CType(226, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblResurtido.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblResurtido.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblResurtido.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblResurtido.Location = New System.Drawing.Point(490, 496)
        Me.lblResurtido.Name = "lblResurtido"
        Me.lblResurtido.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblResurtido.Size = New System.Drawing.Size(17, 17)
        Me.lblResurtido.TabIndex = 52
        '
        '_lblOrden_15
        '
        Me._lblOrden_15.AutoSize = True
        Me._lblOrden_15.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_15, CType(15, Short))
        Me._lblOrden_15.Location = New System.Drawing.Point(514, 474)
        Me._lblOrden_15.Name = "_lblOrden_15"
        Me._lblOrden_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_15.Size = New System.Drawing.Size(61, 13)
        Me._lblOrden_15.TabIndex = 51
        Me._lblOrden_15.Text = "Conciliados"
        '
        'lblConciliado
        '
        Me.lblConciliado.BackColor = System.Drawing.Color.FromArgb(CType(CType(157, Byte), Integer), CType(CType(172, Byte), Integer), CType(CType(189, Byte), Integer))
        Me.lblConciliado.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblConciliado.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblConciliado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblConciliado.Location = New System.Drawing.Point(490, 472)
        Me.lblConciliado.Name = "lblConciliado"
        Me.lblConciliado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblConciliado.Size = New System.Drawing.Size(17, 17)
        Me.lblConciliado.TabIndex = 50
        '
        'lblEstatus
        '
        Me.lblEstatus.BackColor = System.Drawing.SystemColors.Info
        Me.lblEstatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEstatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEstatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblEstatus.Location = New System.Drawing.Point(374, 12)
        Me.lblEstatus.Name = "lblEstatus"
        Me.lblEstatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEstatus.Size = New System.Drawing.Size(181, 21)
        Me.lblEstatus.TabIndex = 0
        Me.lblEstatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblEstatus.Visible = False
        '
        '_lblOrden_14
        '
        Me._lblOrden_14.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_14, CType(14, Short))
        Me._lblOrden_14.Location = New System.Drawing.Point(684, 556)
        Me._lblOrden_14.Name = "_lblOrden_14"
        Me._lblOrden_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_14.Size = New System.Drawing.Size(59, 13)
        Me._lblOrden_14.TabIndex = 64
        Me._lblOrden_14.Text = "Total"
        Me._lblOrden_14.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblOrden_13
        '
        Me._lblOrden_13.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_13, CType(13, Short))
        Me._lblOrden_13.Location = New System.Drawing.Point(684, 524)
        Me._lblOrden_13.Name = "_lblOrden_13"
        Me._lblOrden_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_13.Size = New System.Drawing.Size(59, 13)
        Me._lblOrden_13.TabIndex = 62
        Me._lblOrden_13.Text = "IVA"
        Me._lblOrden_13.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblOrden_12
        '
        Me._lblOrden_12.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_12, CType(12, Short))
        Me._lblOrden_12.Location = New System.Drawing.Point(684, 492)
        Me._lblOrden_12.Name = "_lblOrden_12"
        Me._lblOrden_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_12.Size = New System.Drawing.Size(59, 13)
        Me._lblOrden_12.TabIndex = 60
        Me._lblOrden_12.Text = "Descuento"
        Me._lblOrden_12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblOrden_11
        '
        Me._lblOrden_11.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_11, CType(11, Short))
        Me._lblOrden_11.Location = New System.Drawing.Point(684, 460)
        Me._lblOrden_11.Name = "_lblOrden_11"
        Me._lblOrden_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_11.Size = New System.Drawing.Size(59, 13)
        Me._lblOrden_11.TabIndex = 58
        Me._lblOrden_11.Text = "SubTotal"
        Me._lblOrden_11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblOrden_1
        '
        Me._lblOrden_1.AutoSize = True
        Me._lblOrden_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_1, CType(1, Short))
        Me._lblOrden_1.Location = New System.Drawing.Point(8, 52)
        Me._lblOrden_1.Name = "_lblOrden_1"
        Me._lblOrden_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_1.Size = New System.Drawing.Size(56, 13)
        Me._lblOrden_1.TabIndex = 3
        Me._lblOrden_1.Text = "Proveedor"
        '
        '_lblOrden_0
        '
        Me._lblOrden_0.AutoSize = True
        Me._lblOrden_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblOrden_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblOrden_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrden.SetIndex(Me._lblOrden_0, CType(0, Short))
        Me._lblOrden_0.Location = New System.Drawing.Point(8, 16)
        Me._lblOrden_0.Name = "_lblOrden_0"
        Me._lblOrden_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblOrden_0.Size = New System.Drawing.Size(29, 13)
        Me._lblOrden_0.TabIndex = 1
        Me._lblOrden_0.Text = "Folio"
        '
        'optMoneda
        '
        '
        'btnCancelar
        '
        Me.btnCancelar.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancelar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancelar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancelar.Location = New System.Drawing.Point(269, 602)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancelar.Size = New System.Drawing.Size(109, 36)
        Me.btnCancelar.TabIndex = 141
        Me.btnCancelar.Text = "&Cancelar"
        Me.btnCancelar.UseVisualStyleBackColor = False
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(154, 602)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 140
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(39, 602)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 139
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(384, 602)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 138
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmCXPOrdenCompra
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(893, 650)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.fraApartado)
        Me.Controls.Add(Me.btnAsignarCodigos)
        Me.Controls.Add(Me.btnProv)
        Me.Controls.Add(Me.txtDesctoFinanciero)
        Me.Controls.Add(Me._fraOrden_0)
        Me.Controls.Add(Me.txtPorcDescto)
        Me.Controls.Add(Me.txtTasaIva)
        Me.Controls.Add(Me.txtRemision)
        Me.Controls.Add(Me.txtPedido)
        Me.Controls.Add(Me._fraOrden_3)
        Me.Controls.Add(Me.fraMoneda)
        Me.Controls.Add(Me.fraFecha)
        Me.Controls.Add(Me._fraOrden_5)
        Me.Controls.Add(Me.txtTotal)
        Me.Controls.Add(Me.txtIVA)
        Me.Controls.Add(Me.txtDescuento)
        Me.Controls.Add(Me.txtSubTotal)
        Me.Controls.Add(Me.fraEntregarEn)
        Me.Controls.Add(Me.txtFlex)
        Me.Controls.Add(Me.fraOtrosDatos)
        Me.Controls.Add(Me.fraCostos)
        Me.Controls.Add(Me.txtFolio)
        Me.Controls.Add(Me.dbcProveedor)
        Me.Controls.Add(Me.fraEntrega)
        Me.Controls.Add(Me.mshFlex)
        Me.Controls.Add(Me.lblDescProv)
        Me.Controls.Add(Me.lblDesctoFinanciero)
        Me.Controls.Add(Me.lblPorcDescto)
        Me.Controls.Add(Me.lblTasaIva)
        Me.Controls.Add(Me.lblRemision)
        Me.Controls.Add(Me.lblPedido)
        Me.Controls.Add(Me.txtDescripcion)
        Me.Controls.Add(Me.lblCR)
        Me.Controls.Add(Me._lblOrden_17)
        Me.Controls.Add(Me._lblOrden_16)
        Me.Controls.Add(Me.lblResurtido)
        Me.Controls.Add(Me._lblOrden_15)
        Me.Controls.Add(Me.lblConciliado)
        Me.Controls.Add(Me.lblEstatus)
        Me.Controls.Add(Me._lblOrden_14)
        Me.Controls.Add(Me._lblOrden_13)
        Me.Controls.Add(Me._lblOrden_12)
        Me.Controls.Add(Me._lblOrden_11)
        Me.Controls.Add(Me._lblOrden_1)
        Me.Controls.Add(Me._lblOrden_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(167, 136)
        Me.MaximizeBox = False
        Me.Name = "frmCXPOrdenCompra"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Orden de Compra"
        Me.fraApartado.ResumeLayout(False)
        Me.fraApartado.PerformLayout()
        Me._fraOrden_0.ResumeLayout(False)
        Me._fraOrden_0.PerformLayout()
        Me.fraMoneda.ResumeLayout(False)
        Me.fraMoneda.PerformLayout()
        Me.fraFecha.ResumeLayout(False)
        Me.fraFecha.PerformLayout()
        Me.fraEntregarEn.ResumeLayout(False)
        Me.fraOtrosDatos.ResumeLayout(False)
        Me.fraCostos.ResumeLayout(False)
        Me.fraCostos.PerformLayout()
        Me.fraEntrega.ResumeLayout(False)
        Me.fraEntrega.PerformLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraOrden, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblOrden, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

End Class