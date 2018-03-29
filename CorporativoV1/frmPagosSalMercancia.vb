Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports VB6 = Microsoft.VisualBasic
Imports ADODB
Public Class frmPagosSalMercancia
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtEsDolarFPCambio As System.Windows.Forms.TextBox
    Public WithEvents txtCodFormaPago As System.Windows.Forms.TextBox
    Public WithEvents dbcMoneda As System.Windows.Forms.ComboBox
    Public WithEvents _lblEtiqueta_3 As System.Windows.Forms.Label
    Public WithEvents fraMoneda As System.Windows.Forms.Panel
    Public WithEvents txtFormaOrigen As System.Windows.Forms.TextBox
    Public WithEvents txtImporte As System.Windows.Forms.TextBox
    Public WithEvents txtdoCambio As System.Windows.Forms.Label
    Public WithEvents txtdoTotalPago As System.Windows.Forms.Label
    Public WithEvents txtdoAPagar As System.Windows.Forms.Label
    Public WithEvents txtdoAPagar4Decimales As System.Windows.Forms.Label
    Public WithEvents txtdoTotalPago4Decimales As System.Windows.Forms.Label
    Public WithEvents txtdoCambio4Decimales As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_28 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_29 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_30 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_31 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_32 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_34 As System.Windows.Forms.Label
    Public WithEvents _Marco_3 As System.Windows.Forms.GroupBox
    Public WithEvents txtDolar As System.Windows.Forms.TextBox
    Public WithEvents txtTotal As System.Windows.Forms.Label
    Public WithEvents txtIVA As System.Windows.Forms.Label
    Public WithEvents txtDescuento As System.Windows.Forms.Label
    Public WithEvents txtSubtotal As System.Windows.Forms.Label
    Public WithEvents txtSubtotal4Decimales As System.Windows.Forms.Label
    Public WithEvents txtDescuento4Decimales As System.Windows.Forms.Label
    Public WithEvents txtIVA4Decimales As System.Windows.Forms.Label
    Public WithEvents txtTotal4Decimales As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_36 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_22 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_16 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_15 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_14 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_2 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_1 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_0 As System.Windows.Forms.Label
    Public WithEvents _Marco_1 As System.Windows.Forms.GroupBox
    Public WithEvents msgFormasPago As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents txtmnCambio As System.Windows.Forms.Label
    Public WithEvents txtmnTotalPago As System.Windows.Forms.Label
    Public WithEvents txtmnAPagar As System.Windows.Forms.Label
    Public WithEvents txtmnAPagar4Decimales As System.Windows.Forms.Label
    Public WithEvents txtmnTotalPago4Decimales As System.Windows.Forms.Label
    Public WithEvents txtmnCambio4Decimales As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_4 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_18 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_20 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_19 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_13 As System.Windows.Forms.Label
    Public WithEvents _lblEtiqueta_12 As System.Windows.Forms.Label
    Public WithEvents _Marco_2 As System.Windows.Forms.GroupBox
    Public WithEvents _lblEtiqueta_26 As System.Windows.Forms.Label
    Public WithEvents _Marco_0 As System.Windows.Forms.GroupBox
    Public WithEvents Marco As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblEtiqueta As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    'Programa: Pagos correspondientes a Salida de Mercancía
    'Autor: Rosaura Torres López
    'Fecha de Creación: 30/Mayo/2003


    ' VALORES DE LAS COLUMNAS DEL GRID
    '   0 .- Tecla Rápida
    '   1 .- Descripción
    '   2 .- Importe
    '   3 .- Es Dolar
    '   4 .- Es Cheque
    '   5 .- RequerirDocto
    '   6 .- REquerirAutorizacion
    '   7 .- REstringir Cambio
    '   8 .- Considerar para facturacion
    '   9 .- Considerar para Retiror
    '  10.- Es Tarjeta
    '  11 .- Descontar comision Bancaria
    '  12 .- Porcentaje de Comision
    '  13 .- Porcentaje de iva de comision
    '  14 .- Pago de interes por promocion
    '  15 .- Porcentaje de Intereses
    '  16 .- Porcentaje Iva Intereses
    '  17 .- Codigo Forma de Pago
    '  18 .- vale de Devolucion
    '  19 .- Importe de Comision Bancaria
    '  20 .- Importe de Intereses por Promoción
    '  21 .- Código del Banco

    Const C_ColTECLARAPIDA As Integer = 0
    Const C_COLDESCRIPCION As Integer = 1
    ''''--------------------------------
    Const C_ColIMPORTE As Integer = 2 ' No Cambiar esta constante, porque es usada en el Registro de Cheque, Tarjetas, NC, Vales, para Obtener el Importe de la Forma de Pago en uso, que el usuaaio proporcionó.
    'La cual se comparará con el importe que el usuario esta proporcionando en el registro de documentos.
    'Si se  quiere cambiar, Checar en lor Form. mencionados en el Proc. ValidadarDatos , cuando se asigna valor a la Var. ImporteFP, y cmabiar la columna.
    ''''--------------------------------
    Const C_ColESDOLAR As Integer = 3
    Const C_ColESCHEQUE As Integer = 4
    Const C_ColREQUERIRDOCTO As Integer = 5
    Const C_ColREQUERIRAUT As Integer = 6
    Const C_ColRESTRINGIRCAMBIO As Integer = 7
    Const C_ColCONSIDERARFACT As Integer = 8
    Const C_ColCONSIDERARRETIRO As Integer = 9
    ''''--------------------------------
    Const C_ColESTARJETA As Integer = 10 ''IMPORTANTE, no cambiar este numero a la constantes de Tarjeta, porque es usada en el FOrmulario de Registro de Tarjetas
    'Si se cambia aqui, modificar en la funcion validar datos del form. antes mencionado. cuando se accesa al MSG de las formas de Pago
    ''''--------------------------------
    Const C_ColDESCCOMBANC As Integer = 11
    Const C_COLPORCCOMISION As Integer = 12
    Const C_ColPORCIVACOMISION As Integer = 13
    ''''--------------------------------
    Const C_ColPAGOINTXPROM As Integer = 14
    Const C_ColPORCINTERESES As Integer = 15
    Const C_ColPORCIVAINTERESES As Integer = 16
    ''''--------------------------------
    Const C_ColCODFORMAPAGO As Integer = 17
    ''''--------------------------------
    Const C_ColESVALEDEVOLUCION As Integer = 18 ''IMPORTANTE, no cambiar este numero a la constantes de Devolucion, porque es usada en el FOrmulario de Registro de Notas de Credito y Vales de Devolucion
    'Si se cambia aqui, modificar en la funcion validar datos del form. antes mencionado. cuando se accesa al MSG de las formas de Pago de Pagos
    ''''--------------------------------
    Const C_ColIMPCOMISIONBANCARIA As Integer = 19
    Const C_ColIMPINTERESESPROMOCION As Integer = 20
    ''''--------------------------------
    Const C_ColIMPORTESINREDONDEO As Integer = 21 ' No Cambiar esta constante, porque es usada en el Registro de Cheque, Tarjetas, NC, Vales, para Obtener el Importe de la Forma de Pago en uso, que el usuaaio proporcionó.
    'La cual se comparará con el importe que el usuario esta proporcionando en el registro de documentos.
    'Si se  quiere cambiar, Checar en lor Form. mencionados en el Proc. ValidadarDatos , cuando se asigna valor a la Var. ImporteFP, y cmabiar la columna.
    ''''--------------------------------
    ''''--------------------------------
    Const C_ColCodBanco As Integer = 22 'Codigo del Banco para el filtrado de promociones de tarjetas
    ''''--------------------------------

    Public ExistenFolioReg As Boolean
    Dim I As Integer
    Dim x As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim tecla As Integer
    Dim intCodFormaPago As Integer
    Dim LnContador As Integer
    Dim gnTotRen As Byte
    Dim TotCambioSi As Double
    Dim TotCambioNo As Double
    Dim TipoCambio As Double
    Dim Blanco, Rojo, Amarillo, Azul, Verde, Negro As Object

    'Variables para Guardar los  Importes Forma de PAgo
    Dim CodFormaPago As Integer
    Dim importe As Decimal
    Dim Banco As String
    Dim NoTarjeta As String
    Dim Autorizacion As String
    Dim NoCheque As String
    Dim FolioDev As String
    Dim ComisionBancaria As Decimal
    Dim InteresesPromocion As Decimal
    Dim Escheque As Boolean
    Dim EsTarjeta As Boolean
    Dim EsDevolucion As Boolean
    Dim NumPartidaIngresosFormaDePago As Integer 'Contiene el número consecutivo de Partida para guardar el Ingreso de una Forma de Pago
    Dim mblnImpteFPRCMayorTotalAPagar As Boolean 'Esta variable determina si el importe de las Formas de Pago que resringen el Cambio es MAyor que el Total a pagar por el cliente.
    Public ImporteTecleado As Decimal

    Sub MostrarInteresPromocion(ByRef PorcInteres As Decimal, ByRef PorcIvaInteres As Decimal, ByRef ImporteInteres As Decimal)
        On Error GoTo Merr
        'este proc, pone el Grid de Formas de pago, el importe de interés por promocion que se aplica al uso de tarjetas solamente.
        'Este importe se calcula en el from. de REg de tarjetas, y se pasa como parametros, para que se añadan al form. de Formas de Pago
        With msgFormasPago
            .set_TextMatrix(.Row, C_ColPORCINTERESES, PorcInteres)
            .set_TextMatrix(.Row, C_ColPORCIVAINTERESES, PorcIvaInteres)
            .set_TextMatrix(.Row, C_ColIMPINTERESESPROMOCION, ImporteInteres)
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function CalcularComisionBancaria(ByRef importe As Decimal, ByRef FormaPago As Integer) As Decimal
        'Este Función calcula el importe de Comision Bancaria para una Forma de PAgo.
        'siempre y cuando, se haya especificado en las caracteristicas de la Forma de Pago en uso.
        'Podrá ser usada, desde cualquier forma de Pago que requiera calcular la comision Bancaria (P.E. Cheque., Tarjeta, Vale, etc)
        'EL calculo se hace sobre el Importe Neto que se dió sobre es FOrma de Pago. (Que puede ser pesos o Dolares.)
        Dim DescComision As Boolean
        Dim PorcComision As Decimal
        Dim PorcIvaComision As Decimal
        Dim ComisionBancaria As Decimal
        ComisionBancaria = 0
        gStrSql = "SELECT * FROM CatFormasPago WHERE CodFormaPago = " & FormaPago

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount <> 0 Then
            DescComision = RsGral.Fields("DescontarComisionBanc").Value
            PorcComision = RsGral.Fields("PorcComision").Value
            PorcIvaComision = RsGral.Fields("PorcIvaComision").Value
            If DescComision = True Then
                'Obtener el Importe de Comisión Bancaria.
                ComisionBancaria = FormateoDecimales(importe * (PorcComision / 100))
                'Ahora sumar el Iva al Importe de Comisión Bancaria.
                ComisionBancaria = ComisionBancaria * (1 + (PorcIvaComision / 100))
            End If
            CalcularComisionBancaria = ComisionBancaria
        End If
    End Function

    Sub Encabezado()
        On Error GoTo Errores
        With msgFormasPago
            .set_ColWidth(C_ColTECLARAPIDA, 0, 250)
            .set_ColWidth(C_COLDESCRIPCION, 0, 3050)
            .set_ColWidth(C_ColIMPORTE, 0, 1400)
            .set_ColWidth(C_ColIMPORTESINREDONDEO, 0, 1)
            .set_ColWidth(C_ColESDOLAR, 0, 1)
            .set_ColWidth(C_ColESCHEQUE, 0, 1)
            .set_ColWidth(C_ColREQUERIRDOCTO, 0, 1)
            .set_ColWidth(C_ColREQUERIRAUT, 0, 1)
            .set_ColWidth(C_ColRESTRINGIRCAMBIO, 0, 1)
            .set_ColWidth(C_ColCONSIDERARFACT, 0, 1)
            .set_ColWidth(C_ColCONSIDERARRETIRO, 0, 1)
            .set_ColWidth(C_ColESTARJETA, 0, 1)
            .set_ColWidth(C_ColDESCCOMBANC, 0, 1)
            .set_ColWidth(C_COLPORCCOMISION, 0, 1)
            .set_ColWidth(C_ColPORCIVACOMISION, 0, 1)
            .set_ColWidth(C_ColPAGOINTXPROM, 0, 1)
            .set_ColWidth(C_ColPORCINTERESES, 0, 1)
            .set_ColWidth(C_ColPORCIVAINTERESES, 0, 1)
            .set_ColWidth(C_ColCODFORMAPAGO, 0, 1)
            .set_ColWidth(C_ColESVALEDEVOLUCION, 0, 1)
            .set_ColWidth(C_ColIMPCOMISIONBANCARIA, 0, 0) '1000
            .set_ColWidth(C_ColIMPINTERESESPROMOCION, 0, 0) '1000
            .set_ColWidth(C_ColCodBanco, 0, 0)

            .set_TextMatrix(0, C_ColTECLARAPIDA, " ")
            .set_TextMatrix(0, C_COLDESCRIPCION, "Formas de Pago")
            .set_TextMatrix(0, C_ColIMPORTE, "Importe")
            .set_TextMatrix(0, C_ColIMPORTESINREDONDEO, "Importe con Red")
            .set_TextMatrix(0, C_ColESDOLAR, "ES DOLAR")
            .set_TextMatrix(0, C_ColESCHEQUE, "ES CHEQUE")
            .set_TextMatrix(0, C_ColREQUERIRDOCTO, "REQUIERE DOCUMENTO")
            .set_TextMatrix(0, C_ColREQUERIRAUT, "REQUIERE AUT")
            .set_TextMatrix(0, C_ColRESTRINGIRCAMBIO, "RESRINGIR CAMBIO")
            .set_TextMatrix(0, C_ColCONSIDERARFACT, "CONSIDERAR PARA FACTURACION")
            .set_TextMatrix(0, C_ColCONSIDERARRETIRO, "CONSIDERAR PARA RETIROS")
            .set_TextMatrix(0, C_ColESTARJETA, "ES TARJTEA")
            .set_TextMatrix(0, C_ColDESCCOMBANC, "DESCONTAR COMISON BANCARIA")
            .set_TextMatrix(0, C_COLPORCCOMISION, "PORCENTAJE DE COMISION")
            .set_TextMatrix(0, C_ColPORCIVACOMISION, "PORCENTAJE DE COMISION")
            .set_TextMatrix(0, C_ColPAGOINTXPROM, "PAGO INTERES PROMOCION")
            .set_TextMatrix(0, C_ColPORCINTERESES, "PORCENTAJE INTERES")
            .set_TextMatrix(0, C_ColPORCIVAINTERESES, "PORCENTAJE IVA INTERES")
            .set_TextMatrix(0, C_ColCODFORMAPAGO, "CODIGO FORMA PAGO")
            .set_TextMatrix(0, C_ColESVALEDEVOLUCION, "ES VALE DEVOLUCION")
            .set_TextMatrix(0, C_ColIMPCOMISIONBANCARIA, "iMP COMISION BANCANRIA")
            .set_TextMatrix(0, C_ColIMPINTERESESPROMOCION, "INTERES PROMOCION")
            .set_TextMatrix(0, C_ColCodBanco, "BANCO")

            .Row = 0
            For LnContador = 0 To (.get_Cols() - 1) Step 1
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next LnContador
            .Row = 1
            .WordWrap = False
        End With
Errores:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function ValidaDatos() As Boolean
        On Error GoTo Merr
        If CDec(Numerico(txtdoCambio.Text)) > 0 And Trim(dbcMoneda.Text) = "" Then
            MsgBox("Proporcione la moneda en que se entrega el cambio al cliente.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            dbcMoneda.Focus()
            Exit Function
        End If
        ValidaDatos = True
        Exit Function
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Private Sub dbcmoneda_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMoneda.CursorChanged
        'If FueraChange = True Then Exit Sub
        gStrSql = "SELECT     CodFormaPago, Ltrim(Rtrim(DescFormaPago)) as DescFormaPago " & "From dbo.CatFormasPago " & "WHERE     (EsCheque = 0) AND (EsDevolucion = 0) AND (EsDocumentoInterno = 0) AND (EsTarjeta = 0) AND (Estatus = 'V') AND (DescFormaPago LIKE '" & Trim(dbcMoneda.Text) & "%') " & "ORDER BY CODFORMAPAGO "
        DCChange(gStrSql, tecla, dbcMoneda)
        intCodFormaPago = 0
        txtCodFormaPago.Text = ""
        txtEsDolarFPCambio.Text = ""
    End Sub

    Private Sub dbcmoneda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMoneda.Enter
        Pon_Tool()
        gStrSql = "SELECT     CodFormaPago, Ltrim(Rtrim(DescFormaPago)) as DescFormaPago " & "From dbo.CatFormasPago " & "WHERE     (EsCheque = 0) AND (EsDevolucion = 0) AND (EsDocumentoInterno = 0) AND (EsTarjeta = 0) AND (Estatus = 'V') ORDER BY CODFORMAPAGO "
        DCGotFocus(gStrSql)
    End Sub

    Private Sub dbcmoneda_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcMoneda.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        ElseIf eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then
            dbcmoneda_Leave(dbcMoneda, New System.EventArgs())
            msgFormasPago.Focus()
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcMoneda_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcMoneda.KeyUp
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Up Or eventArgs.KeyCode = System.Windows.Forms.Keys.Down Then
            intCodFormaPago = ObtenerCodFormaPago()
        End If
    End Sub

    Private Sub dbcmoneda_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMoneda.Leave
        gStrSql = "SELECT     CodFormaPago, Ltrim(Rtrim(DescFormaPago)) as DescFormaPago " & "From dbo.CatFormasPago " & "WHERE     (EsCheque = 0) AND (EsDevolucion = 0) AND (EsDocumentoInterno = 0) AND (EsTarjeta = 0) AND (Estatus = 'V') AND (DescFormaPago LIKE '" & Trim(dbcMoneda.Text) & "%') " & "ORDER BY CODFORMAPAGO "
        DCLostFocus(dbcMoneda, gStrSql, intCodFormaPago)
        txtCodFormaPago.Text = CStr(intCodFormaPago)
        ObtenerCodFormaPago()
    End Sub

    Private Sub dbcMoneda_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcMoneda.MouseUp
        intCodFormaPago = ObtenerCodFormaPago()
    End Sub

    Private Sub frmPagosSalMercancia_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmPagosSalMercancia_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmPagosSalMercancia_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Me = Nothing
        'IsNothing(Me)
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub msgFormasPago_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgFormasPago.DblClick
        msgFormasPago_KeyPressEvent(msgFormasPago, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub msgFormasPago_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgFormasPago.Enter
        msgFormasPago.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
        Pon_Tool()
    End Sub

    Private Sub msgFormasPago_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles msgFormasPago.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Escape Then Me.Close()
    End Sub

    Private Sub msgFormasPago_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles msgFormasPago.KeyPressEvent
        'En este Evento, Se muestra el Control para editar el Grid, con los datos que ya tiene el grid
        'Las Columnas Editables son: Codigo, DEscripcion, Cantidad,Descuento
        Dim RequerirAut As Boolean
        With msgFormasPago
            If Trim(.get_TextMatrix(.Row, C_COLDESCRIPCION)) = "" Then Exit Sub
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then 'Para que cuando sea escape, no entre a editar el codigo,simplemente que se regrese al control anterior
                'Validar la Tecla presionada pra que unicamente acepte numeros. Si es Enter no se valida , o no se convierte
                'Se Valida Aqui, des pues de que entro por el IF, para que aunke no sea unna tecla valida, se muestre el TExt para escribir
                eventArgs.keyAscii = ModEstandar.MskCantidad(txtImporte.Text, eventArgs.keyAscii, 10, 2, (txtImporte.SelectionStart))
                Select Case .Col
                    Case C_ColIMPORTE ''-------------- SE EDITA EL IMPORTE---------------------'''''
                        RequerirAut = CBool(.get_TextMatrix(.Row, C_ColREQUERIRAUT))
                        'Validar si la FOrma de Pago requiere autorizacion para usarse, de Ser así, se pide que se de un usuario valido que pueda dar la autorización
                        If RequerirAut = True Then
                            'Pedir el usuario y password para modificar el descto
                            'Para esto se usará la forma: frmAutorizacionConfig.
                            frmAutorizacionConfig.Text = "Autorizacion para utilizar Formas de Pago"
                            frmAutorizacionConfig.ShowDialog()
                            If gblnAutorizacionAceptada = False Then
                                'Si la Peticion no fue aceptada, es decir que el usuario que se proporciono no tiene derecho para autorizar o para modificar
                                'entonces no podrá ser modificado el descuento
                                If gblnSalioSinValidar = False Then 'Si valido el Usuari y Password y no tuvo derecho, mostrar el aviso de ke no puede hacerlo
                                    MsgBox(C_msgSINAUTORIZACION & "Utilizar esta forma de Pago.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "AVISO")
                                End If
                                .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                                .Focus()
                                Exit Sub
                            End If
                        End If
                        '                    Para que se puede editar el importe, Verificar que la columan de Formas de Pago tenga un Valor diferente de ""
                        'En el caso de que la FP sea vale de Devolución, no se edita el Grid de FP, sólo se mostrará la pantalla de registo de vale.
                        If .get_TextMatrix(.Row, C_COLDESCRIPCION) <> "" Then
                            If CBool(Trim(.get_TextMatrix(.Row, C_ColESVALEDEVOLUCION))) = False Then
                                ModEstandar.MSHFlexGridEdit(msgFormasPago, txtImporte, eventArgs.keyAscii)
                                If Len(Trim(txtImporte.Text)) <> 1 Then
                                    ModEstandar.SelTextoTxt(txtImporte)
                                End If
                            Else
                                'frmPVRegNotasCred.Show
                                ValidarImportedeFormaPago()
                            End If
                        End If
                End Select
            End If
        End With
    End Sub

    Private Sub msgFormasPago_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgFormasPago.Leave
        msgFormasPago.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub

    Private Sub frmPagosSalMercancia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
        Encabezado()
        LlenaGrid()
    End Sub

    Private Sub frmPagosSalMercancia_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmPagosSalMercancia_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'Validar la Tecla que se presionó para ver si corresponde con una de las TEclas Rápidas de las FOrmas de Pago
        With msgFormasPago
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, C_COLDESCRIPCION) = "" Then Exit For
                If UCase(Chr(KeyAscii)) = Trim(.get_TextMatrix(I, C_ColTECLARAPIDA)) Then
                    .Row = I
                    .Col = C_ColIMPORTE
                    msgFormasPago_KeyPressEvent(msgFormasPago, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
                    'Poner el Valor de Keyascii=0, para que no se trate de editar dos veces el imprte de la Forma de Pago.
                    'Porque primero pone el foco por aqui, cuando se pulsa una tecla que corresponde a una forma depago, y posteriormente, otravez se hace este paso,
                    'Cuando, sale de aqui y entra al evento KeyPress del Grid de FOrmas de Pago
                    KeyAscii = 0
                    Exit For
                End If
            Next
        End With
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmPagosSalMercancia_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                           Nuevo         Guardar        Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        frmVtasVELiquidacionVendedorExterno.Enabled = True
        '    frmPagosSalMercancia  Nothing
    End Sub

    Sub LlenaGrid()
        On Error GoTo Merr
        Dim I As Integer
        With msgFormasPago
            'gStrSql = "Select * from CatFormasPago Where Estatus = 'V'"
            gStrSql = "Select CodFormaPago, DescFormaPago, TeclaRapida, EsDolar, EsCheque, EsDevolucion, EsDocumentoInterno, RequerirDocto, " & "RequerirAutorizacion, RestringirCambio, ConsiderarParaFacturacion, ConsiderarParaRetiros, EsTarjeta, " & "DescontarComisionBanc, PorcComision, PorcIvaComision, DescCorta, Estatus, IsNull(CodBanco,0) as CodBanco " & "From CatFormasPago Where Estatus = 'V' Order by CodFormaPago "
            .Clear()
            ModEstandar.BorraCmd()
            Cmd.CommandText = "Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                If RsGral.RecordCount > 16 Then
                    .Rows = RsGral.RecordCount + 1
                End If
            Else
                MsgBox("No existen Información almacenada sobre Formas de Pago" & vbNewLine & "Verifique por favor....", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                Exit Sub
            End If
            Encabezado()
            'Poner los valores de los datos recopilados en las columnas TAG correspondientes
            For I = 1 To RsGral.RecordCount
                .set_TextMatrix(I, C_ColTECLARAPIDA, Trim(RsGral.Fields("TeclaRapida").Value))
                .set_TextMatrix(I, C_COLDESCRIPCION, Trim(RsGral.Fields("DescFormaPago").Value))
                .set_TextMatrix(I, C_ColIMPORTE, FormateoDecimales(0))
                .set_TextMatrix(I, C_ColESDOLAR, RsGral.Fields("EsDolar").Value)
                .set_TextMatrix(I, C_ColESCHEQUE, RsGral.Fields("Escheque").Value)
                .set_TextMatrix(I, C_ColREQUERIRDOCTO, RsGral.Fields("RequerirDocto").Value)
                .set_TextMatrix(I, C_ColREQUERIRAUT, RsGral.Fields("RequerirAutorizacion").Value)
                .set_TextMatrix(I, C_ColRESTRINGIRCAMBIO, RsGral.Fields("RestringirCambio").Value)
                .set_TextMatrix(I, C_ColCONSIDERARFACT, RsGral.Fields("ConsiderarParaFacturacion").Value)
                .set_TextMatrix(I, C_ColCONSIDERARRETIRO, RsGral.Fields("ConsiderarparaRetiros").Value)
                .set_TextMatrix(I, C_ColESTARJETA, RsGral.Fields("EsTarjeta").Value)
                .set_TextMatrix(I, C_ColDESCCOMBANC, RsGral.Fields("DescontarComisionBanc").Value)
                .set_TextMatrix(I, C_COLPORCCOMISION, RsGral.Fields("PorcComision").Value)
                .set_TextMatrix(I, C_ColPORCIVACOMISION, RsGral.Fields("PorcIvaComision").Value)
                .set_TextMatrix(I, C_ColCODFORMAPAGO, RsGral.Fields("CodFormaPago").Value)
                .set_TextMatrix(I, C_ColESVALEDEVOLUCION, RsGral.Fields("EsDevolucion").Value)
                .set_TextMatrix(I, C_ColIMPCOMISIONBANCARIA, FormateoDecimales(0))
                .set_TextMatrix(I, C_ColIMPINTERESESPROMOCION, FormateoDecimales(0))
                .set_TextMatrix(I, C_ColCodBanco, RsGral.Fields("CodBanco").Value)
                RsGral.MoveNext()
            Next I
            .Col = 2
            .Row = 1
        End With
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub txtImporte_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Enter
        Pon_Tool()
    End Sub

    Private Sub txtImporte_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtImporte.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Aqui se muestran los datos del control editable, en el Grid
        With msgFormasPago
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    txtImporte.Visible = False
                    txtImporte.Text = ""
                    .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                Case System.Windows.Forms.Keys.Return
                    .set_TextMatrix(.Row, .Col, txtImporte.Text)
                    txtImporte.Text = ""
                    txtImporte.Visible = False
                    msgFormasPago.Col = .Col
                    msgFormasPago.Row = .Row
                    .Focus()
                    'Validar si la Forma de Pago Requiere Guardar Datos Adicionales
                    ValidarImportedeFormaPago()
            End Select
        End With
    End Sub

    Private Sub txtImporte_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImporte.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'En este Evento se validan los datos que se introduzcan al control txtImporte,dependiendo de la columan en que se esté editando
        If KeyAscii = 0 Or KeyAscii = 13 Then GoTo EventExitSub
        With msgFormasPago
            If .Col = C_ColIMPORTE Then
                'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.MskCantidad(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                KeyAscii = ModEstandar.MskCantidad(txtImporte.Text, KeyAscii, 10, 2, (txtImporte.SelectionStart))
            End If
        End With
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtImporte_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        '    'Validar si la Forma de Pago Requiere Guardar Datos Adicionales
        '    ValidarImportedeFormaPago
        ''    CalcularComisionBancaria    'Una vez que se ha distribuido el importe del pago, en las formas de Pago, se verifica si la Forma de Pago requiere se descuente una comision bancaria o un Porcentaje de interes por promoción.
        ''                                'Lo cual se hace en este Procedimiento.

        txtImporte_KeyDown(txtImporte, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Private Sub txtdoCambio_Change()
        If CDbl(Numerico(txtdoCambio.Text)) < 0 Then
            'Fondo Rojo
            txtdoCambio.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
            txtdoCambio.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
        Else
            txtdoCambio.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000018)
            txtdoCambio.ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
        End If
    End Sub

    Sub ValidarImportedeFormaPago()
        'Este Procedimiento verifiac el tipo de Forma de Pago que se está utilizando, para validar si se requiere pedir otros datos .
        'Como por Ejemplo, para cheque pedir número de Documento, Para Tarjetas, pedir el banco, ect, y Asi para todas de acuerdo a sus caracteristicas
        On Error GoTo Merr
        '''Dim ImporteTecleado As Currency
        Dim TipoCambioVenta As Decimal
        Dim CambioalCliente As Decimal

        Dim ImporteAlmacenar As Decimal 'COntiene el importe a almacenar real, descontandole el cambio
        Dim DescripcionFP As String
        Dim cambio As Decimal
        Dim Escheque As Boolean
        Dim EsDolar As Boolean
        Dim RequerirDocto As Boolean
        Dim RequerirAut As Boolean
        Dim RestringirCambio As Boolean
        Dim ConsiderarParaFact As Boolean
        Dim ConsiderarparaRetiros As Boolean
        Dim EsTarjeta As Boolean
        Dim DescComisionBancaria As Boolean
        Dim PorcComision As Decimal
        Dim PorcIvaComision As Decimal
        Dim PagoIntXPromocion As Boolean
        Dim PorcInteres As Decimal
        Dim PorcIvaInteres As Decimal
        Dim CodFormaPago As Integer
        Dim EsValeDev As Boolean
        Dim TotalaPagarD As Decimal

        Dim TotalFPRC As Decimal 'Contiene el total del importe de las formas de pago que Restringen el Cambio
        Dim TotalFPNRC As Decimal 'Contiene el total del importe de las formas de pago que No Restringen el Cambio
        Dim Debe As Decimal 'Contiene el Importe del adeudo del cliente.
        Dim TotalPago As Decimal 'Contiene el total del pago hecho por el cliente. en pesos
        Dim importe As Decimal
        'UPGRADE_WARNING: Couldn't resolve default property of object FormateoDecimales(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TipoCambioVenta = FormateoDecimales(txtDolar)
        Debe = CDec(Numerico(txtdoAPagar4Decimales.Text))
        ImporteTecleado = 0
        With msgFormasPago
            ImporteTecleado = CDec(Numerico(.get_TextMatrix(.Row, C_ColIMPORTE)))
            DescripcionFP = Trim(.get_TextMatrix(.Row, C_COLDESCRIPCION))
            Escheque = CBool(.get_TextMatrix(.Row, C_ColESCHEQUE))
            EsDolar = CBool(.get_TextMatrix(.Row, C_ColESDOLAR))
            RequerirDocto = CBool(.get_TextMatrix(.Row, C_ColREQUERIRDOCTO))
            RequerirAut = CBool(.get_TextMatrix(.Row, C_ColREQUERIRAUT))
            RestringirCambio = CBool(.get_TextMatrix(.Row, C_ColRESTRINGIRCAMBIO))
            ConsiderarParaFact = CBool(.get_TextMatrix(.Row, C_ColCONSIDERARFACT))
            ConsiderarparaRetiros = CBool(.get_TextMatrix(.Row, C_ColCONSIDERARRETIRO))
            EsTarjeta = CBool(.get_TextMatrix(.Row, C_ColESTARJETA))
            DescComisionBancaria = CBool(.get_TextMatrix(.Row, C_ColDESCCOMBANC))
            PorcComision = CDec(.get_TextMatrix(.Row, C_COLPORCCOMISION))
            PorcIvaComision = CDec(.get_TextMatrix(.Row, C_ColPORCIVACOMISION))
            CodFormaPago = CInt(.get_TextMatrix(.Row, C_ColCODFORMAPAGO))
            EsValeDev = CBool(.get_TextMatrix(.Row, C_ColESVALEDEVOLUCION))

            'La Validación de cuando una Forma de Pago requiere autorización para su uso, está en el Evento KeyPress del Grid de Formas de Pago (msgFormasPago)
            If RequerirDocto = True Then
                'Siempre y Cuando esté especificado que la forma de pago requiere un documento.
                If Escheque = True Then
                    'Mandar el Codigo de la forma de pago que se esta usando, para uso posterior en guardar.
                    'frmPVRegCheque.PonerCodFormaPago(CInt(.get_TextMatrix(.Row, C_ColCODFORMAPAGO)))
                    'frmPVRegCheque.ShowDialog()
                End If
                If EsTarjeta = True Then
                    'frmPVRegTarjeta_PV.PonerCodFormaPago(CInt(.get_TextMatrix(.Row, C_ColCODFORMAPAGO)), CInt(.get_TextMatrix(.Row, C_ColCodBanco)))
                    'frmPVRegTarjeta_PV.ShowDialog()
                End If
                If EsValeDev = True Then
                    'Poner la moneda en que está el Vale.
                    'frmPVRegNotasCred.PonerCodFormaPago(CInt(.get_TextMatrix(.Row, C_ColCODFORMAPAGO)), EsDolar)
                    'frmPVRegNotasCred.ShowDialog()
                    'Unload frmPVRegNotasCred
                End If
            End If
            'La Columna de Importe contiene el importe real que el cliente esta pagando, sin descontrar el cambio
            'Esto para despues obtener el  importe del pago total hecho por el cliente

            'En el caso de que la forma de pago sea vale, el valor de importe no se asignará aqui, ya que no es posible teclear un importe,
            'este importe se calculará en el formulario de reg. de Vales , de acuerdo a los vales proporcionados.
            If EsValeDev = False Then
                .set_TextMatrix(.Row, C_ColIMPORTE, VB6.Format(FormateoDecimales(ImporteTecleado), gstrFormatoCantidad))
                .set_TextMatrix(.Row, C_ColIMPORTESINREDONDEO, ImporteTecleado)
            End If
            TotalFPRC = ObtenerTotalFPRC() '(Dólar) Obtiene el Importe totla de las Formas de Pago que    resringen el cambio en Dólares
            TotalFPNRC = ObtenerTotalFPNRC() '(Dólar) Obtiene el Importe totla de las Formas de Pago que NO resringen el cambio en Dólares
            TotalPago = TotalFPNRC + TotalFPRC 'Contiene el Total de Pago

            If System.Math.Round(TotalFPRC, 1) > System.Math.Round(Debe, 1) Then
                mblnImpteFPRCMayorTotalAPagar = True
                MsgBox("EL Importe total de las Formas de Pago que restringen el cambio, es mayor que el Importe a pagar." & vbNewLine & "Verifique por favor..", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
                .Focus()
            End If

            TotalaPagarD = CDbl(txtmnAPagar4Decimales.Text) / TipoCambioVenta

            txtdoTotalPago.Text = VB6.Format(System.Math.Round(TotalPago, 2), gstrFormatoCantidad)
            txtdoTotalPago4Decimales.Text = CStr(TotalPago)

            '''***
            If (CDec(Numerico((txtdoAPagar.Text))) - CDec(Numerico((txtdoTotalPago.Text))) = 0) Then
                txtdoCambio.Text = VB6.Format(System.Math.Round(CDec(Numerico((txtdoAPagar.Text))) - CDec(Numerico((txtdoTotalPago.Text))), 2), gstrFormatoCantidad)
                txtmnTotalPago.Text = VB6.Format(txtmnAPagar.Text, gstrFormatoCantidad)
                txtmnCambio.Text = VB6.Format(System.Math.Round(CDec(Numerico((txtmnAPagar.Text))) - CDec(Numerico((txtmnTotalPago.Text))), 2), gstrFormatoCantidad)
            Else '''***
                txtmnTotalPago.Text = VB6.Format(System.Math.Round(TotalPago * TipoCambioVenta, 2), gstrFormatoCantidad)
                '''en caso de que conincidan solo los pesos ***
                If (CDec(Numerico((txtmnAPagar.Text))) - CDec(Numerico((txtmnTotalPago.Text))) = 0) Then
                    txtmnCambio.Text = VB6.Format(System.Math.Round(CDec(Numerico((txtmnAPagar.Text))) - CDec(Numerico((txtmnTotalPago.Text))), 2), gstrFormatoCantidad)
                    txtdoTotalPago.Text = VB6.Format(txtdoAPagar.Text, gstrFormatoCantidad)
                    txtdoCambio.Text = VB6.Format(System.Math.Round(CDec(Numerico((txtdoAPagar.Text))) - CDec(Numerico((txtdoTotalPago.Text))), 2), gstrFormatoCantidad)
                Else '''***
                    '''anto 12 FEB 2004
                    '''ojo redondeo 2 --> original 1
                    txtmnTotalPago.Text = CStr(System.Math.Round(TotalPago * TipoCambioVenta, 2)) '''OJO        'Se obtiene el Total del Pago en Pesos y re
                    txtmnTotalPago.Text = VB6.Format(txtmnTotalPago.Text, gstrFormatoCantidad) 'Redondeo posteriormente a 1 decimal
                End If '''***
            End If '''***

            'estas cantidades no se redondean, para despues obtener el cambi en 4 decimales
            '''ojo anto 12 feb 2004
            txtmnTotalPago4Decimales.Text = CStr(System.Math.Round(System.Math.Round(TotalPago * TipoCambioVenta, 2), 4))
            '''txtmnTotalPago4Decimales = TotalPago * TipoCambioVenta
            txtmnTotalPago4Decimales.Text = CStr(CDbl(Numerico(txtmnTotalPago4Decimales.Text)))

            'Debe = FormateoDecimales(TotalaPagarD - Round(TotalFPRC, 2))
            Debe = TotalaPagarD - TotalFPRC
            If Debe < 0 Then ' Si el Debe el Menor de Cero, significa que el Total de las FPRC es Mayor que el Importe total del Pago
                Debe = 0
                txtdoCambio.Text = VB6.Format(System.Math.Round(TotalFPNRC) - Debe, gstrFormatoCantidad)
                txtmnCambio.Text = VB6.Format(System.Math.Round(CDbl(txtdoCambio.Text) * TipoCambioVenta, 1), gstrFormatoCantidad)
                txtdoCambio4Decimales.Text = CStr(TotalFPNRC - Debe)
                txtmnCambio4Decimales.Text = CStr(CDbl(txtdoCambio4Decimales.Text) * TipoCambioVenta)
            Else
                txtdoCambio.Text = VB6.Format(CDbl(txtdoTotalPago.Text) - CDbl(txtdoAPagar.Text), gstrFormatoCantidad)
                txtmnCambio.Text = VB6.Format(CDbl(txtmnTotalPago.Text) - CDbl(txtmnAPagar.Text), gstrFormatoCantidad)
                txtdoCambio4Decimales.Text = CStr(System.Math.Round(CDbl(Numerico(txtdoTotalPago4Decimales.Text)) - CDbl(Numerico(txtdoAPagar4Decimales.Text)), 4))
                txtmnCambio4Decimales.Text = CStr(System.Math.Round(CDbl(Numerico(txtmnTotalPago4Decimales.Text)) - CDbl(Numerico(txtmnAPagar4Decimales.Text)), 4))
            End If

            '''End If '''***
            '''End If '''***
            'CHECAR ESTO 25-09-03

            .set_TextMatrix(.Row, C_ColIMPCOMISIONBANCARIA, VB6.Format(CalcularComisionBancaria(CDec(.get_TextMatrix(.Row, C_ColIMPORTE)), CodFormaPago), gstrFormatoCantidad))
            '.TextMatrix(.Row, C_ColIMPINTERESESPROMOCION) = Format( frmPVRegTarjeta.CalcularInteresesPromocion (.TextMatrix(.Row, C_ColIMPORTE), CodFormaPago), gstrFormatoCantidad)
            'CHECAR ESTO 25-09-03

            msgFormasPago.Focus()
        End With
        Exit Sub

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    'Sub ValidarImportedeFormaPago()
    '    'Este Procedimiento verifiac el tipo de Forma de Pago que se está utilizando, para validar si se requiere pedir otros datos .
    '    'Como por Ejemplo, para cheque pedir número de Documento, Para Tarjetas, pedir el banco, ect, y Asi para todas de acuerdo a sus caracteristicas
    '    On Local Error GoTo Merr:
    '    Dim ImporteTecleado As Currency
    '    Dim TipoCambioVenta As Currency
    '    Dim CambioalCliente As Currency
    '
    '    Dim ImporteAlmacenar As Currency 'COntiene el importe a almacenar real, descontandole el cambio
    '    Dim DescripcionFP As String
    '    Dim cambio As Currency
    '    Dim Escheque As Boolean
    '    Dim EsDolar As Boolean
    '    Dim RequerirDocto As Boolean
    '    Dim RequerirAut As Boolean
    '    Dim RestringirCambio As Boolean
    '    Dim ConsiderarParaFact As Boolean
    '    Dim ConsiderarparaRetiros As Boolean
    '    Dim EsTarjeta As Boolean
    '    Dim DescComisionBancaria As Boolean
    '    Dim PorcComision As Currency
    '    Dim PorcIvaComision As Currency
    '    Dim PagoIntXPromocion As Boolean
    '    Dim PorcInteres As Currency
    '    Dim PorcIvaInteres As Currency
    '    Dim CodFormaPago As Integer
    '    Dim EsValeDev As Boolean
    '    Dim TotalaPagarD As Currency
    '
    '    Dim TotalFPRC As Currency   'Contiene el total del importe de las formas de pago que Restringen el Cambio
    '    Dim TotalFPNRC As Currency  'Contiene el total del importe de las formas de pago que No Restringen el Cambio
    '    Dim Debe As Currency        'Contiene el Importe del adeudo del cliente.
    '    Dim TotalPago As Currency   'Contiene el total del pago hecho por el cliente. en pesos
    '    Dim Importe As Currency
    '    TipoCambioVenta = FormateoDecimales(txtDolar)
    '    Debe = Numerico(txtdoAPagar4Decimales)
    '    With msgFormasPago
    '        ImporteTecleado = Numerico(.TextMatrix(.Row, C_ColIMPORTE))
    '        DescripcionFP = Trim(.TextMatrix(.Row, C_ColDESCRIPCION))
    '        Escheque = .TextMatrix(.Row, C_ColESCHEQUE)
    '        EsDolar = .TextMatrix(.Row, C_ColESDOLAR)
    '        RequerirDocto = .TextMatrix(.Row, C_ColREQUERIRDOCTO)
    '        RequerirAut = .TextMatrix(.Row, C_ColREQUERIRAUT)
    '        RestringirCambio = .TextMatrix(.Row, C_ColRESTRINGIRCAMBIO)
    '        ConsiderarParaFact = .TextMatrix(.Row, C_ColCONSIDERARFACT)
    '        ConsiderarparaRetiros = .TextMatrix(.Row, C_ColCONSIDERARRETIRO)
    '        EsTarjeta = .TextMatrix(.Row, C_ColESTARJETA)
    '        DescComisionBancaria = .TextMatrix(.Row, C_ColDESCCOMBANC)
    '        PorcComision = .TextMatrix(.Row, C_COLPORCCOMISION)
    '        PorcIvaComision = .TextMatrix(.Row, C_ColPORCIVACOMISION)
    '        CodFormaPago = .TextMatrix(.Row, C_ColCODFORMAPAGO)
    '        EsValeDev = CBool(.TextMatrix(.Row, C_ColESVALEDEVOLUCION))
    '
    '        'La Validación de cuando una Forma de Pago requiere autorización para su uso, está en el Evento KeyPress del Grid de Formas de Pago (msgFormasPago)
    '        If RequerirDocto = True Then
    '            'Siempre y Cuando esté especificado que la forma de pago requiere un documento.
    '            If Escheque = True Then
    '                'Mandar el Codigo de la forma de pago que se esta usando, para uso posterior en guardar.
    '                frmPVRegCheque.PonerCodFormaPago .TextMatrix(.Row, C_ColCODFORMAPAGO)
    '               frmPVRegCheque.Show vbModal
    '            End If
    '            If EsTarjeta = True Then
    '                frmPVRegTarjeta.PonerCodFormaPago .TextMatrix(.Row, C_ColCODFORMAPAGO)
    '                frmPVRegTarjeta.Show vbModal
    '
    '            End If
    '            If EsValeDev = True Then
    '                'Poner la moneda en que está el Vale.
    '                frmPVRegNotasCred.PonerCodFormaPago .TextMatrix(.Row, C_ColCODFORMAPAGO), EsDolar
    '                frmPVRegNotasCred.Show vbModal
    '                'Unload frmPVRegNotasCred
    '            End If
    '        End If
    '        'La Columna de Importe contiene el importe real que el cliente esta pagando, sin descontrar el cambio
    '        'Esto para despues obtener el  importe del pago total hecho por el cliente
    '
    '        'En el caso de que la forma de pago sea vale, el valor de importe no se asignará aqui, ya que no es posible teclear un importe,
    '        'este importe se calculará en el formulario de reg. de Vales , de acuerdo a los vales proporcionados.
    '        If EsValeDev = False Then
    '            .TextMatrix(.Row, C_ColIMPORTE) = Format(FormateoDecimales(ImporteTecleado), gstrFormatoCantidad)
    '            .TextMatrix(.Row, C_ColIMPORTESINREDONDEO) = ImporteTecleado
    '        End If
    '        TotalFPRC = ObtenerTotalFPRC    '(Dólar) Obtiene el Importe totla de las FOrmas de PAgo que resringen el cambio En Dólares
    '        TotalFPNRC = ObtenerTotalFPNRC  '(Dólar) Obtiene el Importe totla de las FOrmas de PAgo que NO resringen el cambio en Dólares
    '        TotalPago = TotalFPNRC + TotalFPRC 'Contien el Total de Pago
    '
    '        If TotalFPRC > Debe Then
    '            mblnImpteFPRCMayorTotalAPagar = True
    '            MsgBox "EL Importe total de las Formas de Pago que restringen el cambio, es mayor que el Importe a pagar." + vbNewLine + "Verifique por favor..", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '            .FocusRect = flexFocusNone
    '            .SetFocus
    '        End If
    '
    '        TotalaPagarD = txtmnAPagar4Decimales / TipoCambioVenta
    '
    '        txtdoTotalPago = Format(Round((TotalPago), 2), gstrFormatoCantidad)
    '        txtdoTotalPago4Decimales = TotalPago
    '
    '        txtmnTotalPago = Round(TotalPago * TipoCambioVenta, 1)          'Se obtiene el Total del Pago en Pesos y re
    '        txtmnTotalPago = Format(txtmnTotalPago, gstrFormatoCantidad)         'Redondeo posteriormente a 1 decimal
    '
    '        'estas cantidades no se redondean, para despues obtener el cambi en 4 decimales
    '        txtmnTotalPago4Decimales = TotalPago * TipoCambioVenta
    '        txtmnTotalPago4Decimales = CDbl(Numerico(txtmnTotalPago4Decimales))
    '
    '        'Debe = FormateoDecimales(TotalaPagarD - Round(TotalFPRC, 2))
    '        Debe = TotalaPagarD - TotalFPRC
    '        If Debe < 0 Then    ' Si el DEbe el Menor de Cero, significa que el Total de las FPRC es Mayor que el Imñporte Total del PAgo
    '            Debe = 0
    '            txtdoCambio = Format(Round(TotalFPNRC) - Debe, gstrFormatoCantidad)
    '            txtmnCambio = Format(Round(txtdoCambio * TipoCambioVenta, 1), gstrFormatoCantidad)
    '            txtdoCambio4Decimales = TotalFPNRC - Debe
    '            txtmnCambio4Decimales = txtdoCambio4Decimales * TipoCambioVenta
    '        Else
    '            txtdoCambio = Format(txtdoTotalPago - txtdoAPagar, gstrFormatoCantidad)
    '            txtmnCambio = Format(txtmnTotalPago - txtmnAPagar, gstrFormatoCantidad)
    '            txtdoCambio4Decimales = Round(CDbl(Numerico(txtdoTotalPago4Decimales)) - CDbl(Numerico(txtdoAPagar4Decimales)), 4)
    '            txtmnCambio4Decimales = Round(CDbl(Numerico(txtmnTotalPago4Decimales)) - CDbl(Numerico(txtmnAPagar4Decimales)), 4)
    '        End If
    '
    ''CHECAR ESTO 25-09-03
    '
    '                .TextMatrix(.Row, C_ColIMPCOMISIONBANCARIA) = Format(CalcularComisionBancaria(.TextMatrix(.Row, C_ColIMPORTE), CodFormaPago), gstrFormatoCantidad)
    '                '.TextMatrix(.Row, C_ColIMPINTERESESPROMOCION) = Format( frmPVRegTarjeta.CalcularInteresesPromocion (.TextMatrix(.Row, C_ColIMPORTE), CodFormaPago), gstrFormatoCantidad)
    ''CHECAR ESTO 25-09-03
    '        msgFormasPago.SetFocus
    '    End With
    '    Exit Sub
    'Merr:
    '    If Err.Number <> 0 Then ModEstandar.MostrarError
    'End Sub

    Function ObtenerTotalFPRC() As Decimal
        On Error GoTo Merr
        'Esta función contiene el importe total de las formas de pago que restringen el cambio, especificados en el Grid.
        Dim D As Integer
        ObtenerTotalFPRC = 0
        With msgFormasPago
            For D = 1 To .Rows - 1
                If .get_TextMatrix(D, C_COLDESCRIPCION) = "" Then Exit For
                If CBool(.get_TextMatrix(D, C_ColRESTRINGIRCAMBIO)) = True Then
                    If CBool(.get_TextMatrix(D, C_ColESDOLAR)) = True Then

                        ObtenerTotalFPRC = ObtenerTotalFPRC + CDbl(Numerico(.get_TextMatrix(D, C_ColIMPORTESINREDONDEO)))
                    Else
                        'Si es Pesos COmvertir a Dólares
                        'UPGRADE_WARNING: Couldn't resolve default property of object FormateoDecimales(txtDolar). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        ObtenerTotalFPRC = ObtenerTotalFPRC + (CDbl(Numerico(.get_TextMatrix(D, C_ColIMPORTESINREDONDEO))) / FormateoDecimales(txtDolar))
                    End If
                End If
            Next
        End With
        Exit Function
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ObtenerTotalFPNRC() As Decimal
        On Error GoTo Merr
        'Esta función contiene el importe total de las formas de pago que NO restringen el cambio, especificados en el Grid.
        'El importe es en Dólares
        Dim F As Integer
        ObtenerTotalFPNRC = 0
        With msgFormasPago
            For F = 1 To .Rows - 1
                If .get_TextMatrix(F, C_COLDESCRIPCION) = "" Then Exit For
                If CBool(.get_TextMatrix(F, C_ColRESTRINGIRCAMBIO)) = False Then
                    If CBool(.get_TextMatrix(F, C_ColESDOLAR)) = True Then

                        ObtenerTotalFPNRC = ObtenerTotalFPNRC + CDbl(Numerico(.get_TextMatrix(F, C_ColIMPORTESINREDONDEO)))
                    Else
                        'Si son Pesos CONvertir a Dolares
                        ObtenerTotalFPNRC = ObtenerTotalFPNRC + CDbl(CDbl(Numerico(.get_TextMatrix(F, C_ColIMPORTE))) / CDbl(Numerico(txtDolar.Text)))
                        '                    ObtenerTotalFPNRC = ObtenerTotalFPNRC + Round((.TextMatrix(F, C_ColIMPORTE) / Numerico(txtDolar)), 2)
                    End If
                End If
            Next
        End With
        ObtenerTotalFPNRC = ObtenerTotalFPNRC
        Exit Function
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ObtenerTotalPagoPesos() As Object
        '    Esta Funcion obtiene el Total del importe registrado en las formas de pago
        '    On Local Error GoTo MErr:
        '    ObtenerTotalPagoPesos = 0
        '    With msgFormasPago
        '        For i = 1 To .Rows - 1
        '            If .TextMatrix(i, C_ColDESCRIPCION) = "" Then Exit For
        '            ObtenerTotalPagoPesos = ObtenerTotalPagoPesos + Numerico(.TextMatrix(i, C_ColIMPPESOS))
        '        Next
        '    End With
        '    Exit Function
        'MErr:
        '    If Err.Number <> 0 Then ModEstandar.MostrarError
    End Function

    Function ObtenerTotalPagoDolares() As Object
        'Esta Funcion obtiene el Total del importe registrado en las formas de pago en dolares
        '    On Local Error GoTo MErr:
        '    ObtenerTotalPagoDolares = 0
        '    With msgFormasPago
        '        For i = 1 To .Rows - 1
        '            If .TextMatrix(i, C_ColDESCRIPCION) = "" Then Exit For
        '            ObtenerTotalPagoDolares = ObtenerTotalPagoDolares + Numerico(.TextMatrix(i, C_ColIMPREAL))
        '        Next
        '    End With
        '    Exit Function
        'MErr:
        '    If Err.Number <> 0 Then ModEstandar.MostrarError
    End Function

    Function ObtenerTotalCambioPesos() As Object
        '    'Esta Funcion obtiene el Total del Cambio registrado en las formas de pago
        '    On Local Error GoTo MErr:
        '    Dim G As Integer
        '    ObtenerTotalCambioPesos = 0
        '    With msgFormasPago
        '    For G = 1 To .Rows - 1
        '            If .TextMatrix(G, C_ColDESCRIPCION) = "" Then Exit For
        '            ObtenerTotalCambioPesos = ObtenerTotalCambioPesos + Numerico(.TextMatrix(G, C_ColCAMBIO))
        '        Next
        '    End With
        '    Exit Function
        'MErr:
        '    If Err.Number <> 0 Then ModEstandar.MostrarError
    End Function

    Function Guardar() As Object
        'Esta función únicamente define que se guardará, de acuerdo al formulario que haya llamado a Pagos.
        'Estyo está definido, en el Text BOx de Forma Origen. Desde aqui solo se ejecutará la función GUARDAR, del formulario que se requiera.
        Dim FormaLlamado As Object
        FormaLlamado = Trim(txtFormaOrigen.Text)
        'Valida que se haya seleccionado una moneda para el cambio. Siempre y cuando el cambio sea mayor de cero.
        If ValidaDatos() = False Then Exit Function
        Select Case FormaLlamado
            Case "frmVtasVELiquidacionVendedorExterno"
                frmVtasVELiquidacionVendedorExterno.Guardar()
        End Select
    End Function

    Function GuardarIngresos(ByRef FolioMovto As String, ByRef FechaMovto As Date, ByRef Cliente As Integer, ByRef Vendedor As Integer, ByRef TipoIngreso As String, ByRef Moneda As String, ByRef TipoCambio As Decimal, ByRef Estatus As String, ByRef Caja As Integer) As Boolean
        On Error GoTo Merr
        Dim FolioIngreso As String
        Dim Prefijo As String
        Dim Consecutivo As String
        Dim cambio As Decimal
        Dim Total As Decimal
        Dim ComisionBancariaD As Decimal 'La comisión Bancaria se guarda en Dólares
        Dim InteresesPromocionD As Decimal 'El importe de Intereses por Promoción se guarda en Dólares
        Dim CodFPCambio As Integer 'Contiene la Forma de Pago en que se entregará el cambio al cliente. Esto se especifica en el Formulario de PagosSalMcia.
        Dim EsDolarCodFP As Boolean 'Indica si la Forma de Pago en que se da el cambio al Dolar
        gstrProcesoqueGeneraError = "frmPagosSalMercancia (GuardarIngresos)"

        gStrSql = "select Prefijo , Consecutivo + 1 as Consecutivo From CatFolios where DescFolio='INGRESOS' and CodAlmacen = " & gintCodAlmacenGral

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'CambiO = Me.txtdoCambio
        cambio = CDec(Numerico((Me.txtdoCambio4Decimales).Text))

        'Este Cambio se Realizó por la necesidad de guardar siempre un Importe com 4 Decimales. ya que si sacaba del ingrespo en las Formas de Pago, en ocasiones, se tenian importes con 2 decimales.
        'Total = (Me.txtmnTotalPago - Me.txtmnCambio) / Numerico(Me.txtDolar)
        If CDbl(Numerico((Me.txtDolar).Text)) <> 0 Then
            Total = (CDbl(Numerico((Me.txtmnTotalPago4Decimales).Text)) - CDbl(Numerico((Me.txtmnCambio4Decimales).Text))) / CDbl(Numerico((Me.txtDolar).Text))
        Else
            Total = 0
        End If
        Prefijo = Trim(RsGral.Fields("Prefijo").Value)
        Consecutivo = RsGral.Fields("Consecutivo").Value
        FolioIngreso = Prefijo & VB6.Format(CStr(gintCodAlmacen), "00") & VB6.Format(CStr(gintCodCaja), "00") & CStr(Year(FechaMovto)) & VB6.Format(CStr(Month(FechaMovto)), "00") & VB6.Format(CStr(VB.Day(FechaMovto)), "00") & VB6.Format(Consecutivo, "0000")
        'ANtes de Guardar, se debe Obtener el importe total de COmision Bancaria y de Promociones
        ComisionBancariaD = ObtenerTotalComisionBancaria()
        InteresesPromocionD = ObtenerTotalInteresesPromocion()
        CodFPCambio = CInt(Numerico(Trim(txtCodFormaPago.Text)))
        'si no se especificó moneda, significa que no se está otorgando cambio al cliente, por tanto, EsDolar es falso
        If Trim(txtEsDolarFPCambio.Text) = "" Then
            EsDolarCodFP = False
        Else
            EsDolarCodFP = CBool(Trim(txtEsDolarFPCambio.Text))
        End If

        'Guardar los importes generales de la VEnta, Siempre y cuando los haya
        ModStoredProcedures.PR_IEIngresos(FolioIngreso, VB6.Format(FechaMovto, C_FORMATFECHAGUARDAR), (frmVtasVELiquidacionVendedorExterno.txtCodSucMatriz).Text, CStr(Caja), TipoIngreso, FolioMovto, CStr(Cliente), CStr(Moneda), CStr(Total), CStr(ComisionBancariaD), CStr(InteresesPromocionD), CStr(TipoCambio), CStr(Vendedor), Estatus, "01/01/1900", IIf((EsDolarCodFP = False), CStr(cambio), 0), IIf((EsDolarCodFP = True), CStr(cambio), 0), C_INSERCION, "0")
        Cmd.Execute()

        GuardarIngresosFormasdePago(FolioIngreso, FolioMovto, FechaMovto, TipoCambio, Estatus)
        GuardarIngresos = True
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            GuardarIngresos = False
        End If
    End Function

    Function GuardarIngresosFormasdePago(ByRef FolioIngreso As String, ByRef FolioMovto As String, ByRef FechaMovto As Date, ByRef TipoCambio As Decimal, ByRef Estatus As String) As Boolean
        On Error GoTo Merr
        'Este procedimiento, guarda el importe y los datos de cad auna de las formas de pago utilizadas en el pago de la venta.
        'Si se uso cheque, tarjeta o Nota de Crédito, se guardaran los datos relacionados con ellos.
        Dim EsVale As Boolean
        Dim FormaPago As Integer
        Dim importe As Decimal
        Dim SeRegistroCheque As Boolean 'Almacena si ya fue registrada una Forma de Pago en uso, ya que de ser asi, no se ejecutará de nuevo el Proc. de Guardar de esa FOrma de Pago, esto es, para que no se guarden dos veces los mimos datos, por ejemplo en el caso de Tarjeta, como hay dos formas de pago que son tarjeta, la informacion se graba dos veces si no se usa esta variable de bandera-.
        Dim SeRegistroTarjeta As Object
        Dim SeRegistroVale As Object

        gstrProcesoqueGeneraError = "frmPagosSalMercancia (GuardarIngresosFormasdePago)"
        SeRegistroCheque = False
        SeRegistroTarjeta = False
        SeRegistroVale = False
        'El Folio del ingreso de Forma de Pago que se guardará, se obtiene con uan funcion establecida, pero se hará en cada proceso de guardar, antes de grabar.
        'Por ejemplo se obtiene en este proceso, cuando se guarde efectivo o dolar, y para las demas formas de pago, el folio se obtiene dentro del procedimiento de guardar de cada uno de los formulario de forma de pago.
        'Ya que puede ser que se rekiera guardar mas de un registro.
        With msgFormasPago
            For I = 1 To .Rows - 1
                'Si la descripcion es nulo, entonces salir del for, porque ya no hay mas formas de pago.
                If .get_TextMatrix(I, C_COLDESCRIPCION) = "" Then Exit For
                'Unicamente se guardan los datos cuando el importe sea mayor de cero
                If CDbl(Numerico(.get_TextMatrix(I, C_ColIMPORTE))) <> 0 Then
                    'Obtener el tipo de forma de pago para saber como guardar los datos.
                    Escheque = CBool(.get_TextMatrix(I, C_ColESCHEQUE))
                    EsTarjeta = CBool(.get_TextMatrix(I, C_ColESTARJETA))
                    EsVale = CBool(.get_TextMatrix(I, C_ColESVALEDEVOLUCION))
                    FormaPago = CInt(.get_TextMatrix(I, C_ColCODFORMAPAGO))
                    importe = CDec(.get_TextMatrix(I, C_ColIMPORTE)) 'Calcularlo sobre el Importe Total pagado por el CLiente
                    ComisionBancaria = CDec(.get_TextMatrix(I, C_ColIMPCOMISIONBANCARIA))
                    InteresesPromocion = CDec(.get_TextMatrix(I, C_ColIMPINTERESESPROMOCION))
                    If Escheque = True Then
                        'Si es Cehque, ejecutar el Guardar del Form. de Cheque
                        If SeRegistroCheque = False Then
                            'If frmPVRegCheque.Guardar(FolioIngreso, CDate(VB6.Format(FechaMovto, C_FORMATFECHAGUARDAR)), FolioMovto, TipoCambio, Estatus) = False Then
                            '    GuardarIngresosFormasdePago = False
                            '    Exit Function
                            'End If
                            SeRegistroCheque = True
                        End If
                    ElseIf EsTarjeta = True Then
                        If SeRegistroTarjeta = False Then
                            'Si es tarjeta, ejecutar el Guardar del Form. de Tarjeta
                            'If frmPVRegTarjeta_PV.Guardar(FolioIngreso, CDate(VB6.Format(FechaMovto, C_FORMATFECHAGUARDAR)), FolioMovto, TipoCambio, Estatus) = False Then
                            '    GuardarIngresosFormasdePago = False
                            '    Exit Function
                            'End If
                            SeRegistroTarjeta = True
                            gblnPagoVentasconTarjeta = True
                        End If
                    ElseIf EsVale = True Then
                        If SeRegistroVale = False Then
                            'Si es devolucion , ejecutar el Guardar del Form. de Vale de DEvolucion
                            'If frmPVRegNotasCred.Guardar(FolioIngreso, CDate(VB6.Format(FechaMovto, C_FORMATFECHAGUARDAR)), FolioMovto, TipoCambio, Estatus) = False Then
                            '    GuardarIngresosFormasdePago = False
                            '    Exit Function
                            'End If
                            SeRegistroVale = True
                        End If
                    Else
                        'Si no cabe dentro de las opciones antreriores, guardarlo aqui, porque es Efectivo o Dolar.
                        'Obtener el Folio de Ingreso que se debe guardar.
                        NumPartidaIngresosFormaDePago = ObtenerPartidaIngresosFormaDesPago(FolioIngreso)
                        ModStoredProcedures.PR_IEIngresosFormasdePago(FolioIngreso, CStr(NumPartidaIngresosFormaDePago), VB6.Format(FechaMovto, C_FORMATFECHAGUARDAR), FolioMovto, CStr(FormaPago), CStr(importe), CStr(0), CStr(0), "", "", "", "", CStr(ComisionBancaria), CStr(InteresesPromocion), CStr(TipoCambio), Estatus, "01/01/1900", "0", "01/01/1900", CStr(0), C_INSERCION, "0")
                        Cmd.Execute()
                    End If
                End If
            Next
        End With
        GuardarIngresosFormasdePago = True
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            GuardarIngresosFormasdePago = False
        End If
    End Function

    Function ObtenerTotalComisionBancaria() As Decimal
        'Esta Función calcula el importe total de comision bancaria, el cual se saca del GRid de Formas de pago, y se hace una sumatoria
        'EL importe de Comisión debe Calcularse en Pesos.
        ObtenerTotalComisionBancaria = 0
        Dim EsDolar As Boolean
        TipoCambio = CDbl(Numerico(txtDolar.Text))
        With msgFormasPago
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, C_COLDESCRIPCION) = "" Then Exit For
                EsDolar = CBool(.get_TextMatrix(I, C_ColESDOLAR))
                If EsDolar Then
                    ObtenerTotalComisionBancaria = ObtenerTotalComisionBancaria + CDbl(.get_TextMatrix(I, C_ColIMPCOMISIONBANCARIA))
                Else
                    ObtenerTotalComisionBancaria = ObtenerTotalComisionBancaria + (CDbl(.get_TextMatrix(I, C_ColIMPCOMISIONBANCARIA)) / TipoCambio)
                End If
            Next
        End With
    End Function

    Function ObtenerTotalInteresesPromocion() As Decimal
        'Esta Función calcula el importe total de Intereses por Promoción, el cual se saca del GRid de Formas de pago, y se hace una sumatoria
        'EL importe de Interes debe guardarse en dolares
        ObtenerTotalInteresesPromocion = 0
        Dim EsDolar As Boolean
        With msgFormasPago
            For I = 1 To .Rows - 1
                If .get_TextMatrix(I, C_COLDESCRIPCION) = "" Then Exit For
                EsDolar = CBool(.get_TextMatrix(I, C_ColESDOLAR))
                If EsDolar Then
                    ObtenerTotalInteresesPromocion = ObtenerTotalInteresesPromocion + CDbl(.get_TextMatrix(I, C_ColIMPINTERESESPROMOCION))
                Else
                    ObtenerTotalInteresesPromocion = ObtenerTotalInteresesPromocion + (CDbl(.get_TextMatrix(I, C_ColIMPINTERESESPROMOCION)) / TipoCambio)
                End If
            Next
        End With
    End Function

    Function ObtenerFolioIngresoFormaPago() As String
        'EstaFunción obtiene y regresa el Folio de ingreo de formas de pago que se debe almacenar.
        Dim Prefijo As String
        Dim Consecutivo As String

        'Obtener el siguiente folio de ingreso de Formas de Pago, para posteriormente usarlo al dar de alta las formas de pago.
        gStrSql = "select Prefijo , Consecutivo + 1 as Consecutivo From CatFolios where DescFolio= 'INGRESOS FORMAS DE PAGO' "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        Prefijo = Trim(RsGral.Fields("Prefijo").Value)
        Consecutivo = RsGral.Fields("Consecutivo").Value
        ObtenerFolioIngresoFormaPago = Prefijo & VB6.Format(CStr(gintCodAlmacen), "00") & VB6.Format(CStr(gintCodCaja), "00") & CStr(Year(Today)) & VB6.Format(CStr(Month(Today)), "00") & VB6.Format(CStr(VB.Day(Today)), "00") & VB6.Format(Consecutivo, "0000")
    End Function

    Private Sub txtmnCambio_Change()
        If CDbl(Numerico(txtmnCambio.Text)) < 0 Then
            txtmnCambio.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
            txtmnCambio.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
        Else
            txtmnCambio.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000018)
            txtmnCambio.ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
        End If
    End Sub

    Function ObtenerPartidaIngresosFormaDesPago(ByRef FolioIngreso As String) As Integer
        'Esta Función Obtiene el Número de Partida consecutivo para guardar un Ingreso de FOrma de Pago.
        gStrSql = "Select IsNull(max(NumPartida) , 0 ) + 1 as NumPartida from IngresosFormaDePago Where FolioIngreso = '" & FolioIngreso & "'"

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            ObtenerPartidaIngresosFormaDesPago = RsGral.Fields("NumPartida").Value
        End If
    End Function

    Sub InicializaVariables()
        mblnImpteFPRCMayorTotalAPagar = False
        gStrSql = "SELECT     CodFormaPago, Ltrim(Rtrim(DescFormaPago)) as DescFormaPago " & "From dbo.CatFormasPago " & "WHERE   (EsDolar = 0) AND  (EsCheque = 0) AND (EsDevolucion = 0) AND (EsDocumentoInterno = 0) AND (EsTarjeta = 0) AND (Estatus = 'V') " & "ORDER BY CODFORMAPAGO "

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            dbcMoneda.Text = RsGral.Fields("DescFormaPago").Value
            DCLostFocus(dbcMoneda, gStrSql, intCodFormaPago)
            txtCodFormaPago.Text = CStr(intCodFormaPago)
        End If
    End Sub

    Function ObtenerCodFormaPago() As Object
        On Error GoTo Merr
        'ModEstandar.BorraCmd
        gStrSql = "SELECT     CodFormaPago, Ltrim(Rtrim(DescFormaPago)) as DescFormaPago , EsDolar  " & "From dbo.CatFormasPago " & "WHERE     (EsCheque = 0) AND (EsDevolucion = 0) AND (EsDocumentoInterno = 0) AND (EsTarjeta = 0) AND (Estatus = 'V') AND (DescFormaPago LIKE '" & Trim(dbcMoneda.Text) & "') " & "ORDER BY CODFORMAPAGO "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            ObtenerCodFormaPago = RsGral.Fields("CodFormaPago").Value
            txtEsDolarFPCambio.Text = CStr(CBool(RsGral.Fields("EsDolar").Value))
        Else
            ObtenerCodFormaPago = 0
            txtEsDolarFPCambio.Text = ""
        End If
        txtCodFormaPago.Text = ObtenerCodFormaPago

        Exit Function
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ExistenFP() As Boolean
        On Error GoTo Merr
        gStrSql = "Select * from CatFormasPago Where Estatus = 'V'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount <= 0 Then
            MsgBox("No existen información almacenada sobre formas de pago" & vbNewLine & "Verifique por favor....", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            Exit Function
        Else
            ExistenFP = True
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPagosSalMercancia))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._Marco_3 = New System.Windows.Forms.GroupBox()
        Me.txtdoCambio = New System.Windows.Forms.Label()
        Me.txtdoTotalPago = New System.Windows.Forms.Label()
        Me.txtdoAPagar = New System.Windows.Forms.Label()
        Me.txtdoAPagar4Decimales = New System.Windows.Forms.Label()
        Me.txtdoTotalPago4Decimales = New System.Windows.Forms.Label()
        Me.txtdoCambio4Decimales = New System.Windows.Forms.Label()
        Me._lblEtiqueta_28 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_29 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_30 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_31 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_32 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_34 = New System.Windows.Forms.Label()
        Me.txtDolar = New System.Windows.Forms.TextBox()
        Me._Marco_1 = New System.Windows.Forms.GroupBox()
        Me.txtTotal = New System.Windows.Forms.Label()
        Me.txtIVA = New System.Windows.Forms.Label()
        Me.txtDescuento = New System.Windows.Forms.Label()
        Me.txtSubtotal = New System.Windows.Forms.Label()
        Me.txtSubtotal4Decimales = New System.Windows.Forms.Label()
        Me.txtDescuento4Decimales = New System.Windows.Forms.Label()
        Me.txtIVA4Decimales = New System.Windows.Forms.Label()
        Me.txtTotal4Decimales = New System.Windows.Forms.Label()
        Me._lblEtiqueta_36 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_22 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_16 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_15 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_14 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_2 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_1 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_0 = New System.Windows.Forms.Label()
        Me._Marco_2 = New System.Windows.Forms.GroupBox()
        Me.txtmnCambio = New System.Windows.Forms.Label()
        Me.txtmnTotalPago = New System.Windows.Forms.Label()
        Me.txtmnAPagar = New System.Windows.Forms.Label()
        Me.txtmnAPagar4Decimales = New System.Windows.Forms.Label()
        Me.txtmnTotalPago4Decimales = New System.Windows.Forms.Label()
        Me.txtmnCambio4Decimales = New System.Windows.Forms.Label()
        Me._lblEtiqueta_4 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_18 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_20 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_19 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_13 = New System.Windows.Forms.Label()
        Me._lblEtiqueta_12 = New System.Windows.Forms.Label()
        Me._Marco_0 = New System.Windows.Forms.GroupBox()
        Me.txtEsDolarFPCambio = New System.Windows.Forms.TextBox()
        Me.txtCodFormaPago = New System.Windows.Forms.TextBox()
        Me.fraMoneda = New System.Windows.Forms.Panel()
        Me.dbcMoneda = New System.Windows.Forms.ComboBox()
        Me._lblEtiqueta_3 = New System.Windows.Forms.Label()
        Me.txtFormaOrigen = New System.Windows.Forms.TextBox()
        Me.txtImporte = New System.Windows.Forms.TextBox()
        Me.msgFormasPago = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._lblEtiqueta_26 = New System.Windows.Forms.Label()
        Me.Marco = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblEtiqueta = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._Marco_3.SuspendLayout()
        Me._Marco_1.SuspendLayout()
        Me._Marco_2.SuspendLayout()
        Me._Marco_0.SuspendLayout()
        Me.fraMoneda.SuspendLayout()
        CType(Me.msgFormasPago, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Marco, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblEtiqueta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_Marco_3
        '
        Me._Marco_3.BackColor = System.Drawing.SystemColors.Control
        Me._Marco_3.Controls.Add(Me.txtdoCambio)
        Me._Marco_3.Controls.Add(Me.txtdoTotalPago)
        Me._Marco_3.Controls.Add(Me.txtdoAPagar)
        Me._Marco_3.Controls.Add(Me.txtdoAPagar4Decimales)
        Me._Marco_3.Controls.Add(Me.txtdoTotalPago4Decimales)
        Me._Marco_3.Controls.Add(Me.txtdoCambio4Decimales)
        Me._Marco_3.Controls.Add(Me._lblEtiqueta_28)
        Me._Marco_3.Controls.Add(Me._lblEtiqueta_29)
        Me._Marco_3.Controls.Add(Me._lblEtiqueta_30)
        Me._Marco_3.Controls.Add(Me._lblEtiqueta_31)
        Me._Marco_3.Controls.Add(Me._lblEtiqueta_32)
        Me._Marco_3.Controls.Add(Me._lblEtiqueta_34)
        Me._Marco_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Marco_3.Location = New System.Drawing.Point(8, 152)
        Me._Marco_3.Name = "_Marco_3"
        Me._Marco_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Marco_3.Size = New System.Drawing.Size(264, 95)
        Me._Marco_3.TabIndex = 26
        Me._Marco_3.TabStop = False
        Me._Marco_3.Text = " Dólares "
        Me.ToolTip1.SetToolTip(Me._Marco_3, "Totales ( Dólares )")
        '
        'txtdoCambio
        '
        Me.txtdoCambio.BackColor = System.Drawing.SystemColors.Info
        Me.txtdoCambio.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtdoCambio.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtdoCambio.ForeColor = System.Drawing.Color.White
        Me.txtdoCambio.Location = New System.Drawing.Point(120, 64)
        Me.txtdoCambio.Name = "txtdoCambio"
        Me.txtdoCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtdoCambio.Size = New System.Drawing.Size(129, 21)
        Me.txtdoCambio.TabIndex = 35
        Me.txtdoCambio.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtdoTotalPago
        '
        Me.txtdoTotalPago.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(246, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtdoTotalPago.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtdoTotalPago.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtdoTotalPago.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtdoTotalPago.Location = New System.Drawing.Point(120, 40)
        Me.txtdoTotalPago.Name = "txtdoTotalPago"
        Me.txtdoTotalPago.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtdoTotalPago.Size = New System.Drawing.Size(129, 21)
        Me.txtdoTotalPago.TabIndex = 34
        Me.txtdoTotalPago.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtdoAPagar
        '
        Me.txtdoAPagar.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtdoAPagar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtdoAPagar.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtdoAPagar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtdoAPagar.Location = New System.Drawing.Point(120, 16)
        Me.txtdoAPagar.Name = "txtdoAPagar"
        Me.txtdoAPagar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtdoAPagar.Size = New System.Drawing.Size(129, 21)
        Me.txtdoAPagar.TabIndex = 33
        Me.txtdoAPagar.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtdoAPagar4Decimales
        '
        Me.txtdoAPagar4Decimales.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtdoAPagar4Decimales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtdoAPagar4Decimales.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtdoAPagar4Decimales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtdoAPagar4Decimales.Location = New System.Drawing.Point(120, 16)
        Me.txtdoAPagar4Decimales.Name = "txtdoAPagar4Decimales"
        Me.txtdoAPagar4Decimales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtdoAPagar4Decimales.Size = New System.Drawing.Size(105, 21)
        Me.txtdoAPagar4Decimales.TabIndex = 50
        Me.txtdoAPagar4Decimales.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtdoTotalPago4Decimales
        '
        Me.txtdoTotalPago4Decimales.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(246, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtdoTotalPago4Decimales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtdoTotalPago4Decimales.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtdoTotalPago4Decimales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtdoTotalPago4Decimales.Location = New System.Drawing.Point(120, 40)
        Me.txtdoTotalPago4Decimales.Name = "txtdoTotalPago4Decimales"
        Me.txtdoTotalPago4Decimales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtdoTotalPago4Decimales.Size = New System.Drawing.Size(105, 21)
        Me.txtdoTotalPago4Decimales.TabIndex = 49
        Me.txtdoTotalPago4Decimales.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtdoCambio4Decimales
        '
        Me.txtdoCambio4Decimales.BackColor = System.Drawing.SystemColors.Info
        Me.txtdoCambio4Decimales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtdoCambio4Decimales.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtdoCambio4Decimales.ForeColor = System.Drawing.Color.Black
        Me.txtdoCambio4Decimales.Location = New System.Drawing.Point(120, 64)
        Me.txtdoCambio4Decimales.Name = "txtdoCambio4Decimales"
        Me.txtdoCambio4Decimales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtdoCambio4Decimales.Size = New System.Drawing.Size(105, 21)
        Me.txtdoCambio4Decimales.TabIndex = 48
        Me.txtdoCambio4Decimales.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblEtiqueta_28
        '
        Me._lblEtiqueta_28.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_28.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_28.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_28.Location = New System.Drawing.Point(12, 45)
        Me._lblEtiqueta_28.Name = "_lblEtiqueta_28"
        Me._lblEtiqueta_28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_28.Size = New System.Drawing.Size(68, 20)
        Me._lblEtiqueta_28.TabIndex = 32
        Me._lblEtiqueta_28.Text = "Total Pago:"
        '
        '_lblEtiqueta_29
        '
        Me._lblEtiqueta_29.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_29.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_29.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_29.Location = New System.Drawing.Point(12, 70)
        Me._lblEtiqueta_29.Name = "_lblEtiqueta_29"
        Me._lblEtiqueta_29.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_29.Size = New System.Drawing.Size(68, 20)
        Me._lblEtiqueta_29.TabIndex = 31
        Me._lblEtiqueta_29.Text = "Cambio:"
        '
        '_lblEtiqueta_30
        '
        Me._lblEtiqueta_30.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_30.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_30.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_30.Location = New System.Drawing.Point(88, 46)
        Me._lblEtiqueta_30.Name = "_lblEtiqueta_30"
        Me._lblEtiqueta_30.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_30.Size = New System.Drawing.Size(21, 20)
        Me._lblEtiqueta_30.TabIndex = 30
        Me._lblEtiqueta_30.Text = "$"
        '
        '_lblEtiqueta_31
        '
        Me._lblEtiqueta_31.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_31.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_31.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_31.Location = New System.Drawing.Point(88, 70)
        Me._lblEtiqueta_31.Name = "_lblEtiqueta_31"
        Me._lblEtiqueta_31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_31.Size = New System.Drawing.Size(21, 20)
        Me._lblEtiqueta_31.TabIndex = 29
        Me._lblEtiqueta_31.Text = "$"
        '
        '_lblEtiqueta_32
        '
        Me._lblEtiqueta_32.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_32.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_32.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_32.Location = New System.Drawing.Point(88, 20)
        Me._lblEtiqueta_32.Name = "_lblEtiqueta_32"
        Me._lblEtiqueta_32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_32.Size = New System.Drawing.Size(21, 19)
        Me._lblEtiqueta_32.TabIndex = 28
        Me._lblEtiqueta_32.Text = "$"
        '
        '_lblEtiqueta_34
        '
        Me._lblEtiqueta_34.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_34.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_34.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_34.Location = New System.Drawing.Point(12, 21)
        Me._lblEtiqueta_34.Name = "_lblEtiqueta_34"
        Me._lblEtiqueta_34.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_34.Size = New System.Drawing.Size(68, 20)
        Me._lblEtiqueta_34.TabIndex = 27
        Me._lblEtiqueta_34.Text = "A Pagar:"
        '
        'txtDolar
        '
        Me.txtDolar.AcceptsReturn = True
        Me.txtDolar.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.txtDolar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDolar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDolar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDolar.Location = New System.Drawing.Point(544, 16)
        Me.txtDolar.MaxLength = 0
        Me.txtDolar.Name = "txtDolar"
        Me.txtDolar.ReadOnly = True
        Me.txtDolar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDolar.Size = New System.Drawing.Size(73, 21)
        Me.txtDolar.TabIndex = 19
        Me.txtDolar.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDolar, "Tipo de Cambio Dólar")
        '
        '_Marco_1
        '
        Me._Marco_1.BackColor = System.Drawing.SystemColors.Control
        Me._Marco_1.Controls.Add(Me.txtTotal)
        Me._Marco_1.Controls.Add(Me.txtIVA)
        Me._Marco_1.Controls.Add(Me.txtDescuento)
        Me._Marco_1.Controls.Add(Me.txtSubtotal)
        Me._Marco_1.Controls.Add(Me.txtSubtotal4Decimales)
        Me._Marco_1.Controls.Add(Me.txtDescuento4Decimales)
        Me._Marco_1.Controls.Add(Me.txtIVA4Decimales)
        Me._Marco_1.Controls.Add(Me.txtTotal4Decimales)
        Me._Marco_1.Controls.Add(Me._lblEtiqueta_36)
        Me._Marco_1.Controls.Add(Me._lblEtiqueta_22)
        Me._Marco_1.Controls.Add(Me._lblEtiqueta_16)
        Me._Marco_1.Controls.Add(Me._lblEtiqueta_15)
        Me._Marco_1.Controls.Add(Me._lblEtiqueta_14)
        Me._Marco_1.Controls.Add(Me._lblEtiqueta_2)
        Me._Marco_1.Controls.Add(Me._lblEtiqueta_1)
        Me._Marco_1.Controls.Add(Me._lblEtiqueta_0)
        Me._Marco_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Marco_1.Location = New System.Drawing.Point(11, 19)
        Me._Marco_1.Name = "_Marco_1"
        Me._Marco_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Marco_1.Size = New System.Drawing.Size(264, 126)
        Me._Marco_1.TabIndex = 10
        Me._Marco_1.TabStop = False
        Me.ToolTip1.SetToolTip(Me._Marco_1, "Sub Totales")
        '
        'txtTotal
        '
        Me.txtTotal.BackColor = System.Drawing.SystemColors.Info
        Me.txtTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtTotal.Location = New System.Drawing.Point(120, 88)
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotal.Size = New System.Drawing.Size(129, 21)
        Me.txtTotal.TabIndex = 23
        Me.txtTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtIVA
        '
        Me.txtIVA.BackColor = System.Drawing.SystemColors.Info
        Me.txtIVA.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtIVA.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtIVA.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtIVA.Location = New System.Drawing.Point(120, 64)
        Me.txtIVA.Name = "txtIVA"
        Me.txtIVA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtIVA.Size = New System.Drawing.Size(129, 21)
        Me.txtIVA.TabIndex = 22
        Me.txtIVA.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtDescuento
        '
        Me.txtDescuento.BackColor = System.Drawing.SystemColors.Info
        Me.txtDescuento.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDescuento.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDescuento.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDescuento.Location = New System.Drawing.Point(120, 40)
        Me.txtDescuento.Name = "txtDescuento"
        Me.txtDescuento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescuento.Size = New System.Drawing.Size(129, 21)
        Me.txtDescuento.TabIndex = 21
        Me.txtDescuento.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtSubtotal
        '
        Me.txtSubtotal.BackColor = System.Drawing.SystemColors.Info
        Me.txtSubtotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtSubtotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtSubtotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtSubtotal.Location = New System.Drawing.Point(120, 16)
        Me.txtSubtotal.Name = "txtSubtotal"
        Me.txtSubtotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSubtotal.Size = New System.Drawing.Size(129, 21)
        Me.txtSubtotal.TabIndex = 20
        Me.txtSubtotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtSubtotal4Decimales
        '
        Me.txtSubtotal4Decimales.BackColor = System.Drawing.SystemColors.Info
        Me.txtSubtotal4Decimales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtSubtotal4Decimales.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtSubtotal4Decimales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtSubtotal4Decimales.Location = New System.Drawing.Point(120, 16)
        Me.txtSubtotal4Decimales.Name = "txtSubtotal4Decimales"
        Me.txtSubtotal4Decimales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSubtotal4Decimales.Size = New System.Drawing.Size(129, 21)
        Me.txtSubtotal4Decimales.TabIndex = 47
        Me.txtSubtotal4Decimales.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtDescuento4Decimales
        '
        Me.txtDescuento4Decimales.BackColor = System.Drawing.SystemColors.Info
        Me.txtDescuento4Decimales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDescuento4Decimales.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDescuento4Decimales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtDescuento4Decimales.Location = New System.Drawing.Point(120, 40)
        Me.txtDescuento4Decimales.Name = "txtDescuento4Decimales"
        Me.txtDescuento4Decimales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescuento4Decimales.Size = New System.Drawing.Size(129, 21)
        Me.txtDescuento4Decimales.TabIndex = 46
        Me.txtDescuento4Decimales.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtIVA4Decimales
        '
        Me.txtIVA4Decimales.BackColor = System.Drawing.SystemColors.Info
        Me.txtIVA4Decimales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtIVA4Decimales.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtIVA4Decimales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtIVA4Decimales.Location = New System.Drawing.Point(120, 64)
        Me.txtIVA4Decimales.Name = "txtIVA4Decimales"
        Me.txtIVA4Decimales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtIVA4Decimales.Size = New System.Drawing.Size(129, 21)
        Me.txtIVA4Decimales.TabIndex = 45
        Me.txtIVA4Decimales.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtTotal4Decimales
        '
        Me.txtTotal4Decimales.BackColor = System.Drawing.SystemColors.Info
        Me.txtTotal4Decimales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtTotal4Decimales.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtTotal4Decimales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtTotal4Decimales.Location = New System.Drawing.Point(120, 88)
        Me.txtTotal4Decimales.Name = "txtTotal4Decimales"
        Me.txtTotal4Decimales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotal4Decimales.Size = New System.Drawing.Size(129, 21)
        Me.txtTotal4Decimales.TabIndex = 44
        Me.txtTotal4Decimales.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblEtiqueta_36
        '
        Me._lblEtiqueta_36.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_36.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_36.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_36.Location = New System.Drawing.Point(12, 63)
        Me._lblEtiqueta_36.Name = "_lblEtiqueta_36"
        Me._lblEtiqueta_36.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_36.Size = New System.Drawing.Size(60, 20)
        Me._lblEtiqueta_36.TabIndex = 18
        Me._lblEtiqueta_36.Text = "IVA :"
        '
        '_lblEtiqueta_22
        '
        Me._lblEtiqueta_22.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_22.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_22.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_22.Location = New System.Drawing.Point(88, 64)
        Me._lblEtiqueta_22.Name = "_lblEtiqueta_22"
        Me._lblEtiqueta_22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_22.Size = New System.Drawing.Size(21, 21)
        Me._lblEtiqueta_22.TabIndex = 17
        Me._lblEtiqueta_22.Text = "$"
        '
        '_lblEtiqueta_16
        '
        Me._lblEtiqueta_16.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_16.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_16.Location = New System.Drawing.Point(88, 89)
        Me._lblEtiqueta_16.Name = "_lblEtiqueta_16"
        Me._lblEtiqueta_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_16.Size = New System.Drawing.Size(21, 21)
        Me._lblEtiqueta_16.TabIndex = 16
        Me._lblEtiqueta_16.Text = "$"
        '
        '_lblEtiqueta_15
        '
        Me._lblEtiqueta_15.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_15.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_15.Location = New System.Drawing.Point(88, 39)
        Me._lblEtiqueta_15.Name = "_lblEtiqueta_15"
        Me._lblEtiqueta_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_15.Size = New System.Drawing.Size(21, 21)
        Me._lblEtiqueta_15.TabIndex = 15
        Me._lblEtiqueta_15.Text = "$"
        '
        '_lblEtiqueta_14
        '
        Me._lblEtiqueta_14.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_14.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_14.Location = New System.Drawing.Point(88, 16)
        Me._lblEtiqueta_14.Name = "_lblEtiqueta_14"
        Me._lblEtiqueta_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_14.Size = New System.Drawing.Size(21, 21)
        Me._lblEtiqueta_14.TabIndex = 14
        Me._lblEtiqueta_14.Text = "$"
        '
        '_lblEtiqueta_2
        '
        Me._lblEtiqueta_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_2.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_2.Location = New System.Drawing.Point(12, 88)
        Me._lblEtiqueta_2.Name = "_lblEtiqueta_2"
        Me._lblEtiqueta_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_2.Size = New System.Drawing.Size(60, 20)
        Me._lblEtiqueta_2.TabIndex = 13
        Me._lblEtiqueta_2.Text = "Total :"
        '
        '_lblEtiqueta_1
        '
        Me._lblEtiqueta_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_1.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_1.Location = New System.Drawing.Point(12, 40)
        Me._lblEtiqueta_1.Name = "_lblEtiqueta_1"
        Me._lblEtiqueta_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_1.Size = New System.Drawing.Size(76, 20)
        Me._lblEtiqueta_1.TabIndex = 12
        Me._lblEtiqueta_1.Text = "Descuento : "
        '
        '_lblEtiqueta_0
        '
        Me._lblEtiqueta_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_0.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_0.Location = New System.Drawing.Point(12, 16)
        Me._lblEtiqueta_0.Name = "_lblEtiqueta_0"
        Me._lblEtiqueta_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_0.Size = New System.Drawing.Size(60, 20)
        Me._lblEtiqueta_0.TabIndex = 11
        Me._lblEtiqueta_0.Text = "SubTotal :"
        '
        '_Marco_2
        '
        Me._Marco_2.BackColor = System.Drawing.SystemColors.Control
        Me._Marco_2.Controls.Add(Me.txtmnCambio)
        Me._Marco_2.Controls.Add(Me.txtmnTotalPago)
        Me._Marco_2.Controls.Add(Me.txtmnAPagar)
        Me._Marco_2.Controls.Add(Me.txtmnAPagar4Decimales)
        Me._Marco_2.Controls.Add(Me.txtmnTotalPago4Decimales)
        Me._Marco_2.Controls.Add(Me.txtmnCambio4Decimales)
        Me._Marco_2.Controls.Add(Me._lblEtiqueta_4)
        Me._Marco_2.Controls.Add(Me._lblEtiqueta_18)
        Me._Marco_2.Controls.Add(Me._lblEtiqueta_20)
        Me._Marco_2.Controls.Add(Me._lblEtiqueta_19)
        Me._Marco_2.Controls.Add(Me._lblEtiqueta_13)
        Me._Marco_2.Controls.Add(Me._lblEtiqueta_12)
        Me._Marco_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Marco_2.Location = New System.Drawing.Point(8, 272)
        Me._Marco_2.Name = "_Marco_2"
        Me._Marco_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Marco_2.Size = New System.Drawing.Size(264, 103)
        Me._Marco_2.TabIndex = 1
        Me._Marco_2.TabStop = False
        Me._Marco_2.Text = " Moneda Nacional "
        Me.ToolTip1.SetToolTip(Me._Marco_2, "Totales ( Moneda Nacional ) ")
        '
        'txtmnCambio
        '
        Me.txtmnCambio.BackColor = System.Drawing.SystemColors.Info
        Me.txtmnCambio.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtmnCambio.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtmnCambio.ForeColor = System.Drawing.Color.Black
        Me.txtmnCambio.Location = New System.Drawing.Point(120, 72)
        Me.txtmnCambio.Name = "txtmnCambio"
        Me.txtmnCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtmnCambio.Size = New System.Drawing.Size(129, 21)
        Me.txtmnCambio.TabIndex = 36
        Me.txtmnCambio.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtmnTotalPago
        '
        Me.txtmnTotalPago.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtmnTotalPago.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtmnTotalPago.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtmnTotalPago.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtmnTotalPago.Location = New System.Drawing.Point(120, 48)
        Me.txtmnTotalPago.Name = "txtmnTotalPago"
        Me.txtmnTotalPago.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtmnTotalPago.Size = New System.Drawing.Size(129, 21)
        Me.txtmnTotalPago.TabIndex = 25
        Me.txtmnTotalPago.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtmnAPagar
        '
        Me.txtmnAPagar.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtmnAPagar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtmnAPagar.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtmnAPagar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtmnAPagar.Location = New System.Drawing.Point(120, 24)
        Me.txtmnAPagar.Name = "txtmnAPagar"
        Me.txtmnAPagar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtmnAPagar.Size = New System.Drawing.Size(129, 21)
        Me.txtmnAPagar.TabIndex = 24
        Me.txtmnAPagar.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtmnAPagar4Decimales
        '
        Me.txtmnAPagar4Decimales.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtmnAPagar4Decimales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtmnAPagar4Decimales.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtmnAPagar4Decimales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtmnAPagar4Decimales.Location = New System.Drawing.Point(120, 24)
        Me.txtmnAPagar4Decimales.Name = "txtmnAPagar4Decimales"
        Me.txtmnAPagar4Decimales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtmnAPagar4Decimales.Size = New System.Drawing.Size(129, 21)
        Me.txtmnAPagar4Decimales.TabIndex = 53
        Me.txtmnAPagar4Decimales.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtmnTotalPago4Decimales
        '
        Me.txtmnTotalPago4Decimales.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me.txtmnTotalPago4Decimales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtmnTotalPago4Decimales.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtmnTotalPago4Decimales.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtmnTotalPago4Decimales.Location = New System.Drawing.Point(120, 48)
        Me.txtmnTotalPago4Decimales.Name = "txtmnTotalPago4Decimales"
        Me.txtmnTotalPago4Decimales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtmnTotalPago4Decimales.Size = New System.Drawing.Size(129, 21)
        Me.txtmnTotalPago4Decimales.TabIndex = 52
        Me.txtmnTotalPago4Decimales.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtmnCambio4Decimales
        '
        Me.txtmnCambio4Decimales.BackColor = System.Drawing.SystemColors.Info
        Me.txtmnCambio4Decimales.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtmnCambio4Decimales.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtmnCambio4Decimales.ForeColor = System.Drawing.Color.Black
        Me.txtmnCambio4Decimales.Location = New System.Drawing.Point(120, 72)
        Me.txtmnCambio4Decimales.Name = "txtmnCambio4Decimales"
        Me.txtmnCambio4Decimales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtmnCambio4Decimales.Size = New System.Drawing.Size(129, 21)
        Me.txtmnCambio4Decimales.TabIndex = 51
        Me.txtmnCambio4Decimales.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblEtiqueta_4
        '
        Me._lblEtiqueta_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_4.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_4.Location = New System.Drawing.Point(12, 22)
        Me._lblEtiqueta_4.Name = "_lblEtiqueta_4"
        Me._lblEtiqueta_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_4.Size = New System.Drawing.Size(68, 20)
        Me._lblEtiqueta_4.TabIndex = 2
        Me._lblEtiqueta_4.Text = "A Pagar:"
        '
        '_lblEtiqueta_18
        '
        Me._lblEtiqueta_18.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_18.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_18.Location = New System.Drawing.Point(88, 24)
        Me._lblEtiqueta_18.Name = "_lblEtiqueta_18"
        Me._lblEtiqueta_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_18.Size = New System.Drawing.Size(21, 21)
        Me._lblEtiqueta_18.TabIndex = 5
        Me._lblEtiqueta_18.Text = "$"
        '
        '_lblEtiqueta_20
        '
        Me._lblEtiqueta_20.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_20.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_20.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_20.Location = New System.Drawing.Point(88, 71)
        Me._lblEtiqueta_20.Name = "_lblEtiqueta_20"
        Me._lblEtiqueta_20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_20.Size = New System.Drawing.Size(21, 21)
        Me._lblEtiqueta_20.TabIndex = 7
        Me._lblEtiqueta_20.Text = "$"
        '
        '_lblEtiqueta_19
        '
        Me._lblEtiqueta_19.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_19.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_19.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_19.Location = New System.Drawing.Point(88, 46)
        Me._lblEtiqueta_19.Name = "_lblEtiqueta_19"
        Me._lblEtiqueta_19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_19.Size = New System.Drawing.Size(21, 21)
        Me._lblEtiqueta_19.TabIndex = 6
        Me._lblEtiqueta_19.Text = "$"
        '
        '_lblEtiqueta_13
        '
        Me._lblEtiqueta_13.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_13.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_13.Location = New System.Drawing.Point(12, 71)
        Me._lblEtiqueta_13.Name = "_lblEtiqueta_13"
        Me._lblEtiqueta_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_13.Size = New System.Drawing.Size(68, 20)
        Me._lblEtiqueta_13.TabIndex = 4
        Me._lblEtiqueta_13.Text = "Cambio:"
        '
        '_lblEtiqueta_12
        '
        Me._lblEtiqueta_12.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_12.ForeColor = System.Drawing.Color.Black
        Me._lblEtiqueta_12.Location = New System.Drawing.Point(12, 45)
        Me._lblEtiqueta_12.Name = "_lblEtiqueta_12"
        Me._lblEtiqueta_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_12.Size = New System.Drawing.Size(68, 20)
        Me._lblEtiqueta_12.TabIndex = 3
        Me._lblEtiqueta_12.Text = "Total Pago:"
        '
        '_Marco_0
        '
        Me._Marco_0.BackColor = System.Drawing.SystemColors.Control
        Me._Marco_0.Controls.Add(Me.txtEsDolarFPCambio)
        Me._Marco_0.Controls.Add(Me.txtCodFormaPago)
        Me._Marco_0.Controls.Add(Me.fraMoneda)
        Me._Marco_0.Controls.Add(Me.txtFormaOrigen)
        Me._Marco_0.Controls.Add(Me.txtImporte)
        Me._Marco_0.Controls.Add(Me._Marco_3)
        Me._Marco_0.Controls.Add(Me.txtDolar)
        Me._Marco_0.Controls.Add(Me._Marco_1)
        Me._Marco_0.Controls.Add(Me.msgFormasPago)
        Me._Marco_0.Controls.Add(Me._Marco_2)
        Me._Marco_0.Controls.Add(Me._lblEtiqueta_26)
        Me._Marco_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Marco_0.Location = New System.Drawing.Point(8, 0)
        Me._Marco_0.Name = "_Marco_0"
        Me._Marco_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Marco_0.Size = New System.Drawing.Size(633, 380)
        Me._Marco_0.TabIndex = 0
        Me._Marco_0.TabStop = False
        '
        'txtEsDolarFPCambio
        '
        Me.txtEsDolarFPCambio.AcceptsReturn = True
        Me.txtEsDolarFPCambio.BackColor = System.Drawing.SystemColors.Window
        Me.txtEsDolarFPCambio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEsDolarFPCambio.Enabled = False
        Me.txtEsDolarFPCambio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEsDolarFPCambio.Location = New System.Drawing.Point(152, 0)
        Me.txtEsDolarFPCambio.MaxLength = 0
        Me.txtEsDolarFPCambio.Name = "txtEsDolarFPCambio"
        Me.txtEsDolarFPCambio.ReadOnly = True
        Me.txtEsDolarFPCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEsDolarFPCambio.Size = New System.Drawing.Size(89, 19)
        Me.txtEsDolarFPCambio.TabIndex = 43
        Me.txtEsDolarFPCambio.Visible = False
        '
        'txtCodFormaPago
        '
        Me.txtCodFormaPago.AcceptsReturn = True
        Me.txtCodFormaPago.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodFormaPago.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodFormaPago.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodFormaPago.Location = New System.Drawing.Point(16, 0)
        Me.txtCodFormaPago.MaxLength = 0
        Me.txtCodFormaPago.Name = "txtCodFormaPago"
        Me.txtCodFormaPago.ReadOnly = True
        Me.txtCodFormaPago.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodFormaPago.Size = New System.Drawing.Size(33, 19)
        Me.txtCodFormaPago.TabIndex = 42
        Me.txtCodFormaPago.Visible = False
        '
        'fraMoneda
        '
        Me.fraMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.fraMoneda.Controls.Add(Me.dbcMoneda)
        Me.fraMoneda.Controls.Add(Me._lblEtiqueta_3)
        Me.fraMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraMoneda.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMoneda.Location = New System.Drawing.Point(288, 8)
        Me.fraMoneda.Name = "fraMoneda"
        Me.fraMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMoneda.Size = New System.Drawing.Size(201, 33)
        Me.fraMoneda.TabIndex = 39
        '
        'dbcMoneda
        '
        Me.dbcMoneda.Location = New System.Drawing.Point(80, 8)
        Me.dbcMoneda.Name = "dbcMoneda"
        Me.dbcMoneda.Size = New System.Drawing.Size(120, 21)
        Me.dbcMoneda.TabIndex = 40
        '
        '_lblEtiqueta_3
        '
        Me._lblEtiqueta_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblEtiqueta_3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblEtiqueta_3.Location = New System.Drawing.Point(24, 4)
        Me._lblEtiqueta_3.Name = "_lblEtiqueta_3"
        Me._lblEtiqueta_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_3.Size = New System.Drawing.Size(51, 29)
        Me._lblEtiqueta_3.TabIndex = 41
        Me._lblEtiqueta_3.Text = "Moneda Cambio"
        '
        'txtFormaOrigen
        '
        Me.txtFormaOrigen.AcceptsReturn = True
        Me.txtFormaOrigen.BackColor = System.Drawing.SystemColors.Window
        Me.txtFormaOrigen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFormaOrigen.Enabled = False
        Me.txtFormaOrigen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFormaOrigen.Location = New System.Drawing.Point(64, 0)
        Me.txtFormaOrigen.MaxLength = 0
        Me.txtFormaOrigen.Name = "txtFormaOrigen"
        Me.txtFormaOrigen.ReadOnly = True
        Me.txtFormaOrigen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFormaOrigen.Size = New System.Drawing.Size(81, 19)
        Me.txtFormaOrigen.TabIndex = 38
        Me.txtFormaOrigen.Visible = False
        '
        'txtImporte
        '
        Me.txtImporte.AcceptsReturn = True
        Me.txtImporte.BackColor = System.Drawing.SystemColors.Window
        Me.txtImporte.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImporte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImporte.Location = New System.Drawing.Point(344, 136)
        Me.txtImporte.MaxLength = 0
        Me.txtImporte.Name = "txtImporte"
        Me.txtImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImporte.Size = New System.Drawing.Size(69, 20)
        Me.txtImporte.TabIndex = 37
        Me.txtImporte.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtImporte.Visible = False
        '
        'msgFormasPago
        '
        Me.msgFormasPago.DataSource = Nothing
        Me.msgFormasPago.Location = New System.Drawing.Point(288, 48)
        Me.msgFormasPago.Name = "msgFormasPago"
        Me.msgFormasPago.OcxState = CType(resources.GetObject("msgFormasPago.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msgFormasPago.Size = New System.Drawing.Size(333, 324)
        Me.msgFormasPago.TabIndex = 9
        '
        '_lblEtiqueta_26
        '
        Me._lblEtiqueta_26.BackColor = System.Drawing.SystemColors.Control
        Me._lblEtiqueta_26.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblEtiqueta_26.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblEtiqueta_26.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblEtiqueta_26.Location = New System.Drawing.Point(496, 17)
        Me._lblEtiqueta_26.Name = "_lblEtiqueta_26"
        Me._lblEtiqueta_26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblEtiqueta_26.Size = New System.Drawing.Size(42, 20)
        Me._lblEtiqueta_26.TabIndex = 8
        Me._lblEtiqueta_26.Text = "Dólar :"
        '
        'frmPagosSalMercancia
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(646, 386)
        Me.Controls.Add(Me._Marco_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(108, 122)
        Me.MaximizeBox = False
        Me.Name = "frmPagosSalMercancia"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Totales de la Venta"
        Me._Marco_3.ResumeLayout(False)
        Me._Marco_1.ResumeLayout(False)
        Me._Marco_2.ResumeLayout(False)
        Me._Marco_0.ResumeLayout(False)
        Me.fraMoneda.ResumeLayout(False)
        CType(Me.msgFormasPago, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Marco, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblEtiqueta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

End Class