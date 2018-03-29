'**********************************************************************************************************************'
'*PROGRAMA: MODULO DE DERECHOS JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Public Module ModDerechos

    Public Sub ChecaDerechos(ByRef CodUsuario As Integer)
        ''''09/04/03 FERNANDO
        On Error GoTo Errores
        'MenuPrincipal.MenuConfiguracionInicial(("F"))
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ''Dim I As Integer
        ''Dim VENTAS(15) As String
        ''Dim CREDITOINDIVIDUAL(10) As String
        ''Dim CXC(4) As String
        ''Dim COMISIONES(4) As String
        ''Dim OBRA(12) As String
        ''Dim SEGURIDAD(1) As String
        ''''''*MODULO VENTAS
        ''''''***CATALOGOS
        ''VENTAS(0) = "FRMVTSABCCATCLIENTES" ''ABC DE CLIENTES
        ''VENTAS(1) = "FRMVTSABCFTESPROSPECTACION" ''ABC A FUENTES DE PROSPECTACION
        ''VENTAS(2) = "FRMVTSABCEMPRESASPROSPECTADAS" ''ABC A EMPRESAS PROSPECTADAS
        ''VENTAS(3) = "FRMVTSABCCATAGRUPADETALLEOBRA" ''ABC A AGRUPADOR DE DETALLES DE OBRA
        ''VENTAS(4) = "FRMVTSABCCATDETALLEOBRA" ''ABC A DETALLE DE OBRA
        ''VENTAS(5) = "FRMVTSABCVENDEDORES" ''ABC A VENDEDORES
        ''VENTAS(6) = "FRMVTSABCCATGRUPOPAQUETES" ''ABC A GRUPOS DE PAQUETES
        ''VENTAS(7) = "" ''PARAMETROS PARA VENTA DE VIVIENDA
        ''VENTAS(8) = "FRMVTSABCCATESTADOCIVIL" ''ABC A ESTADOS CIVILES
        ''VENTAS(9) = "FRMVTSABCCATVIVIENDAS" ''ABC A VIVIENDAS
        ''VENTAS(10) = "" ''PRECALIFICACION
        ''VENTAS(11) = "" ''ADMINISTRACION GENERAL DE PROSPECTOS Y CLIENTES
        ''''CAMBIO DE UBICACION Y VENDEDORES
        ''VENTAS(12) = "" ''CAMBIO DE UBICACION DE VIVIENDA ASIGNADA
        ''VENTAS(13) = "" ''CAMBIO DE VENDEDOR
        ''VENTAS(14) = "" ''ENTREGA DE VIVIENDAS
        ''''REPORTES DE VENTAS
        ''VENTAS(15) = ""
        '''''''*************************************************************************************
        '''''''*    DEFINICION DE ARREGLOS QUE CONTINEN LOS NOMBRES DE LAS FORMAS QUE CONSTITUYEN  *
        '''''''*    EL MENU DEL SISTEMA FINCAS                                                     *
        '''''''*************************************************************************************
        ''''''*MODULO CREDITO INDIVIDUAL
        ''''''***CATALOGOS
        ''CREDITOINDIVIDUAL(0) = "FRMCIDABCINSTITUCIONES" ''ABC A INSTITUCIONES
        ''CREDITOINDIVIDUAL(1) = "FRMCIDABCTIPOSCREDITO" ''ABC A TIPOS DE CREDITO
        ''CREDITOINDIVIDUAL(2) = "FRMCIDABCGRUPOS" ''ABC A GRUPOS
        ''CREDITOINDIVIDUAL(3) = "FRMCIDABCCATREQPORGRUPOS" ''ABC A REQUISITOS POR GRUPO
        ''CREDITOINDIVIDUAL(4) = "FRMCIDABCCATAREASRESPONSABLES" ''ABC A AREAS RESPONSABLES
        ''CREDITOINDIVIDUAL(5) = "FRMCIDABCCATETAPAINTERNA" ''ABC A ETAPAS INTERNAS
        ''CREDITOINDIVIDUAL(6) = "FRMCIDABCCATDOCTRAMITE" ''ABC A DOCUMENTOS PARA TRAMITE
        ''CREDITOINDIVIDUAL(7) = "" ''RECEPCION DE PAGOS A CREDITOS
        ''CREDITOINDIVIDUAL(8) = "" ''INTEGRACION DE EXPEDIENTES
        '''''REPORTES DE CREDITO INDIVIDUAL
        ''CREDITOINDIVIDUAL(9) = ""
        '''''''''''*************************************************************************
        ''''''*MODULO CUENTAS POR COBRAR
        ''''''***CATALOGOS
        ''CXC(0) = "FRMCXCABCCONCEPTOSCOBRANZA" ''ABC A CONCEPTOS DE COBRANZA
        ''CXC(1) = "FRMCXCABCTIPOSPAGARE" ''ABC A TIPOS DE PAGARE
        ''CXC(2) = "" ''ABC A FORMAS DE PAGO
        ''CXC(3) = "" ''PARAMETROS GENERALES DE COBRANZA
        '''''''''''*************************************************************************
        ''''''*MODULO COMISIONES
        ''''''***CATALOGOS
        ''COMISIONES(0) = "FRMCOMABCMOVIMIENTOS" ''ABC A MOVIMIENTOS DE COMISION
        ''COMISIONES(1) = "FRMCOMABCTABLASCOMISIONES" ''ABC A TABLAS DE COMISIONES
        ''COMISIONES(2) = "FRMCOMABCASIGNACIONTABLAS" ''ABC A ASIGNACION DE TABLAS
        ''COMISIONES(3) = "FRMCOMABCDISTRIBPAGOCOMISION" ''ABC A DISTRIBUCION DE PAGOS DE COMISIONES
        '''''********************************************************************************
        '''''MODULO OBRA
        ''OBRA(0) = "FRMOBRABCPLAZAS" ''ABC A PLAZAS
        ''OBRA(1) = "FRMOBRABCPROYECTOS" ''ABC A PROYECTOS
        ''OBRA(2) = "FRMOBRABCETAPASPROYECTO" ''ABC A ETAPAS POR PROYECTO
        ''OBRA(3) = "FRMOBRABCTIPOSVIVIENDA" ''ABC A TIPOS DE VIVIENDA
        ''OBRA(4) = "FRMOBRABCMODELOSVIVIENDA" ''ABC A MODELOS DE VIVIENDA
        ''OBRA(5) = "FRMOBRABCFACHADAS" ''ABC A FACHADAS
        ''OBRA(6) = "FRMOBRABCFICHAESPECIFICACIONES" ''ABC A FICHAS DE ESPECIFICACIONES TECNICAS
        ''OBRA(7) = "FRMOBRABCPRECIOSVIVIENDAS" ''ABC A PRECIOS DE VIVIENDAS
        ''OBRA(8) = "FRMOBRABCPORCENTAJEOBRA" ''ABC A PORCENTAJES DE AVANCE DE OBRA
        ''OBRA(9) = "FRMOBRABCEMPRESASPROMOTORAS" ''ABC A EMPRESAS PROMOTORAS
        ''OBRA(10) = "" ''REGISTRO DEL AVANCE DE OBRA
        ''OBRA(11) = "" ''REPORTES DE OBRA
        '''''********************************************************************************
        '''''MODULO SEGURIDAD
        ''SEGURIDAD(0) = "FRMABCMODULOS" ''ABC A MODULOS Y FUNCIONES
        ''SEGURIDAD(1) = "FRMABCUSUARIOS" ''ABC A USUARIOS Y ACCESOS
        '''''''***********************************************************************************
        '''''''*PROCESO DE LECTURA DE ACCESOS CORRESPONDIENTES A CADA USUARIO DEPENDIENDO DE LOS *
        '''''''*DERECHOS QUE SE LE HAYAN ASIGNADO                                                *
        '''''''***********************************************************************************
        '''''Explicacion del programa: Lo que se hace es buscar en la tabla de accesos, el codigo de
        '''''usuario, modulo y el nombre de la forma. Para saber si tiene derecho.
        '''''En caso de tener el registro se encontara de lo contrario no.
        '''''Asi si se recorre cada arreglo que continen los nombres de las formas para buscar dichos derechos
        '''''y si tiene derecho las opciones del menu se habilitan de lo contrario se inabilitan.
        ''With MenuPrincipal
        ''       '''*********VENTAS****************
        ''       '''*****CATALOGOS*****************
        ''       For I = 1 To 9
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & VENTAS(I - 1) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuVentasCatalogosOpc(I - 1).Enabled = True
        ''            Else
        ''                .mnuVentasCatalogosOpc(I - 1).Enabled = False
        ''            End If
        ''        Next I
        ''        '''*****PRECALIFICACION Y ADMINISTRACION GENERAL DE PROSPECTOS Y CLIENTES
        ''        For I = 1 To 2
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & VENTAS(I + 9) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuVentasOpc(I).Enabled = True
        ''            Else
        ''                .mnuVentasOpc(I).Enabled = False
        ''            End If
        ''        Next I
        ''        ''''*******CAMBIO DE UBICACION Y VENDEDORES
        ''        For I = 1 To 2
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & VENTAS(I + 11) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuVentasCamVendOpc(I - 1).Enabled = True
        ''            Else
        ''                .mnuVentasCamVendOpc(I - 1).Enabled = False
        ''            End If
        ''        Next I
        ''        '''''*********ENTREGA DE VIVIENDAS
        ''        For I = 1 To 1
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & VENTAS(I + 13) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuVentasOpc(I + 3).Enabled = True
        ''            Else
        ''                .mnuVentasOpc(I + 3).Enabled = False
        ''            End If
        ''        Next I
        ''         '''''*********REPORTES DE VENTAS
        ''        For I = 1 To 1
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & VENTAS(I + 14) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuVentasReportesOpc(I - 1).Enabled = True
        ''            Else
        ''                .mnuVentasReportesOpc(I - 1).Enabled = False
        ''            End If
        ''        Next I
        ''        ''****************TERMINA MODULO DE VENTAS***************************
        ''
        ''        ''*************CREDITO INDIVIDUAL
        ''        '''*********************CATALOGOS
        ''        For I = 1 To 7
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & CREDITOINDIVIDUAL(I - 1) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuGestoriaCatalogosOpc(I - 1).Enabled = True
        ''            Else
        ''                .mnuGestoriaCatalogosOpc(I - 1).Enabled = True
        ''            End If
        ''        Next I
        ''        '''RECEPCION DE PAGOS A CREDITOS E INTEGRACION DE EXPEDIENTES
        ''        For I = 1 To 2
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & CREDITOINDIVIDUAL(I + 6) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuGestoriaOpc(I).Enabled = True
        ''            Else
        ''                .mnuGestoriaOpc(I).Enabled = True
        ''            End If
        ''        Next I
        ''        '''REPORTES DE CREDITO INDIVIDUAL
        ''        For I = 1 To 1
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & CREDITOINDIVIDUAL(I + 8) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuGestoriaReportesOpc(I - 1).Enabled = True
        ''            Else
        ''                .mnuGestoriaReportesOpc(I - 1).Enabled = True
        ''            End If
        ''        Next I
        ''        '''TERMINA MODULO DE CREDITO INDIVIDUAL

        ''
        ''        '''********CUENTAS POR COBRAR
        ''        '''************CATALOGOS
        ''        I = 0
        ''        For I = 1 To 4
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & CXC(I - 1) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuCXCCatalogosOpc(I - 1).Enabled = True
        ''            Else
        ''                .mnuCXCCatalogosOpc(I - 1).Enabled = False
        ''            End If
        ''        Next I
        ''        '''****TERMINA MODULO CUENTAS POR COBRAR
        ''
        ''        '''*********COMISIONES
        ''        '''***********CATALOGOS
        ''        I = 0
        ''        For I = 1 To 4
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & COMISIONES(I - 1) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuComisionesCatalogosOpc(I - 1).Enabled = True
        ''            Else
        ''                .mnuComisionesCatalogosOpc(I - 1).Enabled = False
        ''            End If
        ''        Next I
        ''        ''''******TERMINA MODULO DE COMISIONES
        ''
        ''        ''*****OBRA
        ''        '''CATALOGOS
        ''        For I = 1 To 10
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & OBRA(I - 1) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuObraCatalogosOpc(I - 1).Enabled = True
        ''            Else
        ''                .mnuObraCatalogosOpc(I - 1).Enabled = False
        ''            End If
        ''        Next I
        ''        '''REGISTRO DEL AVANCE DE OBRA
        ''        For I = 1 To 1
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & OBRA(I + 9) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuObraOpc(I).Enabled = True
        ''            Else
        ''                .mnuObraOpc(I).Enabled = False
        ''            End If
        ''        Next I
        ''        '''REPORTES DEL MODULO DE OBRA
        ''        For I = 1 To 1
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & OBRA(I + 10) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuObraReportesOpc(I - 1).Enabled = True
        ''            Else
        ''                .mnuObraReportesOpc(I - 1).Enabled = False
        ''            End If
        ''        Next I
        ''        ''''TERMINA MODULO DE OBRA
        ''
        ''        '''********SEGURIDAD
        ''        I = 0
        ''        For I = 1 To 2
        ''            ModEstandar.BorraCmd
        ''            gStrSql = "Select * From Accesos Where CodUsuario = " & gIntCodUsuario & " And Forma = '" & SEGURIDAD(I - 1) & "'"
        ''            Cmd.CommandText = "Up_Select_Datos"
        ''            Cmd.CommandType = adCmdStoredProc
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        ''            Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        ''            Set RsGral = Cmd.Execute
        ''            If RsGral.RecordCount > 0 Then
        ''                .mnuSeguridadOpc(I - 1).Enabled = True
        ''            Else
        ''                .mnuSeguridadOpc(I - 1).Enabled = False
        ''            End If
        ''        Next I
        ''End With
        ''Erase VENTAS
        ''Erase CREDITOINDIVIDUAL
        ''Erase CXC
        ''Erase COMISIONES
        ''Erase OBRA
        ''Erase SEGURIDAD

Errores:
        If Err.Number <> 0 Then ModErrores.Errores()

    End Sub
End Module