'**********************************************************************************************************************'
'*PROGRAMA: MODULO DE SP'S (PROCEDIMIENTOS ALMACENADOS) JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Module ModStoredProcedures
    ''' ****************************************************************************************************************************************************
    ''' SE MODIFICARON SP PARA MANEJO DE DIAMANTE SUELTO:  CATARTICULOS-ORDENESCOMPRAPRECAT
    ''' 27OCT2010 - MAVF
    '''
    ''' Ver 1.0       Estatus: Aprobado
    ''' ****************************************************************************************************************************************************



    Public Sub PR_IMECatGrupos(ByRef CodGrupo As String, ByRef DescGrupo As String, ByRef importe As String, ByRef PorcTasa As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 08/Mayo/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatGrupos" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescGrupo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(DescGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Importe", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(importe))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcTasa", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(ModEstandar.Numerico(PorcTasa))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECatMarcas(ByRef CodGrupo As String, ByRef CodMArca As String, ByRef DescMarca As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 12/Mayo/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatMarcas" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que regresa, en este caso, sera la nueva clave que le haya asignado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMarca", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodMArca)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescMarca", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(DescMarca)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECatFamilias(ByRef CodGrupo As String, ByRef CodFamilia As String, ByRef DescFamilia As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 13/Mayo/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatFamilias" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que regresa, en este caso, sera la nueva clave que le haya asignado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFamilia", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodFamilia)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescFamilia", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(DescFamilia)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECatModelos(ByRef CodGrupo As String, ByRef CodMArca As String, ByRef CodModelo As String, ByRef DescModelo As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 13/Mayo/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatModelos" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que regresa, en este caso, sera la nueva clave que le haya asignado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMarca", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodMArca)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodModelo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodModelo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescModelo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(DescModelo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECatLineas(ByRef CodGrupo As String, ByRef CodFamilia As String, ByRef COdLinea As String, ByRef DescLinea As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 13/Mayo/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatLineas" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que regresa, en este caso, sera la nueva clave que le haya asignado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFamilia", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodFamilia)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(COdLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescLinea", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(DescLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECatSubLineas(ByRef CodGrupo As String, ByRef CodFamilia As String, ByRef COdLinea As String, ByRef CodSubLinea As String, ByRef DescSubLinea As String, ByRef DescCorta As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 13/Mayo/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatSubLineas" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que regresa, en este caso, sera la nueva clave que le haya asignado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFamilia", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodFamilia)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(COdLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSubLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodSubLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescSubLinea", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(DescSubLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescCorta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(DescCorta)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECatCuentasBancarias(ByRef CodBanco As String, ByRef CtaBancaria As String, ByRef TipoCuenta As String, ByRef Sucursal As String, ByRef CuentaHabiente As String, ByRef LetraFolios As String, ByRef SaldoInicial As String, ByRef ConsecutivoChq As String, ByRef Moneda As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 16/Mayo/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatCuentasBancarias" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodBanco", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodBanco)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CtaBancaria", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 16, Trim(CtaBancaria)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCuenta", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, TipoCuenta))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sucursal", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 4, Format(Sucursal, "0000")))
        Cmd.Parameters.Append(Cmd.CreateParameter("CuentaHabiente", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(CuentaHabiente)))
        Cmd.Parameters.Append(Cmd.CreateParameter("LetraFolios", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, LetraFolios))
        Cmd.Parameters.Append(Cmd.CreateParameter("SaldoInicial", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(SaldoInicial)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConsecutivoChq", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 6, Format(ConsecutivoChq, "000000")))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda))) 'Tipo de Moneda que va Guardar la Cuenta
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub


    Public Sub PR_IMECatArticulos(ByRef CodArticulo As String, ByRef DescArticulo As String, ByRef CodGrupo As String, ByRef CodFamilia As String, ByRef COdLinea As String, ByRef CodSubLinea As String, ByRef CodKilates As String, ByRef CodMArca As String, ByRef CodModelo As String, ByRef CodTipoMaterial As String, ByRef Genero As String, ByRef Movimiento As String, ByRef Crono As String, ByRef CodUnidad As String, ByRef CodAlmacenOrigen As String, ByRef CodProveedor As String, ByRef CodigoArticuloProv As String, ByRef MonedaCompra As String, ByRef PrecioPubDolar As String, ByRef CostoFactura As String, ByRef CostoAdicional As String, ByRef CostoIndirecto As String, ByRef CostoReal As String, ByRef CostoFacturaPesos As String, ByRef CostoAdicionalPesos As String, ByRef CostoIndirectoPesos As String, ByRef PesosFijos As String, ByRef OrigenAnt As String, ByRef CodigoAnt As String, ByRef Adicional As String, ByRef mdsPeso As String, ByRef mdsColor As String, ByRef mdsPureza As String, ByRef mdsCertificado As String, ByRef Func As String, ByRef NumOp As String)

        BorraCmd()
        Cmd.CommandText = "UP_IME_CatArticulos" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado

        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que regresa, en este caso, será el código IDENTITY
        Cmd.Parameters.Append(Cmd.CreateParameter("CodArticulo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(CodArticulo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescArticulo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(DescArticulo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodGrupo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFamilia", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodFamilia))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(COdLinea))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSubLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodSubLinea))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodKilates", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodKilates))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMarca", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodMArca))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodModelo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodModelo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodTipoMaterial", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodTipoMaterial))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Genero", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Genero)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Movimiento", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Movimiento)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Crono", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Crono)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodUnidad", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodUnidad))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacenOrigen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodAlmacenOrigen))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProveedor", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodProveedor))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodigoArticuloProv", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(CodigoArticuloProv)))
        Cmd.Parameters.Append(Cmd.CreateParameter("MonedaCompra", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(MonedaCompra)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PrecioPubDolar", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(PrecioPubDolar))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoFactura", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(CostoFactura))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(CostoAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoIndirecto", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(CostoIndirecto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoReal", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(CostoReal))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoFacturaPesos", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(CostoFacturaPesos))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoAdicionalPesos", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(CostoAdicionalPesos))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoIndirectoPesos", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(CostoIndirectoPesos))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PesosFijos", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(PesosFijos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("OrigenAnt", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(ModEstandar.Numerico(OrigenAnt))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodigoAnt", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(CodigoAnt))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Adicional", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(Adicional)))

        '''27OCT2010 - MAVF
        Cmd.Parameters.Append(Cmd.CreateParameter("mdsPeso", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(mdsPeso))))
        Cmd.Parameters.Append(Cmd.CreateParameter("mdsColor", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(mdsColor)))
        Cmd.Parameters.Append(Cmd.CreateParameter("mdsPureza", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 4, Trim(mdsPureza)))
        Cmd.Parameters.Append(Cmd.CreateParameter("mdsCertificado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(mdsCertificado)))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECatFunciones(ByRef CodModulo As String, ByRef CodFuncion As String, ByRef Descripcion As String, ByRef Forma As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 21/Mayo/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatFunciones"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodModulo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodModulo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFuncion", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodFuncion))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descripcion", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 300, Trim(Descripcion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Forma", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 100, Trim(Forma)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECatModulos(ByRef CodModulo As String, ByRef DescModulo As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 21/Mayo/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatModulos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodModulo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodModulo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DesModulo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(DescModulo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECatUsuarios(ByRef CodUsuario As String, ByRef Nombre As String, ByRef Password As String, ByRef Grupo As String, ByRef CodGrupo As String, ByRef Tipo As String, ByRef CodModulo As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 02/Junio/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatUsuarios" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodUsuario", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodUsuario)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Nombre", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(Nombre)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Password", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 300, Trim(Password)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Grupo", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Grupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodGrupo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Tipo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Tipo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodModulo", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, ModEstandar.Numerico(CodModulo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IEAccesos(ByRef CodUsuario As String, ByRef Forma As String, ByRef CodModulo As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 03/Junio/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IE_Accesos" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodUsuario", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodUsuario)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Forma", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 100, Trim(Forma)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodModulo", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, ModEstandar.Numerico(CodModulo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMEOrdenesCompra(ByRef FolioOrdenCompra As String, ByRef CodProveedor As String, ByRef FechaOrdenCompra As String, ByRef FechaEntrega As String, ByRef Remision As String, ByRef Pedido As String, ByRef Origen As String, ByRef CodGrupo As String, ByRef CostoAdicional As String, ByRef CostoIndirectos As String, ByRef Entregar As String, ByRef SubTotal As String, ByRef Descuento As String, ByRef Iva As String, ByRef Total As String, ByRef Moneda As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef PorcIva As String, ByRef PorcDescto As String, ByRef TipoCambio As String, ByRef TipoCambioEuro As String, ByRef TipoCambioC As String, ByRef TipoCambioEuroC As String, ByRef PorcDesctoFinanciero As String, ByRef FechaCompraEI As String, ByRef FolioApartado As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 06/Junio/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_OrdenesCompra" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioOrdenCompra", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(FolioOrdenCompra)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProveedor", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodProveedor)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaOrdenCompra", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaOrdenCompra), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaEntrega", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaEntrega), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Remision", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Remision)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Pedido", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Pedido)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Origen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(Origen)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(CostoAdicional)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoIndirectos", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(CostoIndirectos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Entregar", ADODB.DataTypeEnum.adLongVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483647, Trim(Entregar)))
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotal", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(SubTotal)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(Descuento)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Iva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(Iva)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Total", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(Total)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcIva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(PorcIva)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcDescto", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(PorcDescto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioEuro", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambioEuro)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioC", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambioC)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioEuroC", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambioEuroC)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcDesctoFinanciero", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(PorcDesctoFinanciero)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCompraEI", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCompraEI), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioApartado", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioApartado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMEOrdenesCompraPreCat(ByRef FolioOrdenCompra As String, ByRef NumPartida As String, ByRef CodArticulo As String, ByRef DescArticulo As String, ByRef CantidadRegistro As String, ByRef CantidadRecepcion As String, ByRef CostoUnitario As String, ByRef Costo As String, ByRef Descuento As String, ByRef PorcDescuento As String, ByRef Iva As String, ByRef CostoAdicional As String, ByRef CostoIndirectos As String, ByRef CodGrupo As String, ByRef CodFamilia As String, ByRef COdLinea As String, ByRef CodSubLinea As String, ByRef CodKilates As String, ByRef CodMArca As String, ByRef CodModelo As String, ByRef CodTipoMaterial As String, ByRef Genero As String, ByRef Movimiento As String, ByRef Crono As String, ByRef CodUnidad As String, ByRef CodAlmacenOrigen As String, ByRef CodProveedor As String, ByRef CodigoArticuloProv As String, ByRef Estatus As String, ByRef Adicional As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 06/Junio/2003; 21/Julio/2003 : CodKilates y Crono
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_OrdenesCompraPreCat" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado

        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue, 4)) 'Valor que regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioOrdenCompra", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(FolioOrdenCompra)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(NumPartida)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodArticulo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodArticulo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescArticulo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(DescArticulo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CantidadRegistro", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CantidadRegistro)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CantidadRecepcion", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CantidadRecepcion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoUnitario", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(CostoUnitario)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Costo", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(Costo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(Descuento)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcDescuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(PorcDescuento)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Iva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(Iva)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(CostoAdicional)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoIndirectos", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(CostoIndirectos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFamilia", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodFamilia)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(COdLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSubLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodSubLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodKilates", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodKilates)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMarca", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodMArca)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodModelo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodModelo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodTipoMaterial", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodTipoMaterial)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Genero", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Genero)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Movimiento", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Movimiento)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Crono", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Crono)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodUnidad", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodUnidad)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacenOrigen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodAlmacenOrigen)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProveedor", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodProveedor)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodigoArticuloProv", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(CodigoArticuloProv)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Adicional", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(Adicional)))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMEOrdenesCompraPreCatAux(ByRef FolioOrdenCompra As String, ByRef NumPartida As String, ByRef CodArticulo As String, ByRef DescArticulo As String, ByRef CantidadRegistro As String, ByRef CantidadRecepcion As String, ByRef CostoUnitario As String, ByRef Costo As String, ByRef Descuento As String, ByRef PorcDescuento As String, ByRef Iva As String, ByRef CostoAdicional As String, ByRef CostoIndirectos As String, ByRef CodGrupo As String, ByRef CodFamilia As String, ByRef COdLinea As String, ByRef CodSubLinea As String, ByRef CodKilates As String, ByRef CodMArca As String, ByRef CodModelo As String, ByRef CodTipoMaterial As String, ByRef Genero As String, ByRef Movimiento As String, ByRef Crono As String, ByRef CodUnidad As String, ByRef CodAlmacenOrigen As String, ByRef CodProveedor As String, ByRef CodigoArticuloProv As String, ByRef Estatus As String, ByRef Adicional As String, ByRef PrecioPubDolar As String, ByRef MonedaPP As String, ByRef OrigenAnt As String, ByRef CodigoAnt As String, ByRef Imagen As String, ByRef mdsPeso As String, ByRef mdsColor As String, ByRef mdsPureza As String, ByRef mdsCertificado As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_OrdenesCompraPreCat"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc

        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue, 4)) 'Valor que regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioOrdenCompra", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(FolioOrdenCompra)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(NumPartida))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodArticulo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(CodArticulo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescArticulo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(DescArticulo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CantidadRegistro", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CantidadRegistro))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CantidadRecepcion", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CantidadRecepcion))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoUnitario", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(CostoUnitario))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Costo", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Costo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Descuento))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcDescuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(PorcDescuento))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Iva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Iva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(CostoAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoIndirectos", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(CostoIndirectos))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodGrupo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFamilia", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodFamilia))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(COdLinea))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSubLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodSubLinea))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodKilates", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodKilates))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMarca", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodMArca))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodModelo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodModelo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodTipoMaterial", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodTipoMaterial))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Genero", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Genero)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Movimiento", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Movimiento)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Crono", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Crono)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodUnidad", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodUnidad))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacenOrigen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodAlmacenOrigen))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProveedor", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(CodProveedor))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodigoArticuloProv", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(CodigoArticuloProv)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Adicional", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(Adicional)))

        Cmd.Parameters.Append(Cmd.CreateParameter("PrecioPubDolar", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(PrecioPubDolar))))
        Cmd.Parameters.Append(Cmd.CreateParameter("MonedaPP", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(MonedaPP)))
        Cmd.Parameters.Append(Cmd.CreateParameter("OrigenAnt", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(Numerico(OrigenAnt))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodigoAnt", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodigoAnt))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Imagen", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4000, Trim(Imagen)))

        '''27OCT2010 - MAVF
        Cmd.Parameters.Append(Cmd.CreateParameter("mdsPeso", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(mdsPeso))))
        Cmd.Parameters.Append(Cmd.CreateParameter("mdsColor", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(mdsColor)))
        Cmd.Parameters.Append(Cmd.CreateParameter("mdsPureza", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 4, Trim(mdsPureza)))
        Cmd.Parameters.Append(Cmd.CreateParameter("mdsCertificado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(mdsCertificado)))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECXPFacturas(ByRef CodProvAcreed As String, ByRef TipoFacturaCxP As String, ByRef TipoGasto As String, ByRef FolioOrdenCompra As String, ByRef FechaRegistro As String, ByRef FolioContraRecibo As String, ByRef FolioFactura As String, ByRef NumDocto As String, ByRef FechaFactura As String, ByRef FechaVencto As String, ByRef FechaRecepcion As String, ByRef DiasPago As String, ByRef FechaPago As String, ByRef SubTotal As String, ByRef Descuento As String, ByRef Retenciones As String, ByRef Iva As String, ByRef Total As String, ByRef IsrRet As String, ByRef IvaRet As String, ByRef Moneda As String, ByRef TipoCambio As String, ByRef TipoCambioEuro As String, ByRef PagoConChq As String, ByRef DescuentoFinanciero As String, ByRef SubTotalDF As String, ByRef IvaDF As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 02/Julio/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CXPFacturas" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue, 4)) 'Valor que regresa en este caso sera el número de documento
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodProvAcreed)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoFacturaCxP", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoFacturaCxP)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoGasto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoGasto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioOrdenCompra", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(FolioOrdenCompra)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaRegistro", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaRegistro), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioContraRecibo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioContraRecibo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioFactura", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioFactura)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumDocto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(NumDocto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFactura", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFactura), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaVencto", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaVencto), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaRecepcion", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaRecepcion), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DiasPago", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Numerico(DiasPago))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaPago", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaPago), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotal", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(SubTotal))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Descuento))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Retenciones", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Retenciones))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Iva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Iva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Total", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Total))))
        Cmd.Parameters.Append(Cmd.CreateParameter("IsrRet", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(IsrRet))))
        Cmd.Parameters.Append(Cmd.CreateParameter("IvaRet", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(IvaRet))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(TipoCambio))))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioEuro", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(TipoCambioEuro))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PagoConChq", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(PagoConChq)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescuentoFinanciero", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(DescuentoFinanciero))))
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotalDF", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(SubTotalDF))))
        Cmd.Parameters.Append(Cmd.CreateParameter("IvaDF", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(IvaDF))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMECXPFacturasDet(ByRef CodProvAcreed As String, ByRef FolioFactura As String, ByRef NumDocto As String, ByRef NumPartida As String, ByRef TipoFacturaCxP As String, ByRef Descripcion As String, ByRef Unidad As String, ByRef Cantidad As String, ByRef Precio As String, ByRef Descuento As String, ByRef PorcDescto As String, ByRef Iva As String, ByRef importe As String, ByRef PorcIva As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 02/Julio/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CXPFacturasDet" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue, 4)) 'Valor que regresa en este caso sera el número de Partida de Artículo(s)
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodProvAcreed)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioFactura", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioFactura)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumDocto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(NumDocto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(NumPartida))))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoFacturaCxP", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoFacturaCxP)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descripcion", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Descripcion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Unidad", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Unidad)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Cantidad", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(ModEstandar.Numerico(Cantidad))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Precio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Precio))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Descuento))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcDescto", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(PorcDescto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Iva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Iva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Importe", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(importe))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcIva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(PorcIva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMENotasCreditoCab(ByRef FolioNotaCredito As String, ByRef FechaNotaCredito As String, ByRef TipoNotaCredito As String, ByRef CodProvAcreed As String, ByRef FolioFactura As String, ByRef Concepto As String, ByRef Moneda As String, ByRef SubTotal As String, ByRef Descuento As String, ByRef Iva As String, ByRef Total As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef TipoCambio As String, ByRef TipoCambioEuro As String, ByRef TipoCambioAplic As String, ByRef FechaAplicacion As String, ByRef FolioPagoBancos As String, ByRef FolioNotaProveedor As String, ByRef TipoCambioEuroAplic As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 08/Julio/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_NotasCreditoCab"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioNotaCredito", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(FolioNotaCredito)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaNotaCredito", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaNotaCredito), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoNotaCredito", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoNotaCredito)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(ModEstandar.Numerico(CodProvAcreed))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioFactura", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioFactura)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Concepto", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(Concepto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotal", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(SubTotal)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(Descuento)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Iva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(Iva)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Total", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(Total)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioEuro", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambioEuro)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioAplic", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambioAplic)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaAplicacion", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaAplicacion), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioPagoBancos", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 13, Trim(FolioPagoBancos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioNotaProveedor", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(FolioNotaProveedor)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioEuroAplic", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(TipoCambioEuroAplic))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMENotasCreditoDet(ByRef FolioNotaCredito As String, ByRef NumPartida As String, ByRef FechaNotaCredito As String, ByRef CodArticulo As String, ByRef Descripcion As String, ByRef Unidad As String, ByRef CantidadDevol As String, ByRef Precio As String, ByRef Descuento As String, ByRef Iva As String, ByRef importe As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 08/Julio/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_NotasCreditoDet" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue, 4)) 'Valor que regresa en este caso sera el número de Partida de Artículo(s)
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioNotaCredito", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(FolioNotaCredito)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(NumPartida))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaNotaCredito", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaNotaCredito), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodArticulo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(CodArticulo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descripcion", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(Descripcion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Unidad", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Unidad)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CantidadDevol", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(CantidadDevol))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Precio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Precio))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Descuento))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Iva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(Iva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Importe", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(importe))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMEProgramacionPagos(ByRef FolioProgramacionP As String, ByRef NumPartida As String, ByRef CodProvAcreed As String, ByRef TipoFacturaCxP As String, ByRef TipoGasto As String, ByRef FolioFactura As String, ByRef FechaFactura As String, ByRef FechaPago As String, ByRef TotalPago As String, ByRef Moneda As String, ByRef TipoCambio As String, ByRef TipoCambioE As String, ByRef DescuentoFinanciero As String, ByRef SubTotalDF As String, ByRef IvaDF As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef TipoPagoProg As String, ByRef Efectivo As String, ByRef PasoBancos As String, ByRef FechaPasoBancos As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 23/Julio/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_ProgramacionPagos" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue, 4)) 'Regresa el número de partida
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioProgramacionP", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioProgramacionP)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(NumPartida))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(CodProvAcreed))))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoFacturaCxP", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoFacturaCxP)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoGasto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoGasto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioFactura", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioFactura)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFactura", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFactura), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaPago", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaPago), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TotalPago", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(TotalPago))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioE", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambioE)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescuentoFinanciero", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(DescuentoFinanciero))))
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotalDF", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(SubTotalDF))))
        Cmd.Parameters.Append(Cmd.CreateParameter("IvaDF", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(IvaDF))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoPagoProg", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoPagoProg)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Efectivo", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Efectivo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PasoBancos", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(PasoBancos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaPasoBancos", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaPasoBancos), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMEPPIntervalos(ByRef FolioProgramacionP As String, ByRef CodProvAcreed As String, ByRef TipoFacturaCxP As String, ByRef TipoGasto As String, ByRef FolioFactura As String, ByRef FechaFactura As String, ByRef FechaPago As String, ByRef TotalPago As String, ByRef Moneda As String, ByRef TipoCambio As String, ByRef TipoCambioE As String, ByRef DescuentoFinanciero As String, ByRef SubTotalDF As String, ByRef IvaDF As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef TipoPagoProg As String, ByRef Efectivo As String, ByRef Frecuencia As String, ByRef TipoIntervalo As String, ByRef Repeticiones As String, ByRef FechaInicio As String, ByRef FechaFin As String, ByRef Periodo As String, ByRef DiaSemana As String, ByRef DiaMes As String, ByRef Mes As String, ByRef Opcion As String, ByRef Cual As String, ByRef Cuando As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 23/Julio/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_PPIntervalos" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioProgramacionP", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioProgramacionP)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(CodProvAcreed))))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoFacturaCxP", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoFacturaCxP)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoGasto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoGasto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioFactura", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioFactura)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFactura", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFactura), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaPago", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaPago), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TotalPago", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(TotalPago))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioE", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambioE)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescuentoFinanciero", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(DescuentoFinanciero))))
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotalDF", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(SubTotalDF))))
        Cmd.Parameters.Append(Cmd.CreateParameter("IvaDF", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(IvaDF))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoPagoProg", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoPagoProg)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Efectivo", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Efectivo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Frecuencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Frecuencia)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoIntervalo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(TipoIntervalo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Repeticiones", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(Repeticiones))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaInicio", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaInicio), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Periodo", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(Periodo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DiaSemana", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 7, Trim(DiaSemana)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DiaMes", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(DiaMes))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Mes", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(Mes))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Opcion", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(Opcion))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Cual", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(Cual))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Cuando", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ModEstandar.Numerico(Cuando))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMEPagos(ByRef FolioProgramacionP As String, ByRef NumPartida As String, ByRef CodProvAcreed As String, ByRef TipoFacturaCxP As String, ByRef TipoGasto As String, ByRef FolioFactura As String, ByRef FechaFactura As String, ByRef FechaPago As String, ByRef TotalPago As String, ByRef Moneda As String, ByRef TipoCambio As String, ByRef TipoCambioE As String, ByRef SubTotalDF As String, ByRef IvaDF As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef TipoPagoProg As String, ByRef Efectivo As String, ByRef PasoBancos As String, ByRef FechaPasoBancos As String, ByRef FolioPagoBancos As String, ByRef PartidaPago As String, ByRef Func As String, ByRef NumOp As String)

        '------------------------------------------------------------------------------------
        'PAIMI 23/Julio/2003
        '------------------------------------------------------------------------------------

        BorraCmd()
        Cmd.CommandText = "UP_IME_Pagos" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue, 4)) 'Regresa el número de partida
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioProgramacionP", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioProgramacionP)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(NumPartida))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(CodProvAcreed))))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoFacturaCxP", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoFacturaCxP)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoGasto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoGasto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioFactura", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioFactura)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFactura", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFactura), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaPago", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaPago), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TotalPago", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(TotalPago))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioE", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(TipoCambioE)))
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotalDF", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(SubTotalDF))))
        Cmd.Parameters.Append(Cmd.CreateParameter("IvaDF", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(IvaDF))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoPagoProg", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoPagoProg)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Efectivo", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Efectivo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PasoBancos", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(PasoBancos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaPasoBancos", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaPasoBancos), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioPagoBancos", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 13, Trim(FolioPagoBancos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PartidaPago", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(ModEstandar.Numerico(PartidaPago))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_CuentasPorPagar(ByRef Moneda As String, ByRef TablaTmp As String, ByRef TablaDestino As String, ByRef FechaIni As String, ByRef FechaFin As String, ByRef cWHERE As String)
        '------------------------------------------------------------------------------------
        'PAIMI 23/Julio/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "UP_CuentasPorPagar" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TablaTMP", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(TablaTmp)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TablaDestino", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(TablaDestino)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIni", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIni), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("cWhere", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2000, Trim(cWHERE)))
    End Sub

    Public Sub PR_VentasSalidaMercancia(ByRef FechaIni As String, ByRef FechaFin As String, ByRef ConImpuesto As String, ByRef NomTablaResultado As String)
        '------------------------------------------------------------------------------------
        'PAIMI 21/Agosto/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "UP_VentasSalidaMercancia" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIni", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIni), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
    End Sub

    Public Sub PR_VentasSalidaMercanciaPorProveedor(ByRef CodProvAcreed As String, ByRef FechaIni As String, ByRef FechaFin As String, ByRef ConImpuesto As String, ByRef NomTablaResultado As String)
        '------------------------------------------------------------------------------------
        'PAIMI 22/Agosto/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "UP_VentasSalidaMercanciaPorProveedor" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodProvAcreed))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIni", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIni), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
    End Sub

    Public Sub PR_VentasSalidaMercanciaClasifArtic(ByRef FechaIni As String, ByRef FechaFin As String, ByRef ConImpuesto As String, ByRef NomTablaResultado As String, ByRef Where As String)
        '------------------------------------------------------------------------------------
        'PAIMI 26/Agosto/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "UP_VentasSalidaMercanciaClasifArtic" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIni", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIni), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Where", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2000, Trim(Where)))
    End Sub

    Public Sub PR_VentasSalidaMercanciaComparativo(ByRef CodSucursal As String, ByRef ConImpuesto As String, ByRef NomTablaResultado As String, ByRef Moneda As String, ByRef Anio As String, ByRef Mes As String, ByRef M1_DiaFin As String, ByRef M2_DiaFin As String, ByRef Sucursales As String)
        '------------------------------------------------------------------------------------
        'PAIMI 04/Septiembre/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "UP_VentasSalidaMercanciaComparativo" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodSucursal))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Anio", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(Anio))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Mes", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(Mes))))
        Cmd.Parameters.Append(Cmd.CreateParameter("M1_DiaFin", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(M1_DiaFin))))
        Cmd.Parameters.Append(Cmd.CreateParameter("M2_DiaFin", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(M2_DiaFin))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sucursales", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Sucursales)))
    End Sub

    Public Sub PR_VentasSalidaMercanciaComparativo_Anual(ByRef ConImpuesto As String, ByRef NomTablaResultado As String, ByRef Moneda As String, ByRef Anio As String, ByRef M1_DiaFin As String, ByRef M2_DiaFin As String, ByRef Sucursales As String)
        '------------------------------------------------------------------------------------
        'PAIMI 04/Septiembre/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "UP_VentasSalidaMercanciaComparativo_Anual" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Anio", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(Anio))))
        Cmd.Parameters.Append(Cmd.CreateParameter("M1_DiaFin", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(M1_DiaFin))))
        Cmd.Parameters.Append(Cmd.CreateParameter("M2_DiaFin", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(M2_DiaFin))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sucursales", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Sucursales)))
    End Sub

    Public Sub PR_VentasSalidaMercanciaUtilidad(ByRef FechaIni As String, ByRef FechaFin As String, ByRef ConImpuesto As String, ByRef NomTablaResultado As String, ByRef Moneda As String, ByRef CodSucursal As String)
        '------------------------------------------------------------------------------------
        'PAIMI 10/Septiembre/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_VentasSalidaMercanciaUtilidad" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIni", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIni), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodSucursal))))
    End Sub

    Public Sub PR_VentasSalidaMercanciaRelojeria(ByRef FechaIni As String, ByRef FechaFin As String, ByRef ConImpuesto As String, ByRef NomTablaResultado As String, ByRef CodRelojeria As String, ByRef CodMArca As String)
        '------------------------------------------------------------------------------------
        'PAIMI 19/Septiembre/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_VentasSalidaMercanciaRelojeria" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIni", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIni), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodRelojeria", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodRelojeria))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMarca", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodMArca))))
    End Sub

    Public Sub PR_VentasSalidaMercanciaRelojMaterial(ByRef FechaIni As String, ByRef FechaFin As String, ByRef ConImpuesto As String, ByRef NomTablaResultado As String, ByRef CodRelojeria As String, ByRef CodMaterial As String)
        '------------------------------------------------------------------------------------
        'PAIMI 22/Septiembre/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_VentasSalidaMercanciaRelojMaterial" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIni", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIni), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodRelojeria", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodRelojeria))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMaterial", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodMaterial))))
    End Sub

    Public Sub PR_VentasSalidaMercanciaFlujoVenta(ByRef FechaIni As String, ByRef FechaFin As String, ByRef ConImpuesto As String, ByRef NomTablaResultado As String, ByRef Moneda As String)
        '------------------------------------------------------------------------------------
        'PAIMI 23/Septiembre/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_VentasSalidaMercanciaFlujoVenta" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIni", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIni), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
    End Sub

    Public Sub PR_VentasSalidaMercanciaPorCliente(ByRef FechaIni As String, ByRef FechaFin As String, ByRef ConImpuesto As String, ByRef NomTablaResultado As String)
        '------------------------------------------------------------------------------------
        'PAIMI 26/Septiembre/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_VentasSalidaMercanciaPorCliente" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIni", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIni), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
    End Sub

    Public Sub PR_VentasSalidaMercanciaPorVendedor(ByRef FechaIni As String, ByRef FechaFin As String, ByRef ConImpuesto As String, ByRef NomTablaResultado As String, ByRef CodJoyeria As String, ByRef CodRelojeria As String, ByRef CodVarios As String)
        '------------------------------------------------------------------------------------
        'PAIMI 01/Octubre/2003
        '------------------------------------------------------------------------------------
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_VentasSalidaMercanciaPorVendedor" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIni", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIni), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(ConImpuesto))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NomTablaResultado", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NomTablaResultado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodJoyeria", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodJoyeria))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodRelojeria", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodRelojeria))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodVarios", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodVarios))))
    End Sub

    Public Sub PR_IMECatComisiones(ByRef FechaPeriodo As String, ByRef PorcComision As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'PAIMI 04/Noviembre/2003
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatComisiones" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaPeriodo", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaPeriodo), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcComision", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ModEstandar.Numerico(PorcComision))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    'Rosaura Torres López - 04/06/03 6:45 p.m.
    Public Sub PR_IEIngresos(ByRef FolioIngreso As String, ByRef FechaIngreso As String, ByRef CodSucursal As String, ByRef CodCaja As String, ByRef TipoIngreso As String, ByRef FolioMovto As String, ByRef CodCliente As String, ByRef Moneda As String, ByRef Total As String, ByRef ComisionBancaria As String, ByRef InteresesPromocion As String, ByRef TipoCambio As String, ByRef CodVendedor As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef cambio As String, ByRef CambioDol As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IE_Ingresos" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En este Caso sera Un Procedimiento Almacenado

        Cmd.Parameters.Append(Cmd.CreateParameter("FolioIngreso", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioIngreso)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIngreso", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaIngreso), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodSucursal)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCaja", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodCaja)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoIngreso", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoIngreso)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioMovto", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(FolioMovto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCliente", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodCliente)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("total", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Total)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ComisionBancaria", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ComisionBancaria)))
        Cmd.Parameters.Append(Cmd.CreateParameter("InteresesPromocion", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(InteresesPromocion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodVendedor", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodVendedor)))
        Cmd.Parameters.Append(Cmd.CreateParameter("estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , CDate(FechaCancel)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Cambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(cambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CambioDol", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(CambioDol)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opcion de Transacción
    End Sub

    'Rosaura Torres López - 04/06/03 6:55 p.m.
    Public Sub PR_IEIngresosFormasdePago(ByRef FolioIngreso As String, ByRef NumPartida As String, ByRef FechaIngreso As String, ByRef FolioMovto As String, ByRef CodFormaPago As String, ByRef importe As String, ByRef CodBanco As String, ByRef CodPlan As String, ByRef NoTarjeta As String, ByRef Autorizacion As String, ByRef NoCheque As String, ByRef FolioDevolucion As String, ByRef ComisionBancaria As String, ByRef InteresesPromocion As String, ByRef TipoCambio As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef PasoBancos As String, ByRef FechaPasoBancos As String, ByRef CodBancoRef As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IE_IngresosFormaDePago" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En este Caso sera Un Procedimiento Almacenado

        Cmd.Parameters.Append(Cmd.CreateParameter("FolioIngreso", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioIngreso)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, Numerico(NumPartida)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaIngreso", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaIngreso), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioMovto", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(FolioMovto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFormaPago", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodFormaPago)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Importe", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(importe)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodBanco", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodBanco))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodPlan", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodPlan))))
        Cmd.Parameters.Append(Cmd.CreateParameter("NoTarjeta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(NoTarjeta)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Autorizacion", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Autorizacion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NoCheque", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(NoCheque)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioDevolucion", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioDevolucion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ComisionBancaria", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ComisionBancaria)))
        Cmd.Parameters.Append(Cmd.CreateParameter("InteresesPromocion", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(InteresesPromocion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , CDate(FechaCancel)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PasoBancos", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(PasoBancos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaPasoBancos", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaPasoBancos), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodBancoRef", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodBancoRef))))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(NumOp))) 'Numero de Opcion de Transacción
    End Sub
    Public Sub PR_IMECatBancos(ByRef CodBanco As String, ByRef DescBanco As String, ByRef ControlInterno As String, ByRef Sucursal As String, ByRef Func As String, ByRef NumOp As String)
        'Rosaura Torres    09/Mayo/2003
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatBancos" 'Nombre del Propcedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor ke regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodBanco", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodBanco)))) 'Codigo del Banco a Guardar
        Cmd.Parameters.Append(Cmd.CreateParameter("DescBanco", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(DescBanco))) 'Descripciod Del Banco
        Cmd.Parameters.Append(Cmd.CreateParameter("ControlInterno", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(ControlInterno))) 'Bandera para saber si el Banco es Interno o Comercial
        Cmd.Parameters.Append(Cmd.CreateParameter("Sucursal", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Sucursal))) 'Bandera para Saber si es una Sucursal
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de transaccion
    End Sub
    Public Sub PR_IMECatTipoMaterial(ByRef CodTipoMaterial As String, ByRef DescTipoMaterial As String, ByRef DescCorta As String, ByRef Func As String, ByRef NumOp As String)
        'Rosaura Torres    12/Mayo/2003
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatTipoMaterial" 'Nombre del Propcedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor ke regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodTipoMaterial", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodTipoMaterial)))) 'Codigo del Tipo de Material
        Cmd.Parameters.Append(Cmd.CreateParameter("DescTipoMaterial", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(DescTipoMaterial))) 'Descripciod del material
        Cmd.Parameters.Append(Cmd.CreateParameter("DescCorta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(DescCorta)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de transaccion
    End Sub
    Public Sub PR_IMECatTalleres(ByRef CodTaller As String, ByRef DescTaller As String, ByRef Responsable As String, ByRef Domicilio As String, ByRef TipoTaller As String, ByRef Func As String, ByRef NumOp As String)
        'Rosaura Torres    12/Mayo/2003
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatTalleres" 'Nombre del Propcedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor ke regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodTaller", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodTaller)))) 'Codigo del Tipo de Material
        Cmd.Parameters.Append(Cmd.CreateParameter("DescTaller", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(DescTaller))) 'Descripcion del material
        Cmd.Parameters.Append(Cmd.CreateParameter("Responsable", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(Responsable))) 'Responsable del taller
        Cmd.Parameters.Append(Cmd.CreateParameter("Domicilio", ADODB.DataTypeEnum.adLongVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483647, Trim(Domicilio))) 'Domicilio del Taller
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoTaller", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoTaller))) 'Tipo del Taller J=Joyeria, R=Relojeria, F=Foraneo
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de transaccion
    End Sub

    Public Sub PR_IMECatAlmacen(ByRef CodAlmacen As String, ByRef DescAlmacen As String, ByRef Responsable As String, ByRef Auxiliar As String, ByRef Domicilio As String, ByRef TipoAlmacen As String, ByRef AlmGral As String, ByRef Func As String, ByRef NumOp As String)
        'Rosaura Torres    12/Mayo/2003
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatAlmacen" 'Nombre del Propcedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor ke regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodAlmacen)))) 'Codigo del Tipo de Material
        Cmd.Parameters.Append(Cmd.CreateParameter("DescAlmacen", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(DescAlmacen))) 'Descripcion del material
        Cmd.Parameters.Append(Cmd.CreateParameter("Responsable", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(Responsable))) 'Responsable del Almacen
        Cmd.Parameters.Append(Cmd.CreateParameter("Auxiliar", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(Auxiliar)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Domicilio", ADODB.DataTypeEnum.adLongVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483647, Trim(Domicilio))) 'Domicilio del Almacen
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoAlmacen", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoAlmacen))) 'Tipo del Almacen J=Joyeria, R=Relojeria, F=Foraneo
        Cmd.Parameters.Append(Cmd.CreateParameter("AlmGral", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(AlmGral))) 'Para Saber si es el Almacen General
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de transaccion
    End Sub

    Public Sub PR_IMECatVendedores(ByRef CodVendedor As String, ByRef DescVendedor As String, ByRef Domicilio As String, ByRef Telefono As String, ByRef Referencias As String, ByRef Comentarios As String, ByRef FechaAlta As String, ByRef Func As String, ByRef NumOp As String)
        'Rosaura Torres    12/Mayo/2003
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatVendedores" 'Nombre del Propcedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor ke regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodVendedor", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodVendedor))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescVendedor", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(DescVendedor)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Domicilio", ADODB.DataTypeEnum.adLongVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483647, Trim(Domicilio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Telefono", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Telefono)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Referencias", ADODB.DataTypeEnum.adLongVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483647, Trim(Referencias)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Comentarios", ADODB.DataTypeEnum.adLongVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483647, Trim(Comentarios)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaAlta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaAlta))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de transaccion
    End Sub

    Public Sub PR_IMECatProvAcreed(ByRef CodProvAcreed As String, ByRef DescProvACreed As String, ByRef Tipo As String, ByRef Nacional As String, ByRef Servicio As String, ByRef AgenciaAduanal As String, ByRef Domicilio As String, ByRef Ciudad As String, ByRef CP As String, ByRef Pais As String, ByRef Telefono As String, ByRef Rfc As String, ByRef Email As String, ByRef TaxId As String, ByRef DiasCredito As String, ByRef DesctoVolumen As String, ByRef DesctoFinanciero As String, ByRef ContactoVentas As String, ByRef TelsVentas As String, ByRef ContactoPagos As String, ByRef TelsPagos As String, ByRef CuentasBancarias As String, ByRef Observaciones As String, ByRef Func As String, ByRef NumOp As String)
        'Rosaura Torres    12/Mayo/2003
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatProvAcreed" 'Nombre del Propcedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor ke regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodProvAcreed))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescProvAcreed", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(DescProvACreed)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Tipo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Tipo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Nacional", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Nacional)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Servicio", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Servicio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("AgenciaAduanal", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(AgenciaAduanal)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Domicilio", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(Domicilio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Ciudad", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(Ciudad)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Cp", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(CP)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Pais", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(Pais)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Telefono", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Telefono)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Rfc", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(Rfc)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Email", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Email)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TaxId", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(TaxId)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DiasCredito", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput,  , CByte(Numerico(DiasCredito))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DesctoVolumen", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(DesctoVolumen))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DesctoFinanciero", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(DesctoFinanciero))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ContactoVentas", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(ContactoVentas)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TelsVentas", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(TelsVentas)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ContactoPagos", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(ContactoPagos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TelsPagos", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(TelsPagos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CuentasBancarias", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(CuentasBancarias)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Observaciones", ADODB.DataTypeEnum.adLongVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483647, Trim(Observaciones)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de transaccion
    End Sub

    Public Sub PR_IMConfiguracionGeneral(ByRef NombreEmpresa As String, ByRef RfcEmpresa As String, ByRef DomicilioEmpresa As String, ByRef TipoCambio As String, ByRef PorcUtilMinOperacion As String, ByRef RutaImagenes As String, ByRef TipoCambioEuro As String, ByRef VigenciaApartados As String, ByRef DriveLocal As String, ByRef TransferenciasentreSucursales As String, ByRef Codificacion As String, ByRef LapsoDifStock As String, ByRef Func As String, ByRef NumOp As String)
        'Rosaura Torres López 15/Mayo/2003
        BorraCmd()
        Cmd.CommandText = "UP_IM_ConfiguracionGeneral" 'Nombre del Propcedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("NombreEmpresa", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 60, Trim(NombreEmpresa)))
        Cmd.Parameters.Append(Cmd.CreateParameter("RfcEmpresa", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(RfcEmpresa)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DomicilioEmpresa", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 65, Trim(DomicilioEmpresa)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcUtilMinOperacion", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(PorcUtilMinOperacion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("RutaImagenes", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 255, Trim(RutaImagenes)))
        '    Cmd.Parameters.Append Cmd.CreateParameter("TasaImpuestoEstatal", adCurrency, adParamInput, 4, CCur(TasaImpuestoEstatal))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioEuro", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(TipoCambioEuro)))
        Cmd.Parameters.Append(Cmd.CreateParameter("VigenciaApartador", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(VigenciaApartados)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DriveLocal", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2, Trim(DriveLocal)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TransferenciasentreSucursales", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TransferenciasentreSucursales)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Codificacion", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Codificacion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("LapsoDifStock", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(LapsoDifStock)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de transaccion
    End Sub


    Public Sub PR_IMECatFormasPago(ByRef CodFormaPago As String, ByRef DescFormaPago As String, ByRef TeclaRapida As String, ByRef EsDolar As String, ByRef Escheque As String, ByRef EsDevolucion As String, ByRef EsDocumentoInterno As String, ByRef RequerirDocto As String, ByRef RequerirAutoriz As String, ByRef RestringirCambio As String, ByRef ConsiderarParaFact As String, ByRef ConsiderarparaRetiros As String, ByRef EsTarjeta As String, ByRef DescontarComBanc As String, ByRef PorcComision As String, ByRef PorcIvaComision As String, ByRef DescCorta As String, ByRef Estatus As String, ByRef CodBanco As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatFormasPago" 'Nombre del Propcedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor ke regresa en este caso sera el codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFormaPago", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodFormaPago))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescFormaPago", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(DescFormaPago)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TeclaRapida", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TeclaRapida)))
        Cmd.Parameters.Append(Cmd.CreateParameter("EsDolar", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(EsDolar)))
        Cmd.Parameters.Append(Cmd.CreateParameter("EsCheque", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Escheque)))
        Cmd.Parameters.Append(Cmd.CreateParameter("EsDevolucion", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(EsDevolucion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("EsDocumentoInterno", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(EsDocumentoInterno)))
        Cmd.Parameters.Append(Cmd.CreateParameter("RequerirDocto", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(RequerirDocto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("RequerirAutoriz", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(RequerirAutoriz)))
        Cmd.Parameters.Append(Cmd.CreateParameter("RestringirCambio", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(RestringirCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConsiderarParaFact", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(ConsiderarParaFact)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConsiderarParaRet", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(ConsiderarparaRetiros)))
        Cmd.Parameters.Append(Cmd.CreateParameter("EsTarjeta", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(EsTarjeta)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescontarComBanc", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(DescontarComBanc)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcComision", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(PorcComision)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcIvaComision", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(PorcIvaComision)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescCorta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(DescCorta)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus ", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodBanco", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodBanco)))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, CStr(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de transaccion
    End Sub


    Public Sub PR_IECatDenominaciones(ByRef CodFormaPago As String, ByRef Denominacion As String, ByRef Func As String, ByRef NumOp As String)
        'Rosaura Torres 16/Mayo/2003
        BorraCmd()
        Cmd.CommandText = "UP_IE_CatDenominaciones" 'Nombre del Propcedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFormaPago", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodFormaPago))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Denominacion", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Denominacion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de transaccion
    End Sub

    'Rosaura Torres López - 04/06/03 6:10 p.m.
    Public Sub PR_IMEMovimientosVentasCab(ByRef FolioVenta As String, ByRef FechaVenta As String, ByRef CodSucursal As String, ByRef CodCaja As String, ByRef CodVendedor As String, ByRef CodCliente As String, ByRef Nombre As String, ByRef Rfc As String, ByRef Condicion As String, ByRef Moneda As String, ByRef TipoCambio As String, ByRef SubTotal As String, ByRef Descuento As String, ByRef Iva As String, ByRef Total As String, ByRef Redondeo As String, ByRef Anticipo As String, ByRef PorcIva As String, ByRef Corte As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef SubTotalAdicional As String, ByRef DescuentoAdicional As String, ByRef IvaAdicional As String, ByRef TotalAdicional As String, ByRef RedondeoAdicional As String, ByRef AnticipoAdicional As String, ByRef EstatusAdicional As String, ByRef FolioFactura As String, ByRef FechaVenctoApartado As String, ByRef TipoMovto As String, ByRef VtaVExt As String, ByRef MonedaAnticipo As String, ByRef ApartadoPorCat As Boolean, ByRef CargaCto As Boolean, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IE_MovimientosVentasCab" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En este Caso sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioVenta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioVenta)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaVenta", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaVenta), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodSucursal))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCaja", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodCaja))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodVendedor", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodVendedor))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCliente", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodCliente))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Nombre", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(Nombre)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Rfc", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(Rfc)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Condicion", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2, Trim(Condicion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(TipoCambio))))
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotal", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(SubTotal))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(Descuento))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Iva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(Iva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Total", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(Total))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Redondeo", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(Redondeo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Anticipo", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(Anticipo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcIva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(PorcIva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Corte", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Corte)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , CDate(FechaCancel)))
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotalAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(SubTotalAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescuentoAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(DescuentoAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("IvaAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(IvaAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("TotalAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(TotalAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("RedondeoAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(RedondeoAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("AnticipoAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(AnticipoAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("EstatusAdicional", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(EstatusAdicional)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioFactura", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioFactura)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaVenctoApartado", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , CDate(FechaVenctoApartado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoMovto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoMovto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("VtaVExt", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(VtaVExt)))
        Cmd.Parameters.Append(Cmd.CreateParameter("MonedaAnticipo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(MonedaAnticipo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ApartadoPorCat", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(ApartadoPorCat)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CargaCto", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(CargaCto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opcion de Transacción
    End Sub

    'Rosaura Torres López - 04/06/03 6:30 p.m.
    Public Sub PR_IE_MovimientosVentasDet(ByRef FolioVenta As String, ByRef NumPartida As String, ByRef CodArticulo As String, ByRef DescArticulo As String, ByRef Cantidad As String, ByRef PorcPromociones As String, ByRef PorcDescuentos As String, ByRef ImptePromociones As String, ByRef ImpteDescuentos As String, ByRef PrecioLista As String, ByRef PrecioListaSinIva As String, ByRef PrecioReal As String, ByRef IvaReal As String, ByRef CostoVenta As String, ByRef ImptePromocionesAdicional As String, ByRef ImpteDescuentosAdicional As String, ByRef PrecioListaAdicional As String, ByRef PrecioListaSinIvaAdicional As String, ByRef PrecioRealAdicional As String, ByRef IvaRealAdicional As String, ByRef TipoMovto As String, ByRef DescArticuloAdicional As String, ByRef Metodo As String, ByRef PorcAdicional As String, ByRef FolioAdicional As String, ByRef FolioFactura As String, ByRef EstatusAdicional As String, ByRef CantidadAdicional As String, ByRef TipoCambioAdicional As String, ByRef MonedaAdicional As String, ByRef CondicionAdicional As String, ByRef PorcIvaAdicional As String, ByRef RedondeoAdicional As String, ByRef AnticipoAdicional As String, ByRef FechaVentaAdicional As String, ByRef CodSucursalAdicional As String, ByRef CodCajaAdicional As String, ByRef CodVendedorAdicional As String, ByRef CodClienteAdicional As String, ByRef Medidas As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IE_MovimientosVentasDet" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En este Caso sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioVenta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioVenta)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(NumPartida))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodArticulo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodArticulo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescArticulo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(DescArticulo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Cantidad", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(Cantidad))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcPromociones", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(PorcPromociones))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcDescuentos", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(PorcDescuentos))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ImptePromociones", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(ImptePromociones))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ImpteDescuentos", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(ImpteDescuentos))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PrecioLista", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(PrecioLista))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PrecioListaSinIva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(PrecioListaSinIva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PrecioReal", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(PrecioReal))))
        Cmd.Parameters.Append(Cmd.CreateParameter("IvaReal", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(IvaReal))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoVenta", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(CostoVenta))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ImptePromocionesAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(ImptePromocionesAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ImpteDescuentosAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(ImpteDescuentosAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PrecioListaAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(PrecioListaAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PrecioListaSinIvaAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(PrecioListaSinIvaAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PrecioRealAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(PrecioRealAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("IvaRealAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(IvaRealAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoMovto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoMovto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("DescArticuloAdicional", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(DescArticuloAdicional)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Metodo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Metodo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(PorcAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioAdicional", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioAdicional)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioFactura", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioFactura)))
        Cmd.Parameters.Append(Cmd.CreateParameter("EstatusAdicional", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(EstatusAdicional)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CantidadAdicional", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CantidadAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(TipoCambioAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("MonedaAdicional", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(MonedaAdicional)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CondicionAdicional", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2, Trim(CondicionAdicional)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcIvaAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(PorcIvaAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("RedondeoAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(RedondeoAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("AnticipoAdicional", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(AnticipoAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaVentaAdicional", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaVentaAdicional), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursalAdicional", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodSucursalAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCajaAdicional", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodCajaAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodVendedorAdicional", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodVendedorAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodClienteAdicional", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodClienteAdicional))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Medidas", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Medidas)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opcion de Transacción
    End Sub

    'Rosaura Torres López 14/Julio/2003
    Public Sub PR_IE_Retiros(ByRef FolioRetiro As String, ByRef FechaRetiro As String, ByRef CodSucursal As String, ByRef CodCaja As String, ByRef TipoRetiro As String, ByRef NickUsuario As String, ByRef Moneda As String, ByRef importe As String, ByRef MotivoRetiro As String, ByRef TipoCambio As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef PasoBancos As String, ByRef FechaPasoBancos As String, ByRef NumPartida As String, ByRef CodFormaPago As String, ByRef ImporteFormaPago As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "dbo.UP_IE_Retiros" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioRetiro", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioRetiro)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaReRetiro", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaRetiro), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodSucursal)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCaja", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodCaja)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoRetiro", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoRetiro)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NickUsuario", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(NickUsuario)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Importe", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(importe)))
        Cmd.Parameters.Append(Cmd.CreateParameter("MotivoRetiro", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(MotivoRetiro)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PasoBancos", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(PasoBancos)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaPasoBancos", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaPasoBancos), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(NumPartida)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFormaPago", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodFormaPago))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ImporteFormaPago", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ImporteFormaPago)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'Rosaura Torres López - 09/julio/2003
    Public Sub PR_IEDevolucionCab(ByRef FolioDevolucion As String, ByRef FechaDevolucion As String, ByRef CodSucursal As String, ByRef CodCaja As String, ByRef FolioVenta As String, ByRef FechaVenta As String, ByRef CodCliente As String, ByRef Nombre As String, ByRef Rfc As String, ByRef CodVendedor As String, ByRef Titular As String, ByRef MotivoDevol As String, ByRef TipoDevol As String, ByRef TotalDevol As String, ByRef Moneda As String, ByRef TipoCambio As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef FechaAplicacion As String, ByRef TotalDocto As String, ByRef TotalxAplicar As String, ByRef RedondeoDev As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IE_DevolucionesCab" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En este Caso sera Un Procedimiento Almacenado

        Cmd.Parameters.Append(Cmd.CreateParameter("FolioDevolucion", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioDevolucion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaDevolucion", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaDevolucion), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodSucursal)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCaja", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodCaja)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioVenta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioVenta)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaVenta", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaVenta), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCliente", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodCliente)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Nombre", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(Nombre)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Rfc", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(Rfc)))
        Cmd.Parameters.Append(Cmd.CreateParameter("codVendedor", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodVendedor)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Titular", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(Titular)))
        Cmd.Parameters.Append(Cmd.CreateParameter("MotivoDevol", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(MotivoDevol)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoDevol", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoDevol)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TotalDevol", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(TotalDevol)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaAplicacion", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 4, Format(CDate(FechaAplicacion), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TotalDocto", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(TotalDocto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TotalxAplicar", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(TotalxAplicar)))
        Cmd.Parameters.Append(Cmd.CreateParameter("RedondeoDev", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(RedondeoDev)))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opcion de Transacción
    End Sub

    'JUAN CARLOS OSUNA CORRALES 21 DE JULIO DE 2003
    Public Sub PR_IMEMovimientosReferencias(ByRef FolioMovto As String, ByRef NumPartida As String, ByRef ImporteDeposito As String, ByRef ReferenciaBanco As String, ByRef ImporteRef As String, ByRef Estatus As String, ByRef TipoReferencia As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_MovimientosReferencias"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioMovto", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 13, Trim(FolioMovto))) 'Folio del Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput,  , CInt(NumPartida))) 'Numero de Partida
        Cmd.Parameters.Append(Cmd.CreateParameter("ImporteDeposito", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ImporteDeposito))) 'Importe del Deposito
        Cmd.Parameters.Append(Cmd.CreateParameter("ReferenciaBanco", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(ReferenciaBanco))) 'Referencia del Banco
        Cmd.Parameters.Append(Cmd.CreateParameter("ImporteRef", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(ImporteRef))) 'Importe de la Referencia
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus))) 'Estatus del Deposito
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoReferencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoReferencia))) 'Tipo de Referencia
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'JUAN CARLOS OSUNA 12 DE MAYO DE 2003
    Public Sub PR_IMECatOrigenAplicRecursos(ByRef CodOrigenAplicR As String, ByRef DescOrigenAplicR As String, ByRef Aplicacion As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatOrigenAplicRecursos" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que Regresa, en este Caso sera el Codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodOrigenAplicR", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodOrigenAplicR)))) 'Codigo del Tipo de Origen y Aplicación
        Cmd.Parameters.Append(Cmd.CreateParameter("DescOrigenAplicR", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(DescOrigenAplicR))) 'Descripción del Tipo de Origen y Aplicación
        Cmd.Parameters.Append(Cmd.CreateParameter("Aplicacion", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Aplicacion))) 'Tipo de Aplicación ya Sea Entrada o Salida
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de Transacción
    End Sub

    'JUAN CARLOS OSUNA 13 DE MAYO DE 2003
    Public Sub PR_IMECatRubrosOrigenAplicRecursos(ByRef CodOrigAplicR As String, ByRef CodRubro As String, ByRef DescRubro As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatRubrosOrigenAplicRecursos" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodOrigAplicR", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodOrigAplicR)))) 'Codigo del Tipo de Origen y Aplicación
        Cmd.Parameters.Append(Cmd.CreateParameter("CodRubro", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodRubro)))) 'Codigo del Rubro
        Cmd.Parameters.Append(Cmd.CreateParameter("DescRubro", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(DescRubro))) 'Descripción del Rubro
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    Public Sub PR_IMECatClientes(ByRef CodCliente As String, ByRef DescCliente As String, ByRef Rfc As String, ByRef Domicilio As String, ByRef Colonia As String, ByRef Ciudad As String, ByRef TelCasa As String, ByRef TelOficina As String, ByRef Fax As String, ByRef CP As String, ByRef Email As String, ByRef TipoCte As String, ByRef FechaNacimiento As String, ByRef FechaNacimientoConyuge As String, ByRef AniversarioBodas As String, ByRef Estatus As String, ByRef Observaciones As String, ByRef FechaAlta As String, ByRef AlmacenVExt As String, ByRef CodSucursal As String, ByRef Conyuge As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatClientes" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado

        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor Que Regresa, En Este Caso Sera el Codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCliente", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Val(CodCliente)))) 'Codigo del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("DescCliente", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(DescCliente))) 'Nombre del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Rfc", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(Rfc))) 'Rfc del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Domicilio", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 65, Trim(Domicilio))) 'Domicilio del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Colonia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(Colonia))) 'Colonia Donde Vive el Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Ciudad", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(Ciudad))) 'Ciudad Donde Radica el Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("TelCasa", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(TelCasa))) 'Telefono de la Casa del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("TelOficina", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(TelOficina))) 'Telefono de la Oficina
        Cmd.Parameters.Append(Cmd.CreateParameter("Fax", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(Fax))) 'Fax del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("CP", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(CP))) 'Codigo Postal del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Email", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Email))) 'Email del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCte", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoCte))) 'Tipo de Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaNacimiento", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaNacimiento)) 'Fecha de Nacimiento del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaNacimientoConyuge", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaNacimientoConyuge)) 'Fecha de Nacimiento del Conyuge del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("AniversarioBodas", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, AniversarioBodas)) 'Aniversario de Bodas del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus))) 'Estatus del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Observaciones", ADODB.DataTypeEnum.adLongVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483647, Trim(Observaciones))) 'Observaciones Especiales
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaAlta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaAlta)) 'Fecha de Alta del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("AlmacenVExt", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(AlmacenVExt))) 'Codigo del Vendedor Externo
        Cmd.Parameters.Append(Cmd.CreateParameter("Codsucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Val(CodSucursal)))) 'Codigo de la sucursal a la que Pertenece el cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Conyuge", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(Conyuge))) 'Nombre del conyuge

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub




    'JUAN CARLOS OSUNA CORRALES 13 DE OCTUBRE DE 2003
    Public Sub PR_IMECatDesctosVExternos(ByRef CodGrupo As String, ByRef NumPartida As String, ByRef ImporteIni As String, ByRef ImporteFin As String, ByRef CodMArca As String, ByRef CodFamilia As String, ByRef PorcDescto As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatDesctosVExternos" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando en este caso sera un procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodGrupo)))) 'Codigo del Grupo
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(NumPartida)))) 'Numero de Partida
        Cmd.Parameters.Append(Cmd.CreateParameter("ImporteIni", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(ImporteIni)))) 'Importe Inicial
        Cmd.Parameters.Append(Cmd.CreateParameter("ImporteFin", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(ImporteFin)))) 'Importe Final
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMarca", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodMArca)))) 'Codigo de la Marca
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFamilia", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodFamilia)))) 'Codigo de la Familia
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcDescto", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(PorcDescto)))) 'Porcentaje de Descuento
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'JUAN CARLOS OSUNA CORRALES 27 DE MAYO DE 2003
    Public Sub PR_IMEConfiguracionGralPV(ByRef CodAlmacen As String, ByRef CapturarCantArts As String, ByRef Redondeo As String, ByRef Transferencia As String, ByRef PosicionDecimal As String, ByRef SimboloMonedaNac As String, ByRef EfectivoMaximo As String, ByRef RutaArchivo As String, ByRef ArchivoInvElectronico As String, ByRef Separador As String, ByRef Espacios As String, ByRef PermitirVtaSinExistencia As String, ByRef ConsultarXDescrip As String, ByRef AutorizCambiarCodigoCapt As String, ByRef AutorizCambiarLineaCapt As String, ByRef AutorizAbandonarCaptIni As String, ByRef IndicarSiProdNoSoportaDescto As String, ByRef AutorizConsultarFoliosVta As String, ByRef AutorizModificarDesctos As String, ByRef UtilMinXOperacion As String, ByRef MsgFiscal As String, ByRef MsgNormal As String, ByRef MsgCredito As String, ByRef MsgDevoluciones As String, ByRef ImpresionTransferencias As String, ByRef TasaIVA As String, ByRef CodigoViejo As Boolean, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_ConfiguracionGralPV" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En este Caso sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(CodAlmacen)))) 'Codigo del Almacen
        Cmd.Parameters.Append(Cmd.CreateParameter("CapturarCantArts", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(CapturarCantArts))) 'Capturar Cantidad de Articulos
        Cmd.Parameters.Append(Cmd.CreateParameter("Redondeo", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(Redondeo)))) 'Redondeo de las Ventas
        Cmd.Parameters.Append(Cmd.CreateParameter("Transferencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Transferencia))) 'Tipo de Transferencia E = Electrónica, D = Diskette
        Cmd.Parameters.Append(Cmd.CreateParameter("PosicionDecimal", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(PosicionDecimal)))) 'Numero de Posiciones Decimales
        Cmd.Parameters.Append(Cmd.CreateParameter("SimboloMonedaNac", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(SimboloMonedaNac))) 'Simbolo de Moneda que Aparecera en los Totales
        Cmd.Parameters.Append(Cmd.CreateParameter("EfectivoMaximo", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(EfectivoMaximo)))) 'Efectivo Maximo en Caja
        '    Cmd.Parameters.Append Cmd.CreateParameter("RutaImagen", adChar, adParamInput, 255, Trim(RutaImagen)) 'Ruta Donde se Encuentran las Imagenes
        Cmd.Parameters.Append(Cmd.CreateParameter("RutaArchivo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 255, Trim(RutaArchivo))) 'Ruta del Archivo de Inventario Electronico
        Cmd.Parameters.Append(Cmd.CreateParameter("ArchivoInvElectronico", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(ArchivoInvElectronico))) 'Nombre del Archivo de Inventario Electronico
        Cmd.Parameters.Append(Cmd.CreateParameter("Separador", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Separador))) 'Separador que se Usara en el Archivo de Inventario
        Cmd.Parameters.Append(Cmd.CreateParameter("Espacios", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(Espacios)))) 'Numero de Espacios que se Usaran
        Cmd.Parameters.Append(Cmd.CreateParameter("PermitirVtaSinExistencia", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(PermitirVtaSinExistencia))) 'Permitir Ventas sin Existencia
        Cmd.Parameters.Append(Cmd.CreateParameter("ConsultarXDescrip", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(ConsultarXDescrip))) 'Permitir Consultar X Descripcion
        Cmd.Parameters.Append(Cmd.CreateParameter("AutorizCambiarCodigoCapt", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(AutorizCambiarCodigoCapt))) 'Autorización Para Cambiar el Codigo Capturado
        Cmd.Parameters.Append(Cmd.CreateParameter("AutorizCambiarLineaCapt", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(AutorizCambiarLineaCapt))) 'Autorización para Cambiar la Linea Capturada
        Cmd.Parameters.Append(Cmd.CreateParameter("AutorizAbandonarCaptIni", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(AutorizAbandonarCaptIni))) 'Autorización para Abandonar la Captura Iniciada
        Cmd.Parameters.Append(Cmd.CreateParameter("IndicarSiProdNoSoportaDescto", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(IndicarSiProdNoSoportaDescto))) 'Indica si el Producto Soporta Descuento
        Cmd.Parameters.Append(Cmd.CreateParameter("AutorizConsultarFoliosVta", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(AutorizConsultarFoliosVta))) 'Autorización Para Consultar Folios de Venta
        Cmd.Parameters.Append(Cmd.CreateParameter("AutorizModificarDesctos", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(AutorizModificarDesctos))) 'Autorizacion para Modificar Descuentos
        Cmd.Parameters.Append(Cmd.CreateParameter("UtilMinXOperacion", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(UtilMinXOperacion)))) 'Utilidad Minima Por Operación
        Cmd.Parameters.Append(Cmd.CreateParameter("MsgFiscal", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(MsgFiscal))) 'Mensaje Fiscal
        Cmd.Parameters.Append(Cmd.CreateParameter("MsgNormal", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(MsgNormal))) 'Mensaje Normal
        Cmd.Parameters.Append(Cmd.CreateParameter("MsgCredito", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(MsgCredito))) 'Mensaje de Ventas
        Cmd.Parameters.Append(Cmd.CreateParameter("MsgDevoluciones", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(MsgDevoluciones))) 'Mensaje de Devoluciones
        Cmd.Parameters.Append(Cmd.CreateParameter("ImpresionTransferencias", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(ImpresionTransferencias)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TasaIVA", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(TasaIVA)))) 'Tasa de IVA
        Cmd.Parameters.Append(Cmd.CreateParameter("CodigoViejo", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(CodigoViejo))) 'Mostrar el código viejo
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opcion de Transacción
    End Sub


    Public Sub PR_IE_StockBasicoSucursal(ByRef Fecha As String, ByRef CodGrupo As String, ByRef CodFamilia As String, ByRef COdLinea As String, ByRef CodSubLinea As String, ByRef CodMArca As String, ByRef CodModelo As String, ByRef Stock As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'Rosaura  04/09/03
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IE_StockBasicoSucursal" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("Fecha", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(Fecha), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFamilia", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodFamilia)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(COdLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSubLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodSubLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMarca", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodMArca)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodModelo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodModelo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("stock", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Stock)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub




    'JUAN CARLOS OSUNA CORRALES 23 DE MAYO DE 2003
    Public Sub PR_IMEConfigTicketVenta(ByRef CodSucursal As String, ByRef Renglon As String, ByRef Etiqueta As String, ByRef Formula As String, ByRef Columna As String, ByRef Saltos As String, ByRef Grupo As String, ByRef Tipo As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_ConfigTicketVenta" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodSucursal)))) 'código sucursal
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(Renglon)))) 'Numero de Renglon
        Cmd.Parameters.Append(Cmd.CreateParameter("Etiqueta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, Trim(Etiqueta))) 'Etiqueta
        Cmd.Parameters.Append(Cmd.CreateParameter("Formula", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 100, Trim(Formula))) 'Formula
        Cmd.Parameters.Append(Cmd.CreateParameter("Columna", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(Columna)))) 'Columna
        Cmd.Parameters.Append(Cmd.CreateParameter("Saltos", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(Saltos)))) 'Saltos
        Cmd.Parameters.Append(Cmd.CreateParameter("Grupo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Grupo))) 'Grupo
        Cmd.Parameters.Append(Cmd.CreateParameter("Tipo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2, Trim(Tipo))) 'Tipo (Credito o Contado)
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opcion de Transacción
    End Sub


    'JUAN CARLOS OSUNA 19 DE MAYO DE 2003
    Public Sub PR_IMEConfigFactura(ByRef CodAlmacen As Object, ByRef RenTotales As String, ByRef RenEmpresa As String, ByRef ColEmpresa As String, ByRef RenRFC As String, ByRef ColRFC As String, ByRef RenFecha As String, ByRef ColFecha As String, ByRef RenFolio As String, ByRef ColFolio As String, ByRef RenCalle As String, ByRef ColCalle As String, ByRef RenColonia As String, ByRef ColColonia As String, ByRef RenCiudad As String, ByRef ColCiudad As String, ByRef RenEstado As String, ByRef ColEstado As String, ByRef RenCP As String, ByRef ColCP As String, ByRef RenTelefono As String, ByRef ColTelefono As String, ByRef RenSubTotal As String, ByRef ColSubTotal As String, ByRef RenIva As String, ByRef ColIva As String, ByRef RenTotal As String, ByRef ColTotal As String, ByRef RenDescto As String, ByRef ColDescto As String, ByRef RenImpLetra As String, ByRef ColImpLetra As String, ByRef RenLugarExped As String, ByRef ColLugarExped As String, ByRef CantRenXDet As String, ByRef RenPrimerPartida As String, ByRef ColCodigo As String, ByRef ColCantidad As String, ByRef ColDescripcion As String, ByRef ColDesctoDetalle As String, ByRef ColPromocion As String, ByRef ColPrecioVenta As String, ByRef ColImporte As String, ByRef ColIVAPartida As String, ByRef Ticket As String, ByRef ColLeyenda As String, ByRef RenLeyenda As String, ByRef Leyenda As String, ByRef TamLetra As String, ByRef LongCliente As String, ByRef LongDireccion As String, ByRef LongColonia As String, ByRef LongCiudad As String, ByRef LongEstado As String, ByRef LongLeyenda As String, ByRef LongProducto As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_ConfigFactura" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera un Procedimiento Almacenado
        'UPGRADE_WARNING: Couldn't resolve default property of object CodAlmacen. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodAlmacen)))) 'Código del Almacen
        Cmd.Parameters.Append(Cmd.CreateParameter("RenTotales", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenTotales)))) 'Numero Total de Renglones de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("RenEmpresa", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenEmpresa)))) 'Numero de Renglon Donde Aparecera en La Factura la Dirección de la Empresa
        Cmd.Parameters.Append(Cmd.CreateParameter("ColEmpresa", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColEmpresa)))) 'Numero de Columna Donde Aparecera en la Factura la Dirección de la Empresa
        Cmd.Parameters.Append(Cmd.CreateParameter("RenRFC", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenRFC)))) 'Numero de Renglon Donde Aparecera el RFC
        Cmd.Parameters.Append(Cmd.CreateParameter("ColRFC", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColRFC)))) 'Numero de Columna Donde Aparecera el RFC
        Cmd.Parameters.Append(Cmd.CreateParameter("RenFecha", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenFecha)))) 'Numero de Renglon Donde Aparecera la Fecha
        Cmd.Parameters.Append(Cmd.CreateParameter("ColFecha", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColFecha)))) 'Numero de Columna Donde Aparecera la Fecha
        Cmd.Parameters.Append(Cmd.CreateParameter("RenFolio", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenFolio)))) 'Numero de Renglon Donde Aparecera el Folio
        Cmd.Parameters.Append(Cmd.CreateParameter("ColFolio", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColFolio)))) 'Numero de Columna Donde Aparecera el Folio
        Cmd.Parameters.Append(Cmd.CreateParameter("RenCalle", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenCalle)))) 'Numero de Renglon Donde Aparecera la Calle
        Cmd.Parameters.Append(Cmd.CreateParameter("ColCalle", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColCalle)))) 'Numero de Columna Donde Aparecera la Calle
        Cmd.Parameters.Append(Cmd.CreateParameter("RenColonia", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenColonia)))) 'Numero de Renglon Donde Aparecera la Colonia
        Cmd.Parameters.Append(Cmd.CreateParameter("ColColonia", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColColonia)))) 'Numero de Columna Donde Aparecera la Colonia
        Cmd.Parameters.Append(Cmd.CreateParameter("RenCiudad", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenCiudad)))) 'Numero de Renglon Donde Aparecera la Ciudad
        Cmd.Parameters.Append(Cmd.CreateParameter("ColCiudad", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColCiudad)))) 'Numero de Columna Donde Aparecera la Ciudad
        Cmd.Parameters.Append(Cmd.CreateParameter("RenEstado", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenEstado)))) 'Numero de Renglon Donde Aparecera el Estado
        Cmd.Parameters.Append(Cmd.CreateParameter("ColEstado", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColEstado)))) 'Numero de Columna Donde Aparecera el Estado
        Cmd.Parameters.Append(Cmd.CreateParameter("RenCP", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenCP)))) 'Numero de Renglon Donde Aparecera el Codigo Postal
        Cmd.Parameters.Append(Cmd.CreateParameter("ColCP", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColCP)))) 'Numero de Columna Donde Aparecera el Codigo Postal
        Cmd.Parameters.Append(Cmd.CreateParameter("RenTelefono", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenTelefono)))) 'Numero de Renglon Donde Aparecera el Telefono
        Cmd.Parameters.Append(Cmd.CreateParameter("ColTelefono", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColTelefono)))) 'Numero de Columna Donde Aparecera el Telefono
        Cmd.Parameters.Append(Cmd.CreateParameter("RenSubTotal", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenSubTotal)))) 'Numero de Renglon Donde Aparecera el SubTotal
        Cmd.Parameters.Append(Cmd.CreateParameter("ColSubTotal", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColSubTotal)))) 'Numero de Columna Donde Aparecera el SubTotal
        Cmd.Parameters.Append(Cmd.CreateParameter("RenIva", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenIva)))) 'Numero de Renglon Donde Aparecera el Iva
        Cmd.Parameters.Append(Cmd.CreateParameter("ColIva", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColIva)))) 'Numero de Columna Donde Aparecera el Iva
        Cmd.Parameters.Append(Cmd.CreateParameter("RenTotal", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenTotal)))) 'Numero de Renglon Donde Aparecera el Total de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("ColTotal", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColTotal)))) 'Numero de Columna Donde Aparecera el Rotal de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("RenDescto", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenDescto)))) 'Numero de Renglon Donde Aparecera el Descuento
        Cmd.Parameters.Append(Cmd.CreateParameter("ColDescto", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColDescto)))) 'Numero de Columna Donde Aparecera el Descuento
        Cmd.Parameters.Append(Cmd.CreateParameter("RenImpLetra", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenImpLetra)))) 'Numero de Renglon Donde Aparecera el Importe Con Letra
        Cmd.Parameters.Append(Cmd.CreateParameter("ColImpLetra", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColImpLetra)))) 'Numero de Columna Donde Aparecera el Importe con Letra
        Cmd.Parameters.Append(Cmd.CreateParameter("RenLugarExped", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenLugarExped)))) 'Numero de Renglon Donde Aparecera el Lugar de Expedición
        Cmd.Parameters.Append(Cmd.CreateParameter("ColLugarExped", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColLugarExped)))) 'Numero de Renglon Donde Aparecera el Lugar de Expedición
        Cmd.Parameters.Append(Cmd.CreateParameter("CantRenXDet", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(CantRenXDet)))) 'Numero de Renglones que Apareceran en el Detalle
        Cmd.Parameters.Append(Cmd.CreateParameter("RenPrimerPartida", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenPrimerPartida)))) 'Numero de Renglon Donde Aparecera la Primer Partida
        Cmd.Parameters.Append(Cmd.CreateParameter("ColCodigo", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColCodigo)))) 'Columna Donde Aparecera el Codigo
        Cmd.Parameters.Append(Cmd.CreateParameter("ColCantidad", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColCantidad)))) 'Columna Donde Aparecera la Cantidad
        Cmd.Parameters.Append(Cmd.CreateParameter("ColDescripcion", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColDescripcion)))) 'Columna Donde Aparecera la Descripción
        Cmd.Parameters.Append(Cmd.CreateParameter("ColDesctoDetalle", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColDesctoDetalle)))) 'Columna Donde Aparecera el Descuento del Detalle
        Cmd.Parameters.Append(Cmd.CreateParameter("ColPromocion", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColPromocion)))) 'Columna Donde Aparecera la Promoción
        Cmd.Parameters.Append(Cmd.CreateParameter("ColPrecioVenta", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColPrecioVenta)))) 'Columna Donde Aparecera el Precio de Venta
        Cmd.Parameters.Append(Cmd.CreateParameter("ColImporte", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColImporte)))) 'Columna Donde Aparecera el Importe
        Cmd.Parameters.Append(Cmd.CreateParameter("ColIVAPartida", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColIVAPartida)))) 'Columna Donde Aparecera el Iva de la Partida
        Cmd.Parameters.Append(Cmd.CreateParameter("Ticket", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Ticket))) 'Bandera Para Saber si Sera un Ticket o Una Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("ColLeyenda", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(ColLeyenda)))) 'Columna Donde Aparecera la Leyenda
        Cmd.Parameters.Append(Cmd.CreateParameter("RenLeyenda", ADODB.DataTypeEnum.adSmallInt, ADODB.ParameterDirectionEnum.adParamInput, 2, CShort(Val(RenLeyenda)))) 'Renglon Donde Aparecera la Leyenda
        Cmd.Parameters.Append(Cmd.CreateParameter("Leyenda", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 200, Trim(Leyenda))) 'Leyenda que Llevara la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("TamLetra", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(TamLetra)))) 'Tamaño de Letra
        Cmd.Parameters.Append(Cmd.CreateParameter("LongCliente", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(LongCliente)))) 'Longitud del Campo Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("LongDireccion", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(LongDireccion)))) 'Longitud del Campo Dirección
        Cmd.Parameters.Append(Cmd.CreateParameter("LongColonia", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(LongColonia)))) 'Longitud del Campo Colonia
        Cmd.Parameters.Append(Cmd.CreateParameter("LongCiudad", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(LongCiudad)))) 'Longitud de la Ciudad
        Cmd.Parameters.Append(Cmd.CreateParameter("LongEstado", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(LongEstado)))) 'Longitud del Estado
        Cmd.Parameters.Append(Cmd.CreateParameter("LongProducto", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(LongProducto)))) 'Longitud de la Descripción del Producto
        Cmd.Parameters.Append(Cmd.CreateParameter("LongLeyenda", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, 1, CByte(Val(LongLeyenda)))) 'Longitud de la Leyenda
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'Procedimiento Almacenado para Guardar Facturas en el Corporativo
    'JUAN CARLOS OSUNA CORRALES 05 DE JUNIO DE 2003
    Public Sub PR_IME_Facturas(ByRef FolioFactura As String, ByRef NumPartida As String, ByRef CodSucursal As String, ByRef CodCaja As String, ByRef FechaFactura As String, ByRef TipoFactura As String, ByRef Condicion As String, ByRef CodCliente As String, ByRef Nombre As String, ByRef Rfc As String, ByRef Moneda As String, ByRef TipoCambio As String, ByRef SubTotal As String, ByRef Descuento As String, ByRef Iva As String, ByRef Total As String, ByRef Redondeo As String, ByRef PorcIva As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef Cantidad As String, ByRef DescEspecial As String, ByRef Precio As String, ByRef importe As String, ByRef PorcIvaP As String, ByRef FacturaAdicional As String, ByRef Origen As String, ByRef DesgloseIva As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_Facturas" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando en este Caso sera un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioFactura", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioFactura))) 'Folio de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Val(NumPartida)))) 'Numero de Partida
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Val(CodSucursal)))) 'Codigo de la Sucursal
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCaja", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Val(CodCaja)))) 'Codigo de la Caja
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFactura", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, CDate(Format(FechaFactura, "mm/dd/yyyy")))) 'Fecha de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoFactura", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoFactura))) 'Tipo de Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("Condicion", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2, Trim(Condicion))) 'Condicion Credito o Contado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCliente", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Val(CodCliente)))) 'Codigo del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Nombre", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(Nombre))) 'Nombre del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Rfc", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(Rfc))) 'Rfc del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda))) 'Moneda con que se Pago
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(TipoCambio)))) 'Tipo de Cambio
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotal", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(SubTotal)))) 'SubTotal de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(Descuento)))) 'Descuento de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("Iva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(Iva)))) 'Iva de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("Total", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(Total)))) 'Total de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("Redondeo", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(Redondeo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcIva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(PorcIva)))) 'Porcentaje de Iva
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus))) 'Estatus de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, CDate(Format(FechaCancel, "mm/dd/yyyy")))) 'Fecha de Cancelación
        Cmd.Parameters.Append(Cmd.CreateParameter("Cantidad", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Val(Cantidad)))) 'Cantidad de Articulos
        Cmd.Parameters.Append(Cmd.CreateParameter("DescEspecial", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 100, Trim(DescEspecial))) 'Cantidad Especial
        Cmd.Parameters.Append(Cmd.CreateParameter("Precio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(Precio)))) 'Precio
        Cmd.Parameters.Append(Cmd.CreateParameter("Importe", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Val(importe)))) 'Importe de la Factura
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcIvaP", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Val(PorcIvaP)))) 'Porcentaje de Iva por Partida
        Cmd.Parameters.Append(Cmd.CreateParameter("FacturaAdicional", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FacturaAdicional))) 'Factura Adicional
        Cmd.Parameters.Append(Cmd.CreateParameter("Origen", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Origen))) 'Para Saber quien hace la factura si el corporativo o el punto de venta
        Cmd.Parameters.Append(Cmd.CreateParameter("DesgloseIva", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(DesgloseIva))) 'Para Saber si la factura va llevar desglose de iva
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    ''Procedimiento Almacenado para Guardar Facturas en el Punto de VENTA
    ''JUAN CARLOS OSUNA CORRALES 05 DE JUNIO DE 2003
    'Public Sub PR_IMEFacturas(FolioFactura As String, NumPartida As String, FechaFactura As String, TipoFactura As String, _
    ''           Condicion As String, CodCliente As String, Nombre As String, Rfc As String, Moneda As String, _
    ''           TipoCambio As String, SubTotal As String, Descuento As String, Iva As String, Total As String, _
    ''           Redondeo As String, PorcIva As String, Estatus As String, FechaCancel As String, Cantidad As String, _
    ''           DescEspecial As String, Precio As String, Importe As String, PorcIvaP As String, Func As String, NumOp As String)
    '    BorraCmdPVenta
    '    CmdPVenta.CommandText = "UP_IME_Facturas" 'Nombre del Procedimiento Almacenado
    '    CmdPVenta.CommandType = adCmdStoredProc 'Tipo de Comando en este Caso sera un Procedimiento Almacenado
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("FolioFactura", adVarChar, adParamInput, 17, Trim(FolioFactura)) 'Folio de la Factura
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("NumPartida", adInteger, adParamInput, 4, CLng(Val(NumPartida))) 'Numero de Partida
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("FechaFactura", adDate, adParamInput, 8, CDate(Format(FechaFactura, "mm/dd/yyyy"))) 'Fecha de la Factura
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("TipoFactura", adChar, adParamInput, 1, Trim(TipoFactura)) 'Tipo de Factura
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Condicion", adChar, adParamInput, 2, Trim(Condicion)) 'Condicion Credito o Contado
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("CodCliente", adInteger, adParamInput, 4, CLng(Val(CodCliente))) 'Codigo del Cliente
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Nombre", adVarChar, adParamInput, 30, Trim(Nombre)) 'Nombre del Cliente
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Rfc", adVarChar, adParamInput, 15, Trim(Rfc)) 'Rfc del Cliente
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Moneda", adChar, adParamInput, 1, Trim(Moneda)) 'Moneda con que se Pago
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("TipoCambio", adCurrency, adParamInput, 4, CCur(Val(TipoCambio))) 'Tipo de Cambio
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("SubTotal", adCurrency, adParamInput, 8, CCur(Val(SubTotal))) 'SubTotal de la Factura
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Descuento", adCurrency, adParamInput, 8, CCur(Val(Descuento))) 'Descuento de la Factura
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Iva", adCurrency, adParamInput, 8, CCur(Val(Iva))) 'Iva de la Factura
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Total", adCurrency, adParamInput, 8, CCur(Val(Total))) 'Total de la Factura
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Redondeo", adCurrency, adParamInput, 4, CCur(Val(Redondeo)))
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("PorcIva", adCurrency, adParamInput, 4, CCur(Val(PorcIva))) 'Porcentaje de Iva
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Estatus", adChar, adParamInput, 1, Trim(Estatus)) 'Estatus de la Factura
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("FechaCancel", adDate, adParamInput, 8, CDate(Format(FechaCancel, "mm/dd/yyyy"))) 'Fecha de Cancelación
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Cantidad", adInteger, adParamInput, 4, CLng(Val(Cantidad))) 'Cantidad de Articulos
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("DescEspecial", adVarChar, adParamInput, 100, Trim(DescEspecial)) 'Cantidad Especial
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Precio", adCurrency, adParamInput, 8, CCur(Val(Precio))) 'Precio
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Importe", adCurrency, adParamInput, 8, CCur(Val(Importe))) 'Importe de la Factura
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("PorcIvaP", adCurrency, adParamInput, 4, CCur(Val(PorcIvaP))) 'Porcentaje de Iva por Partida
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("Func", adChar, adParamInput, 1, Trim(Func)) 'Tipo de Transacción
    '    CmdPVenta.Parameters.Append CmdPVenta.CreateParameter("NumOp", adInteger, adParamInput, 1, CInt(NumOp)) 'Numero de Opción de Transacción
    'End Sub

    'JUAN CARLOS OSUNA CORRALES 14 DE JUNIO DE 2003
    Public Sub PR_IMEFoliosCorporativo(ByRef CodFolio As String, ByRef DescFolio As String, ByRef Prefijo As String, ByRef Consecutivo As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_FoliosCorporativo" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando en Este Caso sera un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFolio", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Val(CodFolio)))) 'Codigo del Folio
        Cmd.Parameters.Append(Cmd.CreateParameter("DescFolio", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(DescFolio))) 'Descripción del Folio
        Cmd.Parameters.Append(Cmd.CreateParameter("Prefijo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Prefijo))) 'Prefijo del Folio
        Cmd.Parameters.Append(Cmd.CreateParameter("Consecutivo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Val(Consecutivo)))) 'Consecutivo del Folio
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'JUAN CARLOS OSUNA CORRALES 16 DE JULIO DE 2003
    Public Sub PR_IMEMovimientosBancarios(ByRef FolioMovto As String, ByRef FechaMovto As String, ByRef Movimiento As String, ByRef TipoMovto As String, ByRef Naturaleza As String, ByRef Moneda As String, ByRef TipoCambio As String, ByRef FormaPago As String, ByRef TipoPago As String, ByRef CodBanco As String, ByRef CtaBancaria As String, ByRef Beneficiario As String, ByRef Concepto As String, ByRef PagoProgramado As String, ByRef FolioProgramacion As String, ByRef PartidaPP As String, ByRef FechaDocto As String, ByRef NoDocto As String, ByRef importe As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef FolioRetiro As String, ByRef Conciliado As String, ByRef FechaConciliacion As String, ByRef Modulo As String, ByRef Referencia As String, ByRef FolioElectronico As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_MovimientosBancarios"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioMovto", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 13, Trim(FolioMovto))) 'Folio del Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaMovto", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaMovto), C_FORMATFECHAGUARDAR))) 'Fecha del Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("Movimiento", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2, Trim(Movimiento))) 'Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoMovto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoMovto))) 'Tipo de Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("Naturaleza", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Naturaleza))) 'Naturaleza del Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda))) 'Moneda
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(TipoCambio))) 'Tipo de Cambio
        Cmd.Parameters.Append(Cmd.CreateParameter("FormaPago", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(FormaPago))) 'Forma de Pago
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoPago", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoPago))) 'Tipo de Pago
        Cmd.Parameters.Append(Cmd.CreateParameter("CodBanco", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput,  , CInt(CodBanco))) 'Codigo del Banco
        Cmd.Parameters.Append(Cmd.CreateParameter("CtaBancaria", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 16, Trim(CtaBancaria))) 'Cuenta Bancaria
        Cmd.Parameters.Append(Cmd.CreateParameter("Beneficiario", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Beneficiario))) 'Beneficiario
        Cmd.Parameters.Append(Cmd.CreateParameter("Concepto", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Concepto))) 'Concepto
        Cmd.Parameters.Append(Cmd.CreateParameter("PagoProgramado", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(PagoProgramado))) 'Pago Programado
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioProgramacion", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioProgramacion))) 'Folio de Programacion
        Cmd.Parameters.Append(Cmd.CreateParameter("PatidaPP", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(PartidaPP))) 'Numero de Partida del Folio de Programación
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaDocto", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaDocto), C_FORMATFECHAGUARDAR))) 'Fecha del Documento
        Cmd.Parameters.Append(Cmd.CreateParameter("NoDocto", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(NoDocto))) 'Numero de Documento
        Cmd.Parameters.Append(Cmd.CreateParameter("Importe", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(importe))) 'Importe del Documento
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus))) 'Estatus del Documento
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR))) 'Fecha de Cancelacion
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioRetiro", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioRetiro))) 'Folio de Retiro
        Cmd.Parameters.Append(Cmd.CreateParameter("Conciliado", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Conciliado))) 'Conciliado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaConciliacion", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaConciliacion), C_FORMATFECHAGUARDAR))) 'Fecha de Conciliacion
        Cmd.Parameters.Append(Cmd.CreateParameter("Modulo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Modulo))) 'Modulo al Que Pertenece
        Cmd.Parameters.Append(Cmd.CreateParameter("Referencia", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 13, Trim(Referencia))) 'Referencia Bancaria
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioElectronico", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 13, Trim(FolioElectronico))) 'FolioElectronico
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'JUAN CARLOS OSUNA CORRALES 16 DE JULIO DE 2003
    Public Sub PR_IMEMovimientosOrigenAplic(ByRef FolioMovto As String, ByRef CodOrigenAplicR As String, ByRef CodRubro As String, ByRef CodOrigenAplicAnte As String, ByRef CodRubroAnte As String, ByRef Aplicacion As String, ByRef importe As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_MovimientosOrigenAplic"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioMovto", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 13, Trim(FolioMovto))) 'Foliodel Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("CodOrigenAplicR", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput,  , CInt(CodOrigenAplicR))) 'Codigo de Origen y Aplicación de Recursos
        Cmd.Parameters.Append(Cmd.CreateParameter("CodRubro", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput,  , CInt(CodRubro))) 'Codigo del Rubro
        Cmd.Parameters.Append(Cmd.CreateParameter("CodOrigenAplicAnte", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput,  , CInt(CodOrigenAplicAnte))) 'Codigo del Rubro
        Cmd.Parameters.Append(Cmd.CreateParameter("CodRubroAnte", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput,  , CInt(CodRubroAnte))) 'Codigo del Rubro
        Cmd.Parameters.Append(Cmd.CreateParameter("Aplicacion", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Aplicacion))) 'Aplicación del Origen y Aplicación
        Cmd.Parameters.Append(Cmd.CreateParameter("Importe", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(importe))) 'Importe del Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus))) 'Estatus del Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR))) 'Fecha de Cancelación
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'JUAN CARLOS OSUNA CORRALES 16 DE JULIO DE 2003
    Public Sub PR_IMEFoliosMovtosBancos(ByRef Prefijo As String, ByRef Consecutivo As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_FoliosMovtosBancos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Prefijo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Prefijo))) 'Prefijo del Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("Consecutivo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput,  , CInt(Consecutivo))) 'Consecutivo
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'JUAN CARLOS OSUNA CORRALES 01 DE AGOSTO DE 2003
    Public Sub PR_IMEEjercicioPeriodo(ByRef Ejercicio As String, ByRef Periodo As String, ByRef Prefijo As String, ByRef Consecutivo As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_EjercicioPeriodo"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Ejercicio", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Ejercicio))) 'Ejercicio
        Cmd.Parameters.Append(Cmd.CreateParameter("Periodo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2, Trim(Periodo))) 'Periodo del Ejercicio
        Cmd.Parameters.Append(Cmd.CreateParameter("Prefijo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Prefijo))) 'Prefijo del Movimiento
        Cmd.Parameters.Append(Cmd.CreateParameter("Consecutivo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Consecutivo))) 'Consecutivo del Prefijo
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'JUAN CARLOS OSUNA CORRALES 04 DE AGOSTO DE 2003
    Public Sub PR_IME_Anticipos(ByRef FolioAnticipo As String, ByRef FechaAnticipo As String, ByRef FolioEgreso As String, ByRef CodProvAcreed As String, ByRef Concepto As String, ByRef Moneda As String, ByRef SubTotal As String, ByRef Descuento As String, ByRef Iva As String, ByRef Total As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef TipoCambio As String, ByRef TipoCambioAplic As String, ByRef FechaAplic As String, ByRef FolioPagoBancos As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_ANTICIPOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioAnticipo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(FolioAnticipo))) 'Folio del Anticipo
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaAnticipo", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaAnticipo), C_FORMATFECHAGUARDAR))) 'Fecha del Anticipo
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioEgreso", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 13, Trim(FolioEgreso))) 'Folio de Egreso
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(CodProvAcreed))) 'Codigo del Proveedor Acreedor
        Cmd.Parameters.Append(Cmd.CreateParameter("Concepto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(Concepto))) 'Concepto del Anticipo
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda))) 'Moneda del Anticipo
        Cmd.Parameters.Append(Cmd.CreateParameter("SubTotal", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(SubTotal)))) 'SubTotal del Anticipo
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(Descuento)))) 'Descuento del Anticipo
        Cmd.Parameters.Append(Cmd.CreateParameter("Iva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(Iva)))) 'Iva del Anticipo
        Cmd.Parameters.Append(Cmd.CreateParameter("Total", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(Total)))) 'Total del Anticipo
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus))) 'Estatus del Anticipo
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR))) 'Fecha de Cancelación
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(TipoCambio)))) 'Tipo de Cambio
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambioAplic", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(TipoCambioAplic)))) 'Tipo de Cambio Aplicado
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaAplic", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(FechaAplic), C_FORMATFECHAGUARDAR))) 'Fecha de Aplicación
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioPagoBancos", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 13, Trim(FolioPagoBancos))) 'Folio del egreso generado en bancos
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'JUAN CARLOS OSUNA CORRALES 05 DE AGOSTO DE 2003
    Public Sub PR_IME_ConfiguracionBancos(ByRef UltCierreBancos As String, ByRef UltCierreConciliacion As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_ConfiguracionBancos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("UltCierreBancos", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(UltCierreBancos), C_FORMATFECHAGUARDAR))) 'Fecha del Ultimo Cierre
        Cmd.Parameters.Append(Cmd.CreateParameter("UltCierreConciliacion", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , Format(CDate(UltCierreConciliacion), C_FORMATFECHAGUARDAR))) 'Fecha de la Ultima Conciliación
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'JUAN CARLOS OSUNA 16 DE MAYO DE 2003
    Public Sub PR_IMECatCajas(ByRef CodCaja As String, ByRef CodSucursal As String, ByRef DescCaja As String, ByRef FechaUltimoCorte As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatCajas" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que Regresa, en este Caso sera el Codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCaja", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodCaja)))) 'Numero de la Caja
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodSucursal)))) 'Numero de la Sucursal
        Cmd.Parameters.Append(Cmd.CreateParameter("DescCaja", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(DescCaja))) 'Descripción de la Caja
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaUltimoCorte", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, CDate(FechaUltimoCorte)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de Transacción
    End Sub

    'JUAN CARLOS OSUNA 16 DE MAYO DE 2003
    Public Sub PR_IMECatFolios(ByRef CodFolio As String, ByRef CodAlmacen As String, ByRef DescFolio As String, ByRef Prefijo As String, ByRef Consecutivo As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatFolios" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("ID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Valor que Regresa en Este Caso Sera el Codigo Identity
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFolio", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodFolio)))) 'Codigo del Folio
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodAlmacen)))) 'Codigo del Almacen
        Cmd.Parameters.Append(Cmd.CreateParameter("DescFolio", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(DescFolio))) 'Descripción del Folio
        Cmd.Parameters.Append(Cmd.CreateParameter("Prefijo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 2, Trim(Prefijo))) 'Prefijo del Folio
        Cmd.Parameters.Append(Cmd.CreateParameter("Consecutivo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Val(Consecutivo)))) 'Consecutivo del Folio
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    Public Sub PR_EstadodeResultados(ByRef FechaInicial As String, ByRef FechaFinal As String, ByRef Moneda As String, ByRef Impuesto As String, ByRef sql1 As String, ByRef sql2 As String, ByRef Tabla As String)
        ModEstandar.BorraCmd()
        Cmd.CommandText = "UP_EstadodeResultados"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaInicial", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, FechaInicial))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFinal", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, FechaFinal))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Impuesto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Impuesto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sql1", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Trim(sql1)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sql2", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Trim(sql2)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Tabla", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(Tabla)))
    End Sub

    Public Sub PR_EstadodeResultadosAnual(ByRef Periodo As String, ByRef Moneda As String, ByRef Impuesto As String, ByRef sql1 As String, ByRef sql2 As String, ByRef Tabla As String)
        ModEstandar.BorraCmd()
        Cmd.CommandText = "UP_EstadodeResultadosAnual"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Periodo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(Periodo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Impuesto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Impuesto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sql1", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Trim(sql1)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sql2", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Trim(sql2)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Tabla", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(Tabla)))
    End Sub

    'Rosaura Torres López
    Public Sub PR_IE_MovtosAlmacenCab(ByRef FolioAlmacen As String, ByRef FechaAlmacen As String, ByRef CodAlmacen As String, ByRef CodAlmacenOrigen As String, ByRef FolioOrdenCompra As String, ByRef CodProvAcreed As String, ByRef Factura As String, ByRef CodALmacenREf As String, ByRef CodMovtoAlm As String, ByRef EntradaSalida As String, ByRef Envia As String, ByRef Entrega As String, ByRef Recibe As String, ByRef Concepto As String, ByRef Estatus As String, ByRef NickUsuario As String, ByRef FechaCancel As String, ByRef NickCancel As String, ByRef ReferenciaDeOrigen As String, ByRef FechaReferencia As String, ByRef CodCliente As String, ByRef FolioApartado As String, ByRef FechaRegresoPrestamo As String, ByRef TipoCambio As String, ByRef FolioVenta As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IE_MovtosAlmacenCab" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado

        Cmd.Parameters.Append(Cmd.CreateParameter("FolioAlmacen", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioAlmacen)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaAlmacen", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 4, Format(CDate(FechaAlmacen), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodAlmacen))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacenOrigen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodAlmacenOrigen))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioOrdenCompra", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(FolioOrdenCompra)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodProvAcreed))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Factura", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(Factura)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacenref", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodALmacenREf))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMovtoAlm", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodMovtoAlm))))
        Cmd.Parameters.Append(Cmd.CreateParameter("entradaSalida", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(EntradaSalida)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Envia", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Envia)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Entrega", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Entrega)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Recibe", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(Recibe)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Concepto", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 150, Trim(Concepto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NickUsuario", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(NickUsuario)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 4, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NickCancel", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(NickCancel)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ReferenciaDeOrigen", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(ReferenciaDeOrigen)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaReferencia", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 4, Format(CDate(FechaReferencia), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("codcliente", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodCliente))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioApartado", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioApartado)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaRegresoPrestamo", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 4, Format(CDate(FechaRegresoPrestamo), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput,  , CDec(TipoCambio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FolioVenta", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioVenta)))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de Transacción
    End Sub

    'Rosaura Torres López - 09/07/03
    Public Sub PR_IE_MovtosAlmacenDet(ByRef FolioAlmacen As String, ByRef NumPartida As String, ByRef FechaAlmacen As String, ByRef CodArticulo As String, ByRef CodAlmacenOrigen As String, ByRef Cantidad As String, ByRef CostoUnitario As String, ByRef PrecioVenta As String, ByRef Descuento As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef Confirmacion As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_IE_MovtosAlmacenDet"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En este Caso sera Un Procedimiento Almacenado

        Cmd.Parameters.Append(Cmd.CreateParameter("FolioAlmacen", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 17, Trim(FolioAlmacen)))
        Cmd.Parameters.Append(Cmd.CreateParameter("NumPartida", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(NumPartida)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaAlmacen", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaAlmacen), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodArticulo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(CodArticulo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacenOrigen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(CodAlmacenOrigen)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Cantidad", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Cantidad)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Costounitario", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(CostoUnitario)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PrecioVenta", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(PrecioVenta)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descuento", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Descuento)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Confirmacion", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 4, CBool(Confirmacion)))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opcion de Transacción
    End Sub

    'Rosaura Torres López
    Public Sub PR_IE_Inventario(ByRef CodAlmacen As String, ByRef AlmacenPropio As String, ByRef CodArticulo As String, ByRef CodAlmacenOrigen As String, ByRef ExistenciaInicial As String, ByRef CostoInicialMN As String, ByRef costoInicialDLL As String, ByRef UltimoCostoMN As String, ByRef UltimoCostoDLL As String, ByRef Entradas As String, ByRef Salidas As String, ByRef Apartados As String, ByRef CodMovtoAlm As String, ByRef FechaMovto As String, ByRef Func As String, ByRef NumOp As String)

        BorraCmd()
        Cmd.CommandText = "UP_IE_iNVENTARIO" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado

        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodAlmacen))))
        Cmd.Parameters.Append(Cmd.CreateParameter("AlmacenPropio", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 4, CBool(AlmacenPropio)))
        Cmd.Parameters.Append(Cmd.CreateParameter("codArticulo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodArticulo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacenOrigen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodAlmacenOrigen))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ExistenciaInicial", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(ExistenciaInicial))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoInicialMN", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(CostoInicialMN))))
        Cmd.Parameters.Append(Cmd.CreateParameter("costoInicialDLL", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(costoInicialDLL))))
        Cmd.Parameters.Append(Cmd.CreateParameter("UltimoCostoMN", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(UltimoCostoMN))))
        Cmd.Parameters.Append(Cmd.CreateParameter("UltimoCostoDLL", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, CDec(Numerico(UltimoCostoDLL))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Entradas", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(Entradas))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Salidas", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(Salidas))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Apartados", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(Apartados))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMovtoAlm", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodMovtoAlm))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaMovto", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaMovto), C_FORMATFECHAGUARDAR)))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de Transacción
    End Sub

    'Rosaura Torres López
    Public Sub PR_I_FoliosAlmacen(ByRef CodAlmacen As String, ByRef ConsecutivoMovtoAlm As String, ByRef Func As String, ByRef NumOp As String)

        BorraCmd()
        Cmd.CommandText = "UP_I_FoliosAlmacen" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("Consecutivo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodAlmacen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(CodAlmacen))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConsecutivoMovtoAlm", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Val(ConsecutivoMovtoAlm))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de Transacción
    End Sub

    Public Sub PR_IMECatPlanesxBanco(ByRef CodBanco As String, ByRef CodPlan As String, ByRef DescPlan As String, ByRef PorcIntereses As String, ByRef PorcIva As String, ByRef Estatus As String, ByRef Func As String, ByRef NumOp As String)
        'Rosaura 24/09/03

        BorraCmd()
        Cmd.CommandText = "UP_IME_CatPlanesxBanco" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("CodBanco", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodBanco))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodPlan", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodPlan))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Descplan", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, Trim(DescPlan)))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcInetereses", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(ModEstandar.Numerico(PorcIntereses))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcIva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(ModEstandar.Numerico(PorcIva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    Public Sub PR_IMEPromocionesVentas(ByRef CodGrupo As String, ByRef CodFamilia As String, ByRef COdLinea As String, ByRef CodSubLinea As String, ByRef CodMArca As String, ByRef CodModelo As String, ByRef CodArticulo As String, ByRef importe As String, ByRef Porcentaje As String, ByRef FechaInicio As String, ByRef FechaFin As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef TipoProm As String, ByRef Proveedor As String, ByRef Renglon As String, ByRef Func As String, ByRef NumOp As String)
        '------------------------------------------------------------------------------------
        'Rosaura  19/07/03
        '------------------------------------------------------------------------------------
        '''MODIFIC.-  21ABR2006 - SECCION DESCTO ART X PROV
        '------------------------------------------------------------------------------------
        BorraCmd()
        Cmd.CommandText = "UP_IME_PromocionesVentas" 'Nombre del Procedimiento almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo del comando que en este caso sera del procedimiento almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodGrupo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodGrupo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodFamilia", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodFamilia)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(COdLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSubLinea", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodSubLinea)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodMarca", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodMArca)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodModelo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodModelo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodArticulo", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(CodArticulo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Importe", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 8, ModEstandar.Numerico(importe)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Porcentaje", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(Porcentaje)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaInicio", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaInicio), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaFin", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaFin), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaCancel), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoProm", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoProm)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodProvAcreed", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(Proveedor)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, ModEstandar.Numerico(Renglon)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Número de Opción de Transacción
    End Sub

    ''''Rosaura Torres López 01/Julio/2003
    '''Public Sub PR_IME_Reparaciones(FolioReparacion As String, _
    ''''    FechaReparacion As String, _
    ''''    TipoMovto As String, _
    ''''    CodSucursal As String, _
    ''''    CodCaja As String, _
    ''''    CodVendedor As String, _
    ''''    CodCliente As String, _
    ''''    Nombre As String, _
    ''''    Rfc As String, _
    ''''    Telefono As String, _
    ''''    MotivoReparacion As String, _
    ''''    ObservacionesTaller As String, _
    ''''    CodTaller As String, _
    ''''    TipoReparacion As String, _
    ''''    Moneda As String, _
    ''''    TipoCambio As String, _
    ''''    CostoReparacion As String, _
    ''''    SubTotal As String, _
    ''''    Iva As String, _
    ''''    ImporteVta As String, _
    ''''    Anticipo As String, Estatus As String, _
    ''''    FechaCancel As String, FechaEntregaTaller As String, _
    ''''    FechaRegreso As String, FechaConfirmacion As String, _
    ''''    FechaEntregaCliente As String, Credito As String, Reparado As String, _
    ''''    Func As String, NumOp As String)
    '''
    '''    BorraCmd
    '''    Cmd.CommandText = "UP_IME_Reparaciones" 'Nombre del Procedimiento Almacenado
    '''    Cmd.CommandType = adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado
    '''    'cmd.Parameters.Append cmd.CreateParameter("ID", adInteger, adParamReturnValue) 'Valor Que Regresa, En Este Caso Sera el Codigo Identity
    '''
    '''    Cmd.Parameters.Append Cmd.CreateParameter("FolioReparacion", adVarChar, adParamInput, 19, Trim(FolioReparacion))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("FechaReparacion", adDate, adParamInput, 8, Format(CDate(FechaReparacion), C_FORMATFECHAGUARDAR))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("TipoMovto", adChar, adParamInput, 1, Trim(TipoMovto))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("CodSucursal", adInteger, adParamInput, 4, CInt(Numerico(CodSucursal)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("CodCaja", adInteger, adParamInput, 4, CInt(Numerico(CodCaja)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("CodVendedor", adInteger, adParamInput, 4, CInt(Numerico(CodVendedor)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("CodCliente", adInteger, adParamInput, 4, CInt(Numerico(CodCliente)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("Nombre", adVarChar, adParamInput, 40, Trim(Nombre))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("Rfc", adVarChar, adParamInput, 15, Trim(Rfc))  'Rfc del Cliente
    '''    Cmd.Parameters.Append Cmd.CreateParameter("Telefono", adVarChar, adParamInput, 48, Trim(Telefono))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("MotivoReparacion", adVarChar, adParamInput, 255, Trim(MotivoReparacion))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("ObservacionesTaller", adVarChar, adParamInput, 255, Trim(ObservacionesTaller))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("CodTaller", adInteger, adParamInput, 4, CInt(Numerico(CodTaller)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("CodTipoReparacion", adInteger, adParamInput, 4, CInt(Numerico(TipoReparacion)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("Moneda", adChar, adParamInput, 1, Trim(Moneda))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("TipoCambio", adCurrency, adParamInput, 4, CCur(Numerico(TipoCambio)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("CostoReparacion", adCurrency, adParamInput, 4, CCur(Numerico(CostoReparacion)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("SubtotalVta", adCurrency, adParamInput, 4, CCur(Numerico(SubTotal)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("IvaVta", adCurrency, adParamInput, 4, CCur(Numerico(Iva)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("ImporteVta", adCurrency, adParamInput, 4, CCur(Numerico(ImporteVta)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("Anticipo", adCurrency, adParamInput, 4, CCur(Numerico(Anticipo)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("Estatus", adChar, adParamInput, 1, Trim(Estatus))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("FechaCancel", adDate, adParamInput, 4, CDate(Format(FechaCancel, C_FORMATFECHAGUARDAR)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("FechaEntregaTaller", adDate, adParamInput, 4, CDate(Format(FechaEntregaTaller, C_FORMATFECHAGUARDAR)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("FechaREgreso", adDate, adParamInput, 4, CDate(Format(FechaRegreso, C_FORMATFECHAGUARDAR)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("FechaConfirmacion", adDate, adParamInput, 4, CDate(Format(FechaConfirmacion, C_FORMATFECHAGUARDAR)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("FechaEntregaCliente", adDate, adParamInput, 4, CDate(Format(FechaEntregaCliente, C_FORMATFECHAGUARDAR)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("Credito", adBoolean, adParamInput, 1, CBool(Trim(Credito)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("REparado", adBoolean, adParamInput, 1, CBool(Trim(Reparado)))
    '''    Cmd.Parameters.Append Cmd.CreateParameter("Func", adChar, adParamInput, 1, Trim(Func)) 'Tipo de Transacción
    '''    Cmd.Parameters.Append Cmd.CreateParameter("NumOp", adInteger, adParamInput, 1, CInt(NumOp)) 'Numero de Opción de Transacción
    '''End Sub

    'Rosaura Torres López - 01/Julio/2003
    '
    Public Sub PR_IME_ReparacionesCorpo(ByRef FolioReparacion As String, ByRef FechaReparacion As String, ByRef TipoMovto As String, ByRef CodSucursal As String, ByRef CodCaja As String, ByRef CodVendedor As String, ByRef CodCliente As String, ByRef Nombre As String, ByRef Rfc As String, ByRef Telefono As String, ByRef MotivoReparacion As String, ByRef ObservacionesTaller As String, ByRef CodTaller As String, ByRef TipoReparacion As String, ByRef Moneda As String, ByRef TipoCambio As String, ByRef CostoReparacion As String, ByRef SubTotal As String, ByRef Iva As String, ByRef ImporteVta As String, ByRef Anticipo As String, ByRef Estatus As String, ByRef FechaCancel As String, ByRef FechaEntregaTaller As String, ByRef FechaRegreso As String, ByRef FechaConfirmacion As String, ByRef FechaEntregaCliente As String, ByRef Credito As String, ByRef Reparado As String, ByRef PorcIva As String, ByRef FechaCorpoEnvio As String, ByRef FechaCorpoRegresa As String, ByRef Bitacora As String, ByRef Func As String, ByRef NumOp As String)

        BorraCmd()
        Cmd.CommandText = "UP_IME_Reparaciones" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado

        Cmd.Parameters.Append(Cmd.CreateParameter("FolioReparacion", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 19, Trim(FolioReparacion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaReparacion", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, 8, Format(CDate(FechaReparacion), C_FORMATFECHAGUARDAR)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoMovto", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(TipoMovto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodSucursal))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCaja", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodCaja))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodVendedor", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodVendedor))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCliente", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodCliente))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Nombre", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, Trim(Nombre)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Rfc", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(Rfc))) 'Rfc del Cliente
        Cmd.Parameters.Append(Cmd.CreateParameter("Telefono", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 48, Trim(Telefono)))
        Cmd.Parameters.Append(Cmd.CreateParameter("MotivoReparacion", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4000, Trim(MotivoReparacion)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ObservacionesTaller", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4000, Trim(ObservacionesTaller)))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodTaller", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodTaller))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CodTipoReparacion", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(TipoReparacion))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("TipoCambio", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(TipoCambio))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CostoReparacion", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(CostoReparacion))))
        Cmd.Parameters.Append(Cmd.CreateParameter("SubtotalVta", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(SubTotal))))
        Cmd.Parameters.Append(Cmd.CreateParameter("IvaVta", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(Iva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ImporteVta", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(ImporteVta))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Anticipo", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(Anticipo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Estatus)))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCancel", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaCancel))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaEntregaTaller", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaEntregaTaller))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaREgreso", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaRegreso))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaConfirmacion", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaConfirmacion))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaEntregaCliente", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaEntregaCliente))
        Cmd.Parameters.Append(Cmd.CreateParameter("Credito", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Trim(Credito))))
        Cmd.Parameters.Append(Cmd.CreateParameter("REparado", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, 1, CBool(Trim(Reparado))))
        Cmd.Parameters.Append(Cmd.CreateParameter("PorcIva", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, 4, CDec(Numerico(PorcIva))))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCorpoEnvio", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaCorpoEnvio))
        Cmd.Parameters.Append(Cmd.CreateParameter("FechaCorpoRegreso", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 23, FechaCorpoRegresa))
        Cmd.Parameters.Append(Cmd.CreateParameter("Bitacora", ADODB.DataTypeEnum.adLongVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483647, Trim(Bitacora)))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de Transacción
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de Opción de Transacción
    End Sub

    'ROSAURA TORRES
    Public Sub PR_InicializaInformacion(ByRef Sucursal As String)
        BorraCmd()
        Cmd.CommandText = "UP_InicializaInformacion" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue)) 'Estatus del }Proceso
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Sucursal)))
    End Sub

    'Rosaura Torres lópez
    Public Sub PR_IMConfiguracionImpresora(ByRef CodSucursal As String, ByRef TicketPrinter As String, ByRef RutaImpresora As String)
        BorraCmd()
        Cmd.CommandText = "UP_IM_ConfiguracionImpresora" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodSucursal", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodSucursal)))) 'Numero de la Sucursal
        Cmd.Parameters.Append(Cmd.CreateParameter("TicketPrinter", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 100, Trim(TicketPrinter)))
        Cmd.Parameters.Append(Cmd.CreateParameter("RutaImpresora", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 100, Trim(RutaImpresora)))
    End Sub

    'Rosaura Torres
    Public Sub PR_InvHojadeControl(ByRef Consulta As String, ByRef CodAlmacen As String, ByRef CodAlmacenOrigen As String, ByRef CodArticulo As String, ByRef ExistenciaFisica As String, ByRef Ajuste As String, ByRef Func As String, ByRef NumOp As String)
        BorraCmd()
        Cmd.CommandText = "UP_InvHojadeControl" 'Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("Consulta", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Trim(Consulta)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ParamCodAlmacen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodAlmacen))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ParamCodAlmacenOrigen", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodAlmacenOrigen))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Param|", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CInt(Numerico(CodArticulo))))
        Cmd.Parameters.Append(Cmd.CreateParameter("ParamExistenciaFisica", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(ExistenciaFisica))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Paramajuste", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(Ajuste))))
        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
        Cmd.Parameters.Append(Cmd.CreateParameter("NumOp", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 1, CShort(NumOp))) 'Numero de opcion de Transacción
    End Sub

    'Antonella Vargas
    Public Sub PR_IME_CatCuentasNotifiaciones(ByRef CodCuentaC As String, ByRef CuentaCorreo As String, ByRef Estatus As String, ByRef Func As String)
        BorraCmd()
        Cmd.CommandText = "UP_IME_CatCuentasNotificaciones" '''Nombre del Procedimiento Almacenado
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc 'Tipo de Comando En Este Caso Sera Un Procedimiento Almacenado
        Cmd.Parameters.Append(Cmd.CreateParameter("CodCuentaC", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Numerico(CodCuentaC))))
        Cmd.Parameters.Append(Cmd.CreateParameter("CuentaCorreo", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 100, Trim(CuentaCorreo)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Estatus", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 100, Trim(Estatus)))

        Cmd.Parameters.Append(Cmd.CreateParameter("Func", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Func))) 'Tipo de transaccion
    End Sub

    Public Sub PR_IME_VentasComparativoMensual(ByRef Tabla As String, ByRef Mes As String, ByRef Año As String, ByRef Dias As String, ByRef ConImpuesto As String, ByRef Moneda As String, ByRef Sucursales As String)
        BorraCmd()
        Cmd.CommandText = "UP_VentasComparativoMensual"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Tabla", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(Tabla)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Mes", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Mes)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Año", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Año)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Dias", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Dias)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ConImpuesto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sucursales", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 2000, Trim(Sucursales)))
    End Sub

    Public Sub PR_IME_VentasComparativoAnual(ByRef Tabla As String, ByRef Año As String, ByRef ConImpuesto As String, ByRef Moneda As String, ByRef Sucursales As String)
        BorraCmd()
        Cmd.CommandText = "UP_VentasComparativoAnual"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Tabla", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(Tabla)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Año", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(Año)))
        Cmd.Parameters.Append(Cmd.CreateParameter("ConImpuesto", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, 4, CShort(ConImpuesto)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Moneda", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Moneda)))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sucursales", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 2000, Trim(Sucursales)))
    End Sub
End Module