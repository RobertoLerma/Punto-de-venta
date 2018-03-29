Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmBancosProcesoDiarioRegistrodeDepositos
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdReferencias As System.Windows.Forms.Button
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents txtFolioIngreso As System.Windows.Forms.TextBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents chkDepositoInterno As System.Windows.Forms.CheckBox
    Public WithEvents cmdOrigAplic_DepInt_Pes As System.Windows.Forms.Button
    Public WithEvents cmdOrigAplic_DepInt_Dol As System.Windows.Forms.Button
    Public WithEvents dtpFechaRetiro As System.Windows.Forms.DateTimePicker
    Public WithEvents txtFolioRetiro As System.Windows.Forms.TextBox
    Public WithEvents txtSucursal As System.Windows.Forms.TextBox
    Public WithEvents txtEnvia As System.Windows.Forms.TextBox
    Public WithEvents txtPesos As System.Windows.Forms.TextBox
    Public WithEvents txtDolares As System.Windows.Forms.TextBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents cmdDesglose As System.Windows.Forms.Button
    Public WithEvents cmdOrigenyAplicacion As System.Windows.Forms.Button
    Public WithEvents txtConcepto As System.Windows.Forms.TextBox
    Public WithEvents txtImporte As System.Windows.Forms.TextBox
    Public WithEvents dbcBanco As System.Windows.Forms.ComboBox
    Public WithEvents dbcCuentaBancaria As System.Windows.Forms.ComboBox
    Public WithEvents lblCancelada As System.Windows.Forms.Label
    Public WithEvents lblMoneda As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents lblCodBanco As System.Windows.Forms.Label
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnGuardar As Button
    Public bandera As Boolean = False
    Public strControlActual As String 'Nombre del control actual

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas 
        Dim I As Integer

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "TXTFOLIOINGRESO"
                strCaptionForm = "Consulta de Registro de Depositos"
                gStrSql = "SELECT FolioMovto AS FOLIO,Concepto AS CONCEPTO, FechaMovto AS FECHA,Importe AS IMPORTE FROM MovimientosBancarios " & "WHERE FolioMovto LIKE '" & txtFolioIngreso.Text & "%' AND Movimiento = '" & C_MOVDEPOSITO & "' AND TipoMovto = '" & C_TIPOMOVINGRESO & "' ORDER BY FechaMovto DESC ,FolioMovto DESC"
            Case "TXTFOLIORETIRO"
                strCaptionForm = "Consulta de Folios de Retiros"
                gStrSql = "SELECT FolioRetiro AS 'FOLIO DEL RETIRO', NickUsuario AS USUARIO, FechaRetiro AS 'FECHA DEL RETIRO' FROM " & "Retiros WHERE FolioRetiro LIKE '" & txtFolioRetiro.Text & "%' AND TipoRetiro = '" & C_RETIROCAJAGENERAL & "' AND Estatus = 'V' AND PasoBancos = 0 GROUP BY FolioRetiro,NickUsuario,FechaRetiro " & "ORDER BY FolioRetiro,FechaRetiro DESC"
            Case Else
                Exit Sub
        End Select

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsjNoExiste(C_msgSINDATOS, gstrNombCortoEmpresa)
            Exit Sub
        End If

        'Carga el formulario de consulta
        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 9000, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTFOLIOINGRESO"
                    'ConfiguraConsultas(FrmConsultas, 9000, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 1400) 'Columna del Folio
                    .set_ColWidth(1, 0, 4150) 'Columna del Concepto del Movimiento
                    .set_ColWidth(2, 0, 1200) 'Columna de la Fecha del Movimiento
                    .set_ColWidth(3, 0, 1800) 'Columna del Importe del Movimiento
                    .set_ColAlignment(2, 4)
                    For I = 1 To FrmConsultas.Flexdet.Rows - 1
                        FrmConsultas.Flexdet.set_TextMatrix(I, 2, Format(FrmConsultas.Flexdet.get_TextMatrix(I, 2), "dd/MMM/yyyy"))
                        FrmConsultas.Flexdet.set_TextMatrix(I, 3, Format(FrmConsultas.Flexdet.get_TextMatrix(I, 3), "###,##0.00"))
                    Next I
                    FrmConsultas.Top = TwipsToPixelsY(3500)
                    FrmConsultas.Left = TwipsToPixelsX(2970)
                Case "TXTFOLIORETIRO"
                    'ConfiguraConsultas(FrmConsultas, 8000, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 2500)
                    .set_ColWidth(1, 0, 2500)
                    .set_ColWidth(2, 0, 2500)
                    .set_ColAlignment(2, 4)
                    For I = 1 To FrmConsultas.Flexdet.Rows - 1
                        FrmConsultas.Flexdet.set_TextMatrix(I, 2, Format(FrmConsultas.Flexdet.get_TextMatrix(I, 2), "dd/MMM/yyyy"))
                    Next I
                    FrmConsultas.Top = TwipsToPixelsY(3500)
                    FrmConsultas.Left = TwipsToPixelsX(3500)
            End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub


    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Dim mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Dim mblnSalir As Boolean 'Para Salir Con el Esc
    Dim FueraChange As Boolean
    Dim intCodBanco As Integer
    Dim intUltFormaPago As Integer
    Dim tecla As Integer
    Dim strCuentaPesos As String
    Dim strCuentaDolares As String
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo

    '''Vars para folios de origen aplicacion cuando son depositos internos Pes/Dol
    Public strFolioPesos As String
    Public strFolioDolares As String
    Public ConsultaDepositos As Boolean

    Function Guardar() As Boolean
        Dim blnTransaccion As Boolean
        'Dim strFolioPesos As String
        'Dim strFolioDolares As String
        Dim Ejercicio As Integer
        Dim Periodo As String
        Dim I As Integer
        On Error GoTo Err_Renamed
        strFolioPesos = ""
        strFolioDolares = ""

        'Do While (VB.Timer() - sglTiempoCambio) <= 2.1
        'Loop
        'System.Windows.Forms.Application.DoEvents()

        If Not mblnNuevo Then
            If chkDepositoInterno.CheckState = 0 Then
                Exit Function
            End If
        End If
        If ValidaDatos() = False Then
            Exit Function
        End If

        If chkDepositoInterno.CheckState = 0 Then
            Cnn.BeginTrans()
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            blnTransaccion = True
            'Generar Folio del Movimiento
            Ejercicio = CInt(Format(Year(CDate(dtpFecha.Value)), "0000"))
            Periodo = Format(Month(CDate(dtpFecha.Value)), "00")
            BuscaEjercicio(dtpFecha.Value)
            gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVINGRESO & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                txtFolioIngreso.Text = C_TIPOMOVINGRESO & Format(Year(CDate(dtpFecha.Value)), "0000") & Format(Month(CDate(dtpFecha.Value)), "00") & Format(VB.Day(CDate(dtpFecha.Value)), "00") & Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
                ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVINGRESO, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
                Cmd.Execute()
            End If
            'Guardar el Movimiento Bancario
            ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioIngreso.Text, Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_MOVDEPOSITO, C_TIPOMOVINGRESO, C_NATURALEZACOMERCIAL, IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(gcurCorpoTIPOCAMBIODOLAR), "", C_TIPOPAGOJOYERIA, CStr(intCodBanco), dbcCuentaBancaria.Text, "", txtConcepto.Text, "0", "", "0", "01/01/1900", "", txtImporte.Text, "V", "01/01/1900", "", CStr(0), "01/01/1900", C_MODULOBANCOS, "", "", C_INSERCION, CStr(0))
            Cmd.Execute()
            'Guardar los Movimientos de Origen y Aplicación
            If Not frmDepositos.GuardarMovimientosOrigenAplicacion("REGISTRO DE DEPOSITOS") Then
                Cnn.RollbackTrans()
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Exit Function
            End If
            If cmdDesglose.Enabled Then
                'Guardar el Desglose del Deposito
                If Not frmDesgloseDepositos.GuardarMovimientosDepositos Then
                    Cnn.RollbackTrans()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    Exit Function
                End If
            ElseIf cmdReferencias.Enabled Then
                'Guardar las Referencias de Vouchers
                If Not frmBancosProcesoDiarioReferenciaVouchers.Guardar Then
                    Cnn.RollbackTrans()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    Exit Function
                End If
            End If
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Cnn.CommitTrans()
            blnTransaccion = False
            MsgBox("Los Datos se Han Guardado con Exito" & Chr(13) & "Se ha Generado el Folio de Ingreso " & txtFolioIngreso.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Limpiar()
        ElseIf chkDepositoInterno.CheckState = 1 Then

            Cnn.BeginTrans()
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            blnTransaccion = True
            'Generar el Movimiento para la Cuenta de Pesos
            If CDbl(Numerico(txtPesos.Text)) > 0 Then
                'Generar Folio del Movimiento
                Ejercicio = CInt(Format(Year(CDate(dtpFecha.Value)), "0000"))
                Periodo = Format(Month(CDate(dtpFecha.Value)), "00")
                BuscaEjercicio(dtpFecha.Value)
                gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVINGRESO & "'"
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute
                If RsGral.RecordCount > 0 Then
                    strFolioPesos = C_TIPOMOVINGRESO & Format(Year(CDate(dtpFechaRetiro.Value)), "0000") & Format(Month(CDate(dtpFechaRetiro.Value)), "00") & Format(VB.Day(CDate(dtpFechaRetiro.Value)), "00") & Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
                    ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVINGRESO, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                End If
                'Guardar el Movimiento Bancario
                ModStoredProcedures.PR_IMEMovimientosBancarios(strFolioPesos, Format(dtpFechaRetiro.Value, C_FORMATFECHAGUARDAR), C_MOVDEPOSITO, C_TIPOMOVINGRESO, C_NATURALEZAINTERNA, C_PESO, CStr(gcurCorpoTIPOCAMBIODOLAR), "", C_TIPOPAGOJOYERIA, CStr(intCodBanco), strCuentaPesos, "", "DEPOSITO INTERNO A LA CAJA PRINCIPAL", "0", "", "0", "01/01/1900", "", txtPesos.Text, "V", "01/01/1900", txtFolioRetiro.Text, CStr(0), "01/01/1900", "B", "", "", C_INSERCION, CStr(0))
                Cmd.Execute()
                'Actualizar la Tabla de Retiros
                For I = 1 To flexDetalle.Rows - 1
                    If Trim(flexDetalle.get_TextMatrix(I, 1)) = "Pesos" Then
                        ModStoredProcedures.PR_IE_Retiros(txtFolioRetiro.Text, "01/01/1900", "0", "0", "", "", "", "0", "", "0", "", "01/01/1900", "1", Format(dtpFechaRetiro.Value, C_FORMATFECHAGUARDAR), "0", flexDetalle.get_TextMatrix(I, 3), "0", C_MODIFICACION, CStr(0))
                        Cmd.Execute()
                    End If
                Next

                'Guardar los Movimientos de Origen y Aplicación
                If Not frmDepositosIntPes.GuardarMovimientosOrigenAplicacion("REGISTRO DE DEPOSITOS PES") Then
                    Cnn.RollbackTrans()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    Exit Function
                End If

            End If
            'Generar el Movimiento para la Cuenta de Dolares
            If CDbl(Numerico(txtDolares.Text)) > 0 Then
                'Generar Folio del Movimiento
                Ejercicio = CInt(Format(Year(CDate(dtpFecha.Value)), "0000"))
                Periodo = Format(Month(CDate(dtpFecha.Value)), "00")
                BuscaEjercicio(dtpFecha.Value)
                gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVINGRESO & "'"
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                RsGral = Cmd.Execute
                If RsGral.RecordCount > 0 Then
                    strFolioDolares = C_TIPOMOVINGRESO & Format(Year(CDate(dtpFechaRetiro.Value)), "0000") & Format(Month(CDate(dtpFechaRetiro.Value)), "00") & Format(VB.Day(CDate(dtpFechaRetiro.Value)), "00") & Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
                    ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVINGRESO, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                End If
                'Guardar el Movimiento Bancario
                ModStoredProcedures.PR_IMEMovimientosBancarios(strFolioDolares, Format(dtpFechaRetiro.Value, C_FORMATFECHAGUARDAR), C_MOVDEPOSITO, C_TIPOMOVINGRESO, C_NATURALEZAINTERNA, C_DOLAR, CStr(gcurCorpoTIPOCAMBIODOLAR), "", C_TIPOPAGOJOYERIA, CStr(intCodBanco), strCuentaDolares, "", "DEPOSITO INTERNO A LA CAJA PRINCIPAL", "0", "", "0", "01/01/1900", "", txtDolares.Text, "V", "01/01/1900", txtFolioRetiro.Text, CStr(0), "01/01/1900", "B", "", "", C_INSERCION, CStr(0))
                Cmd.Execute()
                'Actualizar la Tabla de Retiros
                For I = 1 To flexDetalle.Rows - 1
                    If Trim(flexDetalle.get_TextMatrix(I, 1)) = "Dolares" Then
                        ModStoredProcedures.PR_IE_Retiros(txtFolioRetiro.Text, "01/01/1900", "0", "0", "", "", "", "0", "", "0", "", "01/01/1900", "1", Format(dtpFechaRetiro.Value, C_FORMATFECHAGUARDAR), "0", flexDetalle.get_TextMatrix(I, 3), "0", C_MODIFICACION, CStr(0))
                        Cmd.Execute()
                    End If
                Next

                'Guardar los Movimientos de Origen y Aplicación
                If Not frmDepositosIntDol.GuardarMovimientosOrigenAplicacion("REGISTRO DE DEPOSITOS DOL") Then
                    Cnn.RollbackTrans()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    Exit Function
                End If

            End If

            Me.Cursor = System.Windows.Forms.Cursors.Default
            Cnn.CommitTrans()
            blnTransaccion = False
            If CDbl(Numerico(txtPesos.Text)) > 0 And CDbl(Numerico(txtDolares.Text)) > 0 Then
                MsgBox("Los Datos se Han Guardado con Exito" & Chr(13) & "Se Han Generado Los Siguientes Folios" & Chr(13) & strFolioPesos & " Para la Cuenta de Pesos y" & Chr(13) & strFolioDolares & " Para la Cuenta de Dolares", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            ElseIf CDbl(Numerico(txtPesos.Text)) > 0 And CDbl(Numerico(txtDolares.Text)) = 0 Then
                MsgBox("Los Datos se Han Guardado con Exito" & Chr(13) & "Se Genero el Folio " & strFolioPesos & " Para la Cuenta de Pesos", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            ElseIf CDbl(Numerico(txtDolares.Text)) > 0 And CDbl(Numerico(txtPesos.Text)) = 0 Then
                MsgBox("Los Datos se Han Guardado con Exito" & Chr(13) & "Se Genero el Folio " & strFolioDolares & " Para la Cuenta de Dolares", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            End If
            DesactivaDepositoInterno()
            ActivaDepositoComercial()
            chkDepositoInterno.CheckState = System.Windows.Forms.CheckState.Unchecked
            Limpiar()
        End If
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Sub Encabezado()
        With flexDetalle
            .Col = 0
            .Row = 0
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Forma de Pago"
            .set_ColWidth(0, 0, 3650)
            .Col = 1
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Moneda"
            .set_ColWidth(1, 0, 1000)
            .Col = 2
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Importe"
            .set_ColWidth(2, 0, 2000)
            .Col = 3
            .set_ColWidth(3, 0, 0)
            .Row = 1
            .Col = 0
        End With
    End Sub

    Sub CalculaTotales()
        Dim I As Integer
        txtPesos.Text = "0.00"
        txtDolares.Text = "0.00"
        With flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 1)) = "Pesos" Then
                    txtPesos.Text = CStr(CDbl(Numerico(Format(txtPesos.Text, "#####0.00"))) + CDbl(Numerico(Format(.get_TextMatrix(I, 2), "#####0.00"))))
                ElseIf Trim(.get_TextMatrix(I, 1)) = "Dolares" Then
                    txtDolares.Text = CStr(CDbl(Numerico(Format(txtDolares.Text, "#####0.00"))) + CDbl(Numerico(Format(.get_TextMatrix(I, 2), "#####0.00"))))
                End If
            Next
            txtPesos.Text = Format(txtPesos.Text, "###,##0.00")
            txtDolares.Text = Format(txtDolares.Text, "###,##0.00")
        End With
    End Sub

    Sub LlenaDatosRetiros()
        On Error GoTo Merr
        If (txtFolioRetiro).Text = "" Then
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT R.FolioRetiro, R.FechaRetiro, CA.DescAlmacen, R.NickUsuario, FP.DescFormaPago, " & "ISNULL(CASE FP.EsDolar WHEN 0 THEN R.ImporteFormaPago END,0) AS ImportePesos, " & "ISNULL(CASE FP.EsDolar WHEN 1 THEN R.ImporteFormaPago END,0) AS ImporteDolares, " & "CASE FP.EsDolar WHEN 0 THEN 'P' ELSE 'D' END AS TipoMoneda, R.CodFormaPago " & "FROM Retiros R, CatAlmacen CA, CatFormasPago FP " & "WHERE R.FolioRetiro = '" & Trim(txtFolioRetiro.Text) & "' AND CA.CodAlmacen = R.CodSucursal AND FP.CodFormaPago = R.CodFormaPago " & "AND R.TipoRetiro = '" & C_RETIROCAJAGENERAL & "' AND R.Estatus = 'V' AND R.PasoBancos = 0 ORDER BY TipoMoneda"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtSucursal.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            txtEnvia.Text = Trim(RsGral.Fields("NickUsuario").Value)
            With flexDetalle
                .Row = 1
                Do While Not RsGral.EOF
                    .set_TextMatrix(.Row, 0, Trim(RsGral.Fields("DescFormaPago").Value))
                    If RsGral.Fields("TipoMoneda").Value = C_PESO Then
                        .set_TextMatrix(.Row, 1, "Pesos")
                        .set_TextMatrix(.Row, 2, Format(RsGral.Fields("ImportePesos").Value, "###,##0.00"))
                    ElseIf RsGral.Fields("TipoMoneda").Value = C_DOLAR Then
                        .set_TextMatrix(.Row, 1, "Dolares")
                        .set_TextMatrix(.Row, 2, Format(RsGral.Fields("ImporteDolares").Value, "###,##0.00"))
                    End If
                    .set_TextMatrix(.Row, 3, RsGral.Fields("CodFormaPago").Value)
                    RsGral.MoveNext()
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Row = .Row + 1
                Loop
                .Row = 1
                .Col = 0
                CalculaTotales()
            End With
            If CDec(Numerico((txtPesos.Text))) <= 0 Then cmdOrigAplic_DepInt_Pes.Enabled = False Else cmdOrigAplic_DepInt_Pes.Enabled = True
            If CDec(Numerico((txtDolares.Text))) <= 0 Then cmdOrigAplic_DepInt_Dol.Enabled = False Else cmdOrigAplic_DepInt_Dol.Enabled = True
            mblnNuevo = True
            frmDepositosIntDol.Close()
            frmDepositosIntPes.Close()
            frmDepositosIntPes.Nuevo = False
            frmDepositosIntDol.Nuevo = False
        Else
            MsgBox("Folio de Retiro no Existe, Favor de Verificar ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtFolioRetiro.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub LlenaDatos()

        If (bandera = True) Then
            Exit Sub
        End If

        On Error GoTo Merr
        Dim I As Integer
        Dim Total As Decimal
        Dim Naturaleza As String
        Dim Moneda As String
        Dim FolioRetiro As String
        Dim RsAux As New ADODB.Recordset
        Dim Rs As ADODB.Recordset

        If Trim(txtFolioIngreso.Text) = "" Then
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT * FROM MovimientosBancarios MB,CatBancos CB WHERE MB.FolioMovto = '" & txtFolioIngreso.Text & "' AND MB.Movimiento = '" & C_MOVDEPOSITO & "' AND " & "MB.TipoMovto = '" & C_TIPOMOVINGRESO & "' AND CB.CodBanco = MB.CodBanco"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        Rs = Cmd.Execute
        If Rs.RecordCount > 0 Then
            frmBancosProcesoDiarioReferenciaVouchers.flexDetalle.Clear()
            frmBancosProcesoDiarioReferenciaVouchers.Encabezado()
            gStrSql = "SELECT FolioMovto FROM MovimientosBancarios WHERE Referencia = '" & txtFolioIngreso.Text & "' AND Movimiento = '" & C_MOVCANCELACION & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsAux = Cmd.Execute
            If RsAux.RecordCount > 0 Then
                lblCancelada.Text = "Movimiento de Cancelación : " & RsAux.Fields("FolioMovto").Value
            End If
            txtFolioIngreso.Text = Trim(Rs.Fields("FolioMovto").Value)
            dtpFecha.Value = Format(Rs.Fields("FechaMovto").Value, C_FORMATFECHAMOSTRAR)
            dtpFechaRetiro.Value = Format(Rs.Fields("FechaMovto").Value, C_FORMATFECHAMOSTRAR)
            FueraChange = True
            dbcBanco.Text = Trim(Rs.Fields("DescBanco").Value)
            dbcCuentaBancaria.Text = Trim(Rs.Fields("CtaBancaria").Value)
            FueraChange = False
            txtConcepto.Text = Trim(Rs.Fields("Concepto").Value)
            txtImporte.Text = Format(Rs.Fields("importe").Value, "###,##0.00")
            If Rs.Fields("Moneda").Value = C_PESO Then
                lblMoneda.Text = C_DESCPESOS
                Moneda = C_PESO
            ElseIf Rs.Fields("Moneda").Value = C_DOLAR Then
                lblMoneda.Text = C_DESCDOLARES
                Moneda = C_DOLAR
            End If
            Naturaleza = Rs.Fields("Naturaleza").Value
            FolioRetiro = Rs.Fields("FolioRetiro").Value
            If Trim(Naturaleza) = C_NATURALEZACOMERCIAL Then
                gStrSql = "SELECT * FROM MovimientosOrigenAplic MO,CatOrigenAplicRecursos CO,CatRubrosOrigenAplicRecursos CR " & "WHERE FolioMovto = '" & Trim(txtFolioIngreso.Text) & "' AND CO.CodOrigenAplicR = MO.CodOrigenAplicR AND CR.CodRubro = MO.CodRubro AND CO.CodOrigenAplicR = CR.CodOrigAplicR"
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                Rs = Cmd.Execute
                If Rs.RecordCount > 0 Then
                    With frmDepositos.flexDetalle
                        I = 1
                        .Row = 1
                        frmDepositos.lblTotal.Text = "0.00"
                        Do While Not Rs.EOF
                            .set_TextMatrix(.Row, 0, Format(Rs.Fields("CodOrigenAplicR").Value, "0000"))
                            .set_TextMatrix(.Row, 1, Trim(Rs.Fields("DescOrigenAplicR").Value))
                            .set_TextMatrix(.Row, 2, Format(Rs.Fields("CodRubro").Value, "000000"))
                            .set_TextMatrix(.Row, 3, Trim(Rs.Fields("DescRubro").Value))
                            .set_TextMatrix(.Row, 4, Format(Rs.Fields("importe").Value, "###,##0.00"))
                            With frmDepositos
                                .lblTotal.Text = CStr(CDec(Numerico(Format(.lblTotal.Text, "#####0.00"))) + CDbl(Format(Rs.Fields("importe").Value, "###,##0.00")))
                            End With
                            If .Row = .Rows - 1 Then
                                .Rows = .Rows + 1
                            End If
                            .Row = .Row + 1
                            I = I + 1
                            Rs.MoveNext()
                        Loop
                        frmDepositos.lblTotal.Text = Format(frmDepositos.lblTotal.Text, "###,##0.00")
                        frmDepositos.lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                        frmDepositos.Nuevo = True
                        'frmDepositos.Nuevo = False
                    End With
                End If
                gStrSql = "SELECT * FROM MovimientosReferencias Where FolioMovto = '" & txtFolioIngreso.Text & "'"
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                Rs = Cmd.Execute
                If Rs.RecordCount > 0 Then
                    If Rs.Fields("TipoReferencia").Value = "D" Then
                        With frmDesgloseDepositos.flexDetalle
                            I = 1
                            .Row = 1
                            frmDesgloseDepositos.lblTotal.Text = Format(Rs.Fields("ImporteDeposito").Value, "###,##0.00")
                            frmDesgloseDepositos.lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                            Do While Not Rs.EOF
                                .set_TextMatrix(.Row, 0, Trim(Rs.Fields("ReferenciaBanco").Value))
                                .set_TextMatrix(.Row, 1, Format(Rs.Fields("ImporteRef").Value, "###,##0.00"))
                                If .Row = .Rows - 1 Then
                                    .Rows = .Rows + 1
                                End If
                                .Row = .Row + 1
                                I = I + 1
                                Rs.MoveNext()
                            Loop
                        End With
                        cmdReferencias.Enabled = False
                    ElseIf Rs.Fields("TipoReferencia").Value = "V" Then
                        With frmBancosProcesoDiarioReferenciaVouchers.flexDetalle
                            .set_ColWidth(1, 0, 0)
                            .set_ColWidth(2, 0, 0)
                            .set_ColWidth(3, 0, 0)
                            .Width = VB6.TwipsToPixelsX(4450)
                            frmBancosProcesoDiarioReferenciaVouchers.Panel1.Width = VB6.TwipsToPixelsX(4800)
                            frmBancosProcesoDiarioReferenciaVouchers.cmdAceptar.Left = VB6.TwipsToPixelsX(3160)
                            frmBancosProcesoDiarioReferenciaVouchers.Label2.Visible = False
                            frmBancosProcesoDiarioReferenciaVouchers.cmdImportarVouchers.Visible = False
                            frmBancosProcesoDiarioReferenciaVouchers.Label1.Left = VB6.TwipsToPixelsX(1800)
                            frmBancosProcesoDiarioReferenciaVouchers.lblDeposito.Left = VB6.TwipsToPixelsX(3200)
                            frmBancosProcesoDiarioReferenciaVouchers.Width = VB6.TwipsToPixelsX(5350)
                            I = 1
                            .Row = 1
                            frmBancosProcesoDiarioReferenciaVouchers.lblDeposito.Text = Format(Rs.Fields("ImporteDeposito").Value, "###,##0.00")
                            Do While Not Rs.EOF
                                .set_TextMatrix(.Row, 0, Trim(Rs.Fields("ReferenciaBanco").Value))
                                .set_TextMatrix(.Row, 4, Format(Rs.Fields("ImporteRef").Value, "###,##0.00"))
                                If .Row = .Rows - 1 Then
                                    .Rows = .Rows + 1
                                End If
                                .Row = .Row + 1
                                I = I + 1
                                Rs.MoveNext()
                            Loop
                            If Rs.RecordCount > 10 Then
                                .Rows = Rs.RecordCount + 1
                            Else
                                .Rows = 11
                            End If
                            .Row = 1
                        End With
                        cmdDesglose.Enabled = False
                        frmBancosProcesoDiarioReferenciaVouchers.Nuevo = False
                    End If
                Else
                    cmdDesglose.Enabled = False
                    cmdReferencias.Enabled = False
                End If
                chkDepositoInterno.Enabled = False
                ConsultaDepositos = True
            ElseIf Trim(Naturaleza) = C_NATURALEZAINTERNA Then
                ActivaDepositoInterno()
                txtFolioRetiro.Text = FolioRetiro
                txtFolioRetiro.ReadOnly = True

                If Trim(Moneda) = C_PESO Then
                    BuscaDatosRetiro(FolioRetiro, C_PESO)
                    gStrSql = "SELECT * FROM MovimientosOrigenAplic MO,CatOrigenAplicRecursos CO,CatRubrosOrigenAplicRecursos CR " & "WHERE FolioMovto = '" & Trim(txtFolioIngreso.Text) & "' AND CO.CodOrigenAplicR = MO.CodOrigenAplicR AND CR.CodRubro = MO.CodRubro AND CO.CodOrigenAplicR = CR.CodOrigAplicR"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                    Rs = Cmd.Execute
                    If Rs.RecordCount > 0 Then
                        With frmDepositosIntPes.flexDetalle
                            I = 1
                            .Row = 1
                            frmDepositosIntPes.lblTotal.Text = "0.00"
                            Do While Not Rs.EOF
                                .set_TextMatrix(.Row, 0, Format(Rs.Fields("CodOrigenAplicR").Value, "0000"))
                                .set_TextMatrix(.Row, 1, Trim(Rs.Fields("DescOrigenAplicR").Value))
                                .set_TextMatrix(.Row, 2, Format(Rs.Fields("CodRubro").Value, "000000"))
                                .set_TextMatrix(.Row, 3, Trim(Rs.Fields("DescRubro").Value))
                                .set_TextMatrix(.Row, 4, Format(Rs.Fields("importe").Value, "###,##0.00"))
                                With frmDepositosIntPes
                                    .lblTotal.Text = CStr(CDec(Numerico(Format(.lblTotal.Text, "#####0.00"))) + CDbl(Format(Rs.Fields("importe").Value, "###,##0.00")))
                                End With
                                If .Row = .Rows - 1 Then
                                    .Rows = .Rows + 1
                                End If
                                .Row = .Row + 1
                                I = I + 1
                                Rs.MoveNext()
                            Loop
                            frmDepositosIntPes.lblTotal.Text = Format(frmDepositosIntPes.lblTotal.Text, "###,##0.00")
                            frmDepositosIntPes.lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                            frmDepositosIntPes.Nuevo = True
                        End With
                        cmdOrigAplic_DepInt_Pes.Enabled = True
                        ConsultaDepositos = True
                    End If

                ElseIf Trim(Moneda) = C_DOLAR Then
                    BuscaDatosRetiro(FolioRetiro, C_DOLAR)
                    gStrSql = "SELECT * FROM MovimientosOrigenAplic MO,CatOrigenAplicRecursos CO,CatRubrosOrigenAplicRecursos CR " & "WHERE FolioMovto = '" & Trim(txtFolioIngreso.Text) & "' AND CO.CodOrigenAplicR = MO.CodOrigenAplicR AND CR.CodRubro = MO.CodRubro AND CO.CodOrigenAplicR = CR.CodOrigAplicR"
                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "dbo.Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
                    Rs = Cmd.Execute
                    If Rs.RecordCount > 0 Then
                        With frmDepositosIntDol.flexDetalle
                            I = 1
                            .Row = 1
                            frmDepositosIntDol.lblTotal.Text = "0.00"
                            Do While Not Rs.EOF
                                .set_TextMatrix(.Row, 0, Format(Rs.Fields("CodOrigenAplicR").Value, "0000"))
                                .set_TextMatrix(.Row, 1, Trim(Rs.Fields("DescOrigenAplicR").Value))
                                .set_TextMatrix(.Row, 2, Format(Rs.Fields("CodRubro").Value, "000000"))
                                .set_TextMatrix(.Row, 3, Trim(Rs.Fields("DescRubro").Value))
                                .set_TextMatrix(.Row, 4, Format(Rs.Fields("importe").Value, "###,##0.00"))
                                With frmDepositosIntDol
                                    .lblTotal.Text = CStr(CDec(Numerico(Format(.lblTotal.Text, "#####0.00"))) + CDbl(Format(Rs.Fields("importe").Value, "###,##0.00")))
                                End With
                                If .Row = .Rows - 1 Then
                                    .Rows = .Rows + 1
                                End If
                                .Row = .Row + 1
                                I = I + 1
                                Rs.MoveNext()
                            Loop
                            frmDepositosIntDol.lblTotal.Text = Format(frmDepositosIntDol.lblTotal.Text, "###,##0.00")
                            frmDepositosIntDol.lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                            frmDepositosIntDol.Nuevo = True
                        End With
                        cmdOrigAplic_DepInt_Dol.Enabled = True
                        ConsultaDepositos = True
                    End If
                End If
                chkDepositoInterno.Enabled = False
                cmdDesglose.Enabled = False
                cmdOrigenyAplicacion.Enabled = False
                cmdReferencias.Enabled = False
                ModEstandar.CentrarForma(Me, MDIMenuPrincipalCorpo)
            End If
            mblnNuevo = False
            dtpFecha.Enabled = False
            dtpFechaRetiro.Enabled = False

            '''If CCur(Numerico(txtPesos.text)) <= 0 Then cmdOrigAplic_DepInt_Pes.Enabled = False Else cmdOrigAplic_DepInt_Pes.Enabled = True
            '''If CCur(Numerico(txtDolares.text)) <= 0 Then cmdOrigAplic_DepInt_Dol.Enabled = False Else cmdOrigAplic_DepInt_Dol.Enabled = True

        Else
            MsgBox("Folio de Movimiento de Ingreso no Existe ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Frame1.Enabled = True
            txtFolioIngreso.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub BuscaDatosRetiro(ByRef Folio As String, ByRef Moneda As String)
        On Error GoTo Merr
        If Trim(Moneda) = C_PESO Then
            gStrSql = "SELECT R.FolioRetiro, R.FechaRetiro, CA.DescAlmacen, R.NickUsuario, FP.DescFormaPago, " & "R.ImporteFormaPago, R.CodFormaPago " & "FROM Retiros R, CatAlmacen CA, CatFormasPago FP " & "WHERE R.FolioRetiro = '" & Trim(Folio) & "' AND CA.CodAlmacen = R.CodSucursal AND FP.CodFormaPago = R.CodFormaPago " & "AND R.TipoRetiro = '" & C_RETIROCAJAGENERAL & "' AND R.Estatus = 'V' AND R.PasoBancos = 1 AND FP.EsDolar = 0"
        ElseIf Trim(Moneda) = C_DOLAR Then
            gStrSql = "SELECT R.FolioRetiro, R.FechaRetiro, CA.DescAlmacen, R.NickUsuario, FP.DescFormaPago, " & "R.ImporteFormaPago, R.CodFormaPago " & "FROM Retiros R, CatAlmacen CA, CatFormasPago FP " & "WHERE R.FolioRetiro = '" & Trim(Folio) & "' AND CA.CodAlmacen = R.CodSucursal AND FP.CodFormaPago = R.CodFormaPago " & "AND R.TipoRetiro = '" & C_RETIROCAJAGENERAL & "' AND R.Estatus = 'V' AND R.PasoBancos = 1 AND FP.EsDolar = 1"
        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtSucursal.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            txtEnvia.Text = Trim(RsGral.Fields("NickUsuario").Value)
            With flexDetalle
                .Row = 1
                Do While Not RsGral.EOF
                    .set_TextMatrix(.Row, 0, Trim(RsGral.Fields("DescFormaPago").Value))
                    If Trim(Moneda) = C_PESO Then
                        .set_TextMatrix(.Row, 1, "Pesos")
                        .set_TextMatrix(.Row, 2, Format(RsGral.Fields("ImporteFormaPago").Value, "###,##0.00"))
                    ElseIf Trim(Moneda) = C_DOLAR Then
                        .set_TextMatrix(.Row, 1, "Dolares")
                        .set_TextMatrix(.Row, 2, Format(RsGral.Fields("ImporteFormaPago").Value, "###,##0.00"))
                    End If
                    .set_TextMatrix(.Row, 3, RsGral.Fields("CodFormaPago").Value)
                    RsGral.MoveNext()
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Row = .Row + 1
                Loop
                .Row = 1
                .Col = 0
                CalculaTotales()
            End With
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function ChecaGridDepositos() As Boolean
        Dim I As Integer
        ChecaGridDepositos = False
        With frmDesgloseDepositos.flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                    ChecaGridDepositos = True
                End If
            Next
        End With
    End Function

    Function ExisteBancoPrincipal() As Boolean
        On Error GoTo Merr
        ExisteBancoPrincipal = True
        gStrSql = "SELECT * FROM CatBancos WHERE ControlInterno = 1 AND Sucursal = 0"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            MsgBox("¡¡¡ATENCION!!! No Existe el Banco Caja Principal, Favor de Verificar...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            ExisteBancoPrincipal = False
        Else
            intCodBanco = RsGral.Fields("CodBanco").Value
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ExisteCuenta(ByRef Moneda As String) As Boolean
        On Error GoTo Merr
        ExisteCuenta = True
        gStrSql = "SELECT * FROM CatBancos CB,CatCuentasBancarias CC WHERE CB.ControlInterno = 1 AND " & "CB.Sucursal = 0 AND CB.CodBanco = " & intCodBanco & " AND CB.CodBanco = CC.Codbanco AND CC.Moneda = '" & Trim(Moneda) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            ExisteCuenta = False
        Else
            If Moneda = C_PESO Then
                strCuentaPesos = Trim(RsGral.Fields("CtaBancaria").Value)
            ElseIf Moneda = C_DOLAR Then
                strCuentaDolares = Trim(RsGral.Fields("CtaBancaria").Value)
            End If
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Not BuscaUltimoCierre(dtpFecha.Value) Then
            Exit Function
        End If
        If chkDepositoInterno.CheckState = 0 Then
            If Len(Trim(dbcBanco.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Nombre del Banco", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                dbcBanco.Focus()
                Exit Function
            End If
            If Len(Trim(dbcCuentaBancaria.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Cuenta Bancaria", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                dbcCuentaBancaria.Focus()
                Exit Function
            End If
            If Len(Trim(txtConcepto.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Concepto", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtConcepto.Focus()
                Exit Function
            End If
            If CDbl(Numerico(txtImporte.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Importe del Pago", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtImporte.Focus()
                Exit Function
            End If
            If Not ChecaGrid(frmDepositos) Then
                MsgBox("No se Han Capturado los Movimientos de Origen y Aplicación ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                cmdOrigenyAplicacion_Click(cmdOrigenyAplicacion, New System.EventArgs())
                Exit Function
            End If
            If ChecaGridDepositos() Then
                If Numerico((frmDesgloseDepositos.lblImporte).Text) <> Numerico((frmDesgloseDepositos.lblTotal).Text) Then
                    MsgBox("El Total del Desglose de Depositos no es Igual al Importe del Deposito ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    cmdDesglose_Click(cmdDesglose, New System.EventArgs())
                    Exit Function
                End If
            End If
            If Numerico((frmDepositos.lblImporte).Text) <> Numerico((frmDepositos.lblTotal).Text) Then
                MsgBox("El Total de los Movimientos de Origen y Aplicación no es Igual al Importe del Pago...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                cmdOrigenyAplicacion_Click(cmdOrigenyAplicacion, New System.EventArgs())
                Exit Function
            End If
        ElseIf chkDepositoInterno.CheckState = 1 Then
            If Trim(txtFolioRetiro.Text) = "" Then
                MsgBox("¡¡¡ATENCION!!! Debe seleccionar un folio de retiro para poder registrar el depósito interno...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtFolioRetiro.Focus()
                Exit Function
            End If
            If Trim(txtEnvia.Text) = "" Then
                MsgBox("¡¡¡ATENCION!!! No existe información sobre quién registro el folio del retiro ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtFolioRetiro.Focus()
                Exit Function
            End If
            If Trim(txtSucursal.Text) = "" Then
                MsgBox("¡¡¡ATENCION!!! No existe información sobre qué sucursal generó el folio de retiro...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtFolioRetiro.Focus()
                Exit Function
            End If

            If Not ExisteBancoPrincipal() Then
                Exit Function
            End If
            If CDbl(Numerico(txtPesos.Text)) > 0 Then
                If Not ExisteCuenta(C_PESO) Then
                    MsgBox("¡¡¡ATENCION!!! No Existe una Cuenta de Pesos en el Banco Caja Principal ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    Exit Function
                End If
                If Not ChecaGrid(frmDepositosIntPes) Then
                    MsgBox("No se Han Capturado los Movimientos de Origen y Aplicación ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    cmdOrigAplic_DepInt_Pes_Click(cmdOrigAplic_DepInt_Pes, New System.EventArgs())
                    Exit Function
                End If
                If CDec(Numerico((frmDepositosIntPes.lblImporte).Text)) <> CDec(Numerico((frmDepositosIntPes.lblTotal).Text)) Then
                    MsgBox("El Total de los Movimientos de Origen y Aplicación de la cuenta en pesos no es Igual al Importe del Pago...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    cmdOrigAplic_DepInt_Pes_Click(cmdOrigAplic_DepInt_Pes, New System.EventArgs())
                    Exit Function
                End If
            End If
            If CDbl(Numerico(txtDolares.Text)) > 0 Then
                If Not ExisteCuenta(C_DOLAR) Then
                    MsgBox("¡¡¡ATENCION!!! No Existe una Cuenta de Dolares en el Banco Caja Principal ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    Exit Function
                End If
                If Not ChecaGrid(frmDepositosIntDol) Then
                    MsgBox("No se Han Capturado los Movimientos de Origen y Aplicación ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    cmdOrigAplic_DepInt_Dol_Click(cmdOrigAplic_DepInt_Dol, New System.EventArgs())
                    Exit Function
                End If
                If CDec(Numerico((frmDepositosIntDol.lblImporte).Text)) <> CDec(Numerico((frmDepositosIntDol.lblTotal).Text)) Then
                    MsgBox("El Total de los Movimientos de Origen y Aplicación de la cuenta en dólares no es Igual al Importe del Pago...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    cmdOrigAplic_DepInt_Dol_Click(cmdOrigAplic_DepInt_Dol, New System.EventArgs())
                    Exit Function
                End If
            End If
        End If
        ValidaDatos = True
    End Function

    Sub Limpiar()
        If chkDepositoInterno.Enabled = False Then
            txtFolioRetiro.Text = ""
            DesactivaDepositoInterno()
        End If
        Nuevo()
        'If chkDepositoInterno.Value = 0 Then
        ActivaDepositoComercial()
        DesactivaDepositoInterno()
        txtFolioRetiro.Text = ""
        txtFolioIngreso.Text = ""
        txtFolioIngreso.Focus()
        chkDepositoInterno.CheckState = System.Windows.Forms.CheckState.Unchecked
        'ElseIf chkDepositoInterno.Value = 1 Then
        '    txtFolioRetiro = ""
        '    txtFolioRetiro.SetFocus
        'End If
    End Sub

    Sub Nuevo()

        If (bandera = True) Then
            Exit Sub
        End If

        strFolioPesos = ""
        strFolioDolares = ""
        lblMoneda.Text = ""
        dtpFecha.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        lblCancelada.Text = ""
        dbcBanco.Text = ""
        ' dbcBanco.RowSource = Nothing
        dbcCuentaBancaria.Text = ""
        ' dbcCuentaBancaria.RowSource = Nothing
        txtConcepto.Text = ""
        txtImporte.Text = "0.00"
        Frame1.Enabled = True
        InicializaVariables()
        gblnSalir = True
        frmDepositos.Close()
        gblnSalir = False
        dtpFecha.Enabled = False
        frmDesgloseDepositos.Close()
        frmBancosProcesoDiarioReferenciaVouchers.Close()
        frmDepositosIntPes.Close()
        frmDepositosIntDol.Close()
        frmBancosProcesoDiarioReferenciaVouchers.Nuevo = True
        txtFolioRetiro.ReadOnly = False
        dtpFechaRetiro.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        txtSucursal.Text = ""
        txtEnvia.Text = ""
        flexDetalle.Clear()
        Encabezado()
        txtPesos.Text = "0.00"
        txtDolares.Text = "0.00"
        chkDepositoInterno.Enabled = True
        If chkDepositoInterno.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cmdDesglose.Enabled = True
            cmdOrigenyAplicacion.Enabled = True
            cmdReferencias.Enabled = True
        End If
        frmDepositosIntPes.Nuevo = False
        frmDepositosIntDol.Nuevo = False
        cmdOrigAplic_DepInt_Dol.Enabled = False
        cmdOrigAplic_DepInt_Pes.Enabled = False
        ConsultaDepositos = False
    End Sub

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnSalir = False
        FueraChange = False
        intCodBanco = 0
    End Sub

    Sub ActivaDepositoInterno()
        Me.Height = VB6.TwipsToPixelsY(9400)
        txtFolioRetiro.Enabled = True
        txtFolioRetiro.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        txtSucursal.Enabled = True
        txtSucursal.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        txtEnvia.Enabled = True
        txtEnvia.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        flexDetalle.Enabled = True
        flexDetalle.HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightAlways
        flexDetalle.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        txtPesos.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        txtPesos.Enabled = True
        txtDolares.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        txtDolares.Enabled = True
        cmdOrigAplic_DepInt_Dol.Enabled = False
        cmdOrigAplic_DepInt_Pes.Enabled = False
    End Sub

    Sub ActivaDepositoComercial()
        txtFolioIngreso.Enabled = True
        txtFolioIngreso.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        dbcBanco.Enabled = True
        dbcBanco.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        dbcCuentaBancaria.Enabled = True
        dbcCuentaBancaria.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        txtConcepto.Enabled = True
        txtConcepto.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFFF)
        txtImporte.Enabled = True
        txtImporte.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
        cmdDesglose.Enabled = True
        cmdOrigenyAplicacion.Enabled = True
        cmdReferencias.Enabled = True
    End Sub

    Sub DesactivaDepositoInterno()
        txtFolioRetiro.Enabled = False
        txtFolioRetiro.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        txtSucursal.Enabled = False
        txtSucursal.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        txtEnvia.Enabled = False
        txtEnvia.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        flexDetalle.Enabled = False
        flexDetalle.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        flexDetalle.HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
        txtPesos.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        txtPesos.Enabled = False
        txtDolares.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        txtDolares.Enabled = False
        Me.Height = VB6.TwipsToPixelsY(4425)
    End Sub

    Sub DesactivaDepositoComercial()
        txtFolioIngreso.Enabled = False
        txtFolioIngreso.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        dbcBanco.Enabled = False
        dbcBanco.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        dbcCuentaBancaria.Enabled = False
        dbcCuentaBancaria.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        txtConcepto.Enabled = False
        txtConcepto.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        txtImporte.Enabled = False
        txtImporte.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000B)
        cmdDesglose.Enabled = False
        cmdReferencias.Enabled = False
        cmdOrigenyAplicacion.Enabled = False
    End Sub

    Private Sub chkDepositoInterno_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDepositoInterno.CheckStateChanged
        If chkDepositoInterno.CheckState = 1 Then
            Nuevo()
            ActivaDepositoInterno()
            DesactivaDepositoComercial()
            'ModEstandar.CentrarForma Me, MenuPrincipal
            txtFolioRetiro.Text = ""
            If VB.Right(Trim(txtFolioIngreso.Text), 4) <> "0000" Then
                txtFolioIngreso.Text = C_TIPOMOVINGRESO & Format(Year(CDate(dtpFecha.Value)), "0000") & Format(Month(CDate(dtpFecha.Value)), "00") & Format(VB.Day(CDate(dtpFecha.Value)), "00") & "0000"
            End If
        Else
            DesactivaDepositoInterno()
            ActivaDepositoComercial()
            'ModEstandar.CentrarForma Me, MenuPrincipal
            Nuevo()
            txtFolioRetiro.Text = ""
        End If
    End Sub

    Private Sub chkDepositoInterno_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDepositoInterno.Enter
        Pon_Tool()
    End Sub

    Private Sub chkDepositoInterno_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkDepositoInterno.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            If chkDepositoInterno.CheckState = 1 Then
                mblnSalir = True
                Me.Close()
            End If
        End If
    End Sub

    Private Sub cmdDesglose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDesglose.Click
        'frmDesgloseDepositos.InitializeComponent()
        If Trim(dbcBanco.Text) <> "" And Trim(dbcCuentaBancaria.Text) <> "" Then
            If CDbl(Numerico(txtImporte.Text)) > 0 Then
                If Not mblnNuevo Then
                    frmDesgloseDepositos.cmdAceptar.TabIndex = 0
                    frmDesgloseDepositos.flexDetalle.TabIndex = 1
                    frmDesgloseDepositos.flexDetalle.Enabled = False
                Else
                    frmDesgloseDepositos.cmdAceptar.TabIndex = 1
                    frmDesgloseDepositos.flexDetalle.TabIndex = 0
                    frmDesgloseDepositos.cmdAceptar.Enabled = False
                End If
                frmDesgloseDepositos.Text = "Desglose de Depósitos Bancarios"
                frmDesgloseDepositos.Label1.Text = "Importe del Deposito : "
                frmDesgloseDepositos.Panel1.Text = "Desglose del Depósito"
                frmDesgloseDepositos.lblMoneda.Text = lblMoneda.Text
                frmDesgloseDepositos.lblImporte.Text = txtImporte.Text
                frmDesgloseDepositos.flexDetalle.Col = 0
                frmDesgloseDepositos.flexDetalle.Row = 1
                frmDesgloseDepositos.Tag = "frmDesgloseDepositos"
                frmDesgloseDepositos.ShowDialog()
            Else
                MsgBox("El Importe del Depósito debe ser Mayor que Cero, Favor de Teclear un Importe ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtImporte.Focus()
            End If
        Else
            MsgBox("Favor de Seleccionar Una Cuenta Bancaria Valida ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcCuentaBancaria.Focus()
        End If
    End Sub

    Private Sub cmdDesglose_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDesglose.Enter
        Pon_Tool()
    End Sub

    Private Sub cmdOrigAplic_DepInt_Dol_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOrigAplic_DepInt_Dol.Click
        If frmDepositosIntDol.Nuevo Then
            frmDepositosIntDol.cmdAceptar.TabIndex = 0
            frmDepositosIntDol.flexDetalle.TabIndex = 1
            frmDepositosIntDol.flexDetalle.Enabled = False
        Else
            frmDepositosIntDol.flexDetalle.TabIndex = 0
            frmDepositosIntDol.cmdAceptar.TabIndex = 1
            frmDepositosIntDol.cmdAceptar.Enabled = False
        End If
        frmDepositosIntDol.Tag = "frmDepositosIntDol"
        frmDepositosIntDol.Text = "Origen de Recursos (Registro de Depósitos)"
        frmDepositosIntDol.lblMoneda.Text = lblMoneda.Text
        frmDepositosIntDol.lblFechaMovimiento.Text = dtpFecha.Value
        frmDepositosIntDol.lblImporte.Text = Format(txtDolares.Text, "###,##0.00")
        frmDepositosIntDol.flexDetalle.Col = 0
        frmDepositosIntDol.flexDetalle.Row = 1
        frmDepositosIntDol.ShowDialog()
    End Sub

    Private Sub cmdOrigAplic_DepInt_Pes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOrigAplic_DepInt_Pes.Click
        If frmDepositosIntPes.Nuevo Then
            frmDepositosIntPes.cmdAceptar.TabIndex = 0
            frmDepositosIntPes.flexDetalle.TabIndex = 1
            frmDepositosIntPes.flexDetalle.Enabled = False
        Else
            frmDepositosIntPes.flexDetalle.TabIndex = 0
            frmDepositosIntPes.cmdAceptar.TabIndex = 1
            frmDepositosIntPes.cmdAceptar.Enabled = False
        End If
        frmDepositosIntPes.Tag = "frmDepositosIntPes"
        frmDepositosIntPes.Text = "Origen de Recursos (Registro de Depósitos)"
        frmDepositosIntPes.lblMoneda.Text = lblMoneda.Text
        frmDepositosIntPes.lblFechaMovimiento.Text = dtpFecha.Value
        frmDepositosIntPes.lblImporte.Text = Format(txtPesos.Text, "###,##0.00")
        frmDepositosIntPes.flexDetalle.Col = 0
        frmDepositosIntPes.flexDetalle.Row = 1
        frmDepositosIntPes.ShowDialog()

    End Sub

    Private Sub cmdOrigenyAplicacion_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOrigenyAplicacion.Click
        If Trim(dbcBanco.Text) <> "" And Trim(dbcCuentaBancaria.Text) <> "" Then
            If CDbl(Numerico(txtImporte.Text)) > 0 Then
                If frmDepositos.Nuevo Then
                    frmDepositos.cmdAceptar.TabIndex = 0
                    frmDepositos.flexDetalle.TabIndex = 1
                    frmDepositos.flexDetalle.Enabled = False
                Else
                    frmDepositos.flexDetalle.TabIndex = 0
                    frmDepositos.cmdAceptar.TabIndex = 1
                    frmDepositos.cmdAceptar.Enabled = False
                End If
                frmDepositos.Tag = "frmDepositos"
                frmDepositos.Text = "Origen de Recursos (Registro de Depósitos)"
                frmDepositos.lblMoneda.Text = lblMoneda.Text
                frmDepositos.lblFechaMovimiento.Text = dtpFecha.Value
                frmDepositos.lblImporte.Text = txtImporte.Text
                frmDepositos.flexDetalle.Col = 0
                frmDepositos.flexDetalle.Row = 1
                frmDepositos.ShowDialog()
            Else
                MsgBox("El Importe del Depósito debe ser Mayor que Cero, Favor de Teclear un Importe ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                txtImporte.Focus()
            End If
        Else
            MsgBox("Favor de Seleccionar Una Cuenta Bancaria Valida ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcCuentaBancaria.Focus()
        End If
    End Sub

    Private Sub cmdOrigenyAplicacion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOrigenyAplicacion.Enter
        Pon_Tool()
    End Sub

    Private Sub cmdReferencias_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReferencias.Click
        If Trim(dbcBanco.Text) <> "" And Trim(dbcCuentaBancaria.Text) <> "" Then
            Dim frmBancosProcesoDiarioReferenciaVouchers As frmBancosProcesoDiarioReferenciaVouchers = New frmBancosProcesoDiarioReferenciaVouchers()
            frmBancosProcesoDiarioReferenciaVouchers.InitializeComponent()
            frmBancosProcesoDiarioReferenciaVouchers.lblMoneda.Text = lblMoneda.Text
            frmBancosProcesoDiarioReferenciaVouchers.ShowDialog()
        Else
            MsgBox("Favor de Seleccionar Una Cuenta Bancaria Valida ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            dbcCuentaBancaria.Focus()
        End If
    End Sub

    Private Sub dbcBanco_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcBanco" Then
        '    Exit Sub
        'End If
        dbcCuentaBancaria.Text = ""
        lblMoneda.Text = ""
        gStrSql = "SELECT CodBanco,RTRIM(LTRIM(DescBanco)) DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' AND ControlInterno = 0 ORDER BY DescBanco"
        DCChange(gStrSql, tecla)
        If Trim(dbcCuentaBancaria.Text) = "" Then
            lblMoneda.Text = ""
        End If
        intCodBanco = 0
    End Sub

    Private Sub dbcBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Enter
        gStrSql = "SELECT CodBanco,RTRIM(LTRIM(DescBanco)) DescBanco FROM CatBancos WHERE ControlInterno = 0 ORDER BY DescBanco"
        DCGotFocus(gStrSql, dbcBanco)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcBanco_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.KeyDown
        'tecla = eventSender.keyCode
        'If eventSender.keyCode = System.Windows.Forms.Keys.Escape Then
        '    txtFolioIngreso.Focus()
        'End If
    End Sub

    Private Sub dbcBanco_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.KeyPress
        'eventArgs.keyAscii = ModEstandar.gp_CampoMayusculas(eventArgs.keyAscii)
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Leave
        gStrSql = "SELECT CodBanco,RTRIM(LTRIM(DescBanco)) DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' AND ControlInterno = 0 ORDER BY DescBanco"
        DCLostFocus(dbcBanco, gStrSql, intCodBanco)
        lblCodBanco.Text = CStr(intCodBanco)
    End Sub

    Private Sub dbcCuentaBancaria_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcCuentaBancaria" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,RTRIM(LTRIM(CtaBancaria)) CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCChange(gStrSql, tecla)
        frmBancosProcesoDiarioReferenciaVouchers.Close()
        'intCodBanco = 0
    End Sub

    Private Sub dbcCuentaBancaria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Enter
        gStrSql = "SELECT CodBanco,RTRIM(LTRIM(CtaBancaria)) CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCGotFocus(gStrSql, dbcCuentaBancaria)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcCuentaBancaria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.KeyDown
        'tecla = eventSender.keyCode
        'If eventSender.keyCode = System.Windows.Forms.Keys.Escape Then
        '    dbcBanco.Focus()
        'End If
    End Sub

    Private Sub dbcCuentaBancaria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcCuentaBancaria.KeyPress
        'eventArgs.keyAscii = ModEstandar.gp_CampoMayusculas(eventArgs.keyAscii)
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcCuentaBancaria_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.KeyUp
        Dim Aux As String
        Aux = dbcCuentaBancaria.Text
        If dbcCuentaBancaria.SelectedItem <> 0 Then
            dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        End If
        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub dbcCuentaBancaria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Leave
        On Error GoTo Err_Renamed
        gStrSql = "SELECT CodBanco,RTRIM(LTRIM(CtaBancaria)) CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCLostFocus(dbcCuentaBancaria, gStrSql, intCodBanco)
        gStrSql = "SELECT Moneda FROM CatCuentasBancarias WHERE CtaBancaria = '" & Trim(dbcCuentaBancaria.Text) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMoneda.Visible = True
                lblMoneda.Text = C_DESCPESOS
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMoneda.Visible = True
                lblMoneda.Text = C_DESCDOLARES
            End If
        End If
        frmBancosProcesoDiarioReferenciaVouchers.Close()
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcCuentaBancaria_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcCuentaBancaria.MouseUp
        Dim Aux As String
        Aux = dbcCuentaBancaria.Text
        'If dbcCuentaBancaria.SelectedItem <> 0 Then
        '    dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        'End If
        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub dtpFecha_ValueChanged(sender As Object, e As EventArgs) Handles dtpFecha.ValueChanged
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFecha_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFecha.Click
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFecha_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFecha.KeyPress
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaRetiro_ValueChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaRetiro.ValueChanged
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaRetiro_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaRetiro.Click
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFechaRetiro_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaRetiro.KeyPress
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub FlexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        Pon_Tool()
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeDepositos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeDepositos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeDepositos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If Me.ActiveControl.Name = "txtFolioIngreso" Then
                    If Len(Trim(txtFolioIngreso.Text)) = 13 And VB.Right(txtFolioIngreso.Text, 4) <> "0000" Then
                        Frame1.Enabled = False
                    End If
                End If
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape

                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtFolioIngreso" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeDepositos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeDepositos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        bandera = True
        'ModVariables.frmDesgloseDepositos.InitializeComponent()
        Left = VB6.TwipsToPixelsX(3000)
        Top = VB6.TwipsToPixelsY(170)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        Encabezado()
        InicializaVariables()
        Nuevo()
        BuscaEjercicio(dtpFecha.Value)
    End Sub

    Private Sub frmBancosProcesoDiarioRegistrodeDepositos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
        '    'If Cambios = True And mblnNuevo = False Then
        '    'Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
        '    'Case vbYes: 'Guardar el registro
        '    'If Guardar = False Then
        '    'Cancel = 1
        '    'End If
        '    'Case vbNo: 'No hace nada y permite el cierre del formulario
        '    'Case vbCancel: 'Cancela el cierre del formulario sin guardar
        '    'Cancel = 1
        '    'End Select
        '    'End If
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

    Private Sub frmBancosProcesoDiarioRegistrodeDepositos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
        Me.Hide()

        gblnSalir = True
        'frmDepositos.Close()
        frmDepositos = Nothing
        'frmDepositosIntPes.Close()
        frmDepositosIntPes = Nothing
        'frmDepositosIntDol.Close()
        frmDepositosIntDol = Nothing
        'frmDesgloseDepositos.Close()
        frmDesgloseDepositos = Nothing
        'frmBancosProcesoDiarioReferenciaVouchers.Close()
        frmBancosProcesoDiarioReferenciaVouchers = Nothing
        frmBancosProcesoDiarioImportacionVouchers = Nothing
    End Sub

    Private Sub txtConcepto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConcepto.Enter
        SelTextoTxt(txtConcepto)
        Pon_Tool()
    End Sub

    Private Sub txtConcepto_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConcepto.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "!""#$%&/()=?'¡¿*,;.:<>@+-_")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDolares.Enter
        Pon_Tool()
    End Sub

    Private Sub txtEnvia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEnvia.Enter
        Pon_Tool()
    End Sub

    Private Sub txtFolioIngreso_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioIngreso.TextChanged
        If FueraChange Then Exit Sub
        If Not mblnNuevo Then
            Nuevo()
            txtFolioRetiro.Text = ""
            DesactivaDepositoInterno()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtFolioIngreso_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioIngreso.Enter
        strControlActual = UCase("txtFolioIngreso")
        SelTextoTxt(txtFolioIngreso)
        Pon_Tool()
    End Sub

    Private Sub txtFolioIngreso_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioIngreso.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, C_TIPOMOVINGRESO)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioIngreso_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioIngreso.Leave

        If Me.ActiveControl.Name = "btnBuscar" Then
            Exit Sub
        End If

        If Trim(txtFolioIngreso.Text) = "" Then
            txtFolioIngreso.Text = C_TIPOMOVINGRESO & Format(Year(CDate(dtpFecha.Value)), "0000") & Format(Month(CDate(dtpFecha.Value)), "00") & Format(VB.Day(CDate(dtpFecha.Value)), "00") & "0000"
            Exit Sub
        End If
        If mblnCambiosEnCodigo = True And txtFolioIngreso.Text <> "" And VB.Right(txtFolioIngreso.Text, 4) <> "0000" Then
            LlenaDatos()
            frmDepositos.Hide()
        End If
    End Sub
    Private Sub txtFolioRetiro_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioRetiro.TextChanged
        'If Not mblnNuevo Then
        Nuevo()
        '    mblnNuevo = True
        'End If
        'mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtFolioRetiro_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioRetiro.Enter
        strControlActual = UCase("txtFolioRetiro")
        SelTextoTxt(txtFolioRetiro)
        Pon_Tool()
    End Sub

    Private Sub txtFolioRetiro_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioRetiro.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, "T")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFolioRetiro_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioRetiro.Leave

        If Me.ActiveControl.Name = "btnBuscar" Then
            Exit Sub
        End If

        If Trim(txtFolioRetiro.Text) = "" Then
            Exit Sub
        End If
        If mblnCambiosEnCodigo = True And txtFolioRetiro.Text <> "" And mblnNuevo Then
            LlenaDatosRetiros()
        End If
    End Sub

    Private Sub txtImporte_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.TextChanged
        If txtImporte.Text = "" Then
            txtImporte.Text = "0.00"
        End If
    End Sub

    Private Sub txtImporte_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Enter
        SelTextoTxt(txtImporte)
        Pon_Tool()
    End Sub

    Private Sub txtImporte_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImporte.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.MskCantidad(txtImporte.Text, KeyAscii, 15, 2, (txtImporte.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtImporte_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Leave
        txtImporte.Text = Format(txtImporte.Text, "###,##0.00")
    End Sub

    Private Sub txtPesos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPesos.Enter
        Pon_Tool()
    End Sub

    Private Sub txtSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSucursal.Enter
        Pon_Tool()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBancosProcesoDiarioRegistrodeDepositos))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdReferencias = New System.Windows.Forms.Button()
        Me.txtFolioIngreso = New System.Windows.Forms.TextBox()
        Me.txtFolioRetiro = New System.Windows.Forms.TextBox()
        Me.txtSucursal = New System.Windows.Forms.TextBox()
        Me.txtEnvia = New System.Windows.Forms.TextBox()
        Me.txtPesos = New System.Windows.Forms.TextBox()
        Me.txtDolares = New System.Windows.Forms.TextBox()
        Me.cmdDesglose = New System.Windows.Forms.Button()
        Me.cmdOrigenyAplicacion = New System.Windows.Forms.Button()
        Me.txtConcepto = New System.Windows.Forms.TextBox()
        Me.txtImporte = New System.Windows.Forms.TextBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkDepositoInterno = New System.Windows.Forms.CheckBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.cmdOrigAplic_DepInt_Pes = New System.Windows.Forms.Button()
        Me.cmdOrigAplic_DepInt_Dol = New System.Windows.Forms.Button()
        Me.dtpFechaRetiro = New System.Windows.Forms.DateTimePicker()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me.dbcCuentaBancaria = New System.Windows.Forms.ComboBox()
        Me.lblCancelada = New System.Windows.Forms.Label()
        Me.lblMoneda = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.lblCodBanco = New System.Windows.Forms.Label()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Frame4.SuspendLayout()
        Me.Frame2.SuspendLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame3.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdReferencias
        '
        Me.cmdReferencias.BackColor = System.Drawing.SystemColors.Control
        Me.cmdReferencias.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdReferencias.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReferencias.Location = New System.Drawing.Point(424, 128)
        Me.cmdReferencias.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdReferencias.Name = "cmdReferencias"
        Me.cmdReferencias.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdReferencias.Size = New System.Drawing.Size(112, 36)
        Me.cmdReferencias.TabIndex = 7
        Me.cmdReferencias.Text = "&Referencias Vouchers"
        Me.ToolTip1.SetToolTip(Me.cmdReferencias, "Muestra la Ventana de Captura de Origen.")
        Me.cmdReferencias.UseVisualStyleBackColor = False
        '
        'txtFolioIngreso
        '
        Me.txtFolioIngreso.AcceptsReturn = True
        Me.txtFolioIngreso.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioIngreso.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioIngreso.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioIngreso.Location = New System.Drawing.Point(104, 12)
        Me.txtFolioIngreso.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioIngreso.MaxLength = 13
        Me.txtFolioIngreso.Name = "txtFolioIngreso"
        Me.txtFolioIngreso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioIngreso.Size = New System.Drawing.Size(160, 20)
        Me.txtFolioIngreso.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtFolioIngreso, "Folio del Ingreso.")
        '
        'txtFolioRetiro
        '
        Me.txtFolioRetiro.AcceptsReturn = True
        Me.txtFolioRetiro.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtFolioRetiro.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioRetiro.Enabled = False
        Me.txtFolioRetiro.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioRetiro.Location = New System.Drawing.Point(94, 20)
        Me.txtFolioRetiro.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioRetiro.MaxLength = 17
        Me.txtFolioRetiro.Name = "txtFolioRetiro"
        Me.txtFolioRetiro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioRetiro.Size = New System.Drawing.Size(170, 20)
        Me.txtFolioRetiro.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txtFolioRetiro, "Folio del Retiro.")
        '
        'txtSucursal
        '
        Me.txtSucursal.AcceptsReturn = True
        Me.txtSucursal.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSucursal.Enabled = False
        Me.txtSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSucursal.Location = New System.Drawing.Point(78, 44)
        Me.txtSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSucursal.MaxLength = 40
        Me.txtSucursal.Name = "txtSucursal"
        Me.txtSucursal.ReadOnly = True
        Me.txtSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSucursal.Size = New System.Drawing.Size(458, 20)
        Me.txtSucursal.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.txtSucursal, "Sucursal que hace el Retiro.")
        '
        'txtEnvia
        '
        Me.txtEnvia.AcceptsReturn = True
        Me.txtEnvia.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtEnvia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEnvia.Enabled = False
        Me.txtEnvia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEnvia.Location = New System.Drawing.Point(78, 68)
        Me.txtEnvia.Margin = New System.Windows.Forms.Padding(2)
        Me.txtEnvia.MaxLength = 40
        Me.txtEnvia.Name = "txtEnvia"
        Me.txtEnvia.ReadOnly = True
        Me.txtEnvia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEnvia.Size = New System.Drawing.Size(458, 20)
        Me.txtEnvia.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.txtEnvia, "Usuario que Envia el Retiro.")
        '
        'txtPesos
        '
        Me.txtPesos.AcceptsReturn = True
        Me.txtPesos.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtPesos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPesos.Enabled = False
        Me.txtPesos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPesos.Location = New System.Drawing.Point(95, 13)
        Me.txtPesos.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPesos.MaxLength = 0
        Me.txtPesos.Name = "txtPesos"
        Me.txtPesos.ReadOnly = True
        Me.txtPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPesos.Size = New System.Drawing.Size(76, 20)
        Me.txtPesos.TabIndex = 15
        Me.txtPesos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtPesos, "Total en Pesos.")
        '
        'txtDolares
        '
        Me.txtDolares.AcceptsReturn = True
        Me.txtDolares.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtDolares.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDolares.Enabled = False
        Me.txtDolares.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDolares.Location = New System.Drawing.Point(267, 13)
        Me.txtDolares.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDolares.MaxLength = 0
        Me.txtDolares.Name = "txtDolares"
        Me.txtDolares.ReadOnly = True
        Me.txtDolares.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDolares.Size = New System.Drawing.Size(76, 20)
        Me.txtDolares.TabIndex = 16
        Me.txtDolares.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDolares, "Total en Dolares.")
        '
        'cmdDesglose
        '
        Me.cmdDesglose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDesglose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDesglose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDesglose.Location = New System.Drawing.Point(424, 95)
        Me.cmdDesglose.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdDesglose.Name = "cmdDesglose"
        Me.cmdDesglose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDesglose.Size = New System.Drawing.Size(112, 29)
        Me.cmdDesglose.TabIndex = 6
        Me.cmdDesglose.Text = "&Desglose"
        Me.ToolTip1.SetToolTip(Me.cmdDesglose, "Muestra la Ventana de Captura de Desglose del Deposito.")
        Me.cmdDesglose.UseVisualStyleBackColor = False
        '
        'cmdOrigenyAplicacion
        '
        Me.cmdOrigenyAplicacion.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOrigenyAplicacion.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOrigenyAplicacion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOrigenyAplicacion.Location = New System.Drawing.Point(424, 168)
        Me.cmdOrigenyAplicacion.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdOrigenyAplicacion.Name = "cmdOrigenyAplicacion"
        Me.cmdOrigenyAplicacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOrigenyAplicacion.Size = New System.Drawing.Size(110, 32)
        Me.cmdOrigenyAplicacion.TabIndex = 8
        Me.cmdOrigenyAplicacion.Text = "&Origen"
        Me.ToolTip1.SetToolTip(Me.cmdOrigenyAplicacion, "Muestra la Ventana de Captura de Origen.")
        Me.cmdOrigenyAplicacion.UseVisualStyleBackColor = False
        '
        'txtConcepto
        '
        Me.txtConcepto.AcceptsReturn = True
        Me.txtConcepto.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtConcepto.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConcepto.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConcepto.Location = New System.Drawing.Point(104, 70)
        Me.txtConcepto.Margin = New System.Windows.Forms.Padding(2)
        Me.txtConcepto.MaxLength = 100
        Me.txtConcepto.Name = "txtConcepto"
        Me.txtConcepto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConcepto.Size = New System.Drawing.Size(430, 20)
        Me.txtConcepto.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtConcepto, "Concepto de Deposito.")
        '
        'txtImporte
        '
        Me.txtImporte.AcceptsReturn = True
        Me.txtImporte.BackColor = System.Drawing.SystemColors.Window
        Me.txtImporte.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImporte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImporte.Location = New System.Drawing.Point(104, 94)
        Me.txtImporte.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImporte.MaxLength = 18
        Me.txtImporte.Name = "txtImporte"
        Me.txtImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImporte.Size = New System.Drawing.Size(127, 20)
        Me.txtImporte.TabIndex = 5
        Me.txtImporte.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtImporte, "Importe del Depósito.")
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.dtpFecha)
        Me.Frame4.Controls.Add(Me.txtFolioIngreso)
        Me.Frame4.Controls.Add(Me.Label3)
        Me.Frame4.Controls.Add(Me.Label1)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(7, 2)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(547, 40)
        Me.Frame4.TabIndex = 33
        Me.Frame4.TabStop = False
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(432, 12)
        Me.dtpFecha.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(102, 20)
        Me.dtpFecha.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(391, 15)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(37, 17)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "Fecha :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 34
        Me.Label1.Text = "Folio de Ingreso :"
        '
        'chkDepositoInterno
        '
        Me.chkDepositoInterno.BackColor = System.Drawing.SystemColors.Control
        Me.chkDepositoInterno.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDepositoInterno.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkDepositoInterno.Location = New System.Drawing.Point(15, 193)
        Me.chkDepositoInterno.Margin = New System.Windows.Forms.Padding(2)
        Me.chkDepositoInterno.Name = "chkDepositoInterno"
        Me.chkDepositoInterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDepositoInterno.Size = New System.Drawing.Size(150, 17)
        Me.chkDepositoInterno.TabIndex = 9
        Me.chkDepositoInterno.Text = "Deposito &Interno"
        Me.chkDepositoInterno.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.cmdOrigAplic_DepInt_Pes)
        Me.Frame2.Controls.Add(Me.cmdOrigAplic_DepInt_Dol)
        Me.Frame2.Controls.Add(Me.dtpFechaRetiro)
        Me.Frame2.Controls.Add(Me.txtFolioRetiro)
        Me.Frame2.Controls.Add(Me.txtSucursal)
        Me.Frame2.Controls.Add(Me.txtEnvia)
        Me.Frame2.Controls.Add(Me.flexDetalle)
        Me.Frame2.Controls.Add(Me.Label12)
        Me.Frame2.Controls.Add(Me.Label10)
        Me.Frame2.Controls.Add(Me.Frame3)
        Me.Frame2.Controls.Add(Me.Label9)
        Me.Frame2.Controls.Add(Me.Label8)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(7, 302)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(547, 377)
        Me.Frame2.TabIndex = 25
        Me.Frame2.TabStop = False
        '
        'cmdOrigAplic_DepInt_Pes
        '
        Me.cmdOrigAplic_DepInt_Pes.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOrigAplic_DepInt_Pes.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOrigAplic_DepInt_Pes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOrigAplic_DepInt_Pes.Location = New System.Drawing.Point(173, 319)
        Me.cmdOrigAplic_DepInt_Pes.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdOrigAplic_DepInt_Pes.Name = "cmdOrigAplic_DepInt_Pes"
        Me.cmdOrigAplic_DepInt_Pes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOrigAplic_DepInt_Pes.Size = New System.Drawing.Size(128, 41)
        Me.cmdOrigAplic_DepInt_Pes.TabIndex = 17
        Me.cmdOrigAplic_DepInt_Pes.Text = "Origen Cuenta &Pesos"
        Me.cmdOrigAplic_DepInt_Pes.UseVisualStyleBackColor = False
        '
        'cmdOrigAplic_DepInt_Dol
        '
        Me.cmdOrigAplic_DepInt_Dol.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOrigAplic_DepInt_Dol.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOrigAplic_DepInt_Dol.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOrigAplic_DepInt_Dol.Location = New System.Drawing.Point(345, 319)
        Me.cmdOrigAplic_DepInt_Dol.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdOrigAplic_DepInt_Dol.Name = "cmdOrigAplic_DepInt_Dol"
        Me.cmdOrigAplic_DepInt_Dol.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOrigAplic_DepInt_Dol.Size = New System.Drawing.Size(128, 41)
        Me.cmdOrigAplic_DepInt_Dol.TabIndex = 18
        Me.cmdOrigAplic_DepInt_Dol.Text = "Origen Cuenta &Dólares"
        Me.cmdOrigAplic_DepInt_Dol.UseVisualStyleBackColor = False
        '
        'dtpFechaRetiro
        '
        Me.dtpFechaRetiro.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaRetiro.Location = New System.Drawing.Point(434, 17)
        Me.dtpFechaRetiro.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaRetiro.Name = "dtpFechaRetiro"
        Me.dtpFechaRetiro.Size = New System.Drawing.Size(102, 20)
        Me.dtpFechaRetiro.TabIndex = 11
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(78, 92)
        Me.flexDetalle.Margin = New System.Windows.Forms.Padding(2)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(456, 163)
        Me.flexDetalle.TabIndex = 14
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(12, 24)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(98, 11)
        Me.Label12.TabIndex = 32
        Me.Label12.Text = "Folio de Retiro :"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(12, 49)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(67, 12)
        Me.Label10.TabIndex = 31
        Me.Label10.Text = "Sucursal :"
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.txtPesos)
        Me.Frame3.Controls.Add(Me.txtDolares)
        Me.Frame3.Controls.Add(Me.Label2)
        Me.Frame3.Controls.Add(Me.Label6)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(78, 274)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(456, 41)
        Me.Frame3.TabIndex = 26
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Totales"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(45, 16)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(46, 17)
        Me.Label2.TabIndex = 28
        Me.Label2.Text = "Pesos :"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(211, 19)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(52, 13)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "Dólares :"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(396, 20)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(37, 13)
        Me.Label9.TabIndex = 30
        Me.Label9.Text = "Fecha :"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(12, 72)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(61, 11)
        Me.Label8.TabIndex = 29
        Me.Label8.Text = "Envia :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cmdReferencias)
        Me.Frame1.Controls.Add(Me.chkDepositoInterno)
        Me.Frame1.Controls.Add(Me.txtConcepto)
        Me.Frame1.Controls.Add(Me.txtImporte)
        Me.Frame1.Controls.Add(Me.cmdDesglose)
        Me.Frame1.Controls.Add(Me.dbcBanco)
        Me.Frame1.Controls.Add(Me.dbcCuentaBancaria)
        Me.Frame1.Controls.Add(Me.lblCancelada)
        Me.Frame1.Controls.Add(Me.cmdOrigenyAplicacion)
        Me.Frame1.Controls.Add(Me.lblMoneda)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label7)
        Me.Frame1.Controls.Add(Me.Label11)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(7, 45)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(547, 214)
        Me.Frame1.TabIndex = 19
        Me.Frame1.TabStop = False
        '
        'dbcBanco
        '
        Me.dbcBanco.Location = New System.Drawing.Point(104, 20)
        Me.dbcBanco.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(193, 21)
        Me.dbcBanco.TabIndex = 2
        '
        'dbcCuentaBancaria
        '
        Me.dbcCuentaBancaria.Location = New System.Drawing.Point(104, 45)
        Me.dbcCuentaBancaria.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcCuentaBancaria.Name = "dbcCuentaBancaria"
        Me.dbcCuentaBancaria.Size = New System.Drawing.Size(193, 21)
        Me.dbcCuentaBancaria.TabIndex = 3
        '
        'lblCancelada
        '
        Me.lblCancelada.BackColor = System.Drawing.SystemColors.Control
        Me.lblCancelada.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCancelada.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblCancelada.Location = New System.Drawing.Point(12, 114)
        Me.lblCancelada.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCancelada.Name = "lblCancelada"
        Me.lblCancelada.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCancelada.Size = New System.Drawing.Size(295, 20)
        Me.lblCancelada.TabIndex = 36
        '
        'lblMoneda
        '
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMoneda.Location = New System.Drawing.Point(333, 45)
        Me.lblMoneda.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.Size = New System.Drawing.Size(68, 15)
        Me.lblMoneda.TabIndex = 24
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(12, 21)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(49, 17)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "Banco :"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(12, 47)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(142, 17)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Cuenta Bancaria :"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(12, 72)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(67, 17)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "Concepto :"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(12, 95)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(61, 17)
        Me.Label11.TabIndex = 20
        Me.Label11.Text = "Importe :"
        '
        'lblCodBanco
        '
        Me.lblCodBanco.BackColor = System.Drawing.SystemColors.Control
        Me.lblCodBanco.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCodBanco.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCodBanco.Location = New System.Drawing.Point(186, 208)
        Me.lblCodBanco.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCodBanco.Name = "lblCodBanco"
        Me.lblCodBanco.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCodBanco.Size = New System.Drawing.Size(31, 7)
        Me.lblCodBanco.TabIndex = 37
        Me.lblCodBanco.Visible = False
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackColor = System.Drawing.SystemColors.Control
        Me.btnLimpiar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLimpiar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLimpiar.Location = New System.Drawing.Point(252, 264)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLimpiar.Size = New System.Drawing.Size(109, 36)
        Me.btnLimpiar.TabIndex = 43
        Me.btnLimpiar.Text = "&Nuevo"
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.BackColor = System.Drawing.SystemColors.Control
        Me.btnBuscar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBuscar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBuscar.Location = New System.Drawing.Point(137, 264)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 42
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(22, 264)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 41
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoDiarioRegistrodeDepositos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(566, 305)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.lblCodBanco)
        Me.Controls.Add(Me.btnGuardar)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(191, 248)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoDiarioRegistrodeDepositos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Registro de Depósitos"
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub    'Variables

End Class