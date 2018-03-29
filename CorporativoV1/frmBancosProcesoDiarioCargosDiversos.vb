Option Strict Off
Option Explicit On
Imports ADODB
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBancosProcesoDiarioCargosDiversos
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             REGISTRO DE CARGOS DIVERSOS                                                                  *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      JUEVES 24 DE JULIO DE 2003                                                                   *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public bandera As Boolean = False

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdOrigenyAplicacion As System.Windows.Forms.Button
    Public WithEvents cmdDesglose As System.Windows.Forms.Button
    Public WithEvents txtConcepto As System.Windows.Forms.TextBox
    Public WithEvents txtReferencia As System.Windows.Forms.TextBox
    Public WithEvents txtImporte As System.Windows.Forms.TextBox
    Public WithEvents dbcBanco As System.Windows.Forms.ComboBox
    Public WithEvents dbcCuentaBancaria As System.Windows.Forms.ComboBox
    Public WithEvents lblCancelada As System.Windows.Forms.Label
    Public WithEvents lblMoneda As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFecha As System.Windows.Forms.DateTimePicker
    Public WithEvents txtFolioEgreso As System.Windows.Forms.TextBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox

    Public frmcargos2 As frmBancosProcesoDiarioOrigenyAplicacion = New frmBancosProcesoDiarioOrigenyAplicacion()
    Public WithEvents btnBuscar As Button
    Public frmDesgloseCargosDiversos2 As frmBancosProcesoDiarioDesglosedeDepositos = New frmBancosProcesoDiarioDesglosedeDepositos()


    'Variables
    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Dim mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Dim mblnSalir As Boolean 'Para Salir Con el Esc
    Dim FueraChange As Boolean
    Dim intCodBanco As Integer
    Dim tecla As Integer
    Dim sglTiempoCambio As Single 'Para Esperar un Tiempo

    Public ConsultaCargos As Boolean
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
            Case "TXTFOLIOEGRESO"
                strCaptionForm = "Consulta de Registro de Cargos Diversos"
                gStrSql = "SELECT FolioMovto AS FOLIO,Concepto AS CONCEPTO," & "FechaMovto AS FECHA,Importe AS IMPORTE FROM MovimientosBancarios " & "WHERE FolioMovto LIKE '" & txtFolioEgreso.Text & "%' AND Movimiento = '" & C_MOVCARGOS & "' AND TipoMovto = '" & C_TIPOMOVEGRESO & "' ORDER BY FechaMovto DESC ,FolioMovto DESC"
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
                Case "TXTFOLIOEGRESO"
                    'ConfiguraConsultas(FrmConsultas, 9000, RsGral, strTag, strCaptionForm)
                    .set_ColWidth(0, 0, 1400) 'Columna del Folio
                    .set_ColWidth(1, 0, 4000) 'Columna del Concepto del Movimiento
                    .set_ColWidth(2, 0, 1200) 'Columna de la Fecha del Movimiento
                    .set_ColWidth(3, 0, 1800) 'Columna del Importe del Movimiento
                    .set_ColAlignment(2, 4)
                    For I = 1 To FrmConsultas.Flexdet.Rows - 1
                        FrmConsultas.Flexdet.set_TextMatrix(I, 2, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 2), "dd/MMM/yyyy"))
                        FrmConsultas.Flexdet.set_TextMatrix(I, 3, VB6.Format(FrmConsultas.Flexdet.get_TextMatrix(I, 3), "###,##0.00"))
                    Next I
                    FrmConsultas.Top = VB6.TwipsToPixelsY(3500)
                    FrmConsultas.Left = VB6.TwipsToPixelsX(2970)
            End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function ChecaGridCargos() As Boolean
        Dim I As Integer
        ChecaGridCargos = False
        With frmDesgloseCargosDiversos2.flexDetalle
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                    ChecaGridCargos = True
                End If
            Next
        End With
    End Function

    Function Guardar() As Boolean
        Dim blnTransaccion As Boolean
        Dim I As Integer
        Dim Ejercicio As Integer
        Dim Periodo As String
        On Error GoTo Err_Renamed

        'Do While (VB.Timer() - sglTiempoCambio) <= 2.1
        'Loop
        'System.Windows.Forms.Application.DoEvents()

        If Not mblnNuevo Then
            Exit Function
        End If
        If ValidaDatos() = False Then
            Exit Function
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        'Generar Folio del Movimiento
        Ejercicio = CInt(Format(Year(CDate(dtpFecha.Value)), "0000"))
        Periodo = Format(Month(CDate(dtpFecha.Value)), "00")
        BuscaEjercicio(dtpFecha.Value)
        gStrSql = "SELECT Consecutivo FROM EjercicioPeriodo WHERE Ejercicio = " & Ejercicio & " AND " & "Periodo = '" & Periodo & "' AND Prefijo = '" & C_TIPOMOVEGRESO & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtFolioEgreso.Text = C_TIPOMOVEGRESO & Format(Year(CDate(dtpFecha.Value)), "0000") & Format(Month(CDate(dtpFecha.Value)), "00") & Format(VB.Day(CDate(dtpFecha.Value)), "00") & Format(CStr(RsGral.Fields("Consecutivo").Value + 1), "0000")
            ModStoredProcedures.PR_IMEEjercicioPeriodo(CStr(Ejercicio), Periodo, C_TIPOMOVEGRESO, CStr(RsGral.Fields("Consecutivo").Value + 1), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        'Guardar el Movimiento Bancario
        ModStoredProcedures.PR_IMEMovimientosBancarios(txtFolioEgreso.Text, Format(dtpFecha.Value, C_FORMATFECHAGUARDAR), C_MOVCARGOS, C_TIPOMOVEGRESO, C_NATURALEZACOMERCIAL, IIf(lblMoneda.Text = C_DESCPESOS, C_PESO, C_DOLAR), CStr(gcurCorpoTIPOCAMBIODOLAR), "", C_TIPOPAGOJOYERIA, CStr(intCodBanco), dbcCuentaBancaria.Text, "", txtConcepto.Text, "0", "", "0", "01/01/1900", "", txtImporte.Text, "V", "01/01/1900", "", CStr(0), "01/01/1900", C_MODULOBANCOS, txtReferencia.Text, "", C_INSERCION, CStr(0))
        Cmd.Execute()
        'Guardar los Movimientos de Origen y Aplicación
        If Not frmcargos2.GuardarMovimientosOrigenAplicacion("REGISTRO DE CARGOS") Then
            Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Exit Function
        End If
        'Guardar el Desglose del Deposite
        If Not frmDesgloseCargosDiversos2.GuardarMovimientosDepositos Then
            Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Exit Function
        End If
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("Los Datos se Han Guardado con Exito" & Chr(13) & "Se ha Generado el Folio de Egreso " & txtFolioEgreso.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Limpiar()
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Sub Limpiar()
        Nuevo()
        txtFolioEgreso.Text = ""
        txtFolioEgreso.Focus()
    End Sub

    Sub LlenaDatos()
        On Error GoTo Merr
        Dim I As Integer
        Dim Total As Decimal
        Dim RsAux As New ADODB.Recordset
        If Trim(txtFolioEgreso.Text) = "" Then
            Nuevo()
            Exit Sub
        End If
        gStrSql = "SELECT * FROM MovimientosBancarios MB,CatBancos CB WHERE MB.FolioMovto = '" & txtFolioEgreso.Text & "' AND MB.Movimiento = '" & C_MOVCARGOS & "' AND " & "MB.TipoMovto = '" & C_TIPOMOVEGRESO & "' AND CB.CodBanco = MB.CodBanco"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            gStrSql = "SELECT FolioMovto FROM MovimientosBancarios WHERE Referencia = '" & txtFolioEgreso.Text & "' AND Movimiento = '" & C_MOVCANCELACION & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsAux = Cmd.Execute
            If RsAux.RecordCount > 0 Then
                lblCancelada.Text = "Movimiento de Cancelación : " & RsAux.Fields("FolioMovto").Value
            End If
            txtFolioEgreso.Text = Trim(RsGral.Fields("FolioMovto").Value)
            dtpFecha.Value = Format(RsGral.Fields("FechaMovto").Value, C_FORMATFECHAMOSTRAR)
            FueraChange = True
            dbcBanco.Text = Trim(RsGral.Fields("DescBanco").Value)
            dbcCuentaBancaria.Text = Trim(RsGral.Fields("CtaBancaria").Value)
            FueraChange = False
            txtConcepto.Text = Trim(RsGral.Fields("Concepto").Value)
            txtReferencia.Text = Trim(RsGral.Fields("Referencia").Value)
            txtImporte.Text = Format(RsGral.Fields("importe").Value, "###,##0.00")
            If RsGral.Fields("Moneda").Value = C_PESO Then
                lblMoneda.Text = C_DESCPESOS
            ElseIf RsGral.Fields("Moneda").Value = C_DOLAR Then
                lblMoneda.Text = C_DESCDOLARES
            End If
            gStrSql = "SELECT * FROM MovimientosOrigenAplic MO,CatOrigenAplicRecursos CO,CatRubrosOrigenAplicRecursos CR " & "WHERE FolioMovto = '" & Trim(txtFolioEgreso.Text) & "' AND CO.CodOrigenAplicR = MO.CodOrigenAplicR AND CR.CodRubro = MO.CodRubro AND CO.CodOrigenAplicR = CR.CodOrigAplicR"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                With frmcargos2.flexDetalle
                    I = 1
                    .Row = 1
                    frmcargos2.lblTotal.Text = "0.00"
                    Do While Not RsGral.EOF
                        .set_TextMatrix(.Row, 0, Format(RsGral.Fields("CodOrigenAplicR").Value, "0000"))
                        .set_TextMatrix(.Row, 1, Trim(RsGral.Fields("DescOrigenAplicR").Value))
                        .set_TextMatrix(.Row, 2, Format(RsGral.Fields("CodRubro").Value, "000000"))
                        .set_TextMatrix(.Row, 3, Trim(RsGral.Fields("DescRubro").Value))
                        .set_TextMatrix(.Row, 4, Format(RsGral.Fields("importe").Value, "###,##0.00"))
                        With frmcargos2
                            .lblTotal.Text = CStr(CDec(Numerico(Format(.lblTotal.Text, "#####0.00"))) + CDbl(Format(RsGral.Fields("importe").Value, "###,##0.00")))
                        End With
                        If .Row = .Rows - 1 Then
                            .Rows = .Rows + 1
                        End If
                        .Row = .Row + 1
                        I = I + 1
                        RsGral.MoveNext()
                    Loop
                    frmcargos2.lblTotal.Text = Format(frmcargos2.lblTotal.Text, "###,##0.00")
                    frmcargos2.lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                    frmcargos2.Nuevo = True
                End With
            End If
            gStrSql = "SELECT * FROM MovimientosReferencias Where FolioMovto = '" & txtFolioEgreso.Text & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                With frmDesgloseCargosDiversos2.flexDetalle
                    I = 1
                    .Row = 1
                    frmDesgloseCargosDiversos2.lblTotal.Text = Format(RsGral.Fields("ImporteDeposito").Value, "###,##0.00")
                    frmDesgloseCargosDiversos2.lblTotal.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
                    Do While Not RsGral.EOF
                        .set_TextMatrix(.Row, 0, Trim(RsGral.Fields("ReferenciaBanco").Value))
                        .set_TextMatrix(.Row, 1, Format(RsGral.Fields("ImporteRef").Value, "###,##0.00"))
                        If .Row = .Rows - 1 Then
                            .Rows = .Rows + 1
                        End If
                        .Row = .Row + 1
                        I = I + 1
                        RsGral.MoveNext()
                    Loop
                End With
            End If
            mblnNuevo = False
            dtpFecha.Enabled = False
            ConsultaCargos = True
        Else
            MsgBox("Folio de Movimiento de Egreso no Existe ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Frame1.Enabled = True
            txtFolioEgreso.Focus()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnSalir = False
        FueraChange = False
        intCodBanco = 0
    End Sub

    Sub Nuevo()

        If (bandera = True) Then
            Exit Sub
        End If

        lblMoneda.Text = ""
        dtpFecha.Value = VB6.Format(Now, "dd/mmm/yyyy")
        lblCancelada.Text = ""
        dbcBanco.Text = ""
        'dbcBanco.RowSource = Nothing
        dbcCuentaBancaria.Text = ""
        'dbcCuentaBancaria.RowSource = Nothing
        txtConcepto.Text = ""
        txtReferencia.Text = ""
        txtImporte.Text = "0.00"
        Frame1.Enabled = True
        InicializaVariables()
        gblnSalir = True
        frmcargos2.Hide()
        gblnSalir = False
        frmDesgloseCargosDiversos2.Hide()
        'frmBancosProcesoDiarioOrigenyAplicacion.Nuevo = False
        ConsultaCargos = False
    End Sub

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Not BuscaUltimoCierre(dtpFecha.Value) Then
            Exit Function
        End If
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
        If Len(Trim(txtReferencia.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Referencia", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtReferencia.Focus()
            Exit Function
        End If
        If CDbl(Numerico(txtImporte.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Importe del Pago", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtImporte.Focus()
            Exit Function
        End If
        If ChecaGridCargos() Then
            If Numerico((frmDesgloseCargosDiversos2.lblImporte).Text) <> Numerico((frmDesgloseCargosDiversos2.lblTotal).Text) Then
                MsgBox("El Total del Desglose de los Cargos no es Igual al Importe del Cargo Total ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                cmdDesglose_Click(cmdDesglose, New System.EventArgs())
                Exit Function
            End If
        End If
        If Not ChecaGrid(frmcargos2) Then
            MsgBox("No se Han Capturado los Movimientos de Origen y Aplicación ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            cmdOrigenyAplicacion_Click(cmdOrigenyAplicacion, New System.EventArgs())
            Exit Function
        End If
        If Numerico((frmcargos2.lblImporte).Text) <> Numerico((frmcargos2.lblTotal).Text) Then
            MsgBox("El Total de los Movimientos de Origen y Aplicación no es Igual al Importe del Pago...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            cmdOrigenyAplicacion_Click(cmdOrigenyAplicacion, New System.EventArgs())
            Exit Function
        End If
        If Not ChecaSaldo(CInt(intCodBanco), Trim(dbcCuentaBancaria.Text), CDec(txtImporte.Text)) Then
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub cmdDesglose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDesglose.Click

        If Trim(dbcBanco.Text) <> "" And Trim(dbcCuentaBancaria.Text) <> "" Then
            If CDbl(Numerico(txtImporte.Text)) > 0 Then
                If Not mblnNuevo Then
                    frmDesgloseCargosDiversos2.cmdAceptar.TabIndex = 0
                    frmDesgloseCargosDiversos2.flexDetalle.TabIndex = 1
                    frmDesgloseCargosDiversos2.flexDetalle.Enabled = False
                Else
                    frmDesgloseCargosDiversos2.cmdAceptar.TabIndex = 1
                    frmDesgloseCargosDiversos2.flexDetalle.TabIndex = 0
                    frmDesgloseCargosDiversos2.cmdAceptar.Enabled = False
                End If
                frmDesgloseCargosDiversos2.Text = "Desglose de Cargos Diversos"
                frmDesgloseCargosDiversos2.Label1.Text = "Importe del Cargo : "
                frmDesgloseCargosDiversos2.Panel1.Text = "Desglose del Cargo"
                frmDesgloseCargosDiversos2.lblMoneda.Text = lblMoneda.Text
                frmDesgloseCargosDiversos2.lblImporte.Text = txtImporte.Text
                frmDesgloseCargosDiversos2.flexDetalle.Col = 0
                frmDesgloseCargosDiversos2.flexDetalle.Row = 1
                frmDesgloseCargosDiversos2.Tag = "frmDesgloseCargosDiversos2"
                frmDesgloseCargosDiversos2.ShowDialog()
            Else
                MsgBox("El Importe de los Cargos debe ser Mayor que Cero, Favor de Teclear un Importe ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
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

    Private Sub cmdOrigenyAplicacion_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOrigenyAplicacion.Click
        'frmcargos2.InitializeComponent()
        If Trim(dbcBanco.Text) <> "" And Trim(dbcCuentaBancaria.Text) <> "" Then
            If CDbl(Numerico(txtImporte.Text)) > 0 Then
                If frmcargos2.Nuevo Then
                    frmcargos2.cmdAceptar.TabIndex = 0
                    frmcargos2.flexDetalle.TabIndex = 1
                    frmcargos2.flexDetalle.Enabled = False
                Else
                    frmcargos2.flexDetalle.TabIndex = 0
                    frmcargos2.cmdAceptar.TabIndex = 1
                    frmcargos2.cmdAceptar.Enabled = False
                End If
                frmcargos2.Tag = "frmcargos2"
                frmcargos2.Text = "Aplicación de Recursos (Registro de Cargos Diversos)"
                frmcargos2.lblMoneda.Text = lblMoneda.Text
                frmcargos2.lblFechaMovimiento.Text = dtpFecha.Value
                frmcargos2.lblImporte.Text = txtImporte.Text
                frmcargos2.flexDetalle.Col = 0
                frmcargos2.flexDetalle.Row = 1
                frmcargos2.ShowDialog()
            Else
                MsgBox("El Importe de los Cargos debe ser Mayor que Cero, Favor de Teclear un Importe ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
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

    Private Sub dbcBanco_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcBanco" Then
        '    Exit Sub
        'End If
        dbcCuentaBancaria.Text = ""
        lblMoneda.Text = ""
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' AND ControlInterno = 0 ORDER BY DescBanco"
        DCChange(gStrSql, tecla)
        intCodBanco = 0
    End Sub

    Private Sub dbcBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Enter
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE ControlInterno = 0 ORDER BY DescBanco"
        DCGotFocus(gStrSql, dbcBanco)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcBanco_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcBanco.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            txtFolioEgreso.Focus()
        End If
    End Sub

    Private Sub dbcBanco_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcBanco.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.Leave
        gStrSql = "SELECT CodBanco,DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' AND ControlInterno = 0 ORDER BY DescBanco"
        DCLostFocus(dbcBanco, gStrSql, intCodBanco)
    End Sub

    Private Sub dbcCuentaBancaria_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcCuentaBancaria" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCChange(gStrSql, tecla)
        'intCodBanco = 0
    End Sub

    Private Sub dbcCuentaBancaria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Enter
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
        DCGotFocus(gStrSql, dbcCuentaBancaria)
        Pon_Tool()
        FueraChange = False
    End Sub

    Private Sub dbcCuentaBancaria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCuentaBancaria.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            dbcBanco.Focus()
        End If
    End Sub

    Private Sub dbcCuentaBancaria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcCuentaBancaria.KeyPress
        'eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcCuentaBancaria_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCuentaBancaria.KeyUp
        Dim Aux As String
        Aux = dbcCuentaBancaria.Text
        If dbcCuentaBancaria.SelectedItem <> 0 Then
            dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        End If
        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub dbcCuentaBancaria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCuentaBancaria.Leave
        On Error GoTo Err_Renamed
        gStrSql = "SELECT CodBanco,CtaBancaria FROM CatCuentasBancarias WHERE CtaBancaria LIKE '" & Trim(dbcCuentaBancaria.Text) & "%' AND CodBanco = " & intCodBanco & " ORDER BY CtaBancaria"
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
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcCuentaBancaria_MouseUp(ByVal eventSender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbcCuentaBancaria.MouseUp
        Dim Aux As String

        Aux = dbcCuentaBancaria.Text
        'If dbcCuentaBancaria.SelectedItem <> 0 Then
        dbcCuentaBancaria_Leave(dbcCuentaBancaria, New System.EventArgs())
        'End If

        dbcCuentaBancaria.Text = Aux
    End Sub

    Private Sub dtpFecha_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFecha.CursorChanged
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFecha_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFecha.Click
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub dtpFecha_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFecha.KeyPress
        sglTiempoCambio = VB.Timer()
    End Sub

    Private Sub frmBancosProcesoDiarioCargosDiversos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmBancosProcesoDiarioCargosDiversos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmBancosProcesoDiarioCargosDiversos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        'Dim KeyCode As Integer = eventArgs.KeyCode
        'Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Select Case KeyCode
        '    Case System.Windows.Forms.Keys.Return

        '        If Me.ActiveControl.Name = "txtFolioEgreso" Then
        '            If Len(Trim(txtFolioEgreso.Text)) = 13 And VB.Right(txtFolioEgreso.Text, 4) <> "0000" Then
        '                Frame1.Enabled = False
        '            End If
        '        End If
        '        ModEstandar.AvanzarTab(Me)
        '    Case System.Windows.Forms.Keys.Escape

        '        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtFolioEgreso" Then
        '            ModEstandar.RetrocederTab(Me)
        '        Else
        '            mblnSalir = True
        '            Me.Close()
        '        End If
        'End Select
    End Sub

    Private Sub frmBancosProcesoDiarioCargosDiversos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmBancosProcesoDiarioCargosDiversos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        frmcargos2.InitializeComponent()
        frmDesgloseCargosDiversos2.InitializeComponent()
        bandera = True
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        'gstrMovimiento = "S"
        InicializaVariables()
        Nuevo()
        BuscaEjercicio(dtpFecha.Value)
    End Sub

    Private Sub frmBancosProcesoDiarioCargosDiversos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmBancosProcesoDiarioCargosDiversos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        frmBancosProcesoDiarioRegistrodePagos = Nothing
        gblnSalir = True
        frmcargos2.Close()
        frmcargos2 = Nothing
        IsNothing(Me)
        Me.Hide()
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

    Private Sub txtFolioEgreso_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEgreso.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtFolioEgreso_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEgreso.Enter
        strControlActual = UCase("txtFolioEgreso")
        SelTextoTxt(txtFolioEgreso)
        Pon_Tool()
    End Sub

    Private Sub txtFolioEgreso_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFolioEgreso.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, C_TIPOMOVEGRESO)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If

        If KeyAscii = 13 Then
            txtFolioEgreso_Leave(New Object, New EventArgs)
        End If

    End Sub

    Private Sub txtFolioEgreso_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFolioEgreso.Leave
        'If System.Windows.Forms.Form.ActiveForm.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If (Me.ActiveControl.Name = "btnBuscar") Then
            Exit Sub
        End If

        If mblnCambiosEnCodigo = True And txtFolioEgreso.Text <> "" And VB.Right(txtFolioEgreso.Text, 4) <> "0000" Then
            LlenaDatos()
            frmcargos2.Hide()
            'frmBancosProcesoDiarioCargosDiversos.ZOrder
        End If

        If Trim(txtFolioEgreso.Text) = "" Then
            txtFolioEgreso.Text = C_TIPOMOVEGRESO & Format(Year(CDate(dtpFecha.Value)), "0000") & Format(Month(CDate(dtpFecha.Value)), "00") & Format(VB.Day(CDate(dtpFecha.Value)), "00") & "0000"
            Exit Sub
        End If

    End Sub

    Private Sub txtImporte_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.TextChanged
        If Trim(txtImporte.Text) = "" Then
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
        txtImporte.Text = VB6.Format(txtImporte.Text, "###,##0.00")
    End Sub

    Private Sub txtReferencia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferencia.Enter
        SelTextoTxt(txtReferencia)
        Pon_Tool()
    End Sub

    Private Sub txtReferencia_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtReferencia.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "/-_")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtConcepto = New System.Windows.Forms.TextBox()
        Me.txtReferencia = New System.Windows.Forms.TextBox()
        Me.txtImporte = New System.Windows.Forms.TextBox()
        Me.txtFolioEgreso = New System.Windows.Forms.TextBox()
        Me.cmdOrigenyAplicacion = New System.Windows.Forms.Button()
        Me.cmdDesglose = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me.dbcCuentaBancaria = New System.Windows.Forms.ComboBox()
        Me.lblCancelada = New System.Windows.Forms.Label()
        Me.lblMoneda = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtConcepto
        '
        Me.txtConcepto.AcceptsReturn = True
        Me.txtConcepto.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtConcepto.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtConcepto.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConcepto.Location = New System.Drawing.Point(74, 68)
        Me.txtConcepto.Margin = New System.Windows.Forms.Padding(2)
        Me.txtConcepto.MaxLength = 100
        Me.txtConcepto.Name = "txtConcepto"
        Me.txtConcepto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtConcepto.Size = New System.Drawing.Size(356, 20)
        Me.txtConcepto.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtConcepto, "Concepto de los Cargos.")
        '
        'txtReferencia
        '
        Me.txtReferencia.AcceptsReturn = True
        Me.txtReferencia.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtReferencia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReferencia.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtReferencia.Location = New System.Drawing.Point(74, 92)
        Me.txtReferencia.Margin = New System.Windows.Forms.Padding(2)
        Me.txtReferencia.MaxLength = 10
        Me.txtReferencia.Name = "txtReferencia"
        Me.txtReferencia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReferencia.Size = New System.Drawing.Size(113, 20)
        Me.txtReferencia.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtReferencia, "Referencia Bancaria.")
        '
        'txtImporte
        '
        Me.txtImporte.AcceptsReturn = True
        Me.txtImporte.BackColor = System.Drawing.SystemColors.Window
        Me.txtImporte.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImporte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImporte.Location = New System.Drawing.Point(74, 116)
        Me.txtImporte.Margin = New System.Windows.Forms.Padding(2)
        Me.txtImporte.MaxLength = 18
        Me.txtImporte.Name = "txtImporte"
        Me.txtImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImporte.Size = New System.Drawing.Size(92, 20)
        Me.txtImporte.TabIndex = 6
        Me.txtImporte.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtImporte, "Importe de los Cargos.")
        '
        'txtFolioEgreso
        '
        Me.txtFolioEgreso.AcceptsReturn = True
        Me.txtFolioEgreso.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolioEgreso.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFolioEgreso.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFolioEgreso.Location = New System.Drawing.Point(109, 14)
        Me.txtFolioEgreso.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFolioEgreso.MaxLength = 13
        Me.txtFolioEgreso.Name = "txtFolioEgreso"
        Me.txtFolioEgreso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFolioEgreso.Size = New System.Drawing.Size(117, 20)
        Me.txtFolioEgreso.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtFolioEgreso, "Folio del Egreso.")
        '
        'cmdOrigenyAplicacion
        '
        Me.cmdOrigenyAplicacion.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOrigenyAplicacion.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOrigenyAplicacion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOrigenyAplicacion.Location = New System.Drawing.Point(339, 142)
        Me.cmdOrigenyAplicacion.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdOrigenyAplicacion.Name = "cmdOrigenyAplicacion"
        Me.cmdOrigenyAplicacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOrigenyAplicacion.Size = New System.Drawing.Size(105, 32)
        Me.cmdOrigenyAplicacion.TabIndex = 8
        Me.cmdOrigenyAplicacion.Text = "A&plicación"
        Me.cmdOrigenyAplicacion.UseVisualStyleBackColor = False
        '
        'cmdDesglose
        '
        Me.cmdDesglose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDesglose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDesglose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDesglose.Location = New System.Drawing.Point(339, 106)
        Me.cmdDesglose.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdDesglose.Name = "cmdDesglose"
        Me.cmdDesglose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDesglose.Size = New System.Drawing.Size(105, 32)
        Me.cmdDesglose.TabIndex = 7
        Me.cmdDesglose.Text = "&Desglose"
        Me.cmdDesglose.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cmdOrigenyAplicacion)
        Me.Frame1.Controls.Add(Me.txtConcepto)
        Me.Frame1.Controls.Add(Me.cmdDesglose)
        Me.Frame1.Controls.Add(Me.txtReferencia)
        Me.Frame1.Controls.Add(Me.txtImporte)
        Me.Frame1.Controls.Add(Me.dbcBanco)
        Me.Frame1.Controls.Add(Me.dbcCuentaBancaria)
        Me.Frame1.Controls.Add(Me.lblCancelada)
        Me.Frame1.Controls.Add(Me.lblMoneda)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label6)
        Me.Frame1.Controls.Add(Me.Label7)
        Me.Frame1.Controls.Add(Me.Label11)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(7, 51)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(451, 206)
        Me.Frame1.TabIndex = 12
        Me.Frame1.TabStop = False
        '
        'dbcBanco
        '
        Me.dbcBanco.Location = New System.Drawing.Point(74, 18)
        Me.dbcBanco.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(173, 21)
        Me.dbcBanco.TabIndex = 2
        '
        'dbcCuentaBancaria
        '
        Me.dbcCuentaBancaria.Location = New System.Drawing.Point(74, 43)
        Me.dbcCuentaBancaria.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcCuentaBancaria.Name = "dbcCuentaBancaria"
        Me.dbcCuentaBancaria.Size = New System.Drawing.Size(173, 21)
        Me.dbcCuentaBancaria.TabIndex = 3
        '
        'lblCancelada
        '
        Me.lblCancelada.BackColor = System.Drawing.SystemColors.Control
        Me.lblCancelada.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCancelada.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblCancelada.Location = New System.Drawing.Point(17, 155)
        Me.lblCancelada.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCancelada.Name = "lblCancelada"
        Me.lblCancelada.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCancelada.Size = New System.Drawing.Size(235, 26)
        Me.lblCancelada.TabIndex = 19
        '
        'lblMoneda
        '
        Me.lblMoneda.BackColor = System.Drawing.SystemColors.Control
        Me.lblMoneda.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMoneda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblMoneda.Location = New System.Drawing.Point(274, 42)
        Me.lblMoneda.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMoneda.Size = New System.Drawing.Size(89, 15)
        Me.lblMoneda.TabIndex = 18
        Me.lblMoneda.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(12, 24)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(49, 15)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Banco :"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(12, 46)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(67, 13)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Cuenta Bancaria :"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(13, 71)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(67, 13)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Concepto :"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(4, 95)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(67, 11)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Referencia :"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(4, 115)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(57, 20)
        Me.Label11.TabIndex = 13
        Me.Label11.Text = "Importe :"
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.dtpFecha)
        Me.Frame4.Controls.Add(Me.txtFolioEgreso)
        Me.Frame4.Controls.Add(Me.Label3)
        Me.Frame4.Controls.Add(Me.Label1)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(7, 6)
        Me.Frame4.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(451, 40)
        Me.Frame4.TabIndex = 9
        Me.Frame4.TabStop = False
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFecha.Location = New System.Drawing.Point(324, 14)
        Me.dtpFecha.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(96, 20)
        Me.dtpFecha.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(274, 17)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(46, 14)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Fecha :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(17, 17)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(99, 14)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Folio de Egreso :"
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackColor = System.Drawing.SystemColors.Control
        Me.btnLimpiar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLimpiar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLimpiar.Location = New System.Drawing.Point(122, 275)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLimpiar.Size = New System.Drawing.Size(109, 36)
        Me.btnLimpiar.TabIndex = 46
        Me.btnLimpiar.Text = "&Nuevo"
        Me.btnLimpiar.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(7, 275)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 45
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.BackColor = System.Drawing.SystemColors.Control
        Me.btnBuscar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBuscar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBuscar.Location = New System.Drawing.Point(237, 275)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 47
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmBancosProcesoDiarioCargosDiversos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(462, 321)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(238, 212)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmBancosProcesoDiarioCargosDiversos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Cargos Diversos"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

End Class