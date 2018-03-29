Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmPVConfigCaja
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             CONFIGURACION DE CAJA                                                                        *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      VIERNES 16 DE MAYO DE 2003                                                                   *'
    '*FECHA DE TERMINACION : VIERNES 16 DE MAYO DE 2003                                                                   *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents btnLiberarCaja As System.Windows.Forms.Button
    Public WithEvents txtNumCaja As System.Windows.Forms.TextBox
    Public WithEvents txtDescripcion As System.Windows.Forms.TextBox
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents flexDetalle As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents _lblDisponible_0 As System.Windows.Forms.Label
    Public WithEvents _lblFecha_3 As System.Windows.Forms.Label
    Public WithEvents _lblFechaUltimoCorte_2 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblDisponible As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblFecha As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblFechaUltimoCorte As Microsoft.VisualBasic.Compatibility.VB6.LabelArray



    Public mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Public mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Public mblnSALIR As Boolean 'Para Salir de la Captura Sin Preguntar Por Cambios
    Public FueraChange As Boolean
    Public tecla As Integer
    Public intCodSucursal As Integer
    Public WithEvents Panel3 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public RsAux As ADODB.Recordset

    Dim strControlActual As String 'Nombre del control actual

    Sub Buscar()
        On Error GoTo MErr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas


        'strControlActual = UCase(ControlActivo.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control


        Select Case strControlActual
            Case "TXTNUMCAJA"
                strCaptionForm = "Consulta de Cajas"
                gStrSql = "SELECT RIGHT('00'+LTRIM(CodCaja),2) AS CODIGO, DescCaja AS DESCRIPCION FROM CatCajas  " & "Where CodAlmacen = " & intCodSucursal & " ORDER BY CodCaja"
            Case "TXTDESCRIPCION"
                strCaptionForm = "Consulta de Cajas"
                gStrSql = "SELECT DescCaja AS DESCRIPCION, RIGHT('00'+LTRIM(CodCaja),2) AS CODIGO FROM CatCajas WHERE DescCaja LIKE '" & Trim(txtDescripcion.Text) & "%'  AND CodAlmacen = " & intCodSucursal & " ORDER BY DescCaja"
            Case Else
                'Sale de este sub para QUE no ejecute ninguna opcion
                Exit Sub
        End Select
        If Trim(dbcSucursales.Text) = "" Then
            MsgBox("Es necesario que Seleccione una Sucursal.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            dbcSucursales.Focus()
            Exit Sub
        End If
        strSQL = gStrSql 'Se hace uso de una variable temporal para el query
        'Si hubo cambios y es una modificacion entonces preguntara que si desea gravar los cambios
        If Cambios() = True And mblnNuevo = False Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrCorpoNOMBREEMPRESA)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se carguela consulta
                Case MsgBoxResult.Cancel 'Cancela la consulta
                    Exit Sub
            End Select
        End If
        gStrSql = strSQL 'Se regresa el valor de la variable temporal a la variable original
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsgBox(C_msgSINDATOS & vbNewLine & "Verifique Por Favor....", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            Exit Sub
        End If
        'Carga el formulario de consulta
        'Load(FrmConsultas)
        ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)
        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTNUMCAJA"
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                Case "TXTDESCRIPCION"
                    .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
                    .set_ColWidth(1, 0, 900) 'Columna del Código
            End Select
        End With
        FrmConsultas.Show()
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Eliminar()
        On Error GoTo MErr
        Dim blnTransaccion As Boolean
        gStrSql = "SELECT * FROM CatCajas WHERE CodCaja=" & Val(txtNumCaja.Text) & " and CodAlmacen = " & intCodSucursal
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            MsgBox("Proporcione un Numero de Caja Valido para Eliminar.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
            Exit Sub
        End If
        'Preguntar si desea borrar el registro
        Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "")
            Case MsgBoxResult.No
                Exit Sub
            Case MsgBoxResult.Cancel
                Exit Sub
        End Select
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        ModStoredProcedures.PR_IMECatCajas(txtNumCaja.Text, CStr(intCodSucursal), txtDescripcion.Text, "01/01/1900", C_ELIMINACION, CStr(0))
        Cmd.Execute()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        Nuevo()
        Limpiar()
MErr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Sub

    Function Guardar() As Boolean
        On Error GoTo MErr
        Dim blnTransaccion As Boolean
        Dim I As Integer
        '    Dim rsAlmacenes As Recordset
        Guardar = False
        If Cambios() = False Then
            Limpiar()
            Exit Function
        End If
        'Valida si todos los datos han sido llenados para poder ser guardados
        If ValidaDatos() = False Then
            Exit Function
        End If
        If Val(txtNumCaja.Text) = 0 Then
            mblnNuevo = True
        End If
        Cnn.BeginTrans()
        blnTransaccion = True
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        If mblnNuevo Then
            ModStoredProcedures.PR_IMECatCajas(CStr(Val(txtNumCaja.Text)), CStr(intCodSucursal), txtDescripcion.Text, "01/01/1900", C_INSERCION, CStr(0))
            Cmd.Execute()
            txtNumCaja.Text = Format(Cmd.Parameters("ID").Value, "00")
        Else
            ModStoredProcedures.PR_IMECatCajas(txtNumCaja.Text, CStr(intCodSucursal), txtDescripcion.Text, "01/01/1900", C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        Cnn.CommitTrans()
        Cursor = System.Windows.Forms.Cursors.Default
        blnTransaccion = False
        If mblnNuevo Then
            MsgBox("La Configuración de la Caja Ha sido Grabada Correctamente" & Chr(13) & "Con el Numero de Caja: " & txtNumCaja.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrCorpoNOMBREEMPRESA)
        End If
        Nuevo()
        InicializaVariables()
        Guardar = True
        Limpiar()
MErr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Sub LlenaDatos()
        On Error GoTo MErr
        If Val(txtNumCaja.Text) = 0 Then
            Nuevo()
            '        ModEstandar.AvanzarTab Me
            Exit Sub
        End If

        For I = 0 To 2 - txtNumCaja.TextLength
            txtNumCaja.Text = String.Concat("0" + txtNumCaja.Text)
        Next I

        gStrSql = "SELECT * FROM CatCajas WHERE CodCaja=" & Val(txtNumCaja.Text) & " and CodAlmacen = " & intCodSucursal
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            txtDescripcion.Text = Trim(RsGral.Fields("DescCaja").Value)
            txtDescripcion.Tag = Trim(RsGral.Fields("DescCaja").Value)
            With flexDetalle
                .Row = 1
                .Col = 0
                .Text = "Salida de Mercancia"
                .Col = 1
                .Text = RsGral.Fields("PrefijoSalidasMcia").Value
                .Col = 2
                .Text = Format(RsGral.Fields("ConsecSalidasMcia").Value, "0000")
                .Row = 2
                .Col = 0
                .Text = "Vales de Devolución"
                .Col = 1
                .Text = RsGral.Fields("PrefijoValesDevolucion").Value
                .Col = 2
                .Text = Format(RsGral.Fields("ConsecValesDevolucion").Value, "0000")
                .Row = 3
                .Col = 0
                .Text = "Apartados"
                .Col = 1
                .Text = RsGral.Fields("PrefijoApartados").Value
                .Col = 2
                .Text = Format(RsGral.Fields("ConsecApartados").Value, "0000")
                .Row = 4
                .Col = 0
                .Text = "Retiros"
                .Col = 1
                .Text = RsGral.Fields("PrefijoRetiros").Value
                .Col = 2
                .Text = Format(RsGral.Fields("ConsecRetiros").Value, "0000")
                .Row = 1
                .Col = 0
            End With
            flexDetalle.Enabled = True
            flexDetalle.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
            If RsGral.Fields("Disponible").Value = True Then
                _lblDisponible_0.Text = "Caja Disponible"
                btnLiberarCaja.Enabled = False
            Else
                _lblDisponible_0.Text = "Caja No Disponible"
                btnLiberarCaja.Enabled = True
            End If
            _lblDisponible_0.Visible = True
            If Year(RsGral.Fields("FechaUltimoCorte").Value) <> 1900 Then
                _lblFechaUltimoCorte_2.Visible = True
                _lblFecha_3.Text = Format(RsGral.Fields("FechaUltimoCorte").Value, C_FORMATFECHAMOSTRAR)
                _lblFecha_3.Visible = True
            End If
        Else
            MsjNoExiste("La Caja", gstrCorpoNOMBREEMPRESA)
            Nuevo()
            txtNumCaja.Text = ""
            txtNumCaja.Focus()
        End If
        mblnCambiosEnCodigo = False
        mblnNuevo = False
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function Cambios() As Boolean
        Cambios = True
        If Trim(txtDescripcion.Text) <> txtDescripcion.Tag Then Exit Function
        Cambios = False
    End Function

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        If Len(Trim(dbcSucursales.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Sucursal", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            dbcSucursales.Focus()
            Exit Function
        End If
        If Len(Trim(txtDescripcion.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Descripción", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            txtDescripcion.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Sub Limpiar()
        On Error Resume Next
        'Valida si Hubo Cambios, Pregunta si Desea Guardar
        If Cambios() = True And mblnNuevo = False Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrCorpoNOMBREEMPRESA)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                Case MsgBoxResult.Cancel 'Cancela la accion de limpiar la pantalla
                    Exit Sub
            End Select
        End If
        FueraChange = True
        Nuevo()
        txtNumCaja.Text = ""
        btnLiberarCaja.Enabled = False
        dbcSucursales.Text = ""
        dbcSucursales.Tag = ""
        dbcSucursales.Focus()
    End Sub

    Sub Nuevo()
        On Error GoTo MErr
        InicializaVariables()
        txtNumCaja.Text = ""
        txtNumCaja.Tag = ""
        txtDescripcion.Text = ""
        txtDescripcion.Tag = ""
        dbcSucursales.Text = ""
        InicializaVariables()
        flexDetalle.Clear()
        flexDetalle.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
        Encabezado()
        _lblDisponible_0.Visible = False
        _lblFechaUltimoCorte_2.Visible = False
        _lblFecha_3.Visible = False
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnSALIR = False
        FueraChange = False
    End Sub

    Sub Encabezado()
        With flexDetalle
            .Row = 0
            .Col = 0
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(.Col, 0, 3000)
            .Text = "Concepto"
            .Col = 1
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(.Col, 0, 700)
            .Text = "Prefijo"
            .Col = 2
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(.Col, 0, 1290)
            .Text = "Consecutivo"
            .Row = 1
            .Col = 0
        End With
    End Sub

    Private Sub btnLiberarCaja_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnLiberarCaja.Click
        On Error GoTo Errores
        Dim blnTransaction As Boolean

        gStrSql = "SELECT * FROM CatCajas WHERE CodCaja=" & CInt(txtNumCaja.Text) & " and CodAlmacen = " & intCodSucursal & " And Disponible = 0 "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            Cnn.BeginTrans()
            blnTransaction = True
            ModStoredProcedures.PR_IMECatCajas(txtNumCaja.Text, CStr(intCodSucursal), "", "01/01/1900", C_MODIFICACION, CStr(3))
            Cmd.Execute()
            Cnn.CommitTrans()
            blnTransaction = False

            MsgBox("Sucursal " & intCodSucursal & vbNewLine & "Caja " & CInt(txtNumCaja.Text) & vbNewLine & vbNewLine & "Liberada", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            Nuevo()
            InicializaVariables()
            Limpiar()
        Else
            MsgBox("Esta caja ya esta liberada" & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
        End If

Errores:
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged

        'If FueraChange = True Then Exit Sub
        ''If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        ''    Exit Sub
        ''End If
        'gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        'DCChange(gStrSql, tecla)
        'intCodSucursal = 0
        'mblnNuevo = True
        'Nuevo()
        'txtNumCaja.Text = ""
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        '    If Screen.ActiveForm.ActiveControl.Name <> dbcSucursales.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen Where TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        'Pregunta solo si existieron cambios
        If Cambios() = True Then 'And KeyCode = vbKeyDelete Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrCorpoNOMBREEMPRESA)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        eventSender.KeyCode = 0
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se borre el contenido del text
                Case MsgBoxResult.Cancel 'Cancela la captura
                    txtNumCaja.Focus()
                    eventSender.KeyCode = 0
                    Exit Sub
            End Select
        End If

        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSALIR = True
            Me.Close()
        End If
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
    End Sub

    Private Sub flexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Enter
        flexDetalle.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
    End Sub

    Private Sub flexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexDetalle.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Return Then
            txtNumCaja.Focus()
        End If
    End Sub

    Private Sub flexDetalle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexDetalle.Leave
        flexDetalle.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub

    Private Sub frmPVConfigCaja_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmPVConfigCaja_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmPVConfigCaja_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
                    ModEstandar.RetrocederTab(Me)
                    '            Else
                    '                mblnSalir = True
                    '                Unload Me
                End If
        End Select
    End Sub

    Private Sub frmPVConfigCaja_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmPVConfigCaja_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        InicializaVariables()
        Encabezado()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ModEstandar.Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        '    gStrSql = "SELECT  CodAlmacen, DescAlmacen FROM CatAlmacen  " & _
        ''            "Where codAlmacen in(" & gintCodAlmacen & ")"
        '    ModEstandar.BorraCmd
        '    Cmd.CommandText = "dbo.Up_Select_Datos"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
        '    Set RsAux = Cmd.Execute
        '
        '    If RsAux.RecordCount > 0 Then
        '        dbcSucursales = Trim(RsAux!DescAlmacen)
        '        dbcSucursales_LostFocus
        '    End If
    End Sub

    Private Sub frmPVConfigCaja_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSALIR Then
        '    If Cambios() = True And mblnNuevo = False Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrCorpoNOMBREEMPRESA)
        '            Case MsgBoxResult.Yes 'Guardar el registro
        '                If Guardar() = False Then
        '                    Cancel = 1
        '                End If
        '            Case MsgBoxResult.No 'No hace nada y permite el cierre del formulario
        '            Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
        '                Cancel = 1
        '        End Select
        '    End If
        'Else
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0
        '        Case MsgBoxResult.No
        '            mblnSALIR = False
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmPVConfigCaja_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
    End Sub

    Private Sub txtDescripcion_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDescripcion.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return And mblnNuevo Then
            txtNumCaja.Focus()
        End If
    End Sub

    Private Sub txtNumCaja_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumCaja.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtNumCaja_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumCaja.Enter
        SelTextoTxt(txtNumCaja)
        Pon_Tool()
        strControlActual = "TXTNUMCAJA"
    End Sub

    Private Sub txtNumCaja_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtNumCaja.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Pregunta solo si existieron cambios
        If Cambios() = True And KeyCode = System.Windows.Forms.Keys.Delete Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrCorpoNOMBREEMPRESA)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        KeyCode = 0
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se borre el contenido del text
                Case MsgBoxResult.Cancel 'Cancela la captura
                    txtNumCaja.Focus()
                    KeyCode = 0
                    Exit Sub
            End Select
        End If
    End Sub

    Private Sub txtNumCaja_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNumCaja.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back Then
            KeyAscii = 0
        Else
            'Pregunta solo si existieron cambios
            If Cambios() = True And mblnNuevo = False Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrCorpoNOMBREEMPRESA)
                    Case MsgBoxResult.Yes 'Guardar el registro
                        If Guardar() = False Then
                            KeyAscii = 0
                            GoTo EventExitSub
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        txtNumCaja.Focus()
                        KeyAscii = 0
                        GoTo EventExitSub
                End Select
            End If
        End If
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtNumCaja_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumCaja.Leave
        'If System.Windows.Forms.Form.ActiveForm.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If mblnNuevo Then
            If txtNumCaja.Text = "" Then
                txtNumCaja.Text = "00"
            End If
        End If
        If mblnCambiosEnCodigo = True Then 'si hubo cambios en el codigo hace la consulta
            LlenaDatos()
        End If
    End Sub

    Private Sub txtDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Enter
        SelTextoTxt(txtDescripcion)
        Pon_Tool()
        strControlActual = "TXTDESCRIPCION"
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPVConfigCaja))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtNumCaja = New System.Windows.Forms.TextBox()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.btnLiberarCaja = New System.Windows.Forms.Button()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.flexDetalle = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._lblDisponible_0 = New System.Windows.Forms.Label()
        Me._lblFecha_3 = New System.Windows.Forms.Label()
        Me._lblFechaUltimoCorte_2 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblDisponible = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblFecha = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblFechaUltimoCorte = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDisponible, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFecha, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFechaUltimoCorte, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtNumCaja
        '
        Me.txtNumCaja.AcceptsReturn = True
        Me.txtNumCaja.BackColor = System.Drawing.Color.White
        Me.txtNumCaja.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumCaja.ForeColor = System.Drawing.Color.Black
        Me.txtNumCaja.Location = New System.Drawing.Point(100, 59)
        Me.txtNumCaja.MaxLength = 2
        Me.txtNumCaja.Name = "txtNumCaja"
        Me.txtNumCaja.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumCaja.Size = New System.Drawing.Size(41, 20)
        Me.txtNumCaja.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtNumCaja, "Número de Caja")
        '
        'txtDescripcion
        '
        Me.txtDescripcion.AcceptsReturn = True
        Me.txtDescripcion.BackColor = System.Drawing.Color.White
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescripcion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtDescripcion.Location = New System.Drawing.Point(101, 83)
        Me.txtDescripcion.MaxLength = 30
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(240, 20)
        Me.txtDescripcion.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtDescripcion, "Breve Descripción de la Caja")
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.Color.Black
        Me._Label1_0.Location = New System.Drawing.Point(128, 27)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(60, 17)
        Me._Label1_0.TabIndex = 1
        Me._Label1_0.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_0, "Nombre de la Farmacia Actual")
        '
        'btnLiberarCaja
        '
        Me.btnLiberarCaja.BackColor = System.Drawing.SystemColors.Control
        Me.btnLiberarCaja.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnLiberarCaja.Enabled = False
        Me.btnLiberarCaja.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLiberarCaja.Location = New System.Drawing.Point(273, 294)
        Me.btnLiberarCaja.Name = "btnLiberarCaja"
        Me.btnLiberarCaja.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnLiberarCaja.Size = New System.Drawing.Size(88, 37)
        Me.btnLiberarCaja.TabIndex = 13
        Me.btnLiberarCaja.Text = "Liberar Caj&a"
        Me.btnLiberarCaja.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtNumCaja)
        Me.Frame2.Controls.Add(Me.txtDescripcion)
        Me.Frame2.Controls.Add(Me.dbcSucursales)
        Me.Frame2.Controls.Add(Me._Label1_1)
        Me.Frame2.Controls.Add(Me._Label1_5)
        Me.Frame2.Controls.Add(Me._Label1_0)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(8, 0)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(353, 115)
        Me.Frame2.TabIndex = 0
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Datos Generales de la Caja "
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(186, 23)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(155, 21)
        Me.dbcSucursales.TabIndex = 2
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.Color.Black
        Me._Label1_1.Location = New System.Drawing.Point(10, 64)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(93, 21)
        Me._Label1_1.TabIndex = 3
        Me._Label1_1.Text = "Número de Caja :"
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.ForeColor = System.Drawing.Color.Black
        Me._Label1_5.Location = New System.Drawing.Point(10, 89)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(77, 15)
        Me._Label1_5.TabIndex = 6
        Me._Label1_5.Text = "Descripción :"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.flexDetalle)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 120)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(353, 132)
        Me.Frame1.TabIndex = 7
        Me.Frame1.TabStop = False
        '
        'flexDetalle
        '
        Me.flexDetalle.DataSource = Nothing
        Me.flexDetalle.Location = New System.Drawing.Point(7, 15)
        Me.flexDetalle.Name = "flexDetalle"
        Me.flexDetalle.OcxState = CType(resources.GetObject("flexDetalle.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexDetalle.Size = New System.Drawing.Size(338, 109)
        Me.flexDetalle.TabIndex = 8
        '
        '_lblDisponible_0
        '
        Me._lblDisponible_0.BackColor = System.Drawing.SystemColors.Info
        Me._lblDisponible_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDisponible_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblDisponible_0.Location = New System.Drawing.Point(10, 264)
        Me._lblDisponible_0.Name = "_lblDisponible_0"
        Me._lblDisponible_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDisponible_0.Size = New System.Drawing.Size(109, 17)
        Me._lblDisponible_0.TabIndex = 9
        Me._lblDisponible_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._lblDisponible_0.Visible = False
        '
        '_lblFecha_3
        '
        Me._lblFecha_3.BackColor = System.Drawing.SystemColors.Info
        Me._lblFecha_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFecha_3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblFecha_3.Location = New System.Drawing.Point(264, 265)
        Me._lblFecha_3.Name = "_lblFecha_3"
        Me._lblFecha_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFecha_3.Size = New System.Drawing.Size(93, 17)
        Me._lblFecha_3.TabIndex = 11
        Me._lblFecha_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._lblFecha_3.Visible = False
        '
        '_lblFechaUltimoCorte_2
        '
        Me._lblFechaUltimoCorte_2.BackColor = System.Drawing.SystemColors.Info
        Me._lblFechaUltimoCorte_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFechaUltimoCorte_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblFechaUltimoCorte_2.Location = New System.Drawing.Point(188, 265)
        Me._lblFechaUltimoCorte_2.Name = "_lblFechaUltimoCorte_2"
        Me._lblFechaUltimoCorte_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFechaUltimoCorte_2.Size = New System.Drawing.Size(87, 17)
        Me._lblFechaUltimoCorte_2.TabIndex = 10
        Me._lblFechaUltimoCorte_2.Text = "Ultimo Corte :"
        Me._lblFechaUltimoCorte_2.Visible = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Info
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.Location = New System.Drawing.Point(8, 258)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(353, 27)
        Me.Label2.TabIndex = 12
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(8, 337)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(353, 74)
        Me.Panel3.TabIndex = 69
        '
        'btnSalir
        '
        Me.btnSalir.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.salir
        Me.btnSalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSalir.Location = New System.Drawing.Point(208, 14)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(50, 42)
        Me.btnSalir.TabIndex = 70
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.buscar
        Me.btnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBuscar.Location = New System.Drawing.Point(160, 14)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(50, 42)
        Me.btnBuscar.TabIndex = 67
        Me.btnBuscar.Text = " "
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.grabar
        Me.btnGuardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnGuardar.Location = New System.Drawing.Point(11, 14)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(50, 42)
        Me.btnGuardar.TabIndex = 64
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.nuevo
        Me.btnLimpiar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnLimpiar.Location = New System.Drawing.Point(110, 14)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(50, 42)
        Me.btnLimpiar.TabIndex = 66
        Me.btnLimpiar.Text = " "
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.Eliminar
        Me.btnEliminar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnEliminar.Location = New System.Drawing.Point(61, 14)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(50, 42)
        Me.btnEliminar.TabIndex = 65
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'frmPVConfigCaja
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(369, 420)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.btnLiberarCaja)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me._lblDisponible_0)
        Me.Controls.Add(Me._lblFecha_3)
        Me.Controls.Add(Me._lblFechaUltimoCorte_2)
        Me.Controls.Add(Me.Label2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(369, 198)
        Me.MaximizeBox = False
        Me.Name = "frmPVConfigCaja"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Configuración de Cajas"
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        CType(Me.flexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDisponible, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFecha, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFechaUltimoCorte, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        Eliminar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub
End Class