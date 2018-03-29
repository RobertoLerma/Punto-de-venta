'**********************************************************************************************************************'
'*PROGRAMA: ABC ORIGEN Y APLICACIÓN DE RECURSOS JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018 
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility

Public Class frmAbcOrigenyAplicaciondeRecursos

    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             ABC A ORIGEN Y APLICACION DE RECURSOS                                                        *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      LUNES 12 DE MAYO DE 2003                                                                     *'
    '*FECHA DE TERMINACION : LUNES 12 DE MAYO DE 2003                                                                     *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Dim mblnNuevo As Boolean 'Para Saber si es Nuevo o es Consulta
    Dim mblnCambiosEnCodigo As Boolean 'Por si se Modifica el Código
    Dim mblnEntrada As Boolean 'Tipo de Aplicación Entrada
    Dim mblnSalida As Boolean 'Tipo de Aplicación Salida
    Dim mblnSalir As Boolean 'Para Salir de la Captura Sin Preguntar Por Cambios
    Public strControlActual As String 'Nombre del control actual
    Sub Buscar()
        On Error GoTo MErr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas


        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        'strTag = UCase(Me.Name) & "." & strControlActual 'El tag sera el nombre del formulario + el nombre del control
        strTag = UCase("FRMCORPOABCORIGENYAPLICACIONDERECURSOS" & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "TXTCODIGO"
                strCaptionForm = "Consulta de Tipos de Origen y Aplicación"
                gStrSql = "SELECT RIGHT('0000'+LTRIM(CodOrigenAplicR),4) AS CODIGO, DescOrigenAplicR AS DESCRIPCION FROM CatOrigenAplicRecursos Where Aplicacion  = '" & IIf(_optTipoAplicacion_0.Checked, "E", "S") & "' ORDER BY CodOrigenAplicR"
            Case "TXTDESCRIPCION"
                strCaptionForm = "Consulta de Tipos de Origen y Aplicación"
                gStrSql = "SELECT DescOrigenAplicR AS DESCRIPCION, RIGHT('0000'+LTRIM(CodOrigenAplicR),4) AS CODIGO FROM CatOrigenAplicRecursos WHERE DescOrigenAplicR LIKE '" & Trim(txtDescripcion.Text) & "%' And Aplicacion  = '" & IIf(_optTipoAplicacion_0.Checked, "E", "S") & "' ORDER BY DescOrigenAplicR"
            Case Else
                'Sale de este sub para QUE no ejecute ninguna opcion
                Exit Sub
        End Select
        strSQL = gStrSql 'Se hace uso de una variable temporal para el query
        'Si hubo cambios y es una modificacion entonces preguntara que si desea gravar los cambios
        If Cambios() = True And mblnNuevo = False Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se carguela consulta
                Case MsgBoxResult.Cancel 'Cancela la consulta
                    Exit Sub
            End Select
        End If
        gStrSql = strSQL 'Se regresa el valor de la variavle temporal a la variable original
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
        ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODIGO"
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                Case "TXTDESCRIPCION"
                    .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
                    .set_ColWidth(1, 0, 900) 'Columna del Código
            End Select
        End With
        FrmConsultas.ShowDialog()
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Eliminar()
        On Error GoTo MErr
        Dim blnTransaccion As Boolean
        gStrSql = "SELECT DescOrigenAplicR FROM CatOrigenAplicRecursos WHERE CodOrigenAplicR=" & Val(txtCodigo.Text)
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            MsgBox("Proporcione un código valido para eliminar.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
            Exit Sub
        End If
        'Preguntar si desea borrar el registro
        gStrSql = "SELECT * FROM CatRubrosOrigenAplicRecursos WHERE CodOrigAplicR =" & Val(txtCodigo.Text)
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            MsgBox("Existen Rubros Que Dependen de este Código de Origen y Aplicación" & Chr(13) & "Por lo que No se Puede Eliminar, Elimine Los Rubros que Dependen" & Chr(13) & "de este Codigo en el Catalogo de Rubros.", MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If
        Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, "")
            Case MsgBoxResult.No
                Exit Sub
            Case MsgBoxResult.Cancel
                Exit Sub
        End Select
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        ModStoredProcedures.PR_IMECatOrigenAplicRecursos(txtCodigo.Text, txtDescripcion.Text, IIf(_optTipoAplicacion_0.Checked, "E", "S"), C_ELIMINACION, CStr(0))
        Cmd.Execute()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        MsgBox("El Tipo de Origen y Aplicación a sido Eliminado Correctamente con el Código: " & txtCodigo.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
        Cnn.CommitTrans()
        blnTransaccion = False
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
        Guardar = False
        If Cambios() = False Then
            Limpiar()
            Exit Function
        End If
        'Valida si todos los datos han sido llenados para poder ser guardados
        If ValidaDatos() = False Then
            Exit Function
        End If
        If Val(txtCodigo.Text) = 0 Then
            mblnNuevo = True
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        If mblnNuevo Then
            ModStoredProcedures.PR_IMECatOrigenAplicRecursos(CStr(Val(txtCodigo.Text)), txtDescripcion.Text, IIf(_optTipoAplicacion_0.Checked, "E", "S"), C_INSERCION, CStr(0))
            Cmd.Execute()
            txtCodigo.Text = Format(Cmd.Parameters("ID").Value, "0000")
        Else
            ModStoredProcedures.PR_IMECatOrigenAplicRecursos(txtCodigo.Text, txtDescripcion.Text, IIf(_optTipoAplicacion_0.Checked, "E", "S"), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        If mblnNuevo Then
            MsgBox("El Tipo de Origen y Aplicación a sido Grabado Correctamente con el Código: " & txtCodigo.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
        End If
        Nuevo()
        InicializaVariables()
        Guardar = True
        Limpiar()
MErr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Sub LlenaDatos()
        Try
            'On Error GoTo MErr
            If Val(txtCodigo.Text) = 0 Then
                Nuevo()
                ModEstandar.AvanzarTab(Me)
                Exit Sub
            End If
            'txtCodigo.Text = Format(txtCodigo.Text, "0000")
            For i = 1 To 4 - (txtCodigo.TextLength)
                txtCodigo.Text = String.Concat("0" + txtCodigo.Text)
            Next i

            gStrSql = "SELECT * FROM CatOrigenAplicRecursos WHERE CodOrigenAplicR=" & Val(txtCodigo.Text)
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                txtDescripcion.Text = Trim(RsGral.Fields("DescOrigenAplicR").Value)
                txtDescripcion.Tag = Trim(RsGral.Fields("DescOrigenAplicR").Value)
                If RsGral.Fields("Aplicacion").Value = "E" Then
                    _optTipoAplicacion_0.Checked = True
                    mblnEntrada = True
                    mblnSalida = False
                ElseIf RsGral.Fields("Aplicacion").Value = "S" Then
                    _optTipoAplicacion_1.Checked = True
                    mblnEntrada = False
                    mblnSalida = True
                End If
            Else
                MsjNoExiste("El Origen y Aplicación de Recursos", gstrNombCortoEmpresa)
                Limpiar()
            End If

            txtCodigo.Enabled = False
            mblnCambiosEnCodigo = False
            mblnNuevo = False
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Sub Limpiar()
        On Error Resume Next
        'Valida si Hubo Cambios, Pregunta si Desea Guardar
        If Cambios() = True And mblnNuevo = False Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                Case MsgBoxResult.Cancel 'Cancela la accion de limpiar la pantalla
                    Exit Sub
            End Select
        End If
        txtCodigo.Text = ""
        Nuevo()
        InicializaVariables()
        txtCodigo.Focus()
    End Sub

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mblnEntrada = True
        mblnSalida = False
        mblnSalir = False
    End Sub

    Public Sub Nuevo()
        'On Error GoTo MErr
        Try
            txtCodigo.Enabled = True
            txtCodigo.Text = ""
            txtDescripcion.Text = ""
            txtDescripcion.Tag = ""
            _optTipoAplicacion_0.Checked = True
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()

        End Try
    End Sub

    Function Cambios() As Boolean
        Cambios = True
        If Trim(txtDescripcion.Text) <> txtDescripcion.Tag Then Exit Function
        If _optTipoAplicacion_0.Checked <> mblnEntrada Then Exit Function
        If _optTipoAplicacion_1.Checked <> mblnSalida Then Exit Function
        Cambios = False
    End Function

    Function ValidaDatos() As Boolean
        Dim lSql As String

        ValidaDatos = False
        If Len(Trim(txtDescripcion.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Descripción", MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Function
        End If
        lSql = "Select * From CatOrigenAplicRecursos Where DescOrigenAplicR = '" & Trim(txtDescripcion.Text) & "' "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lSql))
        RsGral = Cmd.Execute
        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount > 0 Then
            MsgBox("Ya existe un agrupador con la misma descripción  en el catálogo" & vbNewLine & "Es necesario que sea diferente" & vbNewLine & "Favor de verificar...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            Exit Function
        End If

        ValidaDatos = True
    End Function

    Private Sub frmCorpoABCOrigenyAplicaciondeRecursos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoABCOrigenyAplicaciondeRecursos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCorpoABCOrigenyAplicaciondeRecursos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtCodigo" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmCorpoABCOrigenyAplicaciondeRecursos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoABCOrigenyAplicaciondeRecursos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InicializaVariables()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ModEstandar.CentrarForma(Me)
        Icono(Me, MDIMenuPrincipalCorpo)
    End Sub

    Private Sub frmCorpoABCOrigenyAplicaciondeRecursos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        'ModEstandar.RestaurarForma(Me, False)
        ''Si se cierra el formulario y existio algun cambio en el registro se
        ''informa al usuario del cabio y si desea guardar el registro, ya sea
        ''que sea nuevo o un registro modificado
        'If Not mblnSalir Then
        '    If Cambios() = True And mblnNuevo = False Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
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

    Private Sub frmCorpoABCOrigenyAplicaciondeRecursos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
    End Sub

    'Private Sub optTipoAplicacion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTipoAplicacion.Enter
    '    Dim Index As Integer = optTipoAplicacion.GetIndex(eventSender)
    '    Select Case Index
    '        Case 0
    '            Pon_Tool()
    '        Case 1
    '            Pon_Tool()
    '    End Select
    'End Sub

    Private Sub txtCodigo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodigo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.Enter
        strControlActual = UCase("txtCodigo")
        SelTextoTxt(txtCodigo)
        Pon_Tool()
    End Sub

    Private Sub txtCodigo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodigo.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000

        If (KeyCode = Keys.Enter) Then
            LlenaDatos()
        End If

        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            KeyCode = 0
        Else
            'Si la tecla presionada fue Delete y Hay cambios, pregunta si se desea guardar
            If Cambios() = True And KeyCode = System.Windows.Forms.Keys.Delete Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes 'Guardar el registro
                        If Guardar() = False Then
                            KeyCode = 0
                            Exit Sub
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se borre el contenido del text
                        Nuevo()
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        txtCodigo.Focus()
                        KeyCode = 0
                        Exit Sub
                End Select
            End If
        End If
    End Sub

    Private Sub txtCodigo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'Si la tecla presionada no es numero regresa un 0
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back Then
            KeyAscii = 0
        Else
            'Pregunta solo si existieron cambios
            If Cambios() = True And mblnNuevo = False Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes 'Guardar el registro
                        If Guardar() = False Then
                            KeyAscii = 0
                            GoTo EventExitSub
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        txtCodigo.Focus()
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

    Private Sub txtCodigo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigo.Leave

        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If

        'txtCodigo.Text = Format(txtCodigo.Text, "0000")
        If mblnCambiosEnCodigo = True Then 'si hubo cambios en el codigo hace la consulta
            LlenaDatos()
        End If
    End Sub

    Private Sub txtDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Enter
        strControlActual = UCase("txtDescripcion")
        SelTextoTxt(txtDescripcion)
        Pon_Tool()
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