'**********************************************************************************************************************'
'*PROGRAMA: VENDEDORES JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class FrmCorpoAbcVendedores
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    'Programa: ABC de Vendedores
    'Autor: Rosaura Torres López
    'Fecha de Creación: 12/Mayo/2003
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtComentarios As System.Windows.Forms.RichTextBox
    Public WithEvents txtReferencias As System.Windows.Forms.RichTextBox
    Public WithEvents txtDomicilio As System.Windows.Forms.RichTextBox
    Public WithEvents txtDescripcion As System.Windows.Forms.TextBox
    Public WithEvents txtCodVendedor As System.Windows.Forms.TextBox
    Public WithEvents txtTelefono As System.Windows.Forms.TextBox
    Public WithEvents _lblVendedores_6 As System.Windows.Forms.Label
    Public WithEvents _lblVendedores_5 As System.Windows.Forms.Label
    Public WithEvents _lblVendedores_1 As System.Windows.Forms.Label
    Public WithEvents _lblVendedores_2 As System.Windows.Forms.Label
    Public WithEvents _lblVendedores_0 As System.Windows.Forms.Label
    Public WithEvents _lblVendedores_3 As System.Windows.Forms.Label
    Public WithEvents _lblVendedores_4 As System.Windows.Forms.Label
    Public WithEvents fraGeneral As System.Windows.Forms.GroupBox
    Public WithEvents lblVendedores As Microsoft.VisualBasic.Compatibility.VB6.LabelArray



    'Estas Variables se declaran de manera local, para evitar conflictos al estar usando
    'la misma variable en distintos modulos, que pueden afectar el valor que hayan tomado en un form. distinto al actual
    Dim mblnNuevo As Boolean 'Para Controlar si un registro es Nuevo o se trata de una consulta
    Dim mblnCambiosEnCodigo As Boolean 'Para Controlar si se han efectuado cambios en el código
    Dim mblnTransaccion As Boolean ' Para el control de transacciones
    Dim mblnSALIR As Boolean 'se usa para cuando un usuario presiona escape en el primer control de formulario
    Public WithEvents Panel1 As Panel
    Public WithEvents Panel3 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents dtpFechaAlta As DateTimePicker
    Public strControlActual As String 'Nombre del control actual

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        dtpFechaAlta.Enabled = False
    End Sub

    Sub Buscar()
        'Esta Función se utilizará para Buscar un dato especifico de un formulario, la cual podrá realizarse por campo Codigo o Campo Descripción,
        ' y se Activará presionando la tecla F3.
        On Error GoTo MErr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas


        'strControlActual = UCase(ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        Select Case strControlActual
            Case "TXTCODVENDEDOR"
                strCaptionForm = "Consulta de Vendedores"
                gStrSql = "SELECT RIGHT('000'+LTRIM(CodVendedor),3) AS CODIGO,DescVendedor AS NOMBRE FROM CatVendedores  ORDER BY CodVendedor"
            Case "TXTDESCRIPCION"
                strCaptionForm = "Consulta de Vendedores"
                gStrSql = "SELECT DescVendedor AS NOMBRE, RIGHT('000'+LTRIM(CodVendedor),3) AS CODIGO FROM CatVendedores WHERE DescVendedor LIKE '" & Trim(txtDescripcion.Text) & "%' ORDER BY DescVendedor"
            Case Else
                'Sale de este sub para ke no ejecute ninguna opcion
                Exit Sub
        End Select

        strSQL = gStrSql 'Se hace uso de una variable temporal para el query

        'Si hubo cambios y es una modificacion entonces preguntara que si desea grabar los cambios
        If Cambios() = True And mblnNuevo = False Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Guardar() = False Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se cargue la consulta
                Case MsgBoxResult.Cancel 'Cancela la consulta
                    Exit Sub
            End Select
        End If

        gStrSql = strSQL 'Se regresa el valor de la variavle temporal a la variable original

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsgBox(C_msgSINDATOS & vbNewLine & "Verifique por favor...", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta 
        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODVENDEDOR"
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
        'On Error GoTo MErr
        Try
            'Screen.MousePointer = vbHourglass Esto se manejará hasta antes de iniciar la transacción
            gStrSql = "SELECT DescVendedor FROM CatVendedores WHERE CodVendedor=" & Val(txtCodVendedor.Text)

            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute

            If RsGral.RecordCount = 0 Then
                MsgBox("Proporcione un Código valido para eliminar.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Mensaje")
                RsGral.Close()
                Exit Sub
            End If

            'Preguntar si desea borrar el registro
            Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, "Mensaje")
                Case MsgBoxResult.No
                    Exit Sub
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select

            Dim fechaAlta = AgregarHoraAFecha(dtpFechaAlta.Value)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Cnn.BeginTrans()
            ModStoredProcedures.PR_IMECatVendedores(Trim(txtCodVendedor.Text), Trim(txtDescripcion.Text), Trim(txtDomicilio.Rtf), Trim(txtTelefono.Text), Trim(txtReferencias.Rtf), Trim(txtComentarios.Rtf), fechaAlta, C_ELIMINACION, CStr(0))
            Cmd.Execute()
            Cnn.CommitTrans()
            MsgBox("El Vendedor ha sido eliminado correctamente con el Código: " & txtCodVendedor.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
            Nuevo()
            Limpiar()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Sub
            'MErr:
        Catch ex As Exception
            Cnn.RollbackTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If Err.Number <> 0 Then ModEstandar.MostrarError()
            End
        End Try
    End Sub


    Public Function Guardar() As Boolean
        'On Error GoTo MErr
        Try
            'Si no se realizaron cambios, entonces no se guardara nada
            'Si el Código  es "", entonces no se validará nada, solamente se saldrá del proc.
            If Cambios() = False And Trim(txtCodVendedor.Text) = "" Then
                Limpiar()
                Exit Function
            End If

            'Validar si todos los datos fueron proporcionados para ser guardados
            If ValidaDatos() = False Then
                Exit Function
            End If

            If Val(txtCodVendedor.Text) = 0 Then
                mblnNuevo = True
            End If

            'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Cnn.BeginTrans()
            mblnTransaccion = True

            Dim fechaAlta = AgregarHoraAFecha(dtpFechaAlta.Value)

            If mblnNuevo = True Then 'Se realizará una insercion
                ModStoredProcedures.PR_IMECatVendedores(Trim(txtCodVendedor.Text), Trim(txtDescripcion.Text), Trim(txtDomicilio.Text), Trim(txtTelefono.Text), Trim(txtReferencias.Text), Trim(txtComentarios.Text), fechaAlta, C_INSERCION, CStr(0))
                Cmd.Execute()
                txtCodVendedor.Text = Format(Cmd.Parameters("ID").Value, "000")
            Else ' Se realizará una Modificación
                ModStoredProcedures.PR_IMECatVendedores(Trim(txtCodVendedor.Text), Trim(txtDescripcion.Text), Trim(txtDomicilio.Text), Trim(txtTelefono.Text), Trim(txtReferencias.Text), Trim(txtComentarios.Text), fechaAlta, C_MODIFICACION, CStr(0))
                Cmd.Execute()
            End If
            Cnn.CommitTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            'Por cuestiones de estetica, el puntero del mouse se cambia inmediatamente despues de cerrar la transacción
            '    mblnTransaccion = False
            If mblnNuevo Then
                MsgBox("El Vendedor ha sido grabado correctamente con el Código: " & txtCodVendedor.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
            Else
                MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
            End If
            'Dejar el Procedimiento Nuevo, sirve para que al usar limpiar,. no pregunte si se desea guardar cambios en el codigo
            Nuevo()
            Guardar = True
            'Limpiar()
            '    Screen.MousePointer = vbDefault
            Exit Function
            'MErr:
        Catch ex As Exception
            '    If mblnTransaccion = True Then
            Cnn.RollbackTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
        Return True
    End Function



    Sub Nuevo()
        'Se deben Limpiar todos los controles del formulario con excepcion del Control de la Llavve principal
        'On Error GoTo MErr
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            txtCodVendedor.Enabled = True
            txtCodVendedor.Text = ""
            txtDescripcion.Text = ""
            txtDescripcion.Tag = ""
            txtDomicilio.Text = ""
            txtDomicilio.Tag = ""
            txtTelefono.Text = ""
            txtTelefono.Tag = ""
            txtReferencias.Text = ""
            txtReferencias.Tag = ""
            txtComentarios.Text = ""
            txtComentarios.Tag = ""
            'dtpFechaAlta.Value = CDate(Today)
            'dtpFechaAlta.Tag = CDate(Today)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Sub
            'MErr:
        Catch ex As Exception
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Sub LlenaDatos()
        'On Error GoTo MErr
        Try
            '    Screen.MousePointer = vbHourglass
            If Val(txtCodVendedor.Text) = 0 Then
                Nuevo()
                'Lo quité porque cuando daba un enter en el codigo y este era cero, saltaba dos controles
                '        ModEstandar.AvanzarTab Me
                Exit Sub
            End If
            'txtCodVendedor.Text = VB6.Format(txtCodVendedor.Text, "000")
            gStrSql = "SELECT CodVendedor,DescVendedor,Domicilio,Telefono,Referencias,Comentarios,dbo.FormatFecha(FechaAlta,5)as FechaAlta FROM CatVendedores WHERE CodVendedor= " & Val(txtCodVendedor.Text)
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute

            If RsGral.RecordCount > 0 Then

                txtDescripcion.Text = Trim(RsGral.Fields("DescVendedor").Value.ToString())
                txtDescripcion.Tag = Trim(RsGral.Fields("DescVendedor").Value.ToString())

                txtDomicilio.Text = Trim(RsGral.Fields("Domicilio").Value.ToString())
                txtDomicilio.Tag = Trim(RsGral.Fields("Domicilio").Value.ToString())

                txtTelefono.Text = Trim(RsGral.Fields("Telefono").Value.ToString())
                txtTelefono.Tag = Trim(RsGral.Fields("Telefono").Value.ToString())

                txtReferencias.Text = Trim(RsGral.Fields("Referencias").Value.ToString())
                txtReferencias.Tag = Trim(RsGral.Fields("Referencias").Value.ToString())

                txtComentarios.Text = Trim(RsGral.Fields("Comentarios").Value.ToString())
                txtComentarios.Tag = Trim(RsGral.Fields("Comentarios").Value.ToString())

                'dtpFechaAlta.Value = Trim(RsGral.Fields("FechaAlta").Value)
                'dtpFechaAlta.Tag = Trim(RsGral.Fields("FechaAlta").Value)
            Else
                MsjNoExiste("El Vendedor", gstrNombCortoEmpresa)
                Limpiar()
            End If

            txtCodVendedor.Enabled = False
            mblnCambiosEnCodigo = False
            mblnNuevo = False
            '    Screen.MousePointer = vbDefault
            Exit Sub
            'MErr:      
        Catch ex As Exception
            'Screen.MousePointer = vbDefault

            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Sub Limpiar()
        'Esta función Limpia todos los controles del formulario.
        'Si hubo Cambios, Pregunta si desea guardarlos.
        'On Error GoTo MErr
        Try
            '    Screen.MousePointer = vbHourglass
            If Cambios() = True And mblnNuevo = False Then 'Si hubo Cambios y se trata de una consulta se hace lo siguiente
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes 'Permite Guardar los cambios en el registro
                        If Guardar() = False Then
                            Exit Sub
                        End If
                    Case MsgBoxResult.No
                    'No hace nada y permite que se limpie la pantalla
                    Case MsgBoxResult.Cancel 'Cancela la acción de limpiar la pantalla
                        Exit Sub
                End Select
            End If

            'txtCodVendedor.Enabled = True
            txtCodVendedor.Text = ""
            Nuevo()
            mblnNuevo = True
            mblnCambiosEnCodigo = False
            txtCodVendedor.Focus()
            '    Screen.MousePointer = vbDefault
            Exit Sub
            'MErr:
        Catch ex As Exception
            '    Screen.MousePointer = vbDefault
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Function Cambios() As Object
        'Esta Función validará si se han efectuado cambios en los controles.
        'lo cual es útil para la funcion de guardar. Se inicializa con True, y si se validan todos los campos y no se ha
        'salido del proc. entonces la variable adquiere el valor de False
        'se validan todos los controles existentes, excepto el de la Clave Principal
        'On Error GoTo MErr
        Try
            '    Screen.MousePointer = vbHourglass
            Cambios = True
            If Trim(txtDescripcion.Text) <> Trim(txtDescripcion.Tag) Then Exit Function
            If Trim(txtDomicilio.Rtf) <> Trim(txtDomicilio.Tag) Then Exit Function
            If Trim(txtTelefono.Text) <> Trim(txtTelefono.Tag) Then Exit Function
            If Trim(txtReferencias.Rtf) <> Trim(txtReferencias.Tag) Then Exit Function
            If Trim(txtComentarios.Rtf) <> Trim(txtComentarios.Tag) Then Exit Function
            'If dtpFechaAlta.Value <> CDate(dtpFechaAlta.Tag) Then Exit Function
            Cambios = False
            '    Screen.MousePointer = vbDefault
            Exit Function
            'MErr:
        Catch ex As Exception
            '    Screen.MousePointer = vbDefault
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Function ValidaDatos() As Object
        'Esta Función Valida que todos los datos en el Formulario se introduzcan, para poder realizar la Alta del registro
        'On Error GoTo MErr
        Try
            '    Screen.MousePointer = vbHourglass
            '    ValidaDatos = False No es necesario especificarlo, ya que la funcion se inicializa con falso
            If Len(Trim(txtDescripcion.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Nombre", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtDescripcion.Focus()
                Exit Function
            End If
            If Len(Trim(txtDomicilio.Rtf)) = 0 Then
                MsgBox(C_msgFALTADATO & "Domicilio", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtDomicilio.Focus()
                Exit Function
            End If
            If Len(Trim(txtTelefono.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Teléfono", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtTelefono.Focus()
                Exit Function
            End If
            '    If Len(Trim(ModEstandar.QuitaEnter(txtReferencias))) = 0 Then
            '        MsgBox C_msgFALTADATO & "Referencias", vbexclamation, gstrNombCortoEmpresa
            '        Me.txtReferencias.SetFocus
            '        Exit Function
            '    End If
            '    If Len(Trim(ModEstandar.QuitaEnter(txtComentarios))) = 0 Then
            '        MsgBox C_msgFALTADATO & "Comentarios", vbexclamation, gstrNombCortoEmpresa
            '        Me.txtComentarios.SetFocus
            '        Exit Function
            '    End If
            ValidaDatos = True
            '    Screen.MousePointer = vbDefault
            Exit Function
            'MErr:
        Catch ex As Exception
            '    Screen.MousePointer = vbDefault
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Private Sub FrmCorpoAbcVendedores_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub FrmCorpoAbcVendedores_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub FrmCorpoAbcVendedores_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        'Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
        Nuevo()

    End Sub

    Private Sub FrmCorpoAbcVendedores_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        ' En este evento del formulario se valida la tecla presionada.
        ' Si es Enter se simula un tab(Avanza al siguiente control)
        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                '            txtDomicilio = ModEstandar.QuitaEnter(txtDomicilio)
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub FrmCorpoAbcVendedores_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub FrmCorpoAbcVendedores_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSALIR Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        '    'Si se cierra el formulario y existio algun cambio en el registro se
        '    'informa al usuario del cabio y si desea guardar el registro, ya sea
        '    'que sea nuevo o un registro modificado
        '    If Cambios() = True Then 'And mblnNuevo = False Then
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
        'Else 'Se quiere salir con escape
        '    mblnSALIR = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub FrmCorpoAbcVendedores_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ' ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ' ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
    End Sub

    Private Sub txtCodVendedor_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodVendedor.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodVendedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodVendedor.Enter
        strControlActual = UCase("txtCodVendedor")
        SelTextoTxt(txtCodVendedor)
        Pon_Tool()
    End Sub

    Private Sub txtCodVendedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodVendedor.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSALIR = True
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
                        txtCodVendedor.Focus()
                        KeyCode = 0
                        Exit Sub
                End Select
            End If
        End If
    End Sub

    Private Sub txtCodVendedor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodVendedor.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'Si la tecla presionada no es numero regresa un 0
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back Then
            KeyAscii = 0
        Else
            'Pregunta solo si existieron cambios
            If Cambios() = True And mblnNuevo = False Then
                'Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                'Case MsgBoxResult.Yes 'Guardar el registro
                If Guardar() = False Then
                    KeyAscii = 0
                    GoTo EventExitSub
                End If
                'Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
                'Case MsgBoxResult.Cancel 'Cancela la captura
                txtCodVendedor.Focus()
                KeyAscii = 0
                GoTo EventExitSub
                'End Select
            End If
        End If
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodVendedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodVendedor.Leave
        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If Val(Trim(txtCodVendedor.Text)) = 0 Then txtCodVendedor.Text = "000"
        If mblnCambiosEnCodigo = True And CDbl(Numerico(txtCodVendedor.Text)) <> 0 Then 'si hubo cambios en el codigo hace la consulta para llenar los datos
            LlenaDatos()
        End If
    End Sub

    Private Sub txtComentarios_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtComentarios.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtComentarios_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtComentarios.Enter
        txtComentarios.SelectionStart = 0
        Pon_Tool()
    End Sub

    Private Sub txtComentarios_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtComentarios.Leave
        If ActiveControl.Text <> Me.Text Then Exit Sub
        '''txtComentarios = ModEstandar.QuitaEnter(txtComentarios)
         'txtComentarios.Text = Trim(txtComentarios.Rtf)
    End Sub

    Private Sub txtDescripcion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Enter
        strControlActual = UCase("txtDescripcion")
        SelTextoTxt(txtDescripcion)
        Pon_Tool()
    End Sub


    Private Sub txtDomicilio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDomicilio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.Enter
        txtDomicilio.SelectionStart = 0
        Pon_Tool()
    End Sub

    Private Sub txtDomicilio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.Leave
        If ActiveControl.Text <> Me.Text Then Exit Sub
        '''txtDomicilio = ModEstandar.QuitaEnter(txtDomicilio)
        'txtDomicilio.Text = (txtDomicilio).Rtf
    End Sub

    Private Sub txtReferencias_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferencias.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtReferencias_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferencias.Enter
        txtReferencias.SelectionStart = 0
        Pon_Tool()
    End Sub

    Private Sub txtReferencias_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferencias.Leave
        If ActiveControl.Text <> Me.Text Then Exit Sub
        '''txtReferencias = ModEstandar.QuitaEnter(txtReferencias)
            'txtReferencias.Text = Trim(txtReferencias.Rtf)
    End Sub

    Private Sub txtTelefono_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTelefono.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtTelefono_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTelefono.Enter
        '    SelTextoTxt txtTelefono
        txtTelefono.SelectionStart = Len(Trim(txtTelefono.Text))
        Pon_Tool()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me.txtCodVendedor = New System.Windows.Forms.TextBox()
        Me.txtTelefono = New System.Windows.Forms.TextBox()
        Me.fraGeneral = New System.Windows.Forms.GroupBox()
        Me.dtpFechaAlta = New System.Windows.Forms.DateTimePicker()
        Me.txtComentarios = New System.Windows.Forms.RichTextBox()
        Me.txtReferencias = New System.Windows.Forms.RichTextBox()
        Me.txtDomicilio = New System.Windows.Forms.RichTextBox()
        Me._lblVendedores_6 = New System.Windows.Forms.Label()
        Me._lblVendedores_5 = New System.Windows.Forms.Label()
        Me._lblVendedores_1 = New System.Windows.Forms.Label()
        Me._lblVendedores_2 = New System.Windows.Forms.Label()
        Me._lblVendedores_0 = New System.Windows.Forms.Label()
        Me._lblVendedores_3 = New System.Windows.Forms.Label()
        Me._lblVendedores_4 = New System.Windows.Forms.Label()
        Me.lblVendedores = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.fraGeneral.SuspendLayout()
        CType(Me.lblVendedores, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtDescripcion
        '
        Me.txtDescripcion.AcceptsReturn = True
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescripcion.Location = New System.Drawing.Point(60, 42)
        Me.txtDescripcion.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescripcion.MaxLength = 40
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(284, 20)
        Me.txtDescripcion.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtDescripcion, "Nombre del Vendedor")
        '
        'txtCodVendedor
        '
        Me.txtCodVendedor.AcceptsReturn = True
        Me.txtCodVendedor.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodVendedor.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodVendedor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodVendedor.Location = New System.Drawing.Point(61, 16)
        Me.txtCodVendedor.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodVendedor.MaxLength = 3
        Me.txtCodVendedor.Name = "txtCodVendedor"
        Me.txtCodVendedor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodVendedor.Size = New System.Drawing.Size(44, 20)
        Me.txtCodVendedor.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtCodVendedor, "Código del Vendedor")
        '
        'txtTelefono
        '
        Me.txtTelefono.AcceptsReturn = True
        Me.txtTelefono.BackColor = System.Drawing.SystemColors.Window
        Me.txtTelefono.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTelefono.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTelefono.Location = New System.Drawing.Point(60, 129)
        Me.txtTelefono.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTelefono.MaxLength = 50
        Me.txtTelefono.Name = "txtTelefono"
        Me.txtTelefono.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTelefono.Size = New System.Drawing.Size(284, 20)
        Me.txtTelefono.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtTelefono, "Teléfono del Vendedor")
        '
        'fraGeneral
        '
        Me.fraGeneral.BackColor = System.Drawing.Color.Silver
        Me.fraGeneral.Controls.Add(Me.dtpFechaAlta)
        Me.fraGeneral.Controls.Add(Me.txtComentarios)
        Me.fraGeneral.Controls.Add(Me.txtReferencias)
        Me.fraGeneral.Controls.Add(Me.txtDomicilio)
        Me.fraGeneral.Controls.Add(Me.txtDescripcion)
        Me.fraGeneral.Controls.Add(Me.txtCodVendedor)
        Me.fraGeneral.Controls.Add(Me.txtTelefono)
        Me.fraGeneral.Controls.Add(Me._lblVendedores_6)
        Me.fraGeneral.Controls.Add(Me._lblVendedores_5)
        Me.fraGeneral.Controls.Add(Me._lblVendedores_1)
        Me.fraGeneral.Controls.Add(Me._lblVendedores_2)
        Me.fraGeneral.Controls.Add(Me._lblVendedores_0)
        Me.fraGeneral.Controls.Add(Me._lblVendedores_3)
        Me.fraGeneral.Controls.Add(Me._lblVendedores_4)
        Me.fraGeneral.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGeneral.Location = New System.Drawing.Point(12, 9)
        Me.fraGeneral.Margin = New System.Windows.Forms.Padding(2)
        Me.fraGeneral.Name = "fraGeneral"
        Me.fraGeneral.Padding = New System.Windows.Forms.Padding(2)
        Me.fraGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGeneral.Size = New System.Drawing.Size(360, 353)
        Me.fraGeneral.TabIndex = 0
        Me.fraGeneral.TabStop = False
        '
        'dtpFechaAlta
        '
        Me.dtpFechaAlta.Enabled = False
        Me.dtpFechaAlta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaAlta.Location = New System.Drawing.Point(239, 16)
        Me.dtpFechaAlta.Name = "dtpFechaAlta"
        Me.dtpFechaAlta.Size = New System.Drawing.Size(106, 20)
        Me.dtpFechaAlta.TabIndex = 14
        '
        'txtComentarios
        '
        Me.txtComentarios.Location = New System.Drawing.Point(11, 276)
        Me.txtComentarios.Margin = New System.Windows.Forms.Padding(2)
        Me.txtComentarios.Name = "txtComentarios"
        Me.txtComentarios.Size = New System.Drawing.Size(327, 60)
        Me.txtComentarios.TabIndex = 12
        Me.txtComentarios.Text = ""
        '
        'txtReferencias
        '
        Me.txtReferencias.Location = New System.Drawing.Point(11, 182)
        Me.txtReferencias.Margin = New System.Windows.Forms.Padding(2)
        Me.txtReferencias.Name = "txtReferencias"
        Me.txtReferencias.Size = New System.Drawing.Size(328, 59)
        Me.txtReferencias.TabIndex = 10
        Me.txtReferencias.Text = ""
        '
        'txtDomicilio
        '
        Me.txtDomicilio.Location = New System.Drawing.Point(60, 68)
        Me.txtDomicilio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDomicilio.Name = "txtDomicilio"
        Me.txtDomicilio.Size = New System.Drawing.Size(285, 53)
        Me.txtDomicilio.TabIndex = 6
        Me.txtDomicilio.Text = ""
        '
        '_lblVendedores_6
        '
        Me._lblVendedores_6.BackColor = System.Drawing.Color.Silver
        Me._lblVendedores_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVendedores_6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblVendedores_6.Location = New System.Drawing.Point(14, 253)
        Me._lblVendedores_6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVendedores_6.Name = "_lblVendedores_6"
        Me._lblVendedores_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVendedores_6.Size = New System.Drawing.Size(67, 11)
        Me._lblVendedores_6.TabIndex = 11
        Me._lblVendedores_6.Text = "Comentarios"
        '
        '_lblVendedores_5
        '
        Me._lblVendedores_5.BackColor = System.Drawing.Color.Silver
        Me._lblVendedores_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVendedores_5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblVendedores_5.Location = New System.Drawing.Point(13, 158)
        Me._lblVendedores_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVendedores_5.Name = "_lblVendedores_5"
        Me._lblVendedores_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVendedores_5.Size = New System.Drawing.Size(68, 11)
        Me._lblVendedores_5.TabIndex = 9
        Me._lblVendedores_5.Text = "Referencias"
        '
        '_lblVendedores_1
        '
        Me._lblVendedores_1.AutoSize = True
        Me._lblVendedores_1.BackColor = System.Drawing.Color.Silver
        Me._lblVendedores_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVendedores_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblVendedores_1.Location = New System.Drawing.Point(170, 19)
        Me._lblVendedores_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVendedores_1.Name = "_lblVendedores_1"
        Me._lblVendedores_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVendedores_1.Size = New System.Drawing.Size(64, 13)
        Me._lblVendedores_1.TabIndex = 13
        Me._lblVendedores_1.Text = "Fecha Alta :"
        '
        '_lblVendedores_2
        '
        Me._lblVendedores_2.AutoSize = True
        Me._lblVendedores_2.BackColor = System.Drawing.Color.Silver
        Me._lblVendedores_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVendedores_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVendedores_2.Location = New System.Drawing.Point(14, 44)
        Me._lblVendedores_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVendedores_2.Name = "_lblVendedores_2"
        Me._lblVendedores_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVendedores_2.Size = New System.Drawing.Size(47, 13)
        Me._lblVendedores_2.TabIndex = 3
        Me._lblVendedores_2.Text = "Nombre:"
        '
        '_lblVendedores_0
        '
        Me._lblVendedores_0.AutoSize = True
        Me._lblVendedores_0.BackColor = System.Drawing.Color.Silver
        Me._lblVendedores_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVendedores_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVendedores_0.Location = New System.Drawing.Point(13, 17)
        Me._lblVendedores_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVendedores_0.Name = "_lblVendedores_0"
        Me._lblVendedores_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVendedores_0.Size = New System.Drawing.Size(43, 13)
        Me._lblVendedores_0.TabIndex = 1
        Me._lblVendedores_0.Text = "Código:"
        '
        '_lblVendedores_3
        '
        Me._lblVendedores_3.AutoSize = True
        Me._lblVendedores_3.BackColor = System.Drawing.Color.Silver
        Me._lblVendedores_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVendedores_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVendedores_3.Location = New System.Drawing.Point(8, 68)
        Me._lblVendedores_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVendedores_3.Name = "_lblVendedores_3"
        Me._lblVendedores_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVendedores_3.Size = New System.Drawing.Size(52, 13)
        Me._lblVendedores_3.TabIndex = 5
        Me._lblVendedores_3.Text = "Domicilio:"
        '
        '_lblVendedores_4
        '
        Me._lblVendedores_4.AutoSize = True
        Me._lblVendedores_4.BackColor = System.Drawing.Color.Silver
        Me._lblVendedores_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVendedores_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVendedores_4.Location = New System.Drawing.Point(5, 132)
        Me._lblVendedores_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVendedores_4.Name = "_lblVendedores_4"
        Me._lblVendedores_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVendedores_4.Size = New System.Drawing.Size(52, 13)
        Me._lblVendedores_4.TabIndex = 7
        Me._lblVendedores_4.Text = "Teléfono:"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.fraGeneral)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(385, 451)
        Me.Panel1.TabIndex = 5
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(12, 366)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(360, 74)
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
        'FrmCorpoAbcVendedores
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(406, 475)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(230, 130)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "FrmCorpoAbcVendedores"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Vendedores"
        Me.fraGeneral.ResumeLayout(False)
        Me.fraGeneral.PerformLayout()
        CType(Me.lblVendedores, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        Eliminar()
    End Sub
End Class