'**********************************************************************************************************************'
'*PROGRAMA: TIPO DE MATERIAL JOYERIA RAMOS 
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.Compatibility
Imports ADODB

Public Class frmCorpoAbcTiposMaterial
    Inherits System.Windows.Forms.Form
    'Programa: ABC de Tipos de Material
    'Autor: Rosaura Torres López
    'Fecha de Creación: 12/Mayo/2003

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtCodTipoMaterial As System.Windows.Forms.TextBox
    Public WithEvents txtDescripcion As System.Windows.Forms.TextBox
    Public WithEvents txtDescCorta As System.Windows.Forms.TextBox
    Public WithEvents _lblFormasPago_7 As System.Windows.Forms.Label
    Public WithEvents _lblMaterial_1 As System.Windows.Forms.Label
    Public WithEvents _lblMaterial_0 As System.Windows.Forms.Label
    Public WithEvents fraGeneral As System.Windows.Forms.GroupBox
    Public WithEvents lblFormasPago As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblMaterial As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    'Estas Variables se declaran de manera local, para evitar conflictos al estar usando
    'la misma variable en distintos modulos, que pueden afectar el valor que hayan tomado en un form. distinto al actual
    Public mblnNuevo As Boolean 'Para Controlar si un registro es Nuevo o se trata de una consulta
    Public mblnCambiosEnCodigo As Boolean 'Para Controlar si se han efectuado cambios en el código
    Public WithEvents Panel1 As Panel
    Public WithEvents Panel3 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public mblnSalir As Boolean 'se usa para cuando un usuario presiona escape en el primer control de formulario

    Public strControlActual As String 'Nombre del control actual
    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
    End Sub

    Sub Nuevo()
        'On Error GoTo Merr
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            txtCodTipoMaterial.Enabled = True
            txtCodTipoMaterial.Text = ""
            txtDescripcion.Enabled = True
            txtDescripcion.Text = ""
            txtDescripcion.Tag = ""
            txtDescCorta.Text = ""
            txtDescCorta.Tag = ""

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Sub
            'Merr:
        Catch ex As Exception
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Sub LlenaDatos()
        'Procedimiento que Muestra los datos correspondientes con una clave proporcionada
        'On Error GoTo Merr

        Try
            If Val(txtCodTipoMaterial.Text) = 0 Then
                Nuevo()
                Exit Sub
            End If

            'txtCodTipoMaterial.Text = Format(txtCodTipoMaterial.Text, "000")

            For i = 1 To 3 - (txtCodTipoMaterial.TextLength)
                txtCodTipoMaterial.Text = String.Concat("0" + txtCodTipoMaterial.Text)
            Next i

            gStrSql = "SELECT CodTipoMaterial,DescTipoMaterial, DescCorta FROM  CatTipoMaterial WHERE CodTipoMaterial= '" & txtCodTipoMaterial.Text & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute

            If RsGral.RecordCount > 0 Then
                txtDescripcion.Text = Trim(RsGral.Fields("DescTipoMaterial").Value)
                txtDescripcion.Tag = Trim(RsGral.Fields("DescTipoMaterial").Value)
                txtDescCorta.Text = Trim(RsGral.Fields("DescCorta").Value)
                txtDescCorta.Tag = Trim(RsGral.Fields("DescCorta").Value)
                If Val(txtCodTipoMaterial.Text) = 1 Then txtDescripcion.Enabled = False Else txtDescripcion.Enabled = True
            Else
                MsjNoExiste("El Material", gstrNombCortoEmpresa)
                Limpiar()
            End If

            mblnCambiosEnCodigo = False
            mblnNuevo = False
            Exit Sub
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Sub Limpiar()
        'On Error GoTo Merr
        Try
            'Esta función Limpia todos los controles del formulario.
            'Si hubo Cambios, Pregunta si desea guardarlos.
            'On Error GoTo Merr
            'Screen.MousePointer = vbHourglass
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

            txtCodTipoMaterial.Text = ""
            Nuevo()
            mblnNuevo = True
            mblnCambiosEnCodigo = False
            txtCodTipoMaterial.Focus()
            '    Screen.MousePointer = vbDefault
            Exit Sub
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Function Cambios() As Object
        'Esta Función validará si se han efectuado cambios en los controles.
        'lo cual es útil para la funcion de guardar. Se inicializa con True, y si se validan todos los campos y no se ha
        'salido del proc. entonces la variable adquiere el valor de False
        'se validan todos los controles existentes, excepto el de la Clave Principal
        'On Error GoTo Merr
        Try
            Cambios = True
            If Trim(txtDescripcion.Text) <> Trim(txtDescripcion.Tag) Then Exit Function
            Cambios = False
            Exit Function
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Function ValidaDatos() As Object
        'Esta Función Valida que todos los datos en el Formulario se introduzcan, para poder realizar la Alta del registro
        'On Error GoTo Merr
        Try
            'ValidaDatos = False No es necesario especificarlo, ya que la funcion se inicializa con falso
            If Len(Trim(txtDescripcion.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Descripción", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtDescripcion.Focus()
                Exit Function
            End If
            If Len(Trim(txtDescCorta.Text)) = 0 And CDbl(Numerico(txtCodTipoMaterial.Text)) <> 1 Then
                MsgBox(C_msgFALTADATO & "Descripción corta", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtDescCorta.Focus()
                Exit Function
            End If
            ValidaDatos = True
            Exit Function
            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Function Guardar() As Boolean
        'On Error GoTo Merr
        Try
            'Si no se realizaron cambios, entonces no se guardara nada
            'Si el Código  es "", entonces no se validará nada, solamente se saldrá del proc.
            If Cambios() = False And Trim(txtCodTipoMaterial.Text) = "" Then
                Limpiar()
                Exit Function
            End If

            'Validar si todos los datos fueron proporcionados para ser guardados
            If ValidaDatos() = False Then
                Exit Function
            End If

            If Val(txtCodTipoMaterial.Text) = 0 Then
                mblnNuevo = True
            End If

            'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Cnn.BeginTrans()

            If mblnNuevo = True Then 'Se realizará una insercion
                ModStoredProcedures.PR_IMECatTipoMaterial(Trim(txtCodTipoMaterial.Text), Trim(txtDescripcion.Text), Trim(txtDescCorta.Text), C_INSERCION, CStr(0))
                Cmd.Execute()
                txtCodTipoMaterial.Text = Format(Cmd.Parameters("ID").Value, "000")
            Else ' Se realizará una Modificación
                ModStoredProcedures.PR_IMECatTipoMaterial(Trim(txtCodTipoMaterial.Text), Trim(txtDescripcion.Text), Trim(txtDescCorta.Text), C_MODIFICACION, CStr(0))
                Cmd.Execute()
            End If
            Cnn.CommitTrans()

            If Trim(Me.Tag) = "FRMCXPJOYERIA" Then
                'frmCXPJoyeria.mblnFueraChange = True
                'frmCXPJoyeria.dbcMaterial.Text = Trim(Me.txtDescripcion.Text)
                'frmCXPJoyeria.dbcMaterial.Tag = frmCXPJoyeria.dbcMaterial.Text
                'frmCXPJoyeria.mintCodMaterial = CInt(Numerico((Me.txtCodTipoMaterial.Text)))
                'frmCXPJoyeria.mblnFueraChange = False
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Me.Close()
                Exit Function
            ElseIf Trim(Me.Tag) = "FRMCXPRELOJERIA" Then
                'frmCXPRelojeria.mblnFueraChange = True
                'frmCXPRelojeria.dbcMaterial.Text = Trim(Me.txtDescripcion.Text)
                'frmCXPRelojeria.dbcMaterial.Tag = frmCXPRelojeria.dbcMaterial.Text
                'frmCXPRelojeria.mintCodMaterial = CInt(Numerico((Me.txtCodTipoMaterial.Text)))
                'frmCXPRelojeria.mblnFueraChange = False
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Me.Close()
                Exit Function
            ElseIf Trim(Me.Tag) = "FRMCXPVARIOS" Then
                'frmCXPVarios.mblnFueraChange = True
                'frmCXPVarios.dbcMaterial.Text = Trim(Me.txtDescripcion.Text)
                'frmCXPVarios.dbcMaterial.Tag = frmCXPVarios.dbcMaterial.Text
                'frmCXPVarios.mintCodMaterial = CInt(Numerico((Me.txtCodTipoMaterial.Text)))
                'frmCXPVarios.mblnFueraChange = False
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Me.Close()
                Exit Function
            End If
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            'Por cuestiones de estética el cambio al puntero del mouse se hace antes de iniciar la transacción y al finalizar la misma.

            If mblnNuevo Then
                MsgBox("El Material ha sido grabado correctamente con el Código: " & txtCodTipoMaterial.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
            Else
                MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
            End If
            'Dejar el Procedimiento Nuevo, sirve para que al usar limpiar,. no pregunte si se desea guardar cambios en el codigo
            Nuevo()
            Guardar = True
            Limpiar()

            Exit Function
            'Merr:
        Catch ex As Exception
            Cnn.RollbackTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Sub Eliminar()
        'On Error GoTo Merr
        Try
            'Screen.MousePointer = vbHourglass Esto se manejará hasta antes de iniciar la transacción
            gStrSql = "SELECT DescTipoMaterial FROM CatTipoMaterial WHERE CodTipoMaterial=" & Val(txtCodTipoMaterial.Text)

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
            If CDbl(Numerico(txtCodTipoMaterial.Text)) = 1 Then
                MsgBox("No es posible eliminar este código de material." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                Exit Sub
            End If
            'Preguntar si desea borrar el registro
            Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, "Mensaje")
                Case MsgBoxResult.No
                    Exit Sub
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Cnn.BeginTrans()

            ModStoredProcedures.PR_IMECatTipoMaterial(Trim(txtCodTipoMaterial.Text), Trim(txtDescripcion.Text), Trim(txtDescCorta.Text), C_ELIMINACION, CStr(0))
            Cmd.Execute()
            MsgBox("El Material ha sido eliminado correctamente con el Código: " & txtCodTipoMaterial.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
            Cnn.CommitTrans()
            Nuevo()
            Limpiar()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Sub
            'Merr:
        Catch ex As Exception
            Cnn.RollbackTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Sub Buscar()
        'Esta Función se utilizará para Buscar un dato especifico de un formulario, la cual podrá realizarse por campo Codigo o Campo Descripción,
        ' y se Activará presionando la tecla F3.
        'On Error GoTo Merr
        Try
            Dim strSQL As String
            Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
            Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas


            'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
            strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

            Select Case strControlActual
                Case "TXTCODTIPOMATERIAL"
                    strCaptionForm = "Consulta de Tipos de Material"
                    gStrSql = "SELECT RIGHT('000'+LTRIM(CodTipoMaterial),3) AS CODIGO, DescTipoMaterial AS DESCRIPCION, lTrim(Rtrim(DescCorta)) as DescCorta FROM CatTipoMaterial ORDER BY CodTipoMaterial"
                Case "TXTDESCRIPCION"
                    strCaptionForm = "Consulta de Tipos de Material"
                    gStrSql = "SELECT DescTipoMaterial AS DESCRIPCION,RIGHT('000'+LTRIM(CodTipoMaterial),3) AS CODIGO, lTrim(Rtrim(DescCorta)) as DescCorta FROM CatTipoMaterial WHERE DescTipoMaterial LIKE '" & Trim(txtDescripcion.Text) & "%' ORDER BY DescTipoMaterial"
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
            ConfiguraConsultas(FrmConsultas, 6400, RsGral, strTag, strCaptionForm)

            With FrmConsultas.Flexdet
                Select Case strControlActual
                    Case "TXTCODTIPOMATERIAL"
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                        .set_ColWidth(1, 0, 4500) 'Columna de la Descripción
                        .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColWidth(2, 0, 1000) 'Columna de la Descripción
                        .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    Case "TXTDESCRIPCION"
                        .set_ColWidth(0, 0, 4500) 'Columna de la Descripción
                        .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColWidth(1, 0, 900) 'Columna del Código
                        .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                        .set_ColWidth(2, 0, 1000) 'Columna del Código
                        .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                End Select
            End With

            FrmConsultas.ShowDialog()

            'Merr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub


    Private Sub frmCorpoAbcTiposMaterial_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoAbcTiposMaterial_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'Desactivar todas las opciones del Menu
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmCorpoAbcTiposMaterial_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs)
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        ' En este evento del formulario se valida la tecla presionada.
        ' Si es Enter se simula un tab(Avanza al siguiente control)
        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmCorpoAbcTiposMaterial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Private Sub frmCorpoAbcTiposMaterial_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs)
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Trim(Me.Tag) = "" Then
        '    'Cuando no ha sido llamado desde otro formulario, sale de forma normal
        '    If Not mblnSalir Then
        '        'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '        ModEstandar.RestaurarForma(Me, False)
        '        'Si se cierra el formulario y existio algun cambio en el registro se
        '        'informa al usuario del cabio y si desea guardar el registro, ya sea
        '        'que sea nuevo o un registro modificado
        '        If Cambios() = True Then 'And mblnNuevo = False Then
        '            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '                Case MsgBoxResult.Yes 'Guardar el registro
        '                    If Guardar() = False Then
        '                        Cancel = 1
        '                    End If
        '                Case MsgBoxResult.No 'No hace nada y permite el cierre del formulario
        '                Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
        '                    Cancel = 1
        '            End Select
        '        End If
        '    Else 'Se quiere salir con escape
        '        mblnSalir = False
        '        Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes 'Sale del Formulario
        '                Cancel = 0
        '            Case MsgBoxResult.No 'No sale del formulario
        '                Cancel = 1
        '        End Select
        '    End If
        'Else
        '    Cancel = 0
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoAbcTiposMaterial_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs)
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'If Trim(Me.Tag) = "FRMCXPJOYERIA" Then
        '    frmCXPJoyeria.Enabled = True
        '    frmCXPJoyeria.dbcMaterial.Focus()
        'ElseIf Trim(Me.Tag) = "FRMCXPRELOJERIA" Then
        '    frmCXPRelojeria.Enabled = True
        '    frmCXPRelojeria.dbcMaterial.Focus()
        'ElseIf Trim(Me.Tag) = "FRMCXPVARIOS" Then
        '    frmCXPVarios.Enabled = True
        '    frmCXPVarios.dbcMaterial.Focus()
        'End If

        'Me = Nothing
    End Sub

    Private Sub txtCodTipoMaterial_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodTipoMaterial.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub


    Private Sub txtCodTipoMaterial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodTipoMaterial.Enter
        strControlActual = UCase("txtCodTipoMaterial")
        SelTextoTxt(txtCodTipoMaterial)
        Pon_Tool()
    End Sub


    Private Sub txtCodTipoMaterial_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodTipoMaterial.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000

        If KeyCode = Keys.Enter Then
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
                        txtCodTipoMaterial.Focus()
                        KeyCode = 0
                        Exit Sub
                End Select
            End If
        End If
    End Sub


    Private Sub txtCodTipoMaterial_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodTipoMaterial.KeyPress
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
                        txtCodTipoMaterial.Focus()
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


    Private Sub txtCodTipoMaterial_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodTipoMaterial.Leave

        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        ' Formatear el campo de codigo para cuando se deje en blanco que muestre "000"
        If CDbl(Numerico(Trim(txtCodTipoMaterial.Text))) = 0 Then txtCodTipoMaterial.Text = "000"
        If mblnCambiosEnCodigo = True And CDbl(Numerico(txtCodTipoMaterial.Text)) <> 0 Then 'si hubo cambios en el codigo hace la consulta para llenar los datos
            LlenaDatos()
        End If
    End Sub


    Private Sub txtDescCorta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        SelTextoTxt(txtDescCorta)
    End Sub


    Private Sub txtDescripcion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.TextChanged
        mblnCambiosEnCodigo = True
    End Sub


    Private Sub txtDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Enter
        strControlActual = UCase("txtDescripcion")
        SelTextoTxt(txtDescripcion)
        Pon_Tool()
    End Sub


    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraGeneral = New System.Windows.Forms.GroupBox()
        Me.txtCodTipoMaterial = New System.Windows.Forms.TextBox()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me.txtDescCorta = New System.Windows.Forms.TextBox()
        Me._lblFormasPago_7 = New System.Windows.Forms.Label()
        Me._lblMaterial_1 = New System.Windows.Forms.Label()
        Me._lblMaterial_0 = New System.Windows.Forms.Label()
        Me.lblFormasPago = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblMaterial = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.fraGeneral.SuspendLayout()
        CType(Me.lblFormasPago, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblMaterial, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraGeneral
        '
        Me.fraGeneral.BackColor = System.Drawing.Color.Silver
        Me.fraGeneral.Controls.Add(Me.txtCodTipoMaterial)
        Me.fraGeneral.Controls.Add(Me.txtDescripcion)
        Me.fraGeneral.Controls.Add(Me.txtDescCorta)
        Me.fraGeneral.Controls.Add(Me._lblFormasPago_7)
        Me.fraGeneral.Controls.Add(Me._lblMaterial_1)
        Me.fraGeneral.Controls.Add(Me._lblMaterial_0)
        Me.fraGeneral.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGeneral.Location = New System.Drawing.Point(14, 12)
        Me.fraGeneral.Margin = New System.Windows.Forms.Padding(2)
        Me.fraGeneral.Name = "fraGeneral"
        Me.fraGeneral.Padding = New System.Windows.Forms.Padding(2)
        Me.fraGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGeneral.Size = New System.Drawing.Size(314, 103)
        Me.fraGeneral.TabIndex = 0
        Me.fraGeneral.TabStop = False
        Me.ToolTip1.SetToolTip(Me.fraGeneral, "Descripción")
        '
        'txtCodTipoMaterial
        '
        Me.txtCodTipoMaterial.AcceptsReturn = True
        Me.txtCodTipoMaterial.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodTipoMaterial.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodTipoMaterial.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodTipoMaterial.Location = New System.Drawing.Point(50, 21)
        Me.txtCodTipoMaterial.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodTipoMaterial.MaxLength = 3
        Me.txtCodTipoMaterial.Name = "txtCodTipoMaterial"
        Me.txtCodTipoMaterial.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodTipoMaterial.Size = New System.Drawing.Size(75, 20)
        Me.txtCodTipoMaterial.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtCodTipoMaterial, "Código del Material")
        '
        'txtDescripcion
        '
        Me.txtDescripcion.AcceptsReturn = True
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescripcion.Location = New System.Drawing.Point(74, 46)
        Me.txtDescripcion.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescripcion.MaxLength = 50
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(195, 20)
        Me.txtDescripcion.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtDescripcion, "Descripción del Material")
        '
        'txtDescCorta
        '
        Me.txtDescCorta.AcceptsReturn = True
        Me.txtDescCorta.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescCorta.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescCorta.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescCorta.Location = New System.Drawing.Point(100, 72)
        Me.txtDescCorta.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescCorta.MaxLength = 3
        Me.txtDescCorta.Name = "txtDescCorta"
        Me.txtDescCorta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescCorta.Size = New System.Drawing.Size(169, 20)
        Me.txtDescCorta.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtDescCorta, "Descripción corta")
        '
        '_lblFormasPago_7
        '
        Me._lblFormasPago_7.AutoSize = True
        Me._lblFormasPago_7.BackColor = System.Drawing.Color.Silver
        Me._lblFormasPago_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormasPago_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblFormasPago_7.Location = New System.Drawing.Point(4, 72)
        Me._lblFormasPago_7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblFormasPago_7.Name = "_lblFormasPago_7"
        Me._lblFormasPago_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormasPago_7.Size = New System.Drawing.Size(94, 13)
        Me._lblFormasPago_7.TabIndex = 5
        Me._lblFormasPago_7.Text = "Descripción Corta:"
        '
        '_lblMaterial_1
        '
        Me._lblMaterial_1.AutoSize = True
        Me._lblMaterial_1.BackColor = System.Drawing.Color.Silver
        Me._lblMaterial_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblMaterial_1.ForeColor = System.Drawing.Color.Black
        Me._lblMaterial_1.Location = New System.Drawing.Point(4, 46)
        Me._lblMaterial_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblMaterial_1.Name = "_lblMaterial_1"
        Me._lblMaterial_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblMaterial_1.Size = New System.Drawing.Size(66, 13)
        Me._lblMaterial_1.TabIndex = 3
        Me._lblMaterial_1.Text = "Descripción:"
        '
        '_lblMaterial_0
        '
        Me._lblMaterial_0.AutoSize = True
        Me._lblMaterial_0.BackColor = System.Drawing.Color.Silver
        Me._lblMaterial_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblMaterial_0.ForeColor = System.Drawing.Color.Black
        Me._lblMaterial_0.Location = New System.Drawing.Point(3, 24)
        Me._lblMaterial_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblMaterial_0.Name = "_lblMaterial_0"
        Me._lblMaterial_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblMaterial_0.Size = New System.Drawing.Size(43, 13)
        Me._lblMaterial_0.TabIndex = 1
        Me._lblMaterial_0.Text = "Código:"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.fraGeneral)
        Me.Panel1.Location = New System.Drawing.Point(9, 7)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(342, 207)
        Me.Panel1.TabIndex = 11
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(14, 120)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(314, 74)
        Me.Panel3.TabIndex = 71
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
        'frmCorpoAbcTiposMaterial
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(360, 223)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(271, 247)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoAbcTiposMaterial"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC  a Tipos de Material"
        Me.fraGeneral.ResumeLayout(False)
        Me.fraGeneral.PerformLayout()
        CType(Me.lblFormasPago, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblMaterial, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private Sub frmCorpoAbcTiposMaterial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeComponent()
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        'Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
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