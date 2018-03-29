'**********************************************************************************************************************'
'*PROGRAMA: SUCURSALES JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018      
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCorpoAbcSucursales
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    'Programa: ABC de Sucursales( en tablas es AlMacenes)
    'Autor: Rosaura Torres López
    'Fecha de Creación: 13/Mayo/2003 5:40
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtDomicilio As System.Windows.Forms.RichTextBox
    Public WithEvents chkAlmGeneral As System.Windows.Forms.CheckBox
    Public WithEvents chkMostrarTodos As System.Windows.Forms.CheckBox
    Public WithEvents optAVendedorExterno As System.Windows.Forms.RadioButton
    Public WithEvents optPropia As System.Windows.Forms.RadioButton
    Public WithEvents fraTipoSucursal As System.Windows.Forms.GroupBox
    Public WithEvents txtResponsable As System.Windows.Forms.TextBox
    Public WithEvents txtAuxiliar As System.Windows.Forms.TextBox
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents txtDescripcion As System.Windows.Forms.TextBox
    Public WithEvents _lblSucursales_4 As System.Windows.Forms.Label
    Public WithEvents _lblSucursales_3 As System.Windows.Forms.Label
    Public WithEvents _lblSucursales_2 As System.Windows.Forms.Label
    Public WithEvents _lblSucursales_0 As System.Windows.Forms.Label
    Public WithEvents _lblSucursales_1 As System.Windows.Forms.Label
    Public WithEvents fraGeneral As System.Windows.Forms.GroupBox
    Public WithEvents lblSucursales As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    'Estas Variables se declaran de manera local, para evitar conflictos al estar usando
    'la misma variable en distintos modulos, que pueden afectar el valor que hayan tomado en un form. distinto al actual
    Dim mblnNuevo As Boolean 'Para Controlar si un registro es Nuevo o se trata de una consulta
    Dim mblnCambiosEnCodigo As Boolean 'Para Controlar si se han efectuado cambios en el código
    Dim mblnSALIR As Boolean 'se usa para cuando un usuario presiona escape en el primer control de formulario
    Dim intAlmGeneral As Integer 'Se Usa para Saber si el Almacen es General
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btnSalir As Button
    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnGuardar As Button
    Friend WithEvents btnLimpiar As Button
    Friend WithEvents btnEliminar As Button
    Friend WithEvents Panel1 As Panel
    Dim FueraChange As Boolean

    Public strControlActual As String 'Nombre del control actual
    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        intAlmGeneral = 0
    End Sub

    Sub Buscar()
        'Esta Función se utilizará para Buscar un dato especifico de un formulario, la cual podrá realizarse por campo Codigo o Campo Descripción,
        ' y se Activará presionando la tecla F3.
        'On Error GoTo MErr
        Try
            Dim strSQL As String
            Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
            Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas

            'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
            strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

            Select Case strControlActual
                Case "TXTCODSUCURSAL"
                    strCaptionForm = "Consulta de Sucursales"
                    If Me.chkMostrarTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
                        gStrSql = "SELECT RIGHT('000'+LTRIM(Codalmacen),3) AS CODIGO,Descalmacen AS DESCRIPCION, " & "CASE TipoAlmacen WHEN 'P' THEN 'Propia' WHEN 'V' THEN 'A Vendedor Externo' END as TIPO " & "From Catalmacen ORDER BY Codalmacen"
                    Else
                        gStrSql = "SELECT RIGHT('000'+LTRIM(Codalmacen),3) AS CODIGO,Descalmacen AS DESCRIPCION, " & "CASE TipoAlmacen WHEN 'P' THEN 'Propia' WHEN 'V' THEN 'A Vendedor Externo' END as TIPO " & "From Catalmacen WHERE TipoAlmacen LIKE " & IIf(optPropia.Checked = True, "'P'", "'%'") & " And TipoAlmacen LIKE " & IIf(optAVendedorExterno.Checked = True, "'V'", "'%'") & " ORDER BY Codalmacen"
                    End If
                Case "TXTDESCRIPCION"
                    strCaptionForm = "Consulta de Sucursales"
                    If Me.chkMostrarTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
                        gStrSql = "SELECT Descalmacen AS DESCRIPCION,  RIGHT('000'+LTRIM(Codalmacen),3) AS CODIGO," & "CASE TipoAlmacen WHEN 'P' THEN 'Propia' WHEN 'V' THEN 'A Vendedor Externo' END as TIPO " & "From Catalmacen WHERE DescAlmacen LIKE '" & Trim(txtDescripcion.Text) & "%' ORDER BY DescAlmacen"
                    Else
                        gStrSql = "SELECT Descalmacen AS DESCRIPCION,  RIGHT('000'+LTRIM(Codalmacen),3) AS CODIGO," & "CASE TipoAlmacen WHEN 'P' THEN 'Propia' WHEN 'V' THEN 'A Vendedor Externo' END as TIPO " & "From Catalmacen WHERE DescAlmacen LIKE '" & Trim(txtDescripcion.Text) & "%' AND TipoAlmacen LIKE " & IIf(optPropia.Checked = True, "'P'", "'%'") & "And TipoAlmacen LIKE " & IIf(optAVendedorExterno.Checked = True, "'V'", "'%'") & " ORDER BY DescAlmacen"
                    End If
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
            ConfiguraConsultas(FrmConsultas, 7400, RsGral, strTag, strCaptionForm)

            With FrmConsultas.Flexdet
                Select Case strControlActual
                    Case "TXTCODSUCURSAL"
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                        .set_ColWidth(2, 0, 1700) 'Columna del Tipo
                    Case "TXTDESCRIPCION"
                        .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
                        .set_ColWidth(1, 0, 900) 'Columna del Código
                        .set_ColWidth(2, 0, 1700) 'Columna del Tipo
                End Select
            End With

            FrmConsultas.ShowDialog()
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Sub Eliminar()
        'On Error GoTo MErr
        Try
            'Screen.MousePointer = vbHourglass Esto se manejará hasta antes de iniciar la transacción
            gStrSql = "SELECT DescAlmacen FROM CatAlmacen WHERE CodAlmacen=" & Val(txtCodSucursal.Text)
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute

            If RsGral.RecordCount = 0 Then
                MsgBox("Proporcione un Código valido para eliminar.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Mensaje")
                'Cnn.RollbackTrans()
                RsGral.Close()
                Exit Sub
            End If

            'Preguntar si desea borrar el registro
            Select Case MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton3, "Mensaje")
                Case MsgBoxResult.No
                    Exit Sub
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Cnn.BeginTrans()
            ModStoredProcedures.PR_IMECatAlmacen(Trim(txtCodSucursal.Text), Trim(txtDescripcion.Text), Trim(txtResponsable.Text), Trim(txtAuxiliar.Text), Trim(txtDomicilio.Rtf), "X", "0", C_ELIMINACION, CStr(0))
            Cmd.Execute()
            Cnn.CommitTrans()
            MsgBox("La sucursal se elimino correctamente con el Código: " & txtCodSucursal.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
            Nuevo()
            Limpiar()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Sub
            'MErr:
        Catch ex As Exception
            Cnn.RollbackTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Public Function Guardar() As Boolean
        'On Error GoTo MErr
        Try
            Dim TipoAlmacen As String

            'txtDomicilio_Leave(txtDomicilio, New System.EventArgs())
            ''Si no se realizaron cambios, entonces no se guardara nada
            'Si el Código es "", entonces no se validará nada, solamente se saldrá del proc.
            If Cambios() = False And Trim(txtCodSucursal.Text) = "" Then
                Limpiar()
                Exit Function
            End If

            'Validar si todos los datos fueron proporcionados para ser guardados
            If ValidaDatos() = False Then
                Exit Function
            End If

            If Val(txtCodSucursal.Text) = 0 Then
                mblnNuevo = True
            End If

            'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Cnn.BeginTrans()
            'se obtiene el Tipo de almacen, segun los optionbutton, P:Propio, V: a vendedor externo
            If optPropia.Checked = True Then
                TipoAlmacen = "P"
            Else
                TipoAlmacen = "V"
            End If

            If mblnNuevo = True Then 'Se realizará una insercion
                ModStoredProcedures.PR_IMECatAlmacen(Trim(txtCodSucursal.Text), Trim(txtDescripcion.Text), Trim(txtResponsable.Text), Trim(txtAuxiliar.Text), Trim(txtDomicilio.Text), TipoAlmacen, IIf(chkAlmGeneral.CheckState = 1, "1", "0"), C_INSERCION, CStr(0))
                Cmd.Execute()
                txtCodSucursal.Text = Format(Cmd.Parameters("ID").Value, "000")

            Else ' Se realizará una Modificación
                ModStoredProcedures.PR_IMECatAlmacen(Trim(txtCodSucursal.Text), Trim(txtDescripcion.Text), Trim(txtResponsable.Text), Trim(txtAuxiliar.Text), Trim(txtDomicilio.Text), TipoAlmacen, IIf(chkAlmGeneral.CheckState = 1, "1", "0"), C_MODIFICACION, CStr(0))
                Cmd.Execute()
            End If
            Cnn.CommitTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            'Por cuestiones de estética el cambio al puntero del mouse se hace antes de iniciar la transacción y al finalizar la misma.
            If mblnNuevo Then
                MsgBox("La Sucursal ha sido grabada correctamente con el Código: " & txtCodSucursal.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
            Else
                MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
            End If
            'Dejar el Procedimiento Nuevo, sirve para que al usar limpiar,. no pregunte si se desea guardar cambios en el codigo
            Nuevo()
            Guardar = True
            Limpiar()
            Exit Function
            'MErr:
        Catch ex As Exception
        End Try
        Cnn.RollbackTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub Nuevo()
        'Se deben Limpiar todos los controles del formulario con excepcion del Control de la Llavve principal
        'On Error GoTo MErr
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            txtCodSucursal.Enabled = True
            txtCodSucursal.Text = ""
            txtDescripcion.Text = ""
            txtDescripcion.Tag = ""
            txtResponsable.Text = ""
            txtResponsable.Tag = ""
            txtAuxiliar.Text = ""
            txtAuxiliar.Tag = ""
            txtDomicilio.Text = ""
            txtDomicilio.Tag = ""
            optPropia.Checked = False
            optPropia.Tag = False
            optAVendedorExterno.Checked = False
            optAVendedorExterno.Tag = False
            chkAlmGeneral.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkAlmGeneral.Enabled = False
            chkMostrarTodos.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkMostrarTodos.Enabled = True
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            InicializaVariables()
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
            If Val(txtCodSucursal.Text) = 0 Then
                Nuevo()
                Exit Sub
            End If

            'txtCodSucursal.Text = Format(String.Concat(txtCodSucursal.Text, "000"))

            For i = 1 To 3 - (txtCodSucursal.TextLength)
                txtCodSucursal.Text = String.Concat("0" + txtCodSucursal.Text)
            Next i

            gStrSql = " SELECT CodAlmacen,DescAlmacen,Responsable,Auxiliar,Domicilio,TipoAlmacen,AlmGral FROM CatAlmacen WHERE CodAlmacen= " & Val(txtCodSucursal.Text)
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                txtDescripcion.Text = Trim(RsGral.Fields("DescAlmacen").Value)
                txtDescripcion.Tag = Trim(RsGral.Fields("DescAlmacen").Value)
                txtResponsable.Text = Trim(RsGral.Fields("Responsable").Value)
                txtResponsable.Tag = Trim(RsGral.Fields("Responsable").Value)
                txtAuxiliar.Text = Trim(RsGral.Fields("Auxiliar").Value)
                txtAuxiliar.Tag = Trim(RsGral.Fields("Auxiliar").Value)
                txtDomicilio.Text = Trim(RsGral.Fields("Domicilio").Value)
                txtDomicilio.Tag = Trim(RsGral.Fields("Domicilio").Value)
                If RsGral.Fields("TipoAlmacen").Value = "P" Then
                    optPropia.Checked = True
                    optPropia.Tag = True
                    optAVendedorExterno.Checked = False
                    optAVendedorExterno.Tag = False
                Else
                    optAVendedorExterno.Checked = True
                    optAVendedorExterno.Tag = True
                    optPropia.Checked = False
                    optPropia.Tag = False
                End If
                If RsGral.Fields("AlmGral").Value = True Then
                    FueraChange = True
                    chkAlmGeneral.CheckState = System.Windows.Forms.CheckState.Checked
                    intAlmGeneral = 1
                    FueraChange = False
                Else
                    chkAlmGeneral.CheckState = System.Windows.Forms.CheckState.Unchecked
                    intAlmGeneral = 0
                End If
            Else
                MsjNoExiste("La Sucursal", gstrNombCortoEmpresa)
                Limpiar()
            End If

            txtCodSucursal.Enabled = False
            mblnCambiosEnCodigo = False
            mblnNuevo = False
            Exit Sub
            'MErr:
        Catch ex As Exception
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

            txtCodSucursal.Text = ""
            Nuevo()
            mblnNuevo = True
            mblnCambiosEnCodigo = False
            ' txtCodSucursal.Focus()
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
            Cambios = True
            If Trim(txtDescripcion.Text) <> Trim(txtDescripcion.Tag) Then Exit Function
            If Trim(txtResponsable.Text) <> Trim(txtResponsable.Tag) Then Exit Function
            If Trim(txtAuxiliar.Text) <> Trim(txtAuxiliar.Tag) Then Exit Function
            If Trim(txtDomicilio.Rtf) <> Trim(txtDomicilio.Tag) Then Exit Function
            If optPropia.Checked <> CBool(optPropia.Tag) Then Exit Function
            If optAVendedorExterno.Checked <> CBool(optAVendedorExterno.Tag) Then Exit Function
            If chkAlmGeneral.CheckState <> intAlmGeneral Then Exit Function
            Cambios = False
            Exit Function
            'MErr:
        Catch ex As Exception
            '    Screen.MousePointer = vbDefault
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
        Return Cambios()
    End Function

    Function ValidaDatos() As Object
        'Esta Función Valida que todos los datos en el Formulario se introduzcan, para poder realizar la Alta del registro
        'On Error GoTo MErr
        Try
            If Len(Trim(txtDescripcion.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Descripción", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtDescripcion.Focus()
                Exit Function
            End If
            'If Len(Trim(txtResponsable.Text)) = 0 Then
            '    MsgBox(C_msgFALTADATO & "Responsable", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            '    Me.txtResponsable.Focus()
            '    Exit Function
            'End If
            'If Len(Trim(txtAuxiliar.Text)) = 0 Then
            '    MsgBox(C_msgFALTADATO & "Auxiliar", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            '    Me.txtAuxiliar.Focus()
            '    Exit Function
            'End If
            If Len(Trim(txtDomicilio.Rtf)) = 0 Then
                MsgBox(C_msgFALTADATO & "Domicilio", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtDomicilio.Focus()
                Exit Function
            End If
            'If optPropia.Checked = False And optAVendedorExterno.Checked = False Then
            '    MsgBox(C_msgFALTADATO & "Tipo de Sucursal", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            '    Exit Function
            'End If

            ValidaDatos = True

            Exit Function
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
        Return ValidaDatos()
    End Function


    ' Private Sub chkAlmGeneral_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAlmGeneral.CheckStateChanged
    'If FueraChange Then Exit Sub
    ' If chkAlmGeneral.CheckState = 1 Then
    '    If ExisteAlmGeneral(CShort(txtCodSucursal.Text)) Then
    '        chkAlmGeneral.CheckState = System.Windows.Forms.CheckState.Unchecked
    '        Exit Sub
    '    End If
    ' End If
    ' End Sub

    Private Sub frmCorpoAbcSucursales_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        'ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        'Me.BringToFront()
    End Sub

    Private Sub frmCorpoAbcSucursales_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        'ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmCorpoAbcSucursales_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        'Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
        Nuevo()
        chkAlmGeneral.Enabled = False
    End Sub

    Private Sub frmCorpoAbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmCorpoAbcSucursales_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoAbcSucursales_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSALIR Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        '    'Si se cierra el formulario y existio algun cambio en el registro se
        '    'informa al usuario del cabio y si desea guardar el registro, ya sea
        '    'que sea nuevo o un registro modificado
        '    If Cambios() = True Then ' And mblnNuevo = False Then 'Quite el MblnNuevo=True, para que siempre pregunte al cerrarse, si se kiere guardar cambios
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

    Private Sub frmCorpoAbcSucursales_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        'ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        'ModEstandar.LimpiaDescBarraEstado()
    End Sub


    Private Sub optAVendedorExterno_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAVendedorExterno.CheckedChanged
        If eventSender.Checked Then
            chkAlmGeneral.Enabled = False
        End If
    End Sub


    Private Sub optPropia_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPropia.CheckedChanged
        If eventSender.Checked Then
            chkAlmGeneral.Enabled = True
        End If
    End Sub


    Private Sub txtAuxiliar_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAuxiliar.TextChanged
        'mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtAuxiliar_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAuxiliar.Enter
        'txtAuxiliar.SelectionStart = Len(Trim(txtAuxiliar.Text))
        Pon_Tool()
    End Sub


    Private Sub txtCodSucursal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Enter
        strControlActual = UCase("txtCodSucursal")
        SelTextoTxt(txtCodSucursal)
        Pon_Tool()
    End Sub

    Private Sub txtCodSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodSucursal.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSALIR = True
            Me.Close()
            KeyCode = 0
        Else
            'Si la tecla presionada fue Delete y Hay cambios, pregunta si se desea guardar
            If Cambios() = True And KeyCode = System.Windows.Forms.Keys.Delete Then
                'Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                '    Case MsgBoxResult.Yes 'Guardar el registro
                If Guardar() = False Then
                    KeyCode = 0
                    Exit Sub
                End If
                'Case MsgBoxResult.No 'No hace nada y permite que se borre el contenido del text
                Nuevo()
                'Case MsgBoxResult.Cancel 'Cancela la captura
                ' txtCodSucursal.Focus()
                KeyCode = 0
                Exit Sub
                'End Select
            End If
        End If
    End Sub

    Private Sub txtCodsucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodSucursal.KeyPress
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
                'txtCodSucursal.Focus()
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

    Private Sub txtCodSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Leave
        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        ' Formatear el campo de codigo para cuando se deje en blanco que muestre "000"
        'If CDbl(Numerico(Trim(txtCodSucursal.Text))) = 0 Then txtCodSucursal.Text = "000"
        If mblnCambiosEnCodigo = True And txtCodSucursal.Text <> "" Then 'si hubo cambios en el codigo hace la consulta para llenar los datos
            LlenaDatos()
        End If
    End Sub


    Private Sub txtDescripcion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.TextChanged
        'mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Enter
        strControlActual = UCase("txtDescripcion")
        Pon_Tool()
    End Sub

    Private Sub txtDescripcion_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Leave
        'If Trim(txtDescripcion.Text) = "" Then Exit Sub
        MuestraDatos()
    End Sub

    Private Sub txtDomicilio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.TextChanged
        'mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDomicilio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.Enter
        'txtDomicilio.SelectionStart = 0
        Pon_Tool()
    End Sub

    Private Sub txtDomicilio_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.Leave
        'txtDomicilio.Text = Trim(txtDomicilio.Text)
    End Sub


    Private Sub txtResponsable_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtResponsable.TextChanged
        'mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtResponsable_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtResponsable.Enter
        'txtResponsable.SelectionStart = Len(Trim(txtResponsable.Text))
        'Pon_Tool()
    End Sub

    Private Sub MuestraDatos()
        'On Error GoTo MErr
        Try
            'txtCodSucursal.Text = Format(String.Concat(txtCodSucursal.Text, "000"))

            For i = 1 To 3 - (txtCodSucursal.TextLength)
                txtCodSucursal.Text = String.Concat("0" + txtCodSucursal.Text)
            Next i


            gStrSql = " SELECT CodAlmacen,DescAlmacen,Responsable,Auxiliar,Domicilio,TipoAlmacen,AlmGral FROM CatAlmacen WHERE DescAlmacen = '" & Trim(txtDescripcion.Text) & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                txtCodSucursal.Text = (String.Concat("000" + RsGral.Fields("CodAlmacen").Value))
                LlenaDatos()
            End If
            mblnCambiosEnCodigo = False
            mblnNuevo = False
            Exit Sub
            'MErr:
        Catch ex As Exception
        End Try
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraTipoSucursal = New System.Windows.Forms.GroupBox()
        Me.optAVendedorExterno = New System.Windows.Forms.RadioButton()
        Me.optPropia = New System.Windows.Forms.RadioButton()
        Me.txtResponsable = New System.Windows.Forms.TextBox()
        Me.txtAuxiliar = New System.Windows.Forms.TextBox()
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me.fraGeneral = New System.Windows.Forms.GroupBox()
        Me.txtDomicilio = New System.Windows.Forms.RichTextBox()
        Me.chkAlmGeneral = New System.Windows.Forms.CheckBox()
        Me.chkMostrarTodos = New System.Windows.Forms.CheckBox()
        Me._lblSucursales_4 = New System.Windows.Forms.Label()
        Me._lblSucursales_3 = New System.Windows.Forms.Label()
        Me._lblSucursales_2 = New System.Windows.Forms.Label()
        Me._lblSucursales_0 = New System.Windows.Forms.Label()
        Me._lblSucursales_1 = New System.Windows.Forms.Label()
        Me.lblSucursales = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.fraTipoSucursal.SuspendLayout()
        Me.fraGeneral.SuspendLayout()
        CType(Me.lblSucursales, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraTipoSucursal
        '
        Me.fraTipoSucursal.BackColor = System.Drawing.Color.Silver
        Me.fraTipoSucursal.Controls.Add(Me.optAVendedorExterno)
        Me.fraTipoSucursal.Controls.Add(Me.optPropia)
        Me.fraTipoSucursal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraTipoSucursal.Location = New System.Drawing.Point(70, 197)
        Me.fraTipoSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.fraTipoSucursal.Name = "fraTipoSucursal"
        Me.fraTipoSucursal.Padding = New System.Windows.Forms.Padding(2)
        Me.fraTipoSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTipoSucursal.Size = New System.Drawing.Size(303, 50)
        Me.fraTipoSucursal.TabIndex = 11
        Me.fraTipoSucursal.TabStop = False
        Me.fraTipoSucursal.Text = "Tipo de Sucursal"
        Me.ToolTip1.SetToolTip(Me.fraTipoSucursal, "Tipo de Sucursal")
        '
        'optAVendedorExterno
        '
        Me.optAVendedorExterno.BackColor = System.Drawing.Color.Silver
        Me.optAVendedorExterno.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAVendedorExterno.ForeColor = System.Drawing.Color.Black
        Me.optAVendedorExterno.Location = New System.Drawing.Point(174, 19)
        Me.optAVendedorExterno.Margin = New System.Windows.Forms.Padding(2)
        Me.optAVendedorExterno.Name = "optAVendedorExterno"
        Me.optAVendedorExterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAVendedorExterno.Size = New System.Drawing.Size(111, 21)
        Me.optAVendedorExterno.TabIndex = 13
        Me.optAVendedorExterno.TabStop = True
        Me.optAVendedorExterno.Text = "Vendedor Externo"
        Me.ToolTip1.SetToolTip(Me.optAVendedorExterno, "Sucursal a Vendedor Externo")
        Me.optAVendedorExterno.UseVisualStyleBackColor = False
        '
        'optPropia
        '
        Me.optPropia.BackColor = System.Drawing.Color.Silver
        Me.optPropia.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPropia.ForeColor = System.Drawing.Color.Black
        Me.optPropia.Location = New System.Drawing.Point(78, 20)
        Me.optPropia.Margin = New System.Windows.Forms.Padding(2)
        Me.optPropia.Name = "optPropia"
        Me.optPropia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPropia.Size = New System.Drawing.Size(74, 20)
        Me.optPropia.TabIndex = 12
        Me.optPropia.TabStop = True
        Me.optPropia.Text = "Propia"
        Me.ToolTip1.SetToolTip(Me.optPropia, "Sucursal Propia")
        Me.optPropia.UseVisualStyleBackColor = False
        '
        'txtResponsable
        '
        Me.txtResponsable.AcceptsReturn = True
        Me.txtResponsable.BackColor = System.Drawing.SystemColors.Window
        Me.txtResponsable.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtResponsable.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtResponsable.Location = New System.Drawing.Point(88, 76)
        Me.txtResponsable.Margin = New System.Windows.Forms.Padding(2)
        Me.txtResponsable.MaxLength = 40
        Me.txtResponsable.Name = "txtResponsable"
        Me.txtResponsable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtResponsable.Size = New System.Drawing.Size(319, 20)
        Me.txtResponsable.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtResponsable, "Responsable  de la Sucursal")
        '
        'txtAuxiliar
        '
        Me.txtAuxiliar.AcceptsReturn = True
        Me.txtAuxiliar.BackColor = System.Drawing.SystemColors.Window
        Me.txtAuxiliar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAuxiliar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAuxiliar.Location = New System.Drawing.Point(88, 105)
        Me.txtAuxiliar.Margin = New System.Windows.Forms.Padding(2)
        Me.txtAuxiliar.MaxLength = 40
        Me.txtAuxiliar.Name = "txtAuxiliar"
        Me.txtAuxiliar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAuxiliar.Size = New System.Drawing.Size(319, 20)
        Me.txtAuxiliar.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtAuxiliar, "Auxiliar de la Sucursal")
        '
        'txtCodSucursal
        '
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.Location = New System.Drawing.Point(70, 26)
        Me.txtCodSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodSucursal.MaxLength = 3
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(42, 20)
        Me.txtCodSucursal.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtCodSucursal, "Código de la Sucursal")
        '
        'txtDescripcion
        '
        Me.txtDescripcion.AcceptsReturn = True
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescripcion.Location = New System.Drawing.Point(88, 53)
        Me.txtDescripcion.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescripcion.MaxLength = 40
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(319, 20)
        Me.txtDescripcion.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtDescripcion, "Descripción de Sucursal")
        '
        'fraGeneral
        '
        Me.fraGeneral.BackColor = System.Drawing.Color.Silver
        Me.fraGeneral.Controls.Add(Me.txtDomicilio)
        Me.fraGeneral.Controls.Add(Me.chkAlmGeneral)
        Me.fraGeneral.Controls.Add(Me.chkMostrarTodos)
        Me.fraGeneral.Controls.Add(Me.fraTipoSucursal)
        Me.fraGeneral.Controls.Add(Me.txtResponsable)
        Me.fraGeneral.Controls.Add(Me.txtAuxiliar)
        Me.fraGeneral.Controls.Add(Me.txtCodSucursal)
        Me.fraGeneral.Controls.Add(Me.txtDescripcion)
        Me.fraGeneral.Controls.Add(Me._lblSucursales_4)
        Me.fraGeneral.Controls.Add(Me._lblSucursales_3)
        Me.fraGeneral.Controls.Add(Me._lblSucursales_2)
        Me.fraGeneral.Controls.Add(Me._lblSucursales_0)
        Me.fraGeneral.Controls.Add(Me._lblSucursales_1)
        Me.fraGeneral.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGeneral.Location = New System.Drawing.Point(11, 11)
        Me.fraGeneral.Margin = New System.Windows.Forms.Padding(2)
        Me.fraGeneral.Name = "fraGeneral"
        Me.fraGeneral.Padding = New System.Windows.Forms.Padding(2)
        Me.fraGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGeneral.Size = New System.Drawing.Size(435, 275)
        Me.fraGeneral.TabIndex = 0
        Me.fraGeneral.TabStop = False
        '
        'txtDomicilio
        '
        Me.txtDomicilio.Location = New System.Drawing.Point(69, 128)
        Me.txtDomicilio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDomicilio.Name = "txtDomicilio"
        Me.txtDomicilio.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.txtDomicilio.Size = New System.Drawing.Size(331, 65)
        Me.txtDomicilio.TabIndex = 10
        Me.txtDomicilio.Text = ""
        '
        'chkAlmGeneral
        '
        Me.chkAlmGeneral.BackColor = System.Drawing.Color.Silver
        Me.chkAlmGeneral.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAlmGeneral.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkAlmGeneral.Location = New System.Drawing.Point(318, 254)
        Me.chkAlmGeneral.Margin = New System.Windows.Forms.Padding(2)
        Me.chkAlmGeneral.Name = "chkAlmGeneral"
        Me.chkAlmGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAlmGeneral.Size = New System.Drawing.Size(114, 17)
        Me.chkAlmGeneral.TabIndex = 14
        Me.chkAlmGeneral.Text = "Almacen General"
        Me.chkAlmGeneral.UseVisualStyleBackColor = False
        '
        'chkMostrarTodos
        '
        Me.chkMostrarTodos.BackColor = System.Drawing.Color.Silver
        Me.chkMostrarTodos.Checked = True
        Me.chkMostrarTodos.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMostrarTodos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMostrarTodos.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkMostrarTodos.Location = New System.Drawing.Point(302, 17)
        Me.chkMostrarTodos.Margin = New System.Windows.Forms.Padding(2)
        Me.chkMostrarTodos.Name = "chkMostrarTodos"
        Me.chkMostrarTodos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMostrarTodos.Size = New System.Drawing.Size(110, 32)
        Me.chkMostrarTodos.TabIndex = 15
        Me.chkMostrarTodos.Text = "Mostrar todas las Sucursales"
        Me.chkMostrarTodos.UseVisualStyleBackColor = False
        '
        '_lblSucursales_4
        '
        Me._lblSucursales_4.AutoSize = True
        Me._lblSucursales_4.BackColor = System.Drawing.Color.Silver
        Me._lblSucursales_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSucursales_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSucursales_4.Location = New System.Drawing.Point(12, 134)
        Me._lblSucursales_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblSucursales_4.Name = "_lblSucursales_4"
        Me._lblSucursales_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSucursales_4.Size = New System.Drawing.Size(52, 13)
        Me._lblSucursales_4.TabIndex = 9
        Me._lblSucursales_4.Text = "Domicilio:"
        '
        '_lblSucursales_3
        '
        Me._lblSucursales_3.AutoSize = True
        Me._lblSucursales_3.BackColor = System.Drawing.Color.Silver
        Me._lblSucursales_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSucursales_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSucursales_3.Location = New System.Drawing.Point(12, 107)
        Me._lblSucursales_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblSucursales_3.Name = "_lblSucursales_3"
        Me._lblSucursales_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSucursales_3.Size = New System.Drawing.Size(43, 13)
        Me._lblSucursales_3.TabIndex = 7
        Me._lblSucursales_3.Text = "Auxiliar:"
        '
        '_lblSucursales_2
        '
        Me._lblSucursales_2.AutoSize = True
        Me._lblSucursales_2.BackColor = System.Drawing.Color.Silver
        Me._lblSucursales_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSucursales_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSucursales_2.Location = New System.Drawing.Point(12, 80)
        Me._lblSucursales_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblSucursales_2.Name = "_lblSucursales_2"
        Me._lblSucursales_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSucursales_2.Size = New System.Drawing.Size(72, 13)
        Me._lblSucursales_2.TabIndex = 5
        Me._lblSucursales_2.Text = "Responsable:"
        '
        '_lblSucursales_0
        '
        Me._lblSucursales_0.AutoSize = True
        Me._lblSucursales_0.BackColor = System.Drawing.Color.Silver
        Me._lblSucursales_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSucursales_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSucursales_0.Location = New System.Drawing.Point(12, 29)
        Me._lblSucursales_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblSucursales_0.Name = "_lblSucursales_0"
        Me._lblSucursales_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSucursales_0.Size = New System.Drawing.Size(43, 13)
        Me._lblSucursales_0.TabIndex = 1
        Me._lblSucursales_0.Text = "Código:"
        '
        '_lblSucursales_1
        '
        Me._lblSucursales_1.AutoSize = True
        Me._lblSucursales_1.BackColor = System.Drawing.Color.Silver
        Me._lblSucursales_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSucursales_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSucursales_1.Location = New System.Drawing.Point(12, 55)
        Me._lblSucursales_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblSucursales_1.Name = "_lblSucursales_1"
        Me._lblSucursales_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSucursales_1.Size = New System.Drawing.Size(66, 13)
        Me._lblSucursales_1.TabIndex = 2
        Me._lblSucursales_1.Text = "Descripción:"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(11, 291)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(435, 74)
        Me.Panel3.TabIndex = 70
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
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.fraGeneral)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(467, 380)
        Me.Panel1.TabIndex = 16
        '
        'frmCorpoAbcSucursales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(491, 404)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(214, 134)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoAbcSucursales"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Sucursales"
        Me.fraTipoSucursal.ResumeLayout(False)
        Me.fraGeneral.ResumeLayout(False)
        Me.fraGeneral.PerformLayout()
        CType(Me.lblSucursales, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
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
End Class