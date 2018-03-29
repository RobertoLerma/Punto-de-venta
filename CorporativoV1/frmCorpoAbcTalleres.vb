'**********************************************************************************************************************'
'*PROGRAMA: TALLERES JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB

Public Class frmCorpoAbcTalleres

    Inherits System.Windows.Forms.Form
    'Programa: ABC a Talleres
    'Autor: Rosaura Torres López
    'Fecha de Creación: 12/Mayo/2003 12:00


    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtDomicilio As System.Windows.Forms.RichTextBox
    Public WithEvents chkMostrarTodos As System.Windows.Forms.CheckBox
    Public WithEvents txtCodTaller As System.Windows.Forms.TextBox
    Public WithEvents txtResponsable As System.Windows.Forms.TextBox
    Public WithEvents optForaneo As System.Windows.Forms.RadioButton
    Public WithEvents optJoyeria As System.Windows.Forms.RadioButton
    Public WithEvents optRelojeria As System.Windows.Forms.RadioButton
    Public WithEvents fraTipoTaller As System.Windows.Forms.GroupBox
    Public WithEvents _lblTalleres_1 As System.Windows.Forms.Label
    Public WithEvents _lblTalleres_0 As System.Windows.Forms.Label
    Public WithEvents _lblTalleres_2 As System.Windows.Forms.Label
    Public WithEvents _lblSucursales_4 As System.Windows.Forms.Label
    Public WithEvents fraGeneral As System.Windows.Forms.GroupBox
    Public WithEvents lblSucursales As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblTalleres As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    'Estas Variables se declaran de manera local, para evitar conflictos al estar usando
    'la misma variable en distintos modulos, que pueden afectar el valor que hayan tomado en un form. distinto al actual
    Dim mblnNuevo As Boolean 'Para Controlar si un registro es Nuevo o se trata de una consulta
    Dim mblnCambiosEnCodigo As Boolean 'Para Controlar si se han efectuado cambios en el código
    Public WithEvents txtDescripcion As TextBox
    Public WithEvents Panel1 As Panel
    Public WithEvents Panel3 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public mblnSALIR As Boolean 'se usa para cuando un usuario presiona escape en el primer control de formulario
    Public strControlActual As String 'Nombre del control actual

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
    End Sub

    Sub Buscar()
        'Esta Función se utilizará para Buscar un dato especifico de un formulario, la cual podrá realizarse por campo Codigo o Campo Descripción,
        ' y se Activará presionando la tecla F3.
        'On Error GoTo MErr
        Try
            Dim strSQL As String
            Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
            Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas

            Dim TipoTaller As String 'Tipo de Taller seleccionado segun los Option Button J-R-F

            'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
            strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

            If optJoyeria.Checked = True Then
                TipoTaller = "J"
            Else
                If optRelojeria.Checked = True Then
                    TipoTaller = "R"
                Else
                    If optForaneo.Checked = True Then
                        TipoTaller = "F"
                    Else
                        TipoTaller = ""
                    End If
                End If
            End If


            Select Case strControlActual
                Case "TXTCODTALLER"
                    strCaptionForm = "Consulta de Talleres"
                    If chkMostrarTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
                        gStrSql = "SELECT RIGHT('00'+LTRIM(CodTaller),2) AS CODIGO,DescTaller AS DESCRIPCION,Responsable as  RESPONSABLE, " & "(CASE TipoTaller WHEN 'J' THEN 'JOYERIA' WHEN 'R' THEN 'RELOJERIA'  WHEN 'F' THEN 'FORANEO' END) AS TIPO " & "FROM CatTalleres ORDER BY CodTaller "
                    Else
                        gStrSql = "SELECT RIGHT('00'+LTRIM(CodTaller),2) AS CODIGO,DescTaller AS DESCRIPCION,Responsable as  RESPONSABLE, " & "(CASE TipoTaller WHEN 'J' THEN 'JOYERIA' WHEN 'R' THEN 'RELOJERIA'  WHEN 'F' THEN 'FORANEO' END) AS TIPO " & "FROM CatTalleres " & "WHERE " & IIf((Trim(TipoTaller) <> ""), "TipoTaller= '" & TipoTaller & "'", "TipoTaller LIKE '%'") & "ORDER BY CodTaller "
                    End If
                Case "TXTDESCRIPCION"
                    strCaptionForm = "Consulta de Talleres"
                    If chkMostrarTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
                        gStrSql = "SELECT DescTaller AS DESCRIPCION, RIGHT('00'+LTRIM(CodTaller),2) AS CODIGO,Responsable as  RESPONSABLE, " & "(CASE TipoTaller WHEN 'J' THEN 'JOYERIA' WHEN 'R' THEN 'RELOJERIA'  WHEN 'F' THEN 'FORANEO' END) AS TIPO " & "FROM CatTalleres WHERE DescTaller LIKE '" & txtDescripcion.Text & "%' " & " ORDER BY DescTaller"
                    Else
                        gStrSql = "SELECT DescTaller AS DESCRIPCION, RIGHT('00'+LTRIM(CodTaller),2) AS CODIGO,Responsable as  RESPONSABLE, " & "(CASE TipoTaller WHEN 'J' THEN 'JOYERIA' WHEN 'R' THEN 'RELOJERIA'  WHEN 'F' THEN 'FORANEO' END) AS TIPO " & "FROM CatTalleres WHERE DescTaller LIKE '" & txtDescripcion.Text & "%' " & " AND " & IIf((Trim(TipoTaller) <> ""), " TipoTaller= '" & TipoTaller & "'", "TipoTaller LIKE '%'") & " ORDER BY DescTaller"
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
            ConfiguraConsultas(FrmConsultas, 10200, RsGral, strTag, strCaptionForm)

            With FrmConsultas.Flexdet
                Select Case strControlActual
                    Case "TXTCODTALLER"
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColWidth(1, 0, 3800) 'Columna de la Descripción
                        .set_ColWidth(2, 0, 4500) 'Columna del Nombre del Responsable
                        .set_ColWidth(3, 0, 1000) 'Columna del Tipo de Taller
                    Case "TXTDESCRIPCION"
                        .set_ColWidth(0, 0, 3800) 'Columna de la Descripción
                        .set_ColWidth(1, 0, 900) 'Columna del Código
                        .set_ColWidth(2, 0, 4500) 'Columna del Nombre del Responsable
                        .set_ColWidth(3, 0, 1000) 'Columna del Tipo de Taller
                End Select
            End With
            ModEstandar.CentrarForma(FrmConsultas)
            FrmConsultas.ShowDialog()
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub
    Sub Eliminar()
        On Error GoTo MErr
        '    Screen.MousePointer = vbHourglass Esto se manejará hasta antes de iniciar la transacción

        gStrSql = "SELECT DescTaller FROM CatTalleres WHERE CodTaller=" & Val(txtCodTaller.Text)

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount = 0 Then
            MsgBox("Proporcione un Código valido para eliminar.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Mensaje")
            'cnn.RollbackTrans()
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
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'El parametro TipoTaller no es requerido en la eliminación, por tanto le estoy mandando un Valor Fijo ("O")
        'cnn.BeginTrans()

        ModStoredProcedures.PR_IMECatTalleres(Trim(txtCodTaller.Text), Trim(txtDescripcion.Text), Trim(txtResponsable.Text), Trim(txtDomicilio.Rtf), "O", C_ELIMINACION, CStr(0))
        Cmd.Execute()
        MsgBox("El Taller ha sido eliminado correctamente con el Código: " & txtCodTaller.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
        'cnn.CommitTrans()
        Nuevo()
        Limpiar()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
MErr:
        'cnn.RollbackTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub
    Function Guardar() As Boolean
        'On Error GoTo MErr
        Try
            Dim TipoTaller As String

            txtDomicilio_Leave(txtDomicilio, New System.EventArgs())
            'Si no se realizaron cambios, entonces no se guardara nada
            'Si el Código  es "", entonces no se validará nada, solamente se saldrá del proc.
            If Cambios() = False And Trim(txtCodTaller.Text) = "" Then
                Limpiar()
                Exit Function
            End If

            'Validar si todos los datos fueron proporcionados para ser guardados
            If ValidaDatos() = False Then
                Exit Function
            End If

            If Val(txtCodTaller.Text) = 0 Then
                mblnNuevo = True
            End If

            If optJoyeria.Checked = True Then
                TipoTaller = "J"
            Else
                If optRelojeria.Checked = True Then
                    TipoTaller = "R"
                Else
                    If optForaneo.Checked = True Then
                        TipoTaller = "F"
                    End If
                End If
            End If
            'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            '  cnn.BeginTrans()

            If mblnNuevo = True Then 'Se realizará una insercion
                ModStoredProcedures.PR_IMECatTalleres(Trim(txtCodTaller.Text), Trim(txtDescripcion.Text), Trim(txtResponsable.Text), Trim(txtDomicilio.Text), TipoTaller, C_INSERCION, CStr(0))
                Cmd.Execute()
                txtCodTaller.Text = Format(Cmd.Parameters("ID").Value, "00")

            Else ' Se realizará una Modificación
                ModStoredProcedures.PR_IMECatTalleres(Trim(txtCodTaller.Text), Trim(txtDescripcion.Text), Trim(txtResponsable.Text), Trim(txtDomicilio.Text), TipoTaller, C_MODIFICACION, CStr(0))
                Cmd.Execute()
            End If

            'cnn.CommitTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            'Por cuestiones de estética el cambio al puntero del mouse se hace antes de iniciar la transacción y al finalizar la misma.

            If mblnNuevo Then
                MsgBox("El Taller ha sido grabado correctamente con el Código: " & txtCodTaller.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
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
            'cnn.RollbackTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function
    Public Sub Nuevo()
        'Se deben Limpiar todos los controles del formulario con excepcion del Control de la Llave principal
        'On Error GoTo MErr
        Try
            'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            txtCodTaller.Enabled = True
            txtCodTaller.Text = ""
            txtDescripcion.Text = ""
            txtDescripcion.Tag = ""
            txtResponsable.Text = ""
            txtResponsable.Tag = ""
            txtDomicilio.Text = ""
            txtDomicilio.Tag = ""
            chkMostrarTodos.Checked = False
            chkMostrarTodos.Tag = False
            optForaneo.Checked = False
            optForaneo.Tag = False
            optRelojeria.Checked = False
            optRelojeria.Tag = False
            optJoyeria.Checked = False
            optJoyeria.Tag = False
            'System.Windows.Forms.Cursor.Current = False
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
            'Screen.MousePointer = vbHourglass
            If Val(txtCodTaller.Text) = 0 Then
                Nuevo()
                'ModEstandar.AvanzarTab Me
                Exit Sub
            End If

            'txtCodTaller.Text = VB6.Format(txtCodTaller.Text, "00")
            For i = 1 To 2 - (txtCodTaller.TextLength)
                txtCodTaller.Text = String.Concat("0" + txtCodTaller.Text)
            Next i

            gStrSql = "SELECT CodTaller,DescTaller,Responsable,Domicilio,TipoTaller FROM  CatTalleres WHERE CodTaller= '" & txtCodTaller.Text & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute

            If RsGral.RecordCount > 0 Then
                txtDescripcion.Text = Trim(RsGral.Fields("DescTaller").Value)
                txtDescripcion.Tag = Trim(RsGral.Fields("DescTaller").Value)
                txtResponsable.Text = Trim(RsGral.Fields("Responsable").Value)
                txtResponsable.Tag = Trim(RsGral.Fields("Responsable").Value)
                txtDomicilio.Text = Trim(RsGral.Fields("Domicilio").Value)
                txtDomicilio.Tag = Trim(RsGral.Fields("Domicilio").Value)
                Select Case RsGral.Fields("TipoTaller").Value
                    Case "J"
                        optJoyeria.Checked = System.Windows.Forms.CheckState.Checked
                        optJoyeria.Tag = System.Windows.Forms.CheckState.Checked
                    Case "R"
                        optRelojeria.Checked = System.Windows.Forms.CheckState.Checked
                        optRelojeria.Tag = System.Windows.Forms.CheckState.Checked
                    Case "F"
                        optForaneo.Checked = System.Windows.Forms.CheckState.Checked
                        optForaneo.Tag = System.Windows.Forms.CheckState.Checked
                End Select
            Else
                MsjNoExiste("El Taller", gstrNombCortoEmpresa)
                Limpiar()
            End If

            txtCodTaller.Enabled = False
            mblnCambiosEnCodigo = False
            mblnNuevo = False
            '    Screen.MousePointer = vbDefault
            Exit Sub
            'MErr:
        Catch ex As Exception
            '    Screen.MousePointer = vbDefault
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

            txtCodTaller.Text = ""
            Nuevo()
            mblnNuevo = True
            mblnCambiosEnCodigo = False
            txtCodTaller.Focus()
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
            If Trim(txtResponsable.Text) <> Trim(txtResponsable.Tag) Then Exit Function
            If Trim(txtDomicilio.Rtf) <> Trim(txtDomicilio.Tag) Then Exit Function
            If optJoyeria.Checked <> CBool(optJoyeria.Tag) Then Exit Function
            If optRelojeria.Checked <> CBool(optRelojeria.Tag) Then Exit Function
            If optForaneo.Checked <> CBool(optForaneo.Tag) Then Exit Function
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
            'Screen.MousePointer = vbHourglass
            'ValidaDatos = False No es necesario especificarlo, ya que la funcion se inicializa con falso
            If Len(Trim(txtDescripcion.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Descripción", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtDescripcion.Focus()
                Exit Function
            End If
            If Len(Trim(txtResponsable.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Responsable", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtResponsable.Focus()
                Exit Function
            End If
            If Len(Trim(txtDomicilio.Rtf)) = 0 Then
                MsgBox(C_msgFALTADATO & "Domicilio", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtDomicilio.Focus()
                Exit Function
            End If
            If optJoyeria.Checked = System.Windows.Forms.CheckState.Unchecked And optRelojeria.Checked = System.Windows.Forms.CheckState.Unchecked And optForaneo.Checked = System.Windows.Forms.CheckState.Unchecked Then
                MsgBox(C_msgFALTADATO & "Tipo de Taller", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                '        Me.optJoyeria.SetFocus
                Exit Function
            End If
            ValidaDatos = True
            'Screen.MousePointer = vbDefault
            Exit Function
            'MErr:
        Catch ex As Exception
            'Screen.MousePointer = vbDefault
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Private Sub frmCorpoAbcTalleres_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoAbcTalleres_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmCorpoAbcTalleres_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        'Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
        Nuevo()
    End Sub

    Private Sub frmCorpoAbcTalleres_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmCorpoAbcTalleres_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoAbcTalleres_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSALIR Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        '    'Si se cierra el formulario y existio algun cambio en el registro se
        '    'informa al usuario del cabio y si desea guardar el registro, ya sea
        '    'que sea nuevo o un registro modificado
        '    If Cambios() = True Then ' And mblnNuevo = False Then
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

    Private Sub frmCorpoAbcTalleres_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        IsNothing(Me)
    End Sub

    Private Sub txtCodtaller_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodTaller.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodtaller_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodTaller.Enter
        strControlActual = UCase("txtCodtaller")
        SelTextoTxt(txtCodTaller)
        Pon_Tool()
    End Sub

    Private Sub txtCodtaller_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodTaller.KeyDown
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
                        ' y se borra también el contenido de todo los controles
                        Nuevo()
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        txtCodTaller.Focus()
                        KeyCode = 0
                        Exit Sub
                End Select
            End If
        End If
    End Sub

    Private Sub txtCodtaller_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodTaller.KeyPress
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
                txtCodTaller.Focus()
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

    Private Sub txtCodtaller_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodTaller.Leave
        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If Val(Trim(txtCodTaller.Text)) = 0 Then txtCodTaller.Text = "00"
        If mblnCambiosEnCodigo = True And CDbl(Numerico(txtCodTaller.Text)) <> 0 Then 'si hubo cambios en el codigo hace la consulta para llenar los datos
            LlenaDatos()
        End If
    End Sub

    Private Sub txtDescripcion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        '    SelTextoTxt txtDescripcion
        txtDescripcion.SelectionStart = Len(Trim(txtDescripcion.Text))
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
        txtDomicilio.Text = Trim(txtDomicilio.Text)
    End Sub

    Private Sub txtResponsable_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtResponsable.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtResponsable_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtResponsable.Enter
        '    SelTextoTxt txtResponsable
        txtResponsable.SelectionStart = Len(Trim(txtResponsable.Text))
        Pon_Tool()
    End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCodTaller = New System.Windows.Forms.TextBox()
        Me.txtResponsable = New System.Windows.Forms.TextBox()
        Me.fraTipoTaller = New System.Windows.Forms.GroupBox()
        Me.optForaneo = New System.Windows.Forms.RadioButton()
        Me.optJoyeria = New System.Windows.Forms.RadioButton()
        Me.optRelojeria = New System.Windows.Forms.RadioButton()
        Me.fraGeneral = New System.Windows.Forms.GroupBox()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me.txtDomicilio = New System.Windows.Forms.RichTextBox()
        Me.chkMostrarTodos = New System.Windows.Forms.CheckBox()
        Me._lblTalleres_1 = New System.Windows.Forms.Label()
        Me._lblTalleres_0 = New System.Windows.Forms.Label()
        Me._lblTalleres_2 = New System.Windows.Forms.Label()
        Me._lblSucursales_4 = New System.Windows.Forms.Label()
        Me.lblSucursales = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblTalleres = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.fraTipoTaller.SuspendLayout()
        Me.fraGeneral.SuspendLayout()
        CType(Me.lblSucursales, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTalleres, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtCodTaller
        '
        Me.txtCodTaller.AcceptsReturn = True
        Me.txtCodTaller.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodTaller.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodTaller.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodTaller.Location = New System.Drawing.Point(60, 28)
        Me.txtCodTaller.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodTaller.MaxLength = 2
        Me.txtCodTaller.Name = "txtCodTaller"
        Me.txtCodTaller.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodTaller.Size = New System.Drawing.Size(71, 20)
        Me.txtCodTaller.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtCodTaller, "Código del Taller")
        '
        'txtResponsable
        '
        Me.txtResponsable.AcceptsReturn = True
        Me.txtResponsable.BackColor = System.Drawing.SystemColors.Window
        Me.txtResponsable.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtResponsable.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtResponsable.Location = New System.Drawing.Point(89, 72)
        Me.txtResponsable.Margin = New System.Windows.Forms.Padding(2)
        Me.txtResponsable.MaxLength = 40
        Me.txtResponsable.Name = "txtResponsable"
        Me.txtResponsable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtResponsable.Size = New System.Drawing.Size(234, 20)
        Me.txtResponsable.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtResponsable, "Responsable  del taller")
        '
        'fraTipoTaller
        '
        Me.fraTipoTaller.BackColor = System.Drawing.Color.Silver
        Me.fraTipoTaller.Controls.Add(Me.optForaneo)
        Me.fraTipoTaller.Controls.Add(Me.optJoyeria)
        Me.fraTipoTaller.Controls.Add(Me.optRelojeria)
        Me.fraTipoTaller.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraTipoTaller.Location = New System.Drawing.Point(89, 156)
        Me.fraTipoTaller.Margin = New System.Windows.Forms.Padding(2)
        Me.fraTipoTaller.Name = "fraTipoTaller"
        Me.fraTipoTaller.Padding = New System.Windows.Forms.Padding(2)
        Me.fraTipoTaller.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTipoTaller.Size = New System.Drawing.Size(256, 46)
        Me.fraTipoTaller.TabIndex = 9
        Me.fraTipoTaller.TabStop = False
        Me.fraTipoTaller.Text = "Tipo de Taller"
        Me.ToolTip1.SetToolTip(Me.fraTipoTaller, "Tipo de Sucursal")
        '
        'optForaneo
        '
        Me.optForaneo.BackColor = System.Drawing.Color.Silver
        Me.optForaneo.Cursor = System.Windows.Forms.Cursors.Default
        Me.optForaneo.ForeColor = System.Drawing.Color.Black
        Me.optForaneo.Location = New System.Drawing.Point(170, 17)
        Me.optForaneo.Margin = New System.Windows.Forms.Padding(2)
        Me.optForaneo.Name = "optForaneo"
        Me.optForaneo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optForaneo.Size = New System.Drawing.Size(74, 24)
        Me.optForaneo.TabIndex = 12
        Me.optForaneo.TabStop = True
        Me.optForaneo.Text = "Foráneo"
        Me.ToolTip1.SetToolTip(Me.optForaneo, "Talle Foráneo")
        Me.optForaneo.UseVisualStyleBackColor = False
        '
        'optJoyeria
        '
        Me.optJoyeria.BackColor = System.Drawing.Color.Silver
        Me.optJoyeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.optJoyeria.ForeColor = System.Drawing.Color.Black
        Me.optJoyeria.Location = New System.Drawing.Point(23, 17)
        Me.optJoyeria.Margin = New System.Windows.Forms.Padding(2)
        Me.optJoyeria.Name = "optJoyeria"
        Me.optJoyeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optJoyeria.Size = New System.Drawing.Size(64, 24)
        Me.optJoyeria.TabIndex = 10
        Me.optJoyeria.TabStop = True
        Me.optJoyeria.Text = "Joyería"
        Me.ToolTip1.SetToolTip(Me.optJoyeria, "Taller de Joyería")
        Me.optJoyeria.UseVisualStyleBackColor = False
        '
        'optRelojeria
        '
        Me.optRelojeria.BackColor = System.Drawing.Color.Silver
        Me.optRelojeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.optRelojeria.ForeColor = System.Drawing.Color.Black
        Me.optRelojeria.Location = New System.Drawing.Point(92, 17)
        Me.optRelojeria.Margin = New System.Windows.Forms.Padding(2)
        Me.optRelojeria.Name = "optRelojeria"
        Me.optRelojeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optRelojeria.Size = New System.Drawing.Size(73, 24)
        Me.optRelojeria.TabIndex = 11
        Me.optRelojeria.TabStop = True
        Me.optRelojeria.Text = "Relojería"
        Me.ToolTip1.SetToolTip(Me.optRelojeria, "Taller de Relojería")
        Me.optRelojeria.UseVisualStyleBackColor = False
        '
        'fraGeneral
        '
        Me.fraGeneral.BackColor = System.Drawing.Color.Silver
        Me.fraGeneral.Controls.Add(Me.txtDescripcion)
        Me.fraGeneral.Controls.Add(Me.txtDomicilio)
        Me.fraGeneral.Controls.Add(Me.chkMostrarTodos)
        Me.fraGeneral.Controls.Add(Me.txtCodTaller)
        Me.fraGeneral.Controls.Add(Me.txtResponsable)
        Me.fraGeneral.Controls.Add(Me.fraTipoTaller)
        Me.fraGeneral.Controls.Add(Me._lblTalleres_1)
        Me.fraGeneral.Controls.Add(Me._lblTalleres_0)
        Me.fraGeneral.Controls.Add(Me._lblTalleres_2)
        Me.fraGeneral.Controls.Add(Me._lblSucursales_4)
        Me.fraGeneral.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGeneral.Location = New System.Drawing.Point(14, 11)
        Me.fraGeneral.Margin = New System.Windows.Forms.Padding(2)
        Me.fraGeneral.Name = "fraGeneral"
        Me.fraGeneral.Padding = New System.Windows.Forms.Padding(2)
        Me.fraGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGeneral.Size = New System.Drawing.Size(410, 221)
        Me.fraGeneral.TabIndex = 0
        Me.fraGeneral.TabStop = False
        '
        'txtDescripcion
        '
        Me.txtDescripcion.Location = New System.Drawing.Point(89, 50)
        Me.txtDescripcion.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescripcion.MaxLength = 30
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.Size = New System.Drawing.Size(234, 20)
        Me.txtDescripcion.TabIndex = 14
        '
        'txtDomicilio
        '
        Me.txtDomicilio.Location = New System.Drawing.Point(89, 94)
        Me.txtDomicilio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDomicilio.Name = "txtDomicilio"
        Me.txtDomicilio.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.txtDomicilio.Size = New System.Drawing.Size(234, 50)
        Me.txtDomicilio.TabIndex = 8
        Me.txtDomicilio.Text = ""
        '
        'chkMostrarTodos
        '
        Me.chkMostrarTodos.BackColor = System.Drawing.Color.Silver
        Me.chkMostrarTodos.Checked = True
        Me.chkMostrarTodos.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMostrarTodos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMostrarTodos.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkMostrarTodos.Location = New System.Drawing.Point(236, 9)
        Me.chkMostrarTodos.Margin = New System.Windows.Forms.Padding(2)
        Me.chkMostrarTodos.Name = "chkMostrarTodos"
        Me.chkMostrarTodos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMostrarTodos.Size = New System.Drawing.Size(169, 40)
        Me.chkMostrarTodos.TabIndex = 13
        Me.chkMostrarTodos.Text = "Mostrar todos los Talleres"
        Me.chkMostrarTodos.UseVisualStyleBackColor = False
        '
        '_lblTalleres_1
        '
        Me._lblTalleres_1.AutoSize = True
        Me._lblTalleres_1.BackColor = System.Drawing.Color.Silver
        Me._lblTalleres_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblTalleres_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblTalleres_1.Location = New System.Drawing.Point(14, 53)
        Me._lblTalleres_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblTalleres_1.Name = "_lblTalleres_1"
        Me._lblTalleres_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblTalleres_1.Size = New System.Drawing.Size(66, 13)
        Me._lblTalleres_1.TabIndex = 3
        Me._lblTalleres_1.Text = "Descripción:"
        '
        '_lblTalleres_0
        '
        Me._lblTalleres_0.AutoSize = True
        Me._lblTalleres_0.BackColor = System.Drawing.Color.Silver
        Me._lblTalleres_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblTalleres_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblTalleres_0.Location = New System.Drawing.Point(14, 31)
        Me._lblTalleres_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblTalleres_0.Name = "_lblTalleres_0"
        Me._lblTalleres_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblTalleres_0.Size = New System.Drawing.Size(43, 13)
        Me._lblTalleres_0.TabIndex = 1
        Me._lblTalleres_0.Text = "Código:"
        '
        '_lblTalleres_2
        '
        Me._lblTalleres_2.AutoSize = True
        Me._lblTalleres_2.BackColor = System.Drawing.Color.Silver
        Me._lblTalleres_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblTalleres_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblTalleres_2.Location = New System.Drawing.Point(14, 75)
        Me._lblTalleres_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblTalleres_2.Name = "_lblTalleres_2"
        Me._lblTalleres_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblTalleres_2.Size = New System.Drawing.Size(72, 13)
        Me._lblTalleres_2.TabIndex = 5
        Me._lblTalleres_2.Text = "Responsable:"
        '
        '_lblSucursales_4
        '
        Me._lblSucursales_4.AutoSize = True
        Me._lblSucursales_4.BackColor = System.Drawing.Color.Silver
        Me._lblSucursales_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSucursales_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSucursales_4.Location = New System.Drawing.Point(14, 97)
        Me._lblSucursales_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblSucursales_4.Name = "_lblSucursales_4"
        Me._lblSucursales_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSucursales_4.Size = New System.Drawing.Size(52, 13)
        Me._lblSucursales_4.TabIndex = 7
        Me._lblSucursales_4.Text = "Domicilio:"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.fraGeneral)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(438, 321)
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
        Me.Panel3.Location = New System.Drawing.Point(14, 236)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(410, 74)
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
        'frmCorpoAbcTalleres
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(461, 344)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(177, 160)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoAbcTalleres"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC  a Talleres"
        Me.fraTipoTaller.ResumeLayout(False)
        Me.fraGeneral.ResumeLayout(False)
        Me.fraGeneral.PerformLayout()
        CType(Me.lblSucursales, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTalleres, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
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