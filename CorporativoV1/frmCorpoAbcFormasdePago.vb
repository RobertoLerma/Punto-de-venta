'**********************************************************************************************************************'
'*PROGRAMA: ABC FORMAS DE PAGO JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA   
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018      
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB

Public Class frmCorpoAbcFormasdePago
    Inherits System.Windows.Forms.Form
    'Programa: ABC de Formas de Pago
    'Autor: Rosaura Torres López
    'Fecha de Creación: 15/Mayo/2003

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtDescCorta As System.Windows.Forms.TextBox
    Public WithEvents txtDescripcion As System.Windows.Forms.TextBox
    Public WithEvents txtCodFormaPago As System.Windows.Forms.TextBox
    Public WithEvents optCancelado As System.Windows.Forms.RadioButton
    Public WithEvents optVigente As System.Windows.Forms.RadioButton
    Public WithEvents optSuspendido As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents Option3 As System.Windows.Forms.RadioButton
    Public WithEvents Option2 As System.Windows.Forms.RadioButton
    Public WithEvents Option1 As System.Windows.Forms.RadioButton
    Public WithEvents chkComisionBanc As System.Windows.Forms.CheckBox
    Public WithEvents chkConsiderarRetiros As System.Windows.Forms.CheckBox
    Public WithEvents chkConsiderarFact As System.Windows.Forms.CheckBox
    Public WithEvents chkReqAutorizacion As System.Windows.Forms.CheckBox
    Public WithEvents chkEsDolar As System.Windows.Forms.CheckBox
    Public WithEvents cboTeclaRapida As System.Windows.Forms.ComboBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents chkReqDocumento As System.Windows.Forms.CheckBox
    Public WithEvents chkEsCheque As System.Windows.Forms.CheckBox
    Public WithEvents chkEsTarjeta As System.Windows.Forms.CheckBox
    Public WithEvents txtPorcIvaComision As System.Windows.Forms.TextBox
    Public WithEvents txtPorcComision As System.Windows.Forms.TextBox
    Public WithEvents chkRestCambio As System.Windows.Forms.CheckBox
    Public WithEvents chkEsDevolucion As System.Windows.Forms.CheckBox
    Public WithEvents chkEsDocumentoInterno As System.Windows.Forms.CheckBox
    Public WithEvents lblBanco As System.Windows.Forms.Label
    Public WithEvents _lblFormasPago_5 As System.Windows.Forms.Label
    Public WithEvents _lblFormasPago_2 As System.Windows.Forms.Label
    Public WithEvents _lblFormasPago_4 As System.Windows.Forms.Label
    Public WithEvents _lblFormasPago_3 As System.Windows.Forms.Label
    Public WithEvents fraConfiguracion As System.Windows.Forms.GroupBox
    Public WithEvents btnDenominaciones As System.Windows.Forms.Button
    Public WithEvents _lblFormasPago_7 As System.Windows.Forms.Label
    Public WithEvents _lblFormasPago_0 As System.Windows.Forms.Label
    Public WithEvents _lblFormasPago_1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents lblFormasPago As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    'Estas Variables se declaran de manera local, para evitar conflictos al estar usando
    'la misma variable en distintos modulos, que pueden afectar el valor que hayan tomado en un form. distinto al actual
    Dim mblnNuevo As Boolean 'Para Controlar si un registro es Nuevo o se trata de una consulta
    Dim mblnCambiosEnCodigo As Boolean 'Para Controlar si se han efectuado cambios en el código
    Dim rsCombo As ADODB.Recordset
    Dim cLetras As Object
    Dim cLetrasEnCombo As String
    Dim maLetra(25) As String
    Dim I As Integer
    Dim Denominacion As Decimal
    Dim mblnSalir As Boolean 'se usa para cuando un usuario presiona escape en el primer control de formulario
    Dim mblnLoadDenominaciones As Boolean 'Indica si el Formulario de Denominaciones está cargado en memoria o no
    Const C_ColDENOMINACION As Integer = 0 'Para el Manejo del GRid
    Dim intCodBanco As Integer
    Dim FueraChange As Boolean
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents dbcBanco As ComboBox
    Public WithEvents btnBuscar As Button
    Public WithEvents btnSalir As Button
    Public Tecla As Integer
    Public strControlActual As String 'Nombre del control actual

    Function EsFormaPagoUsada(ByRef CodFormaPago As Integer) As Boolean
        'On Error GoTo MErr
        Try
            gStrSql = "SELECT  *  From dbo.IngresosFormaDePago Where (CodFormaPago = " & CodFormaPago & ")"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                EsFormaPagoUsada = True
            End If
            Exit Function
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        gblnMostrarDatosGrid = True
        mblnLoadDenominaciones = False
    End Sub

    Private Sub btnDenominaciones_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDenominaciones.Enter
        Pon_Tool()
    End Sub

    Private Sub chkComisionBanc_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkComisionBanc.CheckStateChanged
        If chkComisionBanc.CheckState = System.Windows.Forms.CheckState.Checked Then
            lblFormasPago(3).Enabled = True
            lblFormasPago(4).Enabled = True
            txtPorcComision.Enabled = True
            txtPorcIvaComision.Enabled = True
        Else
            lblFormasPago(3).Enabled = False
            lblFormasPago(4).Enabled = False
            txtPorcComision.Enabled = False
            txtPorcIvaComision.Enabled = False
        End If
        txtPorcComision.Text = "0.00"
        txtPorcIvaComision.Text = "0.00"
    End Sub

    Private Sub chkEsTarjeta_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkEsTarjeta.CheckStateChanged
        If chkEsTarjeta.CheckState = System.Windows.Forms.CheckState.Checked Then
            chkComisionBanc.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkComisionBanc.Enabled = True
            lblBanco.Enabled = True
            dbcBanco.Enabled = True
        Else
            chkComisionBanc.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkComisionBanc.Enabled = False
            lblBanco.Enabled = False
            dbcBanco.Enabled = False
        End If
    End Sub

    Private Sub dbcBanco_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcBanco.EnabledChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcBanco" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodBanco, Ltrim(Rtrim(DescBanco)) as DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' And ControlInterno = 0 ORDER BY DescBanco "
        DCChange(gStrSql, Tecla, dbcBanco)
        intCodBanco = 0
        'If dbcBanco.SelectedItem <> "" Then
        '    Call dbcBanco_Leave(dbcBanco, New System.EventArgs())
        'End If
    End Sub

    Private Sub dbcBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcBanco.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodBanco, Ltrim(Rtrim(DescBanco)) as DescBanco FROM CatBancos Where ControlInterno = 0 ORDER BY DescBanco "
        DCGotFocus(gStrSql, dbcBanco)
    End Sub

    Private Sub dbcBanco_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcBanco.KeyDown
        Tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            eventSender.KeyCode = 0
        End If
    End Sub

    Private Sub dbcBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        intCodBanco = 0
        gStrSql = "SELECT CodBanco, Ltrim(Rtrim(DescBanco)) as DescBanco FROM CatBancos WHERE DescBanco LIKE '" & Trim(dbcBanco.Text) & "%' And ControlInterno = 0 ORDER BY DescBanco "
        DCLostFocus(dbcBanco, gStrSql, intCodBanco)
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub frmCorpoAbcFormasdePago_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoAbcFormasdePago_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmCorpoAbcFormasdePago_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        'Icono(Me, MenuPrincipal)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
        'Nuevo()
        LlenaCombo(cboTeclaRapida)
    End Sub

    Private Sub frmCorpoAbcFormasdePago_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmCorpoAbcFormasdePago_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoAbcFormasdePago_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSalir Then
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
        '    mblnSalir = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoAbcFormasdePago_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'frmCorpoFPAbcDenominaciones.Close()
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
                Case "TXTCODFORMAPAGO"
                    strCaptionForm = "Consulta de Formas de Pago"
                    gStrSql = "SELECT RIGHT('000'+LTRIM(CodFormaPago),3) AS CODIGO,DescFormaPago AS DESCRIPCION FROM CatFormasPago  ORDER BY CodFormaPago"
                Case "TXTDESCRIPCION"
                    strCaptionForm = "Consulta de Formas de Pago"
                    gStrSql = "SELECT DescFormaPago AS DESCRIPCION, RIGHT('000'+LTRIM(CodFormaPago),3) AS CODIGO FROM CatFormasPago WHERE DescFormaPago LIKE '" & Trim(txtDescripcion.Text) & "%' ORDER BY Descripcion"
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
                    Case "TXTCODFORMAPAGO"
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                    Case "TXTDESCRIPCION"
                        .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
                        .set_ColWidth(1, 0, 900) 'Columna del Código
                End Select
            End With
            FrmConsultas.ShowDialog()
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Sub Eliminar()
        ''On Local Error GoTo MErr
        'Try
        '    'Screen.MousePointer = vbHourglass Esto se manejará hasta antes de iniciar la transacción

        '    gStrSql = "SELECT DescFormaPago FROM CatFormasPago WHERE CodFormaPago=" & Val(txtCodFormaPago)

        '    ModEstandar.BorraCmd()
        '    ModEstandar.BorraCmd()
        '    Cmd.CommandText = "dbo.Up_Select_Datos"
        '    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        '    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        '    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        '    RsGral = Cmd.Execute

        '    If RsGral.RecordCount = 0 Then
        '        MsgBox("Proporcione un Código valido para eliminar.", vbExclamation + vbOKOnly, "Mensaje")
        '        'Cnn.RollbackTrans
        '        RsGral.Close()
        '        Exit Sub
        '    End If

        '    'Preguntar si desea borrar el registro
        '    Select Case MsgBox(C_msgBORRAR, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Mensaje")
        '        Case vbNo
        '            Exit Sub
        '        Case vbCancel
        '            Exit Sub
        '    End Select
        '    'Screen.MousePointer = vbHourglass
        '    Cnn.BeginTrans()

        '    'ModStoredProcedures.PR_IMECatFormasPago(Trim(txtCodFormaPago.Text), Trim(txtDescripcion.Text), Trim(cboTeclaRapida.Text), CStr(chkEsDolar.Text), CStr(chkEsCheque.CheckState), CStr(chkEsDevolucion.CheckState), CStr(chkEsDocumentoInterno.CheckState), CStr(chkReqDocumento.CheckState), CStr(chkReqAutorizacion.CheckState), CStr(chkRestCambio.CheckState), CStr(chkConsiderarFact.CheckState), CStr(chkConsiderarRetiros.CheckState), CStr(chkEsTarjeta.CheckState), CStr(chkComisionBanc.CheckState), Trim(txtPorcComision.Text), Trim(txtPorcIvaComision.Text), Trim(txtDescCorta.Text), Trim(CStr(intCodBanco)), C_MODIFICACION, CStr(0))
        '    '        CStr(chkEsCheque.Value), CStr(chkEsDevolucion.Value), CStr(chkEsDocumentoInterno), CStr(chkReqDocumento.Value), CStr(chkReqAutorizacion.Value), CStr(chkRestCambio.Value), CStr(chkConsiderarFact.Value), _
        '    '        CStr(chkConsiderarRetiros.Value), CStr(chkEsTarjeta.Value), CStr(chkComisionBanc.Value), _
        '    '        Trim(txtPorcComision), Trim(txtPorcIvaComision), " ", "", C_ELIMINACION, 0
        '    Cmd.Execute()

        '    Cnn.CommitTrans()
        '    Nuevo()
        '    Limpiar()
        '    'Screen.MousePointer = vbDefault
        '    Exit Sub
        '    'MErr:
        'Catch ex As Exception
        '    'Cnn.RollbackTrans()
        '    'Screen.MousePointer = vbDefault
        '    If Err.Number <> 0 Then ModEstandar.MostrarError()
        'End Try
    End Sub

    Function Guardar() As Boolean
        'On Error GoTo MErr
        Try
            Dim Estatus As String
            ' Si el contenido del control del codigo de la forma de pago,es un espacio en blanco, entonces, se sale de la funcion se guardar.
            ' Ya que no debe hacer nada, ni sikiera validar los datos
            If Trim(txtCodFormaPago.Text) = "" Then
                Exit Function
            End If
            'Validar si todos los datos fueron proporcionados para ser guardados,
            If ValidaDatos() = False Then
                Exit Function
            End If

            If CInt(Numerico(txtCodFormaPago.Text)) = 0 Then
                mblnNuevo = True
            End If
            dbcBanco_Leave(dbcBanco, New System.EventArgs())

            'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Cnn.BeginTrans()

            If optVigente.Checked = True Then
                Estatus = "V"
            ElseIf optSuspendido.Checked = True Then
                Estatus = "S"
            ElseIf optCancelado.Checked = True Then
                Estatus = "C"
            End If

            'Guardar los datos de las Denominaciones
            'Para lo cual primero se debe eliminar todo lo existente en  CatDenominaciones

            If mblnNuevo = True Then 'Se realizará una insercion
                ModStoredProcedures.PR_IMECatFormasPago(Trim(txtCodFormaPago.Text), Trim(txtDescripcion.Text), Trim(cboTeclaRapida.Text), CStr(chkEsDolar.CheckState), CStr(chkEsCheque.CheckState), CStr(chkEsDevolucion.CheckState), CStr(chkEsDocumentoInterno.CheckState), CStr(chkReqDocumento.CheckState), CStr(chkReqAutorizacion.CheckState), CStr(chkRestCambio.CheckState), CStr(chkConsiderarFact.CheckState), CStr(chkConsiderarRetiros.CheckState), CStr(chkEsTarjeta.CheckState), CStr(chkComisionBanc.CheckState), Trim(txtPorcComision.Text), Trim(txtPorcIvaComision.Text), Trim(txtDescCorta.Text), Estatus, Trim(CStr(intCodBanco)), C_INSERCION, CStr(0))
                Cmd.Execute()

                txtCodFormaPago.Text = Format(Cmd.Parameters("ID").Value, "000")

                If mblnLoadDenominaciones = True Then
                    'Si el Grid se Cargó, significa que se pudo haberselo hecho algun cambio, por eso se guarda.
                    'Si no se Mostró, entonces, no modificar las denominaciones , porque no se modifico nada
                    ModStoredProcedures.PR_IECatDenominaciones(CStr(txtCodFormaPago.Text), CStr("10"), C_ELIMINACION, CStr(0))
                    Cmd.Execute()
                    'Ahora realizar la alta de las nuevas denominaciones ,de una por una
                    'With frmCorpoFPAbcDenominaciones
                    '    With .FlexDenominaciones
                    '        For I = 1 To .Rows - 1
                    '            Denominacion = CDec(Numerico(.get_TextMatrix(I, C_ColDENOMINACION)))
                    '            If Denominacion <> 0 Then ' Si es mayor de cetro, guardar la denominacion
                    '                ModStoredProcedures.PR_IECatDenominaciones(CStr(txtCodFormaPago.Text), CStr(Denominacion), C_INSERCION, CStr(0))
                    '                cmd.Execute()
                    '            End If
                    '        Next
                    '        'mblnLoadDenominaciones = True
                    '    End With
                    'End With
                End If

            Else ' Se realizará una Modificación
                If mblnLoadDenominaciones = True Then
                    'Si el Grid se Cargó, significa que se pudo haberselo hecho algun cambio, por eso se guarda.
                    'Si no se Mostró, entonces, no modificar las denominaciones , porque no se modifico nada
                    ModStoredProcedures.PR_IECatDenominaciones(CStr(txtCodFormaPago.Text), CStr("10"), C_ELIMINACION, CStr(0))
                    Cmd.Execute()
                    'Ahora realizar la alta de las nuevas denominaciones ,de un apor una
                    'With frmCorpoFPAbcDenominaciones
                    '    With .FlexDenominaciones
                    '        For I = 1 To .Rows - 1
                    '            Denominacion = CDec(Numerico(.get_TextMatrix(I, C_ColDENOMINACION)))
                    '            If Denominacion <> 0 Then ' Si es mayor de cetro, guardar la denominacion
                    '                ModStoredProcedures.PR_IECatDenominaciones(CStr(txtCodFormaPago.Text), CStr(Denominacion), C_INSERCION, CStr(0))
                    '                cmd.Execute()
                    '            End If
                    '        Next
                    '    End With
                    'End With
                End If
                ModStoredProcedures.PR_IMECatFormasPago(Trim(txtCodFormaPago.Text), Trim(txtDescripcion.Text), Trim(cboTeclaRapida.Text), CStr(chkEsDolar.CheckState), CStr(chkEsCheque.CheckState), CStr(chkEsDevolucion.CheckState), CStr(chkEsDocumentoInterno.CheckState), CStr(chkReqDocumento.CheckState), CStr(chkReqAutorizacion.CheckState), CStr(chkRestCambio.CheckState), CStr(chkConsiderarFact.CheckState), CStr(chkConsiderarRetiros.CheckState), CStr(chkEsTarjeta.CheckState), CStr(chkComisionBanc.CheckState), Trim(txtPorcComision.Text), Trim(txtPorcIvaComision.Text), Trim(txtDescCorta.Text), Estatus, Trim(CStr(intCodBanco)), C_MODIFICACION, CStr(0))
                Cmd.Execute()
            End If

            Cnn.CommitTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            'Por cuestiones de estética el cambio al puntero del mouse se hace antes de iniciar la transacción y al finalizar la misma.

            If mblnNuevo Then
                MsgBox("La Forma de Pago ha sido grabada correctamente con el Código: " & txtCodFormaPago.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
            Else
                'Si se realizaron cambios el el FrmFormasPago, mostrar el mensaje de modificado
                If Cambios() = True Then
                    MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
                End If
            End If
            'Dejar el Procedimiento Nuevo, sirve para que al usar limpiar,. no pregunte si se desea guardar cambios en el codigo
            Nuevo()
            Guardar = True
            Limpiar()
            Exit Function
            'MErr:
        Catch ex As Exception
            Cnn.RollbackTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Sub Nuevo()
        'Se deben Limpiar todos los controles del formulario con excepcion del Control de la Llavve principal
        'On Error GoTo MErr
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            txtDescripcion.Text = ""
            txtDescripcion.Tag = ""
            txtDescCorta.Text = ""
            txtDescCorta.Tag = ""
            cboTeclaRapida.Text = ""
            cboTeclaRapida.Tag = ""
            chkEsDolar.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEsDolar.Tag = System.Windows.Forms.CheckState.Unchecked
            chkEsDevolucion.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEsDevolucion.Tag = System.Windows.Forms.CheckState.Unchecked
            chkEsDocumentoInterno.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEsDocumentoInterno.Tag = System.Windows.Forms.CheckState.Unchecked
            chkEsCheque.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEsCheque.Tag = System.Windows.Forms.CheckState.Unchecked
            chkReqDocumento.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkReqDocumento.Tag = System.Windows.Forms.CheckState.Unchecked
            chkReqAutorizacion.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkReqAutorizacion.Tag = System.Windows.Forms.CheckState.Unchecked
            chkEsTarjeta.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEsTarjeta.Tag = System.Windows.Forms.CheckState.Unchecked
            chkComisionBanc.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkComisionBanc.Tag = System.Windows.Forms.CheckState.Unchecked
            chkConsiderarFact.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkConsiderarFact.Tag = System.Windows.Forms.CheckState.Unchecked
            chkConsiderarRetiros.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkConsiderarRetiros.Tag = System.Windows.Forms.CheckState.Unchecked
            chkRestCambio.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRestCambio.Tag = System.Windows.Forms.CheckState.Unchecked
            txtPorcComision.Text = "0.00"
            txtPorcComision.Tag = "0.00"
            txtPorcIvaComision.Text = "0.00"
            txtPorcIvaComision.Tag = "0.00"
            btnDenominaciones.Enabled = False
            optVigente.Checked = False
            optVigente.Tag = False
            optSuspendido.Checked = False
            optSuspendido.Tag = False
            optCancelado.Checked = False
            optCancelado.Tag = False

            intCodBanco = 0
            FueraChange = True
            lblBanco.Enabled = False
            dbcBanco.Text = ""
            dbcBanco.Tag = ""
            dbcBanco.Enabled = False
            FueraChange = False
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

            If CInt(Numerico((txtCodFormaPago.Text))) = 0 Then
                Nuevo()
                Exit Sub
            End If

            FueraChange = True
            'txtCodFormaPago.Text = VB6.Format(txtCodFormaPago.Text, "000")
            For I = 0 To 2 - txtCodFormaPago.TextLength
                txtCodFormaPago.Text = String.Concat("0" + txtCodFormaPago.Text)
            Next I

            FueraChange = False
            '''gStrSql = "SELECT * FROM  CatFormasPago WHERE CodFormaPago= '" & txtCodFormaPago & "'"
            gStrSql = "SELECT FP.*, B.CodBanco, B.DescBanco " & "FROM CatFormasPago FP Left Outer Join CatBancos B On FP.CodBanco = B.CodBanco " & "WHERE FP.CodFormaPago= '" & CInt(Numerico(txtCodFormaPago.Text)) & "' "
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute

            If RsGral.RecordCount > 0 Then
                txtDescripcion.Text = Trim(RsGral.Fields("DescFormaPago").Value)
                txtDescripcion.Tag = Trim(RsGral.Fields("DescFormaPago").Value)
                txtDescCorta.Text = Trim(RsGral.Fields("DescCorta").Value)
                txtDescCorta.Tag = Trim(RsGral.Fields("DescCorta").Value)
                If RsGral.Fields("EsDolar").Value = True Then
                    chkEsDolar.CheckState = System.Windows.Forms.CheckState.Checked
                    chkEsDolar.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkEsDolar.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkEsDolar.Tag = System.Windows.Forms.CheckState.Unchecked
                End If
                If RsGral.Fields("Escheque").Value = True Then
                    chkEsCheque.CheckState = System.Windows.Forms.CheckState.Checked
                    chkEsCheque.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkEsCheque.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkEsCheque.Tag = System.Windows.Forms.CheckState.Unchecked
                End If
                If RsGral.Fields("EsDevolucion").Value = True Then
                    chkEsDevolucion.CheckState = System.Windows.Forms.CheckState.Checked
                    chkEsDevolucion.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkEsDevolucion.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkEsDevolucion.Tag = System.Windows.Forms.CheckState.Unchecked
                End If
                If RsGral.Fields("EsDocumentoInterno").Value = True Then
                    chkEsDocumentoInterno.CheckState = System.Windows.Forms.CheckState.Checked
                    chkEsDocumentoInterno.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkEsDocumentoInterno.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkEsDocumentoInterno.Tag = System.Windows.Forms.CheckState.Unchecked
                End If
                If RsGral.Fields("EsTarjeta").Value = True Then
                    chkEsTarjeta.CheckState = System.Windows.Forms.CheckState.Checked
                    chkEsTarjeta.Tag = System.Windows.Forms.CheckState.Checked
                    dbcBanco.Text = IIf(IsDBNull(RsGral.Fields("CodBanco").Value), "", RsGral.Fields("DescBanco").Value)
                    dbcBanco.Tag = IIf(IsDBNull(RsGral.Fields("CodBanco").Value), "", RsGral.Fields("DescBanco").Value)
                    intCodBanco = IIf(IsDBNull(RsGral.Fields("CodBanco").Value), 0, RsGral.Fields("CodBanco").Value)
                Else
                    chkEsTarjeta.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkEsTarjeta.Tag = System.Windows.Forms.CheckState.Unchecked
                    dbcBanco.Text = ""
                End If
                If RsGral.Fields("RequerirDocto").Value = True Then
                    chkReqDocumento.CheckState = System.Windows.Forms.CheckState.Checked
                    chkReqDocumento.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkReqDocumento.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkReqDocumento.Tag = System.Windows.Forms.CheckState.Unchecked
                End If
                If RsGral.Fields("RequerirAutorizacion").Value = True Then
                    chkReqAutorizacion.CheckState = System.Windows.Forms.CheckState.Checked
                    chkReqAutorizacion.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkReqAutorizacion.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkReqAutorizacion.Tag = System.Windows.Forms.CheckState.Unchecked
                End If
                If RsGral.Fields("DescontarComisionBanc").Value = True Then
                    chkComisionBanc.CheckState = System.Windows.Forms.CheckState.Checked
                    chkComisionBanc.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkComisionBanc.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkComisionBanc.Tag = System.Windows.Forms.CheckState.Unchecked
                End If
                If RsGral.Fields("ConsiderarParaFacturacion").Value = True Then
                    chkConsiderarFact.CheckState = System.Windows.Forms.CheckState.Checked
                    chkConsiderarFact.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkConsiderarFact.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkConsiderarFact.Tag = System.Windows.Forms.CheckState.Unchecked
                End If
                If RsGral.Fields("ConsiderarparaRetiros").Value = True Then
                    chkConsiderarRetiros.CheckState = System.Windows.Forms.CheckState.Checked
                    chkConsiderarRetiros.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkConsiderarRetiros.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkConsiderarRetiros.Tag = System.Windows.Forms.CheckState.Unchecked
                End If
                If RsGral.Fields("RestringirCambio").Value = True Then
                    chkRestCambio.CheckState = System.Windows.Forms.CheckState.Checked
                    chkRestCambio.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkRestCambio.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkRestCambio.Tag = System.Windows.Forms.CheckState.Unchecked
                End If

                txtPorcComision.Text = Format((RsGral.Fields("PorcComision").Value), "0.00")
                txtPorcComision.Tag = Format((RsGral.Fields("PorcComision").Value), "0.00")
                txtPorcIvaComision.Text = Format((RsGral.Fields("PorcIvaComision").Value), "0.00")
                txtPorcIvaComision.Tag = Format((RsGral.Fields("PorcIvaComision").Value), "0.00")

                If Trim(RsGral.Fields("Estatus").Value) = "V" Then
                    optVigente.Checked = True
                    optVigente.Tag = True
                ElseIf Trim(RsGral.Fields("Estatus").Value) = "S" Then
                    optSuspendido.Checked = True
                    optSuspendido.Tag = True
                ElseIf Trim(RsGral.Fields("Estatus").Value) = "C" Then
                    optCancelado.Checked = True
                    optCancelado.Tag = True
                End If

                LlenaCombo(cboTeclaRapida)
                For I = 0 To cboTeclaRapida.Items.Count - 1
                    cboTeclaRapida.SelectedIndex = (I)
                    If Trim(cboTeclaRapida.Text) = Trim(RsGral.Fields("TeclaRapida").Value) Then
                        cboTeclaRapida.SelectedIndex = cboTeclaRapida.SelectedIndex
                        cboTeclaRapida.Tag = cboTeclaRapida.Text
                        Exit For
                    End If
                Next
                gblnMostrarDatosGrid = True
            Else
                MsjNoExiste("La Forma de Pago", gstrNombCortoEmpresa)
                Limpiar()
            End If

            mblnCambiosEnCodigo = False
            mblnNuevo = False
            Exit Sub
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Sub LlenaCombo(ByRef cboParam As System.Windows.Forms.ComboBox)
        'On Error GoTo MErr
        Try
            Dim lStrSql As String
            Dim I As Object
            Dim J As Integer

            For I = 0 To 25
                maLetra(I) = UCase(Chr(System.Windows.Forms.Keys.A + I))
            Next

            lStrSql = "SELECT TeclaRapida FROM CatFormasPago WHERE TeclaRapida IS NOT NULL  AND COdFormaPago <> " & Numerico(txtCodFormaPago.Text) & " " & "ORDER BY TeclaRapida"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lStrSql))
            rsCombo = Cmd.Execute

            cLetras = ""
            If rsCombo.RecordCount > 0 Then
                rsCombo.MoveFirst()
                For I = 1 To rsCombo.RecordCount
                    cLetras = cLetras & UCase(rsCombo.Fields("TeclaRapida").Value)
                    rsCombo.MoveNext()
                Next I
            End If
            cboTeclaRapida.Items.Clear()
            cLetrasEnCombo = ""
            For I = 0 To 25
                If InStr(1, cLetras, maLetra(I)) = 0 Then
                    cboTeclaRapida.Items.Add(maLetra(I))
                    cLetrasEnCombo = cLetrasEnCombo & maLetra(I)
                End If
            Next I
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModErrores.Errores()
        End Try
    End Sub

    Sub Limpiar()
        'Esta función Limpia todos los controles del formulario.
        'Si hubo Cambios, Pregunta si desea guardarlos.
        'On Error GoTo MErr
        Try
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

            txtCodFormaPago.Text = ""
            Nuevo()
            LlenaCombo(cboTeclaRapida)
            If mblnLoadDenominaciones = True Then
                'frmCorpoFPAbcDenominaciones.FlexDenominaciones.Clear()
                'frmCorpoFPAbcDenominaciones.Visible = False
            End If
            mblnNuevo = True
            mblnCambiosEnCodigo = False
            txtCodFormaPago.Focus()
            Exit Sub
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Sub

    Function Cambios() As Object
        'Esta Función validará si se han efectuado cambios en los controles.
        'lo cual es útil para la funcion de guardar. Se inicializa con True, y si se validan todos los campos y no se ha
        'salido del proc. entonces la variable adquiere el valor de False
        'se validan todos los controles existentes, excepto el de la Clave Principal
        On Error GoTo MErr
        Cambios = True


        If Trim(txtDescripcion.Text) <> Trim(txtDescripcion.Tag) Then Exit Function
        If Trim(txtDescCorta.Text) <> Trim(txtDescCorta.Tag) Then Exit Function
        If Trim(cboTeclaRapida.Text) <> Trim(cboTeclaRapida.Tag) Then Exit Function
        If chkEsDolar.CheckState <> CDbl(chkEsDolar.Tag) Then Exit Function
        If chkEsCheque.CheckState <> CDbl(chkEsCheque.Tag) Then Exit Function
        If chkEsTarjeta.CheckState <> CDbl(chkEsTarjeta.Tag) Then Exit Function
        If chkEsDevolucion.CheckState <> CDbl(chkEsDevolucion.Tag) Then Exit Function
        If chkEsDocumentoInterno.CheckState <> CDbl(chkEsDocumentoInterno.Tag) Then Exit Function
        If chkReqDocumento.CheckState <> CDbl(chkReqDocumento.Tag) Then Exit Function
        If chkReqAutorizacion.CheckState <> CDbl(chkReqAutorizacion.Tag) Then Exit Function
        If chkComisionBanc.CheckState <> CDbl(chkComisionBanc.Tag) Then Exit Function
        If chkConsiderarFact.CheckState <> CDbl(chkConsiderarFact.Tag) Then Exit Function
        If chkRestCambio.CheckState <> CDbl(chkRestCambio.Tag) Then Exit Function
        If txtPorcComision.Text <> txtPorcComision.Tag Then Exit Function
        If txtPorcIvaComision.Text <> txtPorcIvaComision.Tag Then Exit Function
        If optVigente.Checked <> CBool(optVigente.Tag) Then Exit Function
        If optSuspendido.Checked <> CBool(optSuspendido.Tag) Then Exit Function
        If optCancelado.Checked <> CBool(optCancelado.Tag) Then Exit Function
        If Trim(dbcBanco.Text) <> Trim(dbcBanco.Tag) Then Exit Function

        Cambios = False

        Exit Function
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ValidaDatos() As Object
        'Esta Función Valida que todos los datos en el Formulario se introduzcan, para poder realizar la Alta del registro
        On Error GoTo MErr
        '    ValidaDatos = False No es necesario especificarlo, ya que la funcion se inicializa con falso
        If Len(Trim(txtDescripcion.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Descripción", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtDescripcion.Focus()
            Exit Function
        End If
        If Len(Trim(txtDescCorta.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Descripción Corta", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtDescCorta.Focus()
            Exit Function
        End If
        If Len(Trim(txtDescCorta.Text)) < 3 Then
            MsgBox("La Descripción Corta debe ser de  tres Caracteres." & vbNewLine & "Verifique Porfavor.", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtDescCorta.Focus()
            Exit Function
        End If
        If Len(Trim(cboTeclaRapida.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Tecla Rápida", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.cboTeclaRapida.Focus()
            Exit Function
        End If
        If InStr(1, cLetrasEnCombo, cboTeclaRapida.Text) = 0 Then
            MsgBox("Introduzca una letra de las existentes en el combo", MsgBoxStyle.Information, gstrNombCortoEmpresa)
            cboTeclaRapida.Text = ""
            cboTeclaRapida.Focus()
            Exit Function
        End If
        If chkComisionBanc.CheckState = System.Windows.Forms.CheckState.Checked Then
            If CDbl(Numerico(Trim(txtPorcComision.Text))) = 0 Then
                MsgBox(C_msgFALTADATO & "Procentaje de Comisión", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtPorcComision.Focus()
                Exit Function
            End If
        End If
        If CDbl(Numerico(Trim(txtPorcIvaComision.Text))) > 100 Then
            MsgBox("El Procentaje Iva de Comisión debe ser menor o igual a 100." & vbNewLine & "Verifique Porfavor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Me.txtPorcIvaComision.Focus()
            Exit Function
        End If
        If optVigente.Checked = False And optSuspendido.Checked = False And optCancelado.Checked = False Then
            MsgBox("Proporcione el estatus de la forma de pago.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Me.optVigente.Focus()
            Exit Function
        End If
        If (EsFormaPagoUsada(CInt(Numerico(txtCodFormaPago.Text)))) And optCancelado.Checked = True = True Then
            MsgBox("No es posible cancelar esta Forma de Pago, debido a que ya fue utilizada." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Me.optVigente.Focus()
            Exit Function
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object ValidaDatos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ValidaDatos = True
        Exit Function
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Private Sub btnDenominaciones_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDenominaciones.Click
        With frmCorpoFPAbcDenominaciones
            .Visible = True
            .Encabezado()
            If gblnMostrarDatosGrid = True Then
                .LlenaGrid()
            End If
            'Hacer que aparezca selecionado la primer fila del GRid
            With .FlexDenominaciones
                .Focus()
            End With
            .ScrollGrid()
            mblnLoadDenominaciones = True
        End With

    End Sub

    'Private Sub chkEsDivisa_Click()
    '    If chkEsDivisa.Value = vbChecked Then
    '        lblFormasPago(7).Enabled = True
    '        txtTipoCambio.Enabled = True
    '    Else
    '        lblFormasPago(7).Enabled = False
    '        txtTipoCambio.Enabled = False
    '    End If
    '    txtTipoCambio = "1.00"
    'End Sub

    'Private Sub chkPagoIntPromocion_Click()
    '    If chkPagoIntPromocion.Value = vbChecked Then
    '        lblFormasPago(5).Enabled = True
    '        lblFormasPago(6).Enabled = True
    '        txtPorcInteres.Enabled = True
    '        txtPorcIvaInteres.Enabled = True
    '    Else
    '        lblFormasPago(5).Enabled = False
    '        lblFormasPago(6).Enabled = False
    '        txtPorcInteres.Enabled = False
    '        txtPorcIvaInteres.Enabled = False
    '    End If
    '    txtPorcInteres = "0.00"
    '    txtPorcIvaInteres = "0.00"
    'End Sub

    Private Sub txtCodFormaPago_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodFormaPago.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodFormaPago_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodFormaPago.Enter
        strControlActual = UCase("txtCodFormaPago")
        SelTextoTxt(txtCodFormaPago)
        Pon_Tool()
    End Sub

    Private Sub txtCodFormaPago_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodFormaPago.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
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
                        txtCodFormaPago.Focus()
                        KeyCode = 0
                        Exit Sub
                End Select
            End If
            If KeyCode = System.Windows.Forms.Keys.Delete And mblnLoadDenominaciones = True Then
                frmCorpoFPAbcDenominaciones.FlexDenominaciones.Clear()
                frmCorpoFPAbcDenominaciones.Visible = False
            End If
        End If
    End Sub

    Private Sub txtCodFormaPago_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodFormaPago.KeyPress
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
                        txtCodFormaPago.Focus()
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

    Private Sub txtCodFormaPago_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodFormaPago.Leave
        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If Val(Trim(txtCodFormaPago.Text)) = 0 Then txtCodFormaPago.Text = "000"
        btnDenominaciones.Enabled = True
        If mblnCambiosEnCodigo = True And CDbl(Numerico(txtCodFormaPago.Text)) <> 0 Then 'si hubo cambios en el codigo hace la consulta para llenar los datos
            LlenaDatos()
        End If
    End Sub

    Private Sub txtDesccorta_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescCorta.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDescCorta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescCorta.Enter
        SelTextoTxt(txtDescCorta)
        Pon_Tool()
    End Sub
    Private Sub txtDescripcion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Enter
        strControlActual = UCase("txtDescripcion")
        SelTextoTxt(txtDescripcion)
        Pon_Tool()
    End Sub

    Private Sub txtPorcComision_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcComision.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtPorcComision_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcComision.Enter
        SelTextoTxt(txtPorcComision)
        Pon_Tool()
    End Sub

    Private Sub txtPorcComision_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPorcComision.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtPorcComision.Text, KeyAscii, 3, 2, (txtPorcComision.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPorcComision_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcComision.Leave
        ' txtPorcComision.Text = VB6.Format(Numerico(txtPorcComision.Text), "0.00")
        If CDbl(Numerico(Trim(txtPorcComision.Text))) > 100 Then
            MsgBox("El Procentaje de Comisión debe ser menor o igual a 100." & vbNewLine & "Verifique Porfavor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Me.txtPorcComision.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub txtPorcInteres_Change()
        mblnCambiosEnCodigo = True
    End Sub

    'Private Sub txtPorcInteres_GotFocus()
    '    SelTextoTxt txtPorcInteres
    '    Pon_Tool
    'End Sub

    'Private Sub txtPorcInteres_KeyPress(KeyAscii As Integer)
    '    KeyAscii = ModEstandar.MskCantidad(txtPorcInteres, KeyAscii, 3, 2, txtPorcInteres.SelStart)
    'End Sub

    'Private Sub txtPorcInteres_LostFocus()
    '    txtPorcInteres = Format(Numerico(txtPorcInteres), "0.00")
    '    If Numerico(Trim(txtPorcInteres)) > 100 Then
    '        MsgBox "El Procentaje de Interés debe ser menor o igual a 100." + vbNewLine + "Verifique Porfavor...", vbExclamation + vbOKOnly, gstrNombCortoEmpresa
    '        Me.txtPorcInteres.SetFocus
    '        Exit Sub
    '    End If
    'End Sub

    Private Sub txtPorcIvaComision_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcIvaComision.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtPorcIvaComision_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcIvaComision.Enter
        SelTextoTxt(txtPorcIvaComision)
        Pon_Tool()
    End Sub

    Private Sub txtPorcIvaComision_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPorcIvaComision.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtPorcIvaComision.Text, KeyAscii, 3, 2, (txtPorcIvaComision.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPorcIvaComision_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcIvaComision.Leave
        'txtPorcIvaComision.Text = VB6.Format(Numerico(txtPorcIvaComision.Text), "0.00")
        If CDbl(Numerico(Trim(txtPorcIvaComision.Text))) > 100 Then
            MsgBox("El Procentaje Iva de Comisión debe ser menor o igual a 100." & vbNewLine & "Verifique Porfavor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Me.txtPorcIvaComision.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub txtPorcIvaInteres_Change()
        mblnCambiosEnCodigo = True
    End Sub

    'Private Sub txtPorcIvaInteres_GotFocus()
    '    SelTextoTxt txtPorcIvaInteres
    '    Pon_Tool
    'End Sub
    '
    'Private Sub txtPorcIvaInteres_KeyPress(KeyAscii As Integer)
    '    KeyAscii = ModEstandar.MskCantidad(txtPorcIvaInteres, KeyAscii, 3, 2, txtPorcIvaInteres.SelStart)
    'End Sub

    'Private Sub txtPorcIvaInteres_LostFocus()
    '    txtPorcIvaInteres = Format(Numerico(txtPorcIvaInteres), "0.00")
    '    If Numerico(Trim(txtPorcIvaInteres)) > 100 Then
    '        MsgBox "El Procentaje Iva sobre Interés debe ser menor o igual a 100." + vbNewLine + "Verifique Porfavor...", vbExclamation + vbOKOnly, gstrNombCortoEmpresa
    '        Me.txtPorcIvaInteres.SetFocus
    '        Exit Sub
    '    End If
    'End Sub


    'Private Sub txtTipoCambio_Change()
    '    mblnCambiosEnCodigo = True
    'End Sub

    'Private Sub txtTipoCambio_GotFocus()
    '    SelTextoTxt txtTipoCambio
    '    Pon_Tool
    'End Sub

    'Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
    '    KeyAscii = ModEstandar.MskCantidad(txtTipoCambio, KeyAscii, 5, 4, txtTipoCambio.SelStart)
    'End Sub

    'Private Sub txtTipoCambio_LostFocus()
    '    txtTipoCambio = Format(Numerico(txtTipoCambio), "0.00")
    'End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtDescCorta = New System.Windows.Forms.TextBox()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me.txtCodFormaPago = New System.Windows.Forms.TextBox()
        Me.chkComisionBanc = New System.Windows.Forms.CheckBox()
        Me.chkConsiderarRetiros = New System.Windows.Forms.CheckBox()
        Me.chkConsiderarFact = New System.Windows.Forms.CheckBox()
        Me.chkReqAutorizacion = New System.Windows.Forms.CheckBox()
        Me.chkEsDolar = New System.Windows.Forms.CheckBox()
        Me.cboTeclaRapida = New System.Windows.Forms.ComboBox()
        Me.chkReqDocumento = New System.Windows.Forms.CheckBox()
        Me.chkEsCheque = New System.Windows.Forms.CheckBox()
        Me.chkEsTarjeta = New System.Windows.Forms.CheckBox()
        Me.txtPorcIvaComision = New System.Windows.Forms.TextBox()
        Me.txtPorcComision = New System.Windows.Forms.TextBox()
        Me.chkRestCambio = New System.Windows.Forms.CheckBox()
        Me.chkEsDevolucion = New System.Windows.Forms.CheckBox()
        Me.chkEsDocumentoInterno = New System.Windows.Forms.CheckBox()
        Me.lblBanco = New System.Windows.Forms.Label()
        Me.btnDenominaciones = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.fraConfiguracion = New System.Windows.Forms.GroupBox()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.optCancelado = New System.Windows.Forms.RadioButton()
        Me.optVigente = New System.Windows.Forms.RadioButton()
        Me.optSuspendido = New System.Windows.Forms.RadioButton()
        Me.Option3 = New System.Windows.Forms.RadioButton()
        Me.Option2 = New System.Windows.Forms.RadioButton()
        Me.Option1 = New System.Windows.Forms.RadioButton()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me._lblFormasPago_5 = New System.Windows.Forms.Label()
        Me._lblFormasPago_2 = New System.Windows.Forms.Label()
        Me._lblFormasPago_4 = New System.Windows.Forms.Label()
        Me._lblFormasPago_3 = New System.Windows.Forms.Label()
        Me._lblFormasPago_7 = New System.Windows.Forms.Label()
        Me._lblFormasPago_0 = New System.Windows.Forms.Label()
        Me._lblFormasPago_1 = New System.Windows.Forms.Label()
        Me.lblFormasPago = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.fraConfiguracion.SuspendLayout()
        Me.Frame3.SuspendLayout()
        CType(Me.lblFormasPago, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtDescCorta
        '
        Me.txtDescCorta.AcceptsReturn = True
        Me.txtDescCorta.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescCorta.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescCorta.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescCorta.Location = New System.Drawing.Point(105, 84)
        Me.txtDescCorta.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescCorta.MaxLength = 3
        Me.txtDescCorta.Name = "txtDescCorta"
        Me.txtDescCorta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescCorta.Size = New System.Drawing.Size(165, 20)
        Me.txtDescCorta.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtDescCorta, "Descripción corta")
        '
        'txtDescripcion
        '
        Me.txtDescripcion.AcceptsReturn = True
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescripcion.Location = New System.Drawing.Point(85, 54)
        Me.txtDescripcion.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescripcion.MaxLength = 30
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(185, 20)
        Me.txtDescripcion.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtDescripcion, "Descripción de la Forma de Pago")
        '
        'txtCodFormaPago
        '
        Me.txtCodFormaPago.AcceptsReturn = True
        Me.txtCodFormaPago.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodFormaPago.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodFormaPago.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodFormaPago.Location = New System.Drawing.Point(57, 23)
        Me.txtCodFormaPago.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodFormaPago.MaxLength = 3
        Me.txtCodFormaPago.Name = "txtCodFormaPago"
        Me.txtCodFormaPago.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodFormaPago.Size = New System.Drawing.Size(45, 20)
        Me.txtCodFormaPago.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtCodFormaPago, "Código de la Forma de Pago")
        '
        'chkComisionBanc
        '
        Me.chkComisionBanc.BackColor = System.Drawing.Color.Silver
        Me.chkComisionBanc.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkComisionBanc.Enabled = False
        Me.chkComisionBanc.ForeColor = System.Drawing.Color.Black
        Me.chkComisionBanc.Location = New System.Drawing.Point(223, 39)
        Me.chkComisionBanc.Margin = New System.Windows.Forms.Padding(2)
        Me.chkComisionBanc.Name = "chkComisionBanc"
        Me.chkComisionBanc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkComisionBanc.Size = New System.Drawing.Size(131, 17)
        Me.chkComisionBanc.TabIndex = 19
        Me.chkComisionBanc.Text = "Comisión Bancaria"
        Me.ToolTip1.SetToolTip(Me.chkComisionBanc, "Descontar Comisión Bancaria")
        Me.chkComisionBanc.UseVisualStyleBackColor = False
        '
        'chkConsiderarRetiros
        '
        Me.chkConsiderarRetiros.BackColor = System.Drawing.Color.Silver
        Me.chkConsiderarRetiros.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkConsiderarRetiros.ForeColor = System.Drawing.Color.Black
        Me.chkConsiderarRetiros.Location = New System.Drawing.Point(12, 214)
        Me.chkConsiderarRetiros.Margin = New System.Windows.Forms.Padding(2)
        Me.chkConsiderarRetiros.Name = "chkConsiderarRetiros"
        Me.chkConsiderarRetiros.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkConsiderarRetiros.Size = New System.Drawing.Size(140, 17)
        Me.chkConsiderarRetiros.TabIndex = 17
        Me.chkConsiderarRetiros.Text = "Considerar para Retiros"
        Me.ToolTip1.SetToolTip(Me.chkConsiderarRetiros, "Considerar para Retiros")
        Me.chkConsiderarRetiros.UseVisualStyleBackColor = False
        '
        'chkConsiderarFact
        '
        Me.chkConsiderarFact.BackColor = System.Drawing.Color.Silver
        Me.chkConsiderarFact.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkConsiderarFact.ForeColor = System.Drawing.Color.Black
        Me.chkConsiderarFact.Location = New System.Drawing.Point(12, 195)
        Me.chkConsiderarFact.Margin = New System.Windows.Forms.Padding(2)
        Me.chkConsiderarFact.Name = "chkConsiderarFact"
        Me.chkConsiderarFact.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkConsiderarFact.Size = New System.Drawing.Size(169, 17)
        Me.chkConsiderarFact.TabIndex = 16
        Me.chkConsiderarFact.Text = "Considerar para Facturación"
        Me.ToolTip1.SetToolTip(Me.chkConsiderarFact, "Considerar para Facturación")
        Me.chkConsiderarFact.UseVisualStyleBackColor = False
        '
        'chkReqAutorizacion
        '
        Me.chkReqAutorizacion.BackColor = System.Drawing.Color.Silver
        Me.chkReqAutorizacion.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkReqAutorizacion.ForeColor = System.Drawing.Color.Black
        Me.chkReqAutorizacion.Location = New System.Drawing.Point(12, 156)
        Me.chkReqAutorizacion.Margin = New System.Windows.Forms.Padding(2)
        Me.chkReqAutorizacion.Name = "chkReqAutorizacion"
        Me.chkReqAutorizacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkReqAutorizacion.Size = New System.Drawing.Size(140, 17)
        Me.chkReqAutorizacion.TabIndex = 14
        Me.chkReqAutorizacion.Text = "Requerir Autorización"
        Me.ToolTip1.SetToolTip(Me.chkReqAutorizacion, "Requiere de Autorización")
        Me.chkReqAutorizacion.UseVisualStyleBackColor = False
        '
        'chkEsDolar
        '
        Me.chkEsDolar.BackColor = System.Drawing.Color.Silver
        Me.chkEsDolar.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEsDolar.ForeColor = System.Drawing.Color.Black
        Me.chkEsDolar.Location = New System.Drawing.Point(12, 58)
        Me.chkEsDolar.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEsDolar.Name = "chkEsDolar"
        Me.chkEsDolar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEsDolar.Size = New System.Drawing.Size(75, 17)
        Me.chkEsDolar.TabIndex = 9
        Me.chkEsDolar.Text = "Es Dólar"
        Me.ToolTip1.SetToolTip(Me.chkEsDolar, "Es Dólar")
        Me.chkEsDolar.UseVisualStyleBackColor = False
        '
        'cboTeclaRapida
        '
        Me.cboTeclaRapida.BackColor = System.Drawing.SystemColors.Window
        Me.cboTeclaRapida.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboTeclaRapida.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboTeclaRapida.Location = New System.Drawing.Point(92, 30)
        Me.cboTeclaRapida.Margin = New System.Windows.Forms.Padding(2)
        Me.cboTeclaRapida.Name = "cboTeclaRapida"
        Me.cboTeclaRapida.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboTeclaRapida.Size = New System.Drawing.Size(79, 21)
        Me.cboTeclaRapida.Sorted = True
        Me.cboTeclaRapida.TabIndex = 8
        Me.cboTeclaRapida.Text = "cboTeclaRapida"
        Me.ToolTip1.SetToolTip(Me.cboTeclaRapida, "Tecla Rápida")
        '
        'chkReqDocumento
        '
        Me.chkReqDocumento.BackColor = System.Drawing.Color.Silver
        Me.chkReqDocumento.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkReqDocumento.ForeColor = System.Drawing.Color.Black
        Me.chkReqDocumento.Location = New System.Drawing.Point(12, 136)
        Me.chkReqDocumento.Margin = New System.Windows.Forms.Padding(2)
        Me.chkReqDocumento.Name = "chkReqDocumento"
        Me.chkReqDocumento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkReqDocumento.Size = New System.Drawing.Size(125, 17)
        Me.chkReqDocumento.TabIndex = 13
        Me.chkReqDocumento.Text = "Requerir Documento"
        Me.ToolTip1.SetToolTip(Me.chkReqDocumento, "Requerir Documento")
        Me.chkReqDocumento.UseVisualStyleBackColor = False
        '
        'chkEsCheque
        '
        Me.chkEsCheque.BackColor = System.Drawing.Color.Silver
        Me.chkEsCheque.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEsCheque.ForeColor = System.Drawing.Color.Black
        Me.chkEsCheque.Location = New System.Drawing.Point(12, 78)
        Me.chkEsCheque.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEsCheque.Name = "chkEsCheque"
        Me.chkEsCheque.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEsCheque.Size = New System.Drawing.Size(140, 17)
        Me.chkEsCheque.TabIndex = 10
        Me.chkEsCheque.Text = "Es Cheque"
        Me.ToolTip1.SetToolTip(Me.chkEsCheque, "Es Cheque")
        Me.chkEsCheque.UseVisualStyleBackColor = False
        '
        'chkEsTarjeta
        '
        Me.chkEsTarjeta.BackColor = System.Drawing.Color.Silver
        Me.chkEsTarjeta.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEsTarjeta.ForeColor = System.Drawing.Color.Black
        Me.chkEsTarjeta.Location = New System.Drawing.Point(223, 17)
        Me.chkEsTarjeta.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEsTarjeta.Name = "chkEsTarjeta"
        Me.chkEsTarjeta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEsTarjeta.Size = New System.Drawing.Size(125, 17)
        Me.chkEsTarjeta.TabIndex = 18
        Me.chkEsTarjeta.Text = "Es Tarjeta"
        Me.ToolTip1.SetToolTip(Me.chkEsTarjeta, "Es Tarjeta")
        Me.chkEsTarjeta.UseVisualStyleBackColor = False
        '
        'txtPorcIvaComision
        '
        Me.txtPorcIvaComision.AcceptsReturn = True
        Me.txtPorcIvaComision.BackColor = System.Drawing.SystemColors.Window
        Me.txtPorcIvaComision.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPorcIvaComision.Enabled = False
        Me.txtPorcIvaComision.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPorcIvaComision.Location = New System.Drawing.Point(323, 101)
        Me.txtPorcIvaComision.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPorcIvaComision.MaxLength = 0
        Me.txtPorcIvaComision.Name = "txtPorcIvaComision"
        Me.txtPorcIvaComision.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPorcIvaComision.Size = New System.Drawing.Size(44, 20)
        Me.txtPorcIvaComision.TabIndex = 23
        Me.txtPorcIvaComision.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtPorcIvaComision, "Porcentaje de IVA de comisión")
        '
        'txtPorcComision
        '
        Me.txtPorcComision.AcceptsReturn = True
        Me.txtPorcComision.BackColor = System.Drawing.SystemColors.Window
        Me.txtPorcComision.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPorcComision.Enabled = False
        Me.txtPorcComision.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPorcComision.Location = New System.Drawing.Point(307, 74)
        Me.txtPorcComision.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPorcComision.MaxLength = 0
        Me.txtPorcComision.Name = "txtPorcComision"
        Me.txtPorcComision.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPorcComision.Size = New System.Drawing.Size(44, 20)
        Me.txtPorcComision.TabIndex = 21
        Me.txtPorcComision.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtPorcComision, "Porcentaje de Comisión Bancaria")
        '
        'chkRestCambio
        '
        Me.chkRestCambio.BackColor = System.Drawing.Color.Silver
        Me.chkRestCambio.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkRestCambio.ForeColor = System.Drawing.Color.Black
        Me.chkRestCambio.Location = New System.Drawing.Point(12, 176)
        Me.chkRestCambio.Margin = New System.Windows.Forms.Padding(2)
        Me.chkRestCambio.Name = "chkRestCambio"
        Me.chkRestCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkRestCambio.Size = New System.Drawing.Size(110, 17)
        Me.chkRestCambio.TabIndex = 15
        Me.chkRestCambio.Text = "Restringir Cambio"
        Me.ToolTip1.SetToolTip(Me.chkRestCambio, "Restringir Cambio")
        Me.chkRestCambio.UseVisualStyleBackColor = False
        '
        'chkEsDevolucion
        '
        Me.chkEsDevolucion.BackColor = System.Drawing.Color.Silver
        Me.chkEsDevolucion.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEsDevolucion.ForeColor = System.Drawing.Color.Black
        Me.chkEsDevolucion.Location = New System.Drawing.Point(12, 98)
        Me.chkEsDevolucion.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEsDevolucion.Name = "chkEsDevolucion"
        Me.chkEsDevolucion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEsDevolucion.Size = New System.Drawing.Size(140, 17)
        Me.chkEsDevolucion.TabIndex = 11
        Me.chkEsDevolucion.Text = "Es Devolución"
        Me.ToolTip1.SetToolTip(Me.chkEsDevolucion, "Es Devolución")
        Me.chkEsDevolucion.UseVisualStyleBackColor = False
        '
        'chkEsDocumentoInterno
        '
        Me.chkEsDocumentoInterno.BackColor = System.Drawing.Color.Silver
        Me.chkEsDocumentoInterno.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEsDocumentoInterno.ForeColor = System.Drawing.Color.Black
        Me.chkEsDocumentoInterno.Location = New System.Drawing.Point(12, 117)
        Me.chkEsDocumentoInterno.Margin = New System.Windows.Forms.Padding(2)
        Me.chkEsDocumentoInterno.Name = "chkEsDocumentoInterno"
        Me.chkEsDocumentoInterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEsDocumentoInterno.Size = New System.Drawing.Size(98, 17)
        Me.chkEsDocumentoInterno.TabIndex = 12
        Me.chkEsDocumentoInterno.Text = "Es Documento Interno"
        Me.ToolTip1.SetToolTip(Me.chkEsDocumentoInterno, "Es Documento Interno")
        Me.chkEsDocumentoInterno.UseVisualStyleBackColor = False
        '
        'lblBanco
        '
        Me.lblBanco.BackColor = System.Drawing.Color.Silver
        Me.lblBanco.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBanco.Enabled = False
        Me.lblBanco.ForeColor = System.Drawing.Color.Black
        Me.lblBanco.Location = New System.Drawing.Point(203, 136)
        Me.lblBanco.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblBanco.Name = "lblBanco"
        Me.lblBanco.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBanco.Size = New System.Drawing.Size(57, 15)
        Me.lblBanco.TabIndex = 36
        Me.lblBanco.Text = "Banco :"
        Me.ToolTip1.SetToolTip(Me.lblBanco, "Nombre de la Farmacia Actual")
        '
        'btnDenominaciones
        '
        Me.btnDenominaciones.BackColor = System.Drawing.SystemColors.Control
        Me.btnDenominaciones.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnDenominaciones.Enabled = False
        Me.btnDenominaciones.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDenominaciones.Location = New System.Drawing.Point(317, 11)
        Me.btnDenominaciones.Margin = New System.Windows.Forms.Padding(2)
        Me.btnDenominaciones.Name = "btnDenominaciones"
        Me.btnDenominaciones.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnDenominaciones.Size = New System.Drawing.Size(114, 42)
        Me.btnDenominaciones.TabIndex = 30
        Me.btnDenominaciones.Text = "&Denominaciones"
        Me.ToolTip1.SetToolTip(Me.btnDenominaciones, "Denominaciones en Forma de Pago")
        Me.btnDenominaciones.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.Color.Silver
        Me.Frame1.Controls.Add(Me.btnSalir)
        Me.Frame1.Controls.Add(Me.txtDescCorta)
        Me.Frame1.Controls.Add(Me.btnBuscar)
        Me.Frame1.Controls.Add(Me.btnLimpiar)
        Me.Frame1.Controls.Add(Me.txtDescripcion)
        Me.Frame1.Controls.Add(Me.btnEliminar)
        Me.Frame1.Controls.Add(Me.txtCodFormaPago)
        Me.Frame1.Controls.Add(Me.fraConfiguracion)
        Me.Frame1.Controls.Add(Me.btnGuardar)
        Me.Frame1.Controls.Add(Me.btnDenominaciones)
        Me.Frame1.Controls.Add(Me._lblFormasPago_7)
        Me.Frame1.Controls.Add(Me._lblFormasPago_0)
        Me.Frame1.Controls.Add(Me._lblFormasPago_1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(11, 8)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(436, 493)
        Me.Frame1.TabIndex = 28
        Me.Frame1.TabStop = False
        '
        'fraConfiguracion
        '
        Me.fraConfiguracion.BackColor = System.Drawing.Color.Silver
        Me.fraConfiguracion.Controls.Add(Me.dbcBanco)
        Me.fraConfiguracion.Controls.Add(Me.Frame3)
        Me.fraConfiguracion.Controls.Add(Me.Option3)
        Me.fraConfiguracion.Controls.Add(Me.Option2)
        Me.fraConfiguracion.Controls.Add(Me.Option1)
        Me.fraConfiguracion.Controls.Add(Me.chkComisionBanc)
        Me.fraConfiguracion.Controls.Add(Me.chkConsiderarRetiros)
        Me.fraConfiguracion.Controls.Add(Me.chkConsiderarFact)
        Me.fraConfiguracion.Controls.Add(Me.chkReqAutorizacion)
        Me.fraConfiguracion.Controls.Add(Me.chkEsDolar)
        Me.fraConfiguracion.Controls.Add(Me.cboTeclaRapida)
        Me.fraConfiguracion.Controls.Add(Me.Frame2)
        Me.fraConfiguracion.Controls.Add(Me.chkReqDocumento)
        Me.fraConfiguracion.Controls.Add(Me.chkEsCheque)
        Me.fraConfiguracion.Controls.Add(Me.chkEsTarjeta)
        Me.fraConfiguracion.Controls.Add(Me.txtPorcIvaComision)
        Me.fraConfiguracion.Controls.Add(Me.txtPorcComision)
        Me.fraConfiguracion.Controls.Add(Me.chkRestCambio)
        Me.fraConfiguracion.Controls.Add(Me.chkEsDevolucion)
        Me.fraConfiguracion.Controls.Add(Me.chkEsDocumentoInterno)
        Me.fraConfiguracion.Controls.Add(Me.lblBanco)
        Me.fraConfiguracion.Controls.Add(Me._lblFormasPago_5)
        Me.fraConfiguracion.Controls.Add(Me._lblFormasPago_2)
        Me.fraConfiguracion.Controls.Add(Me._lblFormasPago_4)
        Me.fraConfiguracion.Controls.Add(Me._lblFormasPago_3)
        Me.fraConfiguracion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraConfiguracion.Location = New System.Drawing.Point(10, 136)
        Me.fraConfiguracion.Margin = New System.Windows.Forms.Padding(2)
        Me.fraConfiguracion.Name = "fraConfiguracion"
        Me.fraConfiguracion.Padding = New System.Windows.Forms.Padding(2)
        Me.fraConfiguracion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraConfiguracion.Size = New System.Drawing.Size(385, 256)
        Me.fraConfiguracion.TabIndex = 6
        Me.fraConfiguracion.TabStop = False
        Me.fraConfiguracion.Text = " Configuración "
        '
        'dbcBanco
        '
        Me.dbcBanco.FormattingEnabled = True
        Me.dbcBanco.Location = New System.Drawing.Point(244, 135)
        Me.dbcBanco.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(123, 21)
        Me.dbcBanco.TabIndex = 37
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.Color.Silver
        Me.Frame3.Controls.Add(Me.optCancelado)
        Me.Frame3.Controls.Add(Me.optVigente)
        Me.Frame3.Controls.Add(Me.optSuspendido)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(193, 158)
        Me.Frame3.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(188, 74)
        Me.Frame3.TabIndex = 24
        Me.Frame3.TabStop = False
        Me.Frame3.Text = " Estatus "
        '
        'optCancelado
        '
        Me.optCancelado.BackColor = System.Drawing.Color.Silver
        Me.optCancelado.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCancelado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCancelado.Location = New System.Drawing.Point(72, 52)
        Me.optCancelado.Margin = New System.Windows.Forms.Padding(2)
        Me.optCancelado.Name = "optCancelado"
        Me.optCancelado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCancelado.Size = New System.Drawing.Size(85, 17)
        Me.optCancelado.TabIndex = 27
        Me.optCancelado.TabStop = True
        Me.optCancelado.Text = "Cancelado"
        Me.optCancelado.UseVisualStyleBackColor = False
        '
        'optVigente
        '
        Me.optVigente.BackColor = System.Drawing.Color.Silver
        Me.optVigente.Cursor = System.Windows.Forms.Cursors.Default
        Me.optVigente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optVigente.Location = New System.Drawing.Point(72, 13)
        Me.optVigente.Margin = New System.Windows.Forms.Padding(2)
        Me.optVigente.Name = "optVigente"
        Me.optVigente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optVigente.Size = New System.Drawing.Size(85, 17)
        Me.optVigente.TabIndex = 25
        Me.optVigente.TabStop = True
        Me.optVigente.Text = "Vigente"
        Me.optVigente.UseVisualStyleBackColor = False
        '
        'optSuspendido
        '
        Me.optSuspendido.BackColor = System.Drawing.Color.Silver
        Me.optSuspendido.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSuspendido.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSuspendido.Location = New System.Drawing.Point(72, 32)
        Me.optSuspendido.Margin = New System.Windows.Forms.Padding(2)
        Me.optSuspendido.Name = "optSuspendido"
        Me.optSuspendido.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSuspendido.Size = New System.Drawing.Size(85, 15)
        Me.optSuspendido.TabIndex = 26
        Me.optSuspendido.TabStop = True
        Me.optSuspendido.Text = "Suspendido"
        Me.optSuspendido.UseVisualStyleBackColor = False
        '
        'Option3
        '
        Me.Option3.BackColor = System.Drawing.SystemColors.Control
        Me.Option3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Option3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Option3.Location = New System.Drawing.Point(258, 188)
        Me.Option3.Margin = New System.Windows.Forms.Padding(2)
        Me.Option3.Name = "Option3"
        Me.Option3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Option3.Size = New System.Drawing.Size(61, 11)
        Me.Option3.TabIndex = 33
        Me.Option3.TabStop = True
        Me.Option3.Text = "Cancelado"
        Me.Option3.UseVisualStyleBackColor = False
        '
        'Option2
        '
        Me.Option2.BackColor = System.Drawing.SystemColors.Control
        Me.Option2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Option2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Option2.Location = New System.Drawing.Point(253, 174)
        Me.Option2.Margin = New System.Windows.Forms.Padding(2)
        Me.Option2.Name = "Option2"
        Me.Option2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Option2.Size = New System.Drawing.Size(61, 14)
        Me.Option2.TabIndex = 32
        Me.Option2.TabStop = True
        Me.Option2.Text = "Vigente"
        Me.Option2.UseVisualStyleBackColor = False
        '
        'Option1
        '
        Me.Option1.BackColor = System.Drawing.SystemColors.Control
        Me.Option1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Option1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Option1.Location = New System.Drawing.Point(258, 169)
        Me.Option1.Margin = New System.Windows.Forms.Padding(2)
        Me.Option1.Name = "Option1"
        Me.Option1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Option1.Size = New System.Drawing.Size(61, 11)
        Me.Option1.TabIndex = 31
        Me.Option1.TabStop = True
        Me.Option1.Text = "Suspendido"
        Me.Option1.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(180, 17)
        Me.Frame2.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(2, 176)
        Me.Frame2.TabIndex = 29
        Me.Frame2.TabStop = False
        '
        '_lblFormasPago_5
        '
        Me._lblFormasPago_5.AutoSize = True
        Me._lblFormasPago_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblFormasPago_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormasPago_5.Enabled = False
        Me._lblFormasPago_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblFormasPago_5.Location = New System.Drawing.Point(199, 174)
        Me._lblFormasPago_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblFormasPago_5.Name = "_lblFormasPago_5"
        Me._lblFormasPago_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormasPago_5.Size = New System.Drawing.Size(48, 13)
        Me._lblFormasPago_5.TabIndex = 34
        Me._lblFormasPago_5.Text = "Estatus :"
        '
        '_lblFormasPago_2
        '
        Me._lblFormasPago_2.AutoSize = True
        Me._lblFormasPago_2.BackColor = System.Drawing.Color.Silver
        Me._lblFormasPago_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormasPago_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblFormasPago_2.Location = New System.Drawing.Point(12, 32)
        Me._lblFormasPago_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblFormasPago_2.Name = "_lblFormasPago_2"
        Me._lblFormasPago_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormasPago_2.Size = New System.Drawing.Size(77, 13)
        Me._lblFormasPago_2.TabIndex = 7
        Me._lblFormasPago_2.Text = "Tecla Rápida :"
        '
        '_lblFormasPago_4
        '
        Me._lblFormasPago_4.AutoSize = True
        Me._lblFormasPago_4.BackColor = System.Drawing.Color.Silver
        Me._lblFormasPago_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormasPago_4.Enabled = False
        Me._lblFormasPago_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblFormasPago_4.Location = New System.Drawing.Point(203, 103)
        Me._lblFormasPago_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblFormasPago_4.Name = "_lblFormasPago_4"
        Me._lblFormasPago_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormasPago_4.Size = New System.Drawing.Size(116, 13)
        Me._lblFormasPago_4.TabIndex = 22
        Me._lblFormasPago_4.Text = "% de IVA de Comisión :"
        '
        '_lblFormasPago_3
        '
        Me._lblFormasPago_3.AutoSize = True
        Me._lblFormasPago_3.BackColor = System.Drawing.Color.Silver
        Me._lblFormasPago_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormasPago_3.Enabled = False
        Me._lblFormasPago_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblFormasPago_3.Location = New System.Drawing.Point(220, 76)
        Me._lblFormasPago_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblFormasPago_3.Name = "_lblFormasPago_3"
        Me._lblFormasPago_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormasPago_3.Size = New System.Drawing.Size(81, 13)
        Me._lblFormasPago_3.TabIndex = 20
        Me._lblFormasPago_3.Text = "% de Comisión :"
        '
        '_lblFormasPago_7
        '
        Me._lblFormasPago_7.AutoSize = True
        Me._lblFormasPago_7.BackColor = System.Drawing.Color.Silver
        Me._lblFormasPago_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormasPago_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblFormasPago_7.Location = New System.Drawing.Point(8, 84)
        Me._lblFormasPago_7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblFormasPago_7.Name = "_lblFormasPago_7"
        Me._lblFormasPago_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormasPago_7.Size = New System.Drawing.Size(94, 13)
        Me._lblFormasPago_7.TabIndex = 4
        Me._lblFormasPago_7.Text = "Descripción Corta:"
        '
        '_lblFormasPago_0
        '
        Me._lblFormasPago_0.AutoSize = True
        Me._lblFormasPago_0.BackColor = System.Drawing.Color.Silver
        Me._lblFormasPago_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormasPago_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblFormasPago_0.Location = New System.Drawing.Point(8, 25)
        Me._lblFormasPago_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblFormasPago_0.Name = "_lblFormasPago_0"
        Me._lblFormasPago_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormasPago_0.Size = New System.Drawing.Size(46, 13)
        Me._lblFormasPago_0.TabIndex = 0
        Me._lblFormasPago_0.Text = "Código :"
        '
        '_lblFormasPago_1
        '
        Me._lblFormasPago_1.AutoSize = True
        Me._lblFormasPago_1.BackColor = System.Drawing.Color.Silver
        Me._lblFormasPago_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormasPago_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblFormasPago_1.Location = New System.Drawing.Point(8, 58)
        Me._lblFormasPago_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblFormasPago_1.Name = "_lblFormasPago_1"
        Me._lblFormasPago_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormasPago_1.Size = New System.Drawing.Size(69, 13)
        Me._lblFormasPago_1.TabIndex = 2
        Me._lblFormasPago_1.Text = "Descripción :"
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Location = New System.Drawing.Point(236, 405)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(93, 35)
        Me.btnLimpiar.TabIndex = 59
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(138, 405)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(93, 35)
        Me.btnEliminar.TabIndex = 58
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.Location = New System.Drawing.Point(39, 404)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(93, 35)
        Me.btnGuardar.TabIndex = 57
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(98, 445)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(93, 35)
        Me.btnBuscar.TabIndex = 60
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnSalir
        '
        Me.btnSalir.Location = New System.Drawing.Point(197, 446)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(93, 35)
        Me.btnSalir.TabIndex = 60
        Me.btnSalir.Text = "Salir"
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'frmCorpoAbcFormasdePago
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(458, 512)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 15)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoAbcFormasdePago"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Formas de Pago"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.fraConfiguracion.ResumeLayout(False)
        Me.fraConfiguracion.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        CType(Me.lblFormasPago, System.ComponentModel.ISupportInitialize).EndInit()
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

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub
End Class