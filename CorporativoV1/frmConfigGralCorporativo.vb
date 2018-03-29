Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmConfigGralCorporativo
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    'Programa: ABC de Bancos
    'Autor: Rosaura Torres López
    'Fecha de Creación: 14/Mayo/2003
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdCuentasNotificaciones As System.Windows.Forms.Button
    Public WithEvents optHoras As System.Windows.Forms.RadioButton
    Public WithEvents optMinutos As System.Windows.Forms.RadioButton
    Public WithEvents txtCodificacion As System.Windows.Forms.TextBox
    Public WithEvents txtDiferenciaStock As System.Windows.Forms.TextBox
    Public WithEvents _Label_8 As System.Windows.Forms.Label
    Public WithEvents _Label_7 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents txtTipoCambioDolar As System.Windows.Forms.TextBox
    Public WithEvents txtTipoCambioEuro As System.Windows.Forms.TextBox
    Public WithEvents _Label_5 As System.Windows.Forms.Label
    Public WithEvents _Label_6 As System.Windows.Forms.Label
    Public WithEvents fraTiposCambio As System.Windows.Forms.GroupBox
    Public WithEvents txtUtilidadMinima As System.Windows.Forms.TextBox
    Public WithEvents txtVigenciaApartado As System.Windows.Forms.TextBox
    Public WithEvents _Label_0 As System.Windows.Forms.Label
    Public WithEvents _Label_3 As System.Windows.Forms.Label
    Public WithEvents fraVentas As System.Windows.Forms.GroupBox
    Public WithEvents txtNombreEmpresa As System.Windows.Forms.TextBox
    Public WithEvents txtDomicilioEmpresa As System.Windows.Forms.TextBox
    Public WithEvents txtRFCEmpresa As System.Windows.Forms.TextBox
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents fraInformacionEmpresa As System.Windows.Forms.GroupBox
    Public WithEvents chkTransferenciasEntreSucursales As System.Windows.Forms.CheckBox
    Public WithEvents cboDriveLocal As System.Windows.Forms.ComboBox
    Public WithEvents btnDirectorioImagenes As System.Windows.Forms.Button
    Public WithEvents txtDirectorioImagenes As System.Windows.Forms.TextBox
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label_1 As System.Windows.Forms.Label
    Public WithEvents fraImagenes As System.Windows.Forms.GroupBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray



    Private Const C_TITLEIMAGES As String = "Seleccione el Directorio de Imágenes ..."


    Dim mblnCambiosEnCodigo As Boolean
    Dim mblnFueraChange As Boolean
    Dim tecla As Integer
    Dim mblnNuevo As Boolean
    Dim I As Integer
    Dim rsLocal As ADODB.Recordset
    Dim mblnSalir As Boolean
    Public mintCuentas As Integer
    Dim mblnMuestraFormaCuentasN As Boolean

    ' Para Manejar el FlexGrid
    Const C_ColCODGRUPO As Integer = 0
    Const C_ColDESGRUPO As Integer = 1
    Const C_ColIMPORTE As Integer = 2
    Const C_ColIEPS As Integer = 3
    Const C_COLIMPORTETAG As Integer = 4
    Public WithEvents btnGuardar As Button
    Const C_ColIEPSTAG As Integer = 5

    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        mintCuentas = 0
        mblnMuestraFormaCuentasN = False
    End Sub

    Sub LlenaDatos()
        On Error GoTo Merr
        Dim FormatoDif As String
        gStrSql = "SELECT * FROM  ConfiguracionGeneral"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            txtNombreEmpresa.Text = Trim(RsGral.Fields("NombreEmp").Value)
            txtNombreEmpresa.Tag = Trim(RsGral.Fields("NombreEmp").Value)
            txtRFCEmpresa.Text = Trim(RsGral.Fields("RFCEmp").Value)
            txtRFCEmpresa.Tag = Trim(RsGral.Fields("RFCEmp").Value)
            txtDomicilioEmpresa.Text = Trim(RsGral.Fields("DomicilioEmp").Value)
            txtDomicilioEmpresa.Tag = Trim(RsGral.Fields("DomicilioEmp").Value)
            txtTipoCambioDolar.Text = VB6.Format(RsGral.Fields("TipoCambio").Value, "0.00")
            txtTipoCambioDolar.Tag = VB6.Format(RsGral.Fields("TipoCambio").Value, "0.00")
            txtTipoCambioEuro.Text = VB6.Format(RsGral.Fields("TipoCambioEuro").Value, "0.00")
            txtTipoCambioEuro.Tag = VB6.Format(RsGral.Fields("TipoCambioEuro").Value, "0.00")
            txtVigenciaApartado.Text = Numerico(RsGral.Fields("VigenciaApartados").Value)
            txtVigenciaApartado.Tag = Numerico(RsGral.Fields("VigenciaApartados").Value)
            txtUtilidadMinima.Text = VB6.Format(Numerico(RsGral.Fields("PorcUtilMinOperacion").Value), "0.00")
            txtUtilidadMinima.Tag = VB6.Format(Numerico(RsGral.Fields("PorcUtilMinOperacion").Value), "0.00")
            txtDirectorioImagenes.Text = Trim(RsGral.Fields("RutaImagenes").Value)
            txtDirectorioImagenes.Tag = Trim(RsGral.Fields("RutaImagenes").Value)
            cboDriveLocal.Text = Trim(RsGral.Fields("DriveLocal").Value)
            cboDriveLocal.Tag = Trim(RsGral.Fields("DriveLocal").Value)
            txtCodificacion.Text = Trim(RsGral.Fields("Codificacion").Value)
            txtCodificacion.Tag = Trim(RsGral.Fields("Codificacion").Value)
            FormatoDif = (Trim(RsGral.Fields("LapsoDifStock").Value))
            If FormatoDif = "H1" Then
                optHoras.Checked = True
                optHoras.Tag = 1
                optMinutos.Tag = 0
            ElseIf FormatoDif = "M1" Then
                optMinutos.Checked = True
                optMinutos.Tag = 1
                optHoras.Tag = 0
            End If
            txtDiferenciaStock.Text = Mid(Trim(RsGral.Fields("LapsoDifStock").Value), 2, Len(Trim(RsGral.Fields("LapsoDifStock").Value)))
            txtDiferenciaStock.Tag = txtDiferenciaStock.Text

            If RsGral.Fields("TransferenciasentreSucursales").Value = False Then
                chkTransferenciasEntreSucursales.Tag = System.Windows.Forms.CheckState.Unchecked
                chkTransferenciasEntreSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkTransferenciasEntreSucursales.Tag = System.Windows.Forms.CheckState.Checked
                chkTransferenciasEntreSucursales.CheckState = System.Windows.Forms.CheckState.Checked
            End If
        Else
            MsgBox("No existe informacion de configuración almacenada" & vbNewLine & "Verifique Porfavor...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If

        mblnCambiosEnCodigo = False
        mblnNuevo = False
        Exit Sub
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub


    Function Cambios() As Object
        'Esta Función validará si se han efectuado cambios en los controles.
        'lo cual es útil para la funcion de guardar. Se inicializa con True, y si se validan todos los campos y no se ha
        'salido del proc. entonces la variable adquiere el valor de False
        'se validan todos los controles existentes, excepto el de la Clave Principal
        On Error GoTo Merr
        Cambios = True

        If Trim(txtNombreEmpresa.Text) <> Trim(txtNombreEmpresa.Tag) Then Exit Function
        If Trim(txtDomicilioEmpresa.Text) <> Trim(txtDomicilioEmpresa.Tag) Then Exit Function
        If Trim(txtRFCEmpresa.Text) <> Trim(txtRFCEmpresa.Tag) Then Exit Function
        If Trim(txtTipoCambioDolar.Text) <> Trim(txtTipoCambioDolar.Tag) Then Exit Function
        If Trim(txtTipoCambioEuro.Text) <> Trim(txtTipoCambioEuro.Tag) Then Exit Function
        If Trim(txtVigenciaApartado.Text) <> Trim(txtVigenciaApartado.Tag) Then Exit Function
        If Numerico(Trim(txtUtilidadMinima.Text)) <> Numerico(Trim(txtUtilidadMinima.Tag)) Then Exit Function
        If Trim(txtDirectorioImagenes.Text) <> Trim(txtDirectorioImagenes.Tag) Then Exit Function
        If Trim(cboDriveLocal.Text) <> Trim(cboDriveLocal.Tag) Then Exit Function
        If chkTransferenciasEntreSucursales.CheckState <> CDbl(chkTransferenciasEntreSucursales.Tag) Then Exit Function
        If Trim(txtCodificacion.Tag) <> Trim(txtCodificacion.Text) Then Exit Function
        'If Trim(txtDiferenciaStock) <> Trim(txtDiferenciaStock.Tag) Then Exit Function

        If optHoras.Checked <> CBool(optHoras.Tag) Then Exit Function
        If optMinutos.Checked <> CBool(optMinutos.Tag) Then Exit Function
        If Trim(txtDiferenciaStock.Text) <> Trim(txtDiferenciaStock.Tag) Then Exit Function

        Cambios = False
        Exit Function
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function ValidaDatos() As Object
        'Esta Función Valida que todos los datos en el Formulario se introduzcan, para poder realizar la Alta del registro
        On Error GoTo Merr
        '    ValidaDatos = False No es necesario especificarlo, ya que la funcion se inicializa con falso
        If Trim(txtNombreEmpresa.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Nombre de la Empresa", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtNombreEmpresa.Focus()
            Exit Function
        End If
        If Trim(txtDomicilioEmpresa.Text) = "" Then
            MsgBox(C_msgFALTADATO & "Domicilio de la empresa", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtDomicilioEmpresa.Focus()
            Exit Function
        End If
        If Trim(txtRFCEmpresa.Text) = "" Then
            MsgBox(C_msgFALTADATO & "RFC de la empresa", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtRFCEmpresa.Focus()
            Exit Function
        End If

        If ModEstandar.valida_RFCC(Trim(txtRFCEmpresa.Text)) = False Then
            MsgBox("Proporcione un RFC válido para la empresa.", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtRFCEmpresa.Focus()
            Exit Function
        End If
        If CDbl(Numerico(txtTipoCambioDolar.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Tipo de cambio del dólar", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtTipoCambioDolar.Focus()
            Exit Function
        End If
        If CDbl(Numerico(txtTipoCambioEuro.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Tipo de cambio del euro.", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtTipoCambioEuro.Focus()
            Exit Function
        End If
        If CDbl(Numerico(txtVigenciaApartado.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Vigencia de apartados", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtVigenciaApartado.Focus()
            Exit Function
        End If

        If CDbl(Numerico(txtUtilidadMinima.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Utilidad mínima por operación", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtUtilidadMinima.Focus()
            Exit Function
        End If
        If Len(Trim(txtDirectorioImagenes.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Directorio de imágenes", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtDirectorioImagenes.Focus()
            Exit Function
        End If
        If CDbl(Numerico(Trim(txtUtilidadMinima.Text))) > 100 Then
            MsgBox("EL porcentaje de utilidad mínima por operación debe ser menor o igual a 100." & vbNewLine & "Verifique Por favor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Me.txtUtilidadMinima.Focus()
            Exit Function
        End If
        If Len(Trim(txtCodificacion.Text)) = 0 Then
            MsgBox(C_msgFALTADATO & "Clave de codificación", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtCodificacion.Focus()
            Exit Function
        End If
        If CDbl(Numerico(Trim(txtDiferenciaStock.Text))) <= 0 Then
            MsgBox(C_msgFALTADATO & "Lapso de tiempo para verificar existencia y stock.", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtDiferenciaStock.Focus()
            Exit Function
        End If
        If optHoras.Checked = True Then
            If CDbl(Numerico(txtDiferenciaStock.Text)) > 24 Then
                MsgBox("El número de horas no debe exceder 24." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                txtDiferenciaStock.Focus()
                Exit Function
            End If
        Else
            If CDbl(Numerico(txtDiferenciaStock.Text)) > 60 Then
                MsgBox("El número de minutos no debe exceder 60." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                txtDiferenciaStock.Focus()
                Exit Function
            End If
        End If

        ValidaDatos = True
        Exit Function
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Function Guardar() As Boolean
        On Error GoTo Merr
        Dim CodGrupo As String
        Dim DescGrupo As String
        Dim importe As String
        Dim IEPS As String
        Dim LapsoStock As String

        'Si no se realizaron cambios, entonces no se guardara nada
        If (Cambios() = False And (Not frmConfigCuentasNotific.mblnCambios)) Then
            Me.Close()
            Exit Function
        End If

        'Validar si todos los datos fueron proporcionados para ser guardados
        If ValidaDatos() = False Then
            Exit Function
        End If

        If optHoras.Checked = True Then
            LapsoStock = "H" & Trim(txtDiferenciaStock.Text)
        Else
            LapsoStock = "M" & Trim(txtDiferenciaStock.Text)
        End If

        'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Cnn.BeginTrans()
        'Aqui se da de alta lo referente a la Tabla de Configuracion General
        ModStoredProcedures.PR_IMConfiguracionGeneral(Trim(txtNombreEmpresa.Text), Trim(txtRFCEmpresa.Text), Trim(txtDomicilioEmpresa.Text), CStr(txtTipoCambioDolar.Text), Trim(txtUtilidadMinima.Text), Trim(txtDirectorioImagenes.Text), Trim(txtTipoCambioEuro.Text), Trim(txtVigenciaApartado.Text), Trim(cboDriveLocal.Text), CStr(chkTransferenciasEntreSucursales.CheckState), Trim(txtCodificacion.Text), LapsoStock, C_INSERCION, CStr(0))
        Cmd.Execute()

        If frmConfigCuentasNotific.mblnCambios Then
            If Not frmConfigCuentasNotific.Guardar Then
                Cnn.RollbackTrans()
                MsgBox("")
                ModEstandar.MostrarError()
            End If
        End If
        Cnn.CommitTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        'Por cuestiones de estética el cambio al puntero del mouse se hace antes de iniciar la transacción y al finalizar la misma.

        MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
        Guardar = True
        ModCorporativo.CargarDatosConfiguracionCorpo()

        LlenaDatos()
        Me.Close()
        Exit Function
Merr:
        Cnn.RollbackTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function


    Private Sub btnDirectorioImagenes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDirectorioImagenes.Click
        Dim Titulo As String
        Dim frmConfigGralCorporativo As New frmConfigGralCorporativo()
        Titulo = ModEstandar.BuscarRutaCarpeta(frmConfigGralCorporativo, C_TITLEIMAGES)

        If Titulo <> "" Then
            Me.txtDirectorioImagenes.Text = Trim(Titulo)
        End If
    End Sub

    Function BuscarRuta() As Object
        On Error GoTo Merr
        Dim shlBusca As New Shell32.Shell
        Dim fldRecurso As Shell32.Folder
        Dim lngOpciones As Integer
        Dim Titulo As String
        Titulo = "seleccione la imagen"
        lngOpciones = BIF_RETURNONLYFSDIRS
        fldRecurso = shlBusca.BrowseForFolder(Me.Handle.ToInt32, Trim(Titulo), lngOpciones, False)
        If Not fldRecurso Is Nothing Then
            If Trim(fldRecurso.Items.Item.Path) <> "" Then
                BuscarRuta = fldRecurso.Items.Item.Path
            End If
        End If
        Exit Function
Merr:
        'MostrarError "Ocurrió un error al intentar abrir la busqueda de las carpetas"
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function




    Private Sub cmdCuentasNotificaciones_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCuentasNotificaciones.Click
        mblnMuestraFormaCuentasN = True
        frmConfigCuentasNotific.Show()
    End Sub

    Private Sub frmConfigGralCorporativo_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmConfigGralCorporativo_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmConfigGralCorporativo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
        LlenaCombo(cboDriveLocal)
        LlenaDatos()
    End Sub

    Private Sub frmConfigGralCorporativo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmConfigGralCorporativo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        '    KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmConfigGralCorporativo_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSalir Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        '    'Si se cierra el formulario y existio algun cambio en el registro se
        '    'informa al usuario del cabio y si desea guardar el registro, ya sea
        '    'que sea nuevo o un registro modificado
        '    If Cambios() = True And mblnNuevo = False Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes 'Guardar el registro
        '                If Guardar() = False Then
        '                    Cancel = 1
        '                Else '''Una vez guardada la info se procede a descargar el formulario de cuentas de correo
        '                    If mblnMuestraFormaCuentasN Then frmConfigCuentasNotific.Close()
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

    Private Sub frmConfigGralCorporativo_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub


    Private Sub Label2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Label2.Click
        frmEliminarDatos.Show()
    End Sub

    Private Sub txtCodificacion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodificacion.Enter
        SelTextoTxt(txtCodificacion)
        Pon_Tool()
    End Sub

    Private Sub txtCodificacion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodificacion.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoLetras(KeyAscii)
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Private Sub txtDiferenciaStock_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDiferenciaStock.Enter
        SelTextoTxt(txtDiferenciaStock)
        Pon_Tool()
    End Sub

    Private Sub txtDiferenciaStock_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDiferenciaStock.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If optHoras.Checked = True Then
                If CDbl(Numerico(txtDiferenciaStock.Text)) > 24 Then
                    MsgBox("El número de horas no debe exceder 24." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                    txtDiferenciaStock.Focus()
                End If
            Else
                If CDbl(Numerico(txtDiferenciaStock.Text)) > 60 Then
                    MsgBox("El número de minutos no debe exceder 60." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                    txtDiferenciaStock.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub txtDiferenciaStock_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDiferenciaStock.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Private Sub txtDiferenciaStock_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDiferenciaStock.Leave
        If optHoras.Checked = True Then
            If CDbl(Numerico(txtDiferenciaStock.Text)) > 24 Then
                MsgBox("El número de horas no debe exceder 24." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                txtDiferenciaStock.Focus()
            End If
        Else
            If CDbl(Numerico(txtDiferenciaStock.Text)) > 60 Then
                MsgBox("El número de minutos no debe exceder 60." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                txtDiferenciaStock.Focus()
            End If
        End If
    End Sub


    Private Sub txtDirectorioImagenes_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDirectorioImagenes.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDirectorioImagenes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDirectorioImagenes.Enter
        SelTextoTxt(txtDirectorioImagenes)
        Pon_Tool()
    End Sub

    Private Sub txtDomicilioEmpresa_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilioEmpresa.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDomicilioEmpresa_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilioEmpresa.Enter
        SelTextoTxt(txtDomicilioEmpresa)
        Pon_Tool()
    End Sub


    Private Sub txtNombreEmpresa_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombreEmpresa.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtNombreEmpresa_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombreEmpresa.Enter
        SelTextoTxt(txtNombreEmpresa)
        Pon_Tool()
    End Sub


    Private Sub txtRFCEmpresa_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRFCEmpresa.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtRFCEmpresa_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRFCEmpresa.Enter
        SelTextoTxt(txtRFCEmpresa)
        Pon_Tool()
    End Sub

    Private Sub txtRFCEmpresa_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRFCEmpresa.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Back Then GoTo EventExitSub
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        KeyAscii = ModEstandar.Valida_RFC(txtRFCEmpresa.Text, KeyAscii, Len(txtRFCEmpresa.Text) + 1)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambioDolar_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioDolar.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtTipoCambioDolar_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioDolar.Enter
        SelTextoTxt(txtTipoCambioDolar)
        Pon_Tool()
    End Sub

    Private Sub txtTipoCambioDolar_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambioDolar.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtTipoCambioDolar.Text, KeyAscii, 8, 2, (txtTipoCambioDolar.SelectionStart))
        '    KeyAscii = ModEstandar.gp_CampoNumerico(KeyAscii, ".")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambioDolar_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioDolar.Leave
        txtTipoCambioDolar.Text = VB6.Format(txtTipoCambioDolar.Text, "0.00")
    End Sub

    Private Sub txtTipoCambioEuro_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtTipoCambioEuro_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.Enter
        SelTextoTxt(txtTipoCambioEuro)
        Pon_Tool()
    End Sub

    Private Sub txtTipoCambioEuro_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTipoCambioEuro.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtTipoCambioEuro.Text, KeyAscii, 8, 2, (txtTipoCambioEuro.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTipoCambioEuro_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipoCambioEuro.Leave
        txtTipoCambioEuro.Text = VB6.Format(txtTipoCambioEuro.Text, "0.00")
    End Sub

    Private Sub txtUtilidadMinima_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUtilidadMinima.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtUtilidadMinima_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUtilidadMinima.Enter
        SelTextoTxt(txtUtilidadMinima)
        Pon_Tool()
    End Sub

    Private Sub txtUtilidadMinima_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUtilidadMinima.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtUtilidadMinima.Text, KeyAscii, 3, 2, (txtUtilidadMinima.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtUtilidadMinima_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUtilidadMinima.Leave
        txtUtilidadMinima.Text = VB6.Format(Numerico(txtUtilidadMinima.Text), "0.00")
        If CDbl(Numerico(Trim(txtUtilidadMinima.Text))) > 100 Then
            MsgBox("EL Porcentaje de Utilidad Mínima por Operación debe ser menor o igual a 100." & vbNewLine & "Verifique Porfavor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Me.txtUtilidadMinima.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub txtVigenciaApartado_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVigenciaApartado.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtVigenciaApartado_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVigenciaApartado.Enter
        SelTextoTxt(txtVigenciaApartado)
        Pon_Tool()
    End Sub

    Private Sub txtVigenciaApartado_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtVigenciaApartado.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtVigenciaApartado.Text, KeyAscii, 5, 0, (txtVigenciaApartado.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtVigenciaApartado_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVigenciaApartado.Leave
        txtVigenciaApartado.Text = Numerico(txtVigenciaApartado.Text)
    End Sub

    Sub LlenaCombo(ByRef cboParam As System.Windows.Forms.ComboBox)
        On Error GoTo Merr
        Dim lStrSql As String
        Dim I As Object
        Dim J As Integer
        Dim maLetra(25) As Object
        For I = 0 To 25
            maLetra(I) = UCase(Chr(System.Windows.Forms.Keys.A + I))
        Next
        cboParam.Items.Clear()
        For I = 0 To 25
            cboParam.Items.Add(maLetra(I) + ":")
        Next I
        Exit Sub
Merr:
        If Err.Number <> 0 Then ModErrores.Errores()
    End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCodificacion = New System.Windows.Forms.TextBox()
        Me.txtDiferenciaStock = New System.Windows.Forms.TextBox()
        Me.txtTipoCambioDolar = New System.Windows.Forms.TextBox()
        Me.txtTipoCambioEuro = New System.Windows.Forms.TextBox()
        Me.txtUtilidadMinima = New System.Windows.Forms.TextBox()
        Me.txtVigenciaApartado = New System.Windows.Forms.TextBox()
        Me.txtNombreEmpresa = New System.Windows.Forms.TextBox()
        Me.txtDomicilioEmpresa = New System.Windows.Forms.TextBox()
        Me.txtRFCEmpresa = New System.Windows.Forms.TextBox()
        Me.cboDriveLocal = New System.Windows.Forms.ComboBox()
        Me.btnDirectorioImagenes = New System.Windows.Forms.Button()
        Me.txtDirectorioImagenes = New System.Windows.Forms.TextBox()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.cmdCuentasNotificaciones = New System.Windows.Forms.Button()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.optHoras = New System.Windows.Forms.RadioButton()
        Me.optMinutos = New System.Windows.Forms.RadioButton()
        Me._Label_8 = New System.Windows.Forms.Label()
        Me._Label_7 = New System.Windows.Forms.Label()
        Me.fraTiposCambio = New System.Windows.Forms.GroupBox()
        Me._Label_5 = New System.Windows.Forms.Label()
        Me._Label_6 = New System.Windows.Forms.Label()
        Me.fraVentas = New System.Windows.Forms.GroupBox()
        Me._Label_0 = New System.Windows.Forms.Label()
        Me._Label_3 = New System.Windows.Forms.Label()
        Me.fraInformacionEmpresa = New System.Windows.Forms.GroupBox()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me.fraImagenes = New System.Windows.Forms.GroupBox()
        Me.chkTransferenciasEntreSucursales = New System.Windows.Forms.CheckBox()
        Me._Label_1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.fraTiposCambio.SuspendLayout()
        Me.fraVentas.SuspendLayout()
        Me.fraInformacionEmpresa.SuspendLayout()
        Me.fraImagenes.SuspendLayout()
        CType(Me.Label, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCodificacion
        '
        Me.txtCodificacion.AcceptsReturn = True
        Me.txtCodificacion.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodificacion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodificacion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodificacion.Location = New System.Drawing.Point(112, 24)
        Me.txtCodificacion.MaxLength = 10
        Me.txtCodificacion.Name = "txtCodificacion"
        Me.txtCodificacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodificacion.Size = New System.Drawing.Size(97, 20)
        Me.txtCodificacion.TabIndex = 23
        Me.txtCodificacion.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCodificacion, "Clave de codificación")
        '
        'txtDiferenciaStock
        '
        Me.txtDiferenciaStock.AcceptsReturn = True
        Me.txtDiferenciaStock.BackColor = System.Drawing.SystemColors.Window
        Me.txtDiferenciaStock.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDiferenciaStock.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDiferenciaStock.Location = New System.Drawing.Point(440, 24)
        Me.txtDiferenciaStock.MaxLength = 2
        Me.txtDiferenciaStock.Name = "txtDiferenciaStock"
        Me.txtDiferenciaStock.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDiferenciaStock.Size = New System.Drawing.Size(57, 20)
        Me.txtDiferenciaStock.TabIndex = 27
        Me.txtDiferenciaStock.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDiferenciaStock, "Lapso de tiempo para comprobar la diferencia entre la existencia y el stock")
        '
        'txtTipoCambioDolar
        '
        Me.txtTipoCambioDolar.AcceptsReturn = True
        Me.txtTipoCambioDolar.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipoCambioDolar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambioDolar.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambioDolar.Location = New System.Drawing.Point(416, 24)
        Me.txtTipoCambioDolar.MaxLength = 0
        Me.txtTipoCambioDolar.Name = "txtTipoCambioDolar"
        Me.txtTipoCambioDolar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambioDolar.Size = New System.Drawing.Size(89, 20)
        Me.txtTipoCambioDolar.TabIndex = 16
        Me.txtTipoCambioDolar.Text = "0.00"
        Me.txtTipoCambioDolar.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambioDolar, "Tipo de Cambio del Dólar")
        '
        'txtTipoCambioEuro
        '
        Me.txtTipoCambioEuro.AcceptsReturn = True
        Me.txtTipoCambioEuro.BackColor = System.Drawing.SystemColors.Window
        Me.txtTipoCambioEuro.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTipoCambioEuro.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTipoCambioEuro.Location = New System.Drawing.Point(192, 24)
        Me.txtTipoCambioEuro.MaxLength = 0
        Me.txtTipoCambioEuro.Name = "txtTipoCambioEuro"
        Me.txtTipoCambioEuro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTipoCambioEuro.Size = New System.Drawing.Size(89, 20)
        Me.txtTipoCambioEuro.TabIndex = 14
        Me.txtTipoCambioEuro.Text = "0.00"
        Me.txtTipoCambioEuro.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtTipoCambioEuro, "Tipo de Cambio del Dólar")
        '
        'txtUtilidadMinima
        '
        Me.txtUtilidadMinima.AcceptsReturn = True
        Me.txtUtilidadMinima.BackColor = System.Drawing.SystemColors.Window
        Me.txtUtilidadMinima.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUtilidadMinima.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUtilidadMinima.Location = New System.Drawing.Point(192, 24)
        Me.txtUtilidadMinima.MaxLength = 0
        Me.txtUtilidadMinima.Name = "txtUtilidadMinima"
        Me.txtUtilidadMinima.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUtilidadMinima.Size = New System.Drawing.Size(89, 20)
        Me.txtUtilidadMinima.TabIndex = 10
        Me.txtUtilidadMinima.Tag = "0.00"
        Me.txtUtilidadMinima.Text = "0.00"
        Me.txtUtilidadMinima.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtUtilidadMinima, "% de Utilidad Mínima por Operación")
        '
        'txtVigenciaApartado
        '
        Me.txtVigenciaApartado.AcceptsReturn = True
        Me.txtVigenciaApartado.BackColor = System.Drawing.SystemColors.Window
        Me.txtVigenciaApartado.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVigenciaApartado.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVigenciaApartado.Location = New System.Drawing.Point(416, 24)
        Me.txtVigenciaApartado.MaxLength = 0
        Me.txtVigenciaApartado.Name = "txtVigenciaApartado"
        Me.txtVigenciaApartado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVigenciaApartado.Size = New System.Drawing.Size(89, 20)
        Me.txtVigenciaApartado.TabIndex = 12
        Me.txtVigenciaApartado.Text = "0.00"
        Me.txtVigenciaApartado.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtVigenciaApartado, "Tipo de Cambio del Dólar")
        '
        'txtNombreEmpresa
        '
        Me.txtNombreEmpresa.AcceptsReturn = True
        Me.txtNombreEmpresa.BackColor = System.Drawing.SystemColors.Window
        Me.txtNombreEmpresa.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombreEmpresa.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombreEmpresa.Location = New System.Drawing.Point(88, 32)
        Me.txtNombreEmpresa.MaxLength = 60
        Me.txtNombreEmpresa.Name = "txtNombreEmpresa"
        Me.txtNombreEmpresa.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombreEmpresa.Size = New System.Drawing.Size(417, 20)
        Me.txtNombreEmpresa.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtNombreEmpresa, "Nombre de la Empresa")
        '
        'txtDomicilioEmpresa
        '
        Me.txtDomicilioEmpresa.AcceptsReturn = True
        Me.txtDomicilioEmpresa.BackColor = System.Drawing.SystemColors.Window
        Me.txtDomicilioEmpresa.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDomicilioEmpresa.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDomicilioEmpresa.Location = New System.Drawing.Point(88, 56)
        Me.txtDomicilioEmpresa.MaxLength = 65
        Me.txtDomicilioEmpresa.Name = "txtDomicilioEmpresa"
        Me.txtDomicilioEmpresa.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDomicilioEmpresa.Size = New System.Drawing.Size(417, 20)
        Me.txtDomicilioEmpresa.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtDomicilioEmpresa, "Domicilio de la Empresa")
        '
        'txtRFCEmpresa
        '
        Me.txtRFCEmpresa.AcceptsReturn = True
        Me.txtRFCEmpresa.BackColor = System.Drawing.SystemColors.Window
        Me.txtRFCEmpresa.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRFCEmpresa.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRFCEmpresa.Location = New System.Drawing.Point(89, 80)
        Me.txtRFCEmpresa.MaxLength = 0
        Me.txtRFCEmpresa.Name = "txtRFCEmpresa"
        Me.txtRFCEmpresa.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRFCEmpresa.Size = New System.Drawing.Size(240, 20)
        Me.txtRFCEmpresa.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtRFCEmpresa, "RFC de la Empresa")
        '
        'cboDriveLocal
        '
        Me.cboDriveLocal.BackColor = System.Drawing.SystemColors.Window
        Me.cboDriveLocal.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDriveLocal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDriveLocal.Location = New System.Drawing.Point(131, 56)
        Me.cboDriveLocal.Name = "cboDriveLocal"
        Me.cboDriveLocal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDriveLocal.Size = New System.Drawing.Size(65, 21)
        Me.cboDriveLocal.Sorted = True
        Me.cboDriveLocal.TabIndex = 20
        Me.ToolTip1.SetToolTip(Me.cboDriveLocal, "Tecla Rápida")
        '
        'btnDirectorioImagenes
        '
        Me.btnDirectorioImagenes.BackColor = System.Drawing.SystemColors.Control
        Me.btnDirectorioImagenes.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnDirectorioImagenes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDirectorioImagenes.Location = New System.Drawing.Point(488, 24)
        Me.btnDirectorioImagenes.Name = "btnDirectorioImagenes"
        Me.btnDirectorioImagenes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnDirectorioImagenes.Size = New System.Drawing.Size(25, 21)
        Me.btnDirectorioImagenes.TabIndex = 31
        Me.btnDirectorioImagenes.Text = "..."
        Me.ToolTip1.SetToolTip(Me.btnDirectorioImagenes, "Seleccionar Directorio")
        Me.btnDirectorioImagenes.UseVisualStyleBackColor = False
        '
        'txtDirectorioImagenes
        '
        Me.txtDirectorioImagenes.AcceptsReturn = True
        Me.txtDirectorioImagenes.BackColor = System.Drawing.SystemColors.Window
        Me.txtDirectorioImagenes.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDirectorioImagenes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDirectorioImagenes.Location = New System.Drawing.Point(131, 24)
        Me.txtDirectorioImagenes.MaxLength = 255
        Me.txtDirectorioImagenes.Name = "txtDirectorioImagenes"
        Me.txtDirectorioImagenes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDirectorioImagenes.Size = New System.Drawing.Size(361, 20)
        Me.txtDirectorioImagenes.TabIndex = 18
        Me.ToolTip1.SetToolTip(Me.txtDirectorioImagenes, "Directorio de Imágenes")
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.ForeColor = System.Drawing.Color.Black
        Me._Label1_3.Location = New System.Drawing.Point(8, 56)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(121, 21)
        Me._Label1_3.TabIndex = 19
        Me._Label1_3.Text = "Drive Local :"
        Me.ToolTip1.SetToolTip(Me._Label1_3, "Nombre de la Farmacia Actual")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cmdCuentasNotificaciones)
        Me.Frame1.Controls.Add(Me.Frame2)
        Me.Frame1.Controls.Add(Me.fraTiposCambio)
        Me.Frame1.Controls.Add(Me.fraVentas)
        Me.Frame1.Controls.Add(Me.fraInformacionEmpresa)
        Me.Frame1.Controls.Add(Me.fraImagenes)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 6)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(543, 471)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'cmdCuentasNotificaciones
        '
        Me.cmdCuentasNotificaciones.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCuentasNotificaciones.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCuentasNotificaciones.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCuentasNotificaciones.Location = New System.Drawing.Point(358, 263)
        Me.cmdCuentasNotificaciones.Name = "cmdCuentasNotificaciones"
        Me.cmdCuentasNotificaciones.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCuentasNotificaciones.Size = New System.Drawing.Size(170, 30)
        Me.cmdCuentasNotificaciones.TabIndex = 33
        Me.cmdCuentasNotificaciones.Text = "Cuentas de N&otificaciones"
        Me.cmdCuentasNotificaciones.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optHoras)
        Me.Frame2.Controls.Add(Me.optMinutos)
        Me.Frame2.Controls.Add(Me.txtCodificacion)
        Me.Frame2.Controls.Add(Me.txtDiferenciaStock)
        Me.Frame2.Controls.Add(Me._Label_8)
        Me.Frame2.Controls.Add(Me._Label_7)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(11, 392)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(521, 64)
        Me.Frame2.TabIndex = 30
        Me.Frame2.TabStop = False
        '
        'optHoras
        '
        Me.optHoras.BackColor = System.Drawing.SystemColors.Control
        Me.optHoras.Cursor = System.Windows.Forms.Cursors.Default
        Me.optHoras.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optHoras.Location = New System.Drawing.Point(376, 40)
        Me.optHoras.Name = "optHoras"
        Me.optHoras.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optHoras.Size = New System.Drawing.Size(57, 17)
        Me.optHoras.TabIndex = 26
        Me.optHoras.TabStop = True
        Me.optHoras.Text = "Hrs."
        Me.optHoras.UseVisualStyleBackColor = False
        '
        'optMinutos
        '
        Me.optMinutos.BackColor = System.Drawing.SystemColors.Control
        Me.optMinutos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMinutos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMinutos.Location = New System.Drawing.Point(376, 16)
        Me.optMinutos.Name = "optMinutos"
        Me.optMinutos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMinutos.Size = New System.Drawing.Size(57, 17)
        Me.optMinutos.TabIndex = 25
        Me.optMinutos.TabStop = True
        Me.optMinutos.Text = "Min."
        Me.optMinutos.UseVisualStyleBackColor = False
        '
        '_Label_8
        '
        Me._Label_8.BackColor = System.Drawing.SystemColors.Control
        Me._Label_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_8.Location = New System.Drawing.Point(32, 24)
        Me._Label_8.Name = "_Label_8"
        Me._Label_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_8.Size = New System.Drawing.Size(74, 21)
        Me._Label_8.TabIndex = 22
        Me._Label_8.Text = "Codificación:"
        '
        '_Label_7
        '
        Me._Label_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_7.Location = New System.Drawing.Point(266, 24)
        Me._Label_7.Name = "_Label_7"
        Me._Label_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_7.Size = New System.Drawing.Size(104, 13)
        Me._Label_7.TabIndex = 24
        Me._Label_7.Text = "Lapso Dif. Stock"
        '
        'fraTiposCambio
        '
        Me.fraTiposCambio.BackColor = System.Drawing.SystemColors.Control
        Me.fraTiposCambio.Controls.Add(Me.txtTipoCambioDolar)
        Me.fraTiposCambio.Controls.Add(Me.txtTipoCambioEuro)
        Me.fraTiposCambio.Controls.Add(Me._Label_5)
        Me.fraTiposCambio.Controls.Add(Me._Label_6)
        Me.fraTiposCambio.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraTiposCambio.Location = New System.Drawing.Point(11, 200)
        Me.fraTiposCambio.Name = "fraTiposCambio"
        Me.fraTiposCambio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTiposCambio.Size = New System.Drawing.Size(521, 57)
        Me.fraTiposCambio.TabIndex = 29
        Me.fraTiposCambio.TabStop = False
        Me.fraTiposCambio.Text = "Tipos de  Cambio "
        '
        '_Label_5
        '
        Me._Label_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_5.Location = New System.Drawing.Point(296, 24)
        Me._Label_5.Name = "_Label_5"
        Me._Label_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_5.Size = New System.Drawing.Size(113, 21)
        Me._Label_5.TabIndex = 15
        Me._Label_5.Text = "Tipo de Cambio Dólar : "
        '
        '_Label_6
        '
        Me._Label_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_6.Location = New System.Drawing.Point(48, 24)
        Me._Label_6.Name = "_Label_6"
        Me._Label_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_6.Size = New System.Drawing.Size(129, 21)
        Me._Label_6.TabIndex = 13
        Me._Label_6.Text = "Tipo de Cambio del Euro : "
        '
        'fraVentas
        '
        Me.fraVentas.BackColor = System.Drawing.SystemColors.Control
        Me.fraVentas.Controls.Add(Me.txtUtilidadMinima)
        Me.fraVentas.Controls.Add(Me.txtVigenciaApartado)
        Me.fraVentas.Controls.Add(Me._Label_0)
        Me.fraVentas.Controls.Add(Me._Label_3)
        Me.fraVentas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraVentas.Location = New System.Drawing.Point(11, 136)
        Me.fraVentas.Name = "fraVentas"
        Me.fraVentas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraVentas.Size = New System.Drawing.Size(521, 57)
        Me.fraVentas.TabIndex = 8
        Me.fraVentas.TabStop = False
        Me.fraVentas.Text = " Ventas "
        '
        '_Label_0
        '
        Me._Label_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_0.Location = New System.Drawing.Point(8, 24)
        Me._Label_0.Name = "_Label_0"
        Me._Label_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_0.Size = New System.Drawing.Size(177, 21)
        Me._Label_0.TabIndex = 9
        Me._Label_0.Text = "% de Utilidad Mínima por Operación :"
        '
        '_Label_3
        '
        Me._Label_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_3.Location = New System.Drawing.Point(296, 24)
        Me._Label_3.Name = "_Label_3"
        Me._Label_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_3.Size = New System.Drawing.Size(113, 21)
        Me._Label_3.TabIndex = 11
        Me._Label_3.Text = "Vigencia de Apartado :"
        '
        'fraInformacionEmpresa
        '
        Me.fraInformacionEmpresa.BackColor = System.Drawing.SystemColors.Control
        Me.fraInformacionEmpresa.Controls.Add(Me.txtNombreEmpresa)
        Me.fraInformacionEmpresa.Controls.Add(Me.txtDomicilioEmpresa)
        Me.fraInformacionEmpresa.Controls.Add(Me.txtRFCEmpresa)
        Me.fraInformacionEmpresa.Controls.Add(Me._Label1_0)
        Me.fraInformacionEmpresa.Controls.Add(Me._Label1_1)
        Me.fraInformacionEmpresa.Controls.Add(Me._Label1_2)
        Me.fraInformacionEmpresa.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraInformacionEmpresa.Location = New System.Drawing.Point(11, 16)
        Me.fraInformacionEmpresa.Name = "fraInformacionEmpresa"
        Me.fraInformacionEmpresa.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraInformacionEmpresa.Size = New System.Drawing.Size(521, 113)
        Me.fraInformacionEmpresa.TabIndex = 1
        Me.fraInformacionEmpresa.TabStop = False
        Me.fraInformacionEmpresa.Text = " Información General de la Empresa "
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(23, 56)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(65, 13)
        Me._Label1_0.TabIndex = 4
        Me._Label1_0.Text = "Domicilio : "
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(24, 32)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(65, 21)
        Me._Label1_1.TabIndex = 2
        Me._Label1_1.Text = "Nombre :"
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_2.Location = New System.Drawing.Point(24, 80)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(65, 21)
        Me._Label1_2.TabIndex = 6
        Me._Label1_2.Text = "RFC :"
        '
        'fraImagenes
        '
        Me.fraImagenes.BackColor = System.Drawing.SystemColors.Control
        Me.fraImagenes.Controls.Add(Me.chkTransferenciasEntreSucursales)
        Me.fraImagenes.Controls.Add(Me.cboDriveLocal)
        Me.fraImagenes.Controls.Add(Me.btnDirectorioImagenes)
        Me.fraImagenes.Controls.Add(Me.txtDirectorioImagenes)
        Me.fraImagenes.Controls.Add(Me._Label1_3)
        Me.fraImagenes.Controls.Add(Me._Label_1)
        Me.fraImagenes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraImagenes.Location = New System.Drawing.Point(11, 296)
        Me.fraImagenes.Name = "fraImagenes"
        Me.fraImagenes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraImagenes.Size = New System.Drawing.Size(521, 89)
        Me.fraImagenes.TabIndex = 28
        Me.fraImagenes.TabStop = False
        '
        'chkTransferenciasEntreSucursales
        '
        Me.chkTransferenciasEntreSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTransferenciasEntreSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTransferenciasEntreSucursales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTransferenciasEntreSucursales.Location = New System.Drawing.Point(280, 62)
        Me.chkTransferenciasEntreSucursales.Name = "chkTransferenciasEntreSucursales"
        Me.chkTransferenciasEntreSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTransferenciasEntreSucursales.Size = New System.Drawing.Size(225, 21)
        Me.chkTransferenciasEntreSucursales.TabIndex = 21
        Me.chkTransferenciasEntreSucursales.Text = "&Permitir trransferencias entre sucursales"
        Me.chkTransferenciasEntreSucursales.UseVisualStyleBackColor = False
        '
        '_Label_1
        '
        Me._Label_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label_1.Location = New System.Drawing.Point(8, 24)
        Me._Label_1.Name = "_Label_1"
        Me._Label_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_1.Size = New System.Drawing.Size(121, 21)
        Me._Label_1.TabIndex = 17
        Me._Label_1.Text = "Ruta de Archivos :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(9, 281)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(161, 17)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "Especificaciones generales"
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(26, 497)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(110, 42)
        Me.btnGuardar.TabIndex = 34
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmConfigGralCorporativo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(558, 551)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(237, 150)
        Me.MaximizeBox = False
        Me.Name = "frmConfigGralCorporativo"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Configuración General del Corporativo"
        Me.Frame1.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.fraTiposCambio.ResumeLayout(False)
        Me.fraTiposCambio.PerformLayout()
        Me.fraVentas.ResumeLayout(False)
        Me.fraVentas.PerformLayout()
        Me.fraInformacionEmpresa.ResumeLayout(False)
        Me.fraInformacionEmpresa.PerformLayout()
        Me.fraImagenes.ResumeLayout(False)
        Me.fraImagenes.PerformLayout()
        CType(Me.Label, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub
End Class