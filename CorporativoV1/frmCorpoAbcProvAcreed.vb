'**********************************************************************************************************************'
'*PROGRAMA: ABC PROVEEDORES Y ACREEDORES JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018      
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB

Public Class frmCorpoAbcProvAcreed

    Inherits System.Windows.Forms.Form
    'Programa: ABC de Proveedores y Acreedores
    'Autor: Rosaura Torres López
    'Fecha de Creación: 13/Mayo/2003

    Public components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkMostrarTodos As System.Windows.Forms.CheckBox
    Public WithEvents txtCodProvAcreed As System.Windows.Forms.TextBox
    Public WithEvents txtNombre As System.Windows.Forms.TextBox
    Public WithEvents txtEmail As System.Windows.Forms.TextBox
    Public WithEvents txtRFC As System.Windows.Forms.TextBox
    Public WithEvents Frame9 As System.Windows.Forms.GroupBox
    Public WithEvents txtPais As System.Windows.Forms.TextBox
    Public WithEvents txtCodigoPostal As System.Windows.Forms.TextBox
    Public WithEvents txtTelefonos As System.Windows.Forms.TextBox
    Public WithEvents txtDomicilio As System.Windows.Forms.TextBox
    Public WithEvents txtLocalidad As System.Windows.Forms.TextBox
    Public WithEvents _lblProvAcreed_18 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_9 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_7 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_5 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_4 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_1 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_3 As System.Windows.Forms.Label
    Public WithEvents Frame8 As System.Windows.Forms.GroupBox
    Public WithEvents rtbObservaciones As System.Windows.Forms.RichTextBox
    Public WithEvents txtContactoPagos As System.Windows.Forms.TextBox
    Public WithEvents txtCtasBancarias As System.Windows.Forms.TextBox
    Public WithEvents txtTelefonosPagos As System.Windows.Forms.TextBox
    Public WithEvents txtTelefonosVentas As System.Windows.Forms.TextBox
    Public WithEvents txtDescVolumen As System.Windows.Forms.TextBox
    Public WithEvents txtDiasCredito As System.Windows.Forms.TextBox
    Public WithEvents txtContactoVentas As System.Windows.Forms.TextBox
    Public WithEvents txtDescFinanciero As System.Windows.Forms.TextBox
    Public WithEvents txtTAXID As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_17 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_16 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_10 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_6 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_15 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_14 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_13 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_12 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_11 As System.Windows.Forms.Label
    Public WithEvents Frame10 As System.Windows.Forms.GroupBox
    Public WithEvents optEmpresa As System.Windows.Forms.RadioButton
    Public WithEvents optPersonal As System.Windows.Forms.RadioButton
    Public WithEvents fraServicio As System.Windows.Forms.Panel
    Public WithEvents chkAgenciaAduanal As System.Windows.Forms.CheckBox
    Public WithEvents optAcreedor As System.Windows.Forms.RadioButton
    Public WithEvents optProveedor As System.Windows.Forms.RadioButton
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents fraTipo As System.Windows.Forms.Panel
    Public WithEvents Frame7 As System.Windows.Forms.GroupBox
    Public WithEvents optExtranjero As System.Windows.Forms.RadioButton
    Public WithEvents optNacional As System.Windows.Forms.RadioButton
    Public WithEvents Frame6 As System.Windows.Forms.GroupBox
    Public WithEvents fraNacional As System.Windows.Forms.Panel
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents _lblProvAcreed_2 As System.Windows.Forms.Label
    Public WithEvents _lblProvAcreed_0 As System.Windows.Forms.Label
    Public WithEvents lblProvAcreed As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    'Estas Variables se declaran de manera local, para evitar conflictos al estar usando
    'la misma variable en distintos modulos, que pueden afectar el valor que hayan tomado en un form. distinto al actual
    Dim mblnNuevo As Boolean 'Para Controlar si un registro es Nuevo o se trata de una consulta
    Dim mblnCambiosEnCodigo As Boolean 'Para Controlar si se han efectuado cambios en el código
    Dim I As Integer
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
        'mblnNuevo = True
        'mblnCambiosEnCodigo = False

        mblnNuevo = True
        mblnCambiosEnCodigo = False
        'intAlmGeneral = 0
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
                Case "TXTCODPROVACREED"
                    strCaptionForm = "Consulta de Proveedores/Acreedores"
                    'Si se kiere que se muestren todos los Prov/Acreed se hace lo siguiente
                    If chkMostrarTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
                        gStrSql = "SELECT RIGHT('000'+LTRIM(CodProvAcreed),3) AS CODIGO,DescProvAcreed AS NOMBRE, " & "CASE Tipo WHEN 'P' THEN 'Proveedor' WHEN 'A' THEN 'Acreedor  ' END + '    ' + " & "CASE Nacional WHEN 0 THEN 'Extranjero' WHEN 1 THEN 'Nacional  ' END + '    ' + " & "CASE Servicio WHEN 'P' THEN 'Personal' WHEN 'E' THEN 'Empresa '  END as CLASIFICACION, " & "CASE AgenciaAduanal WHEN 0 THEN 'No' WHEN 1 THEN 'Si'  END as AGENCIAADUANAL " & "FROM CatProvAcreed "

                    Else ' De lo Contrario se hara la consulta con los datos seleccionados

                        gStrSql = "SELECT RIGHT('000'+LTRIM(CodProvAcreed),3) AS CODIGO,DescProvAcreed AS NOMBRE, " & "CASE Tipo WHEN 'P' THEN 'Proveedor' WHEN 'A' THEN 'Acreedor  ' END + '    ' + " & "CASE Nacional WHEN 0 THEN 'Extranjero' WHEN 1 THEN 'Nacional  ' END + '    ' + " & "CASE Servicio WHEN 'P' THEN 'Personal' WHEN 'E' THEN 'Empresa '  END as CLASIFICACION, " & "CASE AgenciaAduanal WHEN 0 THEN 'No' WHEN 1 THEN 'Si'  END as AGENCIAADUANAL " & "FROM CatProvAcreed " & "WHERE Tipo= " & IIf(optProveedor.Checked = True, "'P'", "'A'") & " AND Nacional= " & IIf(optNacional.Checked = True, "'1'", "'0'") & " " & "AND Servicio= " & IIf(optEmpresa.Checked = True, "'E'", "'P'") & " ORDER BY CodProvAcreed"
                    End If

                Case "TXTNOMBRE"
                    strCaptionForm = "Consulta de Proveedores/Acreedores"
                    If chkMostrarTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
                        gStrSql = "SELECT  DescProvAcreed AS NOMBRE, RIGHT('000'+LTRIM(CodProvAcreed),3) AS CODIGO, " & "CASE Tipo WHEN 'P' THEN 'Proveedor' WHEN 'A' THEN 'Acreedor  ' END + '    ' + " & "CASE Nacional WHEN 0 THEN 'Extranjero' WHEN 1 THEN 'Nacional  ' END + '    ' + " & "CASE Servicio WHEN 'P' THEN 'Personal' WHEN 'E' THEN 'Empresa '  END as CLASIFICACION, " & "CASE AgenciaAduanal WHEN 0 THEN 'No' WHEN 1 THEN 'Si'  END as AGENCIAADUANAL " & "FROM CatProvAcreed " & "WHERE DescProvAcreed LIKE '" & Trim(txtNombre.Text) & "%' ORDER BY DescProvAcreed"

                    Else
                        gStrSql = "SELECT DescProvAcreed AS NOMBRE, RIGHT('000'+LTRIM(CodProvAcreed),3) AS CODIGO, " & "CASE Tipo WHEN 'P' THEN 'Proveedor' WHEN 'A' THEN 'Acreedor  ' END + '    ' + " & "CASE Nacional WHEN 0 THEN 'Extranjero' WHEN 1 THEN 'Nacional  ' END + '    ' + " & "CASE Servicio WHEN 'P' THEN 'Personal' WHEN 'E' THEN 'Empresa '  END as CLASIFICACION, " & "CASE AgenciaAduanal WHEN 0 THEN 'No' WHEN 1 THEN 'Si'  END as AGENCIAADUANAL " & "FROM CatProvAcreed " & "WHERE DescProvAcreed LIKE '" & Trim(txtNombre.Text) & "%' " & "AND Tipo= " & IIf(optProveedor.Checked = True, "'P'", "'A'") & " AND Nacional= " & IIf(optNacional.Checked = True, "'1'", "'0'") & " " & "AND Servicio= " & IIf(optEmpresa.Checked = True, "'E'", "'P'") & " ORDER BY DescProvAcreed"
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
            ConfiguraConsultas(FrmConsultas, 11100, RsGral, strTag, strCaptionForm)

            With FrmConsultas.Flexdet
                Select Case strControlActual
                    Case "TXTCODPROVACREED"
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColWidth(1, 0, 5500) 'Columna de la Descripción
                        .set_ColWidth(2, 0, 2700) 'Columna del Tipo
                        .set_ColWidth(3, 0, 2000) 'Columna de Nacional
                        .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter) 'Alinear la COlumna de AgenciaAduanal a center-center
                        'Con el Siguiente Código, se hace que las columnas con el mismo nombre se agrupen
                        '                .MergeCells = flexMergeFree  'Para que sólo las filas con igual contenido se junten
                        '                For i = 0 To RsGral.RecordCount Step 1
                        '                    .MergeRow(0) = True
                        '                Next

                    Case "TXTNOMBRE"
                        .set_ColWidth(0, 0, 5500) 'Columna de la Descripción
                        .set_ColWidth(1, 0, 900) 'Columna del Código
                        .set_ColWidth(2, 0, 2700) 'Columna del Tipo
                        .set_ColWidth(3, 0, 2000) 'Columna de Nacional
                        .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter) 'Alinear la COlumna de AgenciaAduanal a center-center
                        '                'Con el Siguiente Código, se hace que las columnas con el mismo nombre se agrupen
                        '                .MergeCells = flexMergeRestrictRows
                        '                For i = 0 To RsGral.RecordCount Step 1
                        '                .MergeRow(0) = True
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
        'On Error GoTo MErr
        Try
            'Screen.MousePointer = vbHourglass Esto se manejará hasta antes de iniciar la transacción

            gStrSql = "SELECT DescProvAcreed FROM CatProvAcreed WHERE CodProvAcreed=" & Val(txtCodProvAcreed.Text)

            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            RsGral = Cmd.Execute

            If RsGral.RecordCount = 0 Then
                MsgBox("Proporcione un Código valido para eliminar.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Mensaje")
                'Cnn.RollbackTrans
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
            Cnn.BeginTrans()

            ModStoredProcedures.PR_IMECatProvAcreed(Trim(txtCodProvAcreed.Text), Trim(txtNombre.Text), "X", CStr(2), "y", CStr(chkAgenciaAduanal.CheckState), Trim(txtDomicilio.Text), Trim(txtLocalidad.Text), Trim(txtCodigoPostal.Text), Trim(txtPais.Text), Trim(txtTelefonos.Text), Trim(txtRFC.Text), Trim(txtEmail.Text), Trim(txtTAXID.Text), Trim(txtDiasCredito.Text), CStr(txtDescVolumen.Text), CStr(txtDescFinanciero.Text), Trim(txtContactoVentas.Text), Trim(txtTelefonosVentas.Text), Trim(txtContactoPagos.Text), Trim(txtTelefonosPagos.Text), Trim(txtCtasBancarias.Text), Trim(Me.rtbObservaciones.Text), C_ELIMINACION, CStr(0))
            Cmd.Execute()
            MsgBox("El  proveedor/acreedor ha sido eliminado correctamente con el código: " & txtCodProvAcreed.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Cnn.CommitTrans()
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
    Function Guardar() As Boolean
        'On Error GoTo MErr
        Try
            Dim Tipo As String 'Contiene el tipo (Proveedor - Acreedor)
            Dim Servicio As String 'El tipo de servicio Personal (P) o de la Empresa (E)
            Dim Nacional As Boolean 'Para saber si es nacional o extranjero

            'Si no se realizaron cambios, entonces no se guardara nada
            'Si el Código  es "", entonces no se validará nada, solamente se saldrá del proc.
            If Cambios() = False And Trim(txtCodProvAcreed.Text) = "" Then
                Limpiar()
                Exit Function
            End If

            'Validar si todos los datos fueron proporcionados para ser guardados
            If ValidaDatos() = False Then
                Exit Function
            End If

            If Val(txtCodProvAcreed.Text) = 0 Then
                mblnNuevo = True
            End If

            'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Cnn.BeginTrans()

            If optProveedor.Checked = True Then
                Tipo = "P"
            Else
                Tipo = "A"
            End If
            If optNacional.Checked = True Then
                Nacional = True
            Else
                Nacional = False
            End If
            If optEmpresa.Checked = True Then
                Servicio = "E"
            Else
                Servicio = "P"
            End If
            If mblnNuevo = True Then 'Se realizará una insercion
                ModStoredProcedures.PR_IMECatProvAcreed(Trim(txtCodProvAcreed.Text), Trim(txtNombre.Text), Tipo, CStr(Nacional), Servicio, CStr(chkAgenciaAduanal.CheckState), Mid(Trim(txtDomicilio.Text), 1, 150), Trim(txtLocalidad.Text), Trim(txtCodigoPostal.Text), Trim(txtPais.Text), Trim(txtTelefonos.Text), Trim(txtRFC.Text), Trim(txtEmail.Text), Trim(txtTAXID.Text), Trim(txtDiasCredito.Text), CStr(txtDescVolumen.Text), CStr(txtDescFinanciero.Text), Trim(txtContactoVentas.Text), Trim(txtTelefonosVentas.Text), Trim(txtContactoPagos.Text), Trim(txtTelefonosPagos.Text), Trim(txtCtasBancarias.Text), Trim(rtbObservaciones.Text), C_INSERCION, CStr(0))
                Cmd.Execute()
                txtCodProvAcreed.Text = Format(Cmd.Parameters("ID").Value, "000")

            Else ' Se realizará una Modificación
                ModStoredProcedures.PR_IMECatProvAcreed(Trim(txtCodProvAcreed.Text), Trim(txtNombre.Text), Tipo, CStr(Nacional), Servicio, CStr(chkAgenciaAduanal.CheckState), Mid(Trim(txtDomicilio.Text), 1, 150), Trim(txtLocalidad.Text), Trim(txtCodigoPostal.Text), Trim(txtPais.Text), Trim(txtTelefonos.Text), Trim(txtRFC.Text), Trim(txtEmail.Text), Trim(txtTAXID.Text), Trim(txtDiasCredito.Text), CStr(txtDescVolumen.Text), CStr(txtDescFinanciero.Text), Trim(txtContactoVentas.Text), Trim(txtTelefonosVentas.Text), Trim(txtContactoPagos.Text), Trim(txtTelefonosPagos.Text), Trim(txtCtasBancarias.Text), Trim(rtbObservaciones.Text), C_MODIFICACION, CStr(0))
                Cmd.Execute()
            End If
            Cnn.CommitTrans()

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            'If Trim(Me.Tag) = "FRMCXPAGREGARPAGO" Then
            '    With frmCXPAgregarPago
            '        .mblnFueraChange = True
            '        .dbcProveedor.Text = Trim(Me.txtNombre.Text)
            '        .dbcProveedor.Tag = .dbcProveedor.Text
            '        .mintCodProveedor = CInt(Numerico((Me.txtCodProvAcreed.Text)))
            '        .mblnFueraChange = False
            '        Me.Close()
            '        Exit Function
            '    End With
            'End If
            'Por cuestiones de estética el cambio al puntero del mouse se hace antes de iniciar la transacción y al finalizar la misma.
            If mblnNuevo Then
                MsgBox("El " & IIf((Tipo = "P"), "proveedor", "acreedor") & " ha sido grabado correctamente con el código: " & txtCodProvAcreed.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
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
            InicializaVariables()
            txtCodProvAcreed.Text = ""
            txtCodProvAcreed.Enabled = True
            txtNombre.Text = ""
            txtNombre.Tag = ""
            txtDomicilio.Text = ""
            txtDomicilio.Tag = ""
            txtLocalidad.Text = ""
            txtLocalidad.Tag = ""
            txtCodigoPostal.Text = ""
            txtCodigoPostal.Tag = ""
            txtPais.Text = ""
            txtPais.Tag = ""
            txtTelefonos.Text = ""
            txtTelefonos.Tag = ""
            txtRFC.Text = ""
            txtRFC.Tag = ""
            txtEmail.Text = ""
            txtEmail.Tag = ""
            txtDiasCredito.Text = "0"
            txtDiasCredito.Tag = "0"
            txtDescVolumen.Text = "0.00"
            txtDescVolumen.Tag = "0.00"
            txtDescFinanciero.Text = "0.00"
            txtDescFinanciero.Tag = "0.00"
            txtTAXID.Text = ""
            txtTAXID.Tag = ""
            txtContactoVentas.Text = ""
            txtContactoVentas.Tag = ""
            txtTelefonosVentas.Text = ""
            txtTelefonosVentas.Tag = ""
            txtContactoPagos.Text = ""
            txtContactoPagos.Tag = ""
            txtTelefonosPagos.Text = ""
            txtTelefonosPagos.Tag = ""
            txtCtasBancarias.Text = ""
            txtCtasBancarias.Tag = ""
            optProveedor.Checked = True
            optProveedor.Tag = True
            optAcreedor.Checked = False
            optAcreedor.Tag = False
            optNacional.Checked = True
            optNacional.Tag = True
            optExtranjero.Checked = False
            optExtranjero.Tag = False
            optEmpresa.Checked = True
            optEmpresa.Tag = True
            optPersonal.Checked = False
            optPersonal.Tag = False
            chkAgenciaAduanal.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkAgenciaAduanal.Tag = System.Windows.Forms.CheckState.Unchecked
            rtbObservaciones.Text = ""
            rtbObservaciones.Tag = ""
            'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
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
            If Val(txtCodProvAcreed.Text) = 0 Then
                Nuevo()
                'ModEstandar.AvanzarTab Me
                Exit Sub
            End If

            'txtCodProvAcreed.Text = Format(txtCodProvAcreed.Text, "000")

            For I = 1 To 3 - (txtCodProvAcreed.TextLength)
                txtCodProvAcreed.Text = String.Concat("0" + txtCodProvAcreed.Text)
            Next I

            gStrSql = "SELECT CodProvAcreed,DescProvAcreed,Tipo,Nacional,Servicio,AgenciaAduanal, Domicilio, " & "Ciudad, Cp, Pais, Telefono, Rfc, Email, TaxId, DiasCredito, DesctoVolumen, DesctoFinanciero, " & "ContactoVentas,TelsVentas, ContactoPagos, TelsPagos,CuentasBancarias, Observaciones " & "From CatProvAcreed WHERE CodProvAcreed= " & Val(txtCodProvAcreed.Text)
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute

            If RsGral.RecordCount > 0 Then
                txtNombre.Text = Trim(RsGral.Fields("DescProvACreed").Value)
                txtNombre.Tag = Trim(RsGral.Fields("DescProvACreed").Value)
                txtDomicilio.Text = Trim(RsGral.Fields("Domicilio").Value)
                txtDomicilio.Tag = Trim(RsGral.Fields("Domicilio").Value)
                txtLocalidad.Text = Trim(RsGral.Fields("Ciudad").Value)
                txtLocalidad.Tag = Trim(RsGral.Fields("Ciudad").Value)
                txtCodigoPostal.Text = Trim(RsGral.Fields("CP").Value)
                txtCodigoPostal.Tag = Trim(RsGral.Fields("CP").Value)
                txtPais.Text = Trim(RsGral.Fields("Pais").Value)
                txtPais.Tag = Trim(RsGral.Fields("Pais").Value)
                txtTelefonos.Text = Trim(RsGral.Fields("Telefono").Value)
                txtTelefonos.Tag = Trim(RsGral.Fields("Telefono").Value)
                txtRFC.Text = Trim(RsGral.Fields("Rfc").Value)
                txtRFC.Tag = Trim(RsGral.Fields("Rfc").Value)
                txtEmail.Text = Trim(RsGral.Fields("Email").Value)
                txtEmail.Tag = Trim(RsGral.Fields("Email").Value)
                txtTAXID.Text = Trim(RsGral.Fields("TaxId").Value)
                txtTAXID.Tag = Trim(RsGral.Fields("TaxId").Value)
                txtDiasCredito.Text = Trim(RsGral.Fields("DiasCredito").Value)
                txtDiasCredito.Tag = Trim(RsGral.Fields("DiasCredito").Value)
                txtDescVolumen.Text = Format(CDec(RsGral.Fields("DesctoVolumen").Value), "0.00")
                txtDescVolumen.Tag = Format(CDec(RsGral.Fields("DesctoVolumen").Value), "0.00")
                txtDescFinanciero.Text = Format(CDec(RsGral.Fields("DesctoFinanciero").Value), "0.00")
                txtDescFinanciero.Tag = Format(CDec(RsGral.Fields("DesctoFinanciero").Value), "0.00")
                txtContactoVentas.Text = Trim(RsGral.Fields("ContactoVentas").Value)
                txtContactoVentas.Tag = Trim(RsGral.Fields("ContactoVentas").Value)
                txtTelefonosVentas.Text = Trim(RsGral.Fields("TelsVentas").Value)
                txtTelefonosVentas.Tag = Trim(RsGral.Fields("TelsVentas").Value)
                txtContactoPagos.Text = Trim(RsGral.Fields("ContactoPagos").Value)
                txtContactoPagos.Tag = Trim(RsGral.Fields("ContactoPagos").Value)
                txtTelefonosPagos.Text = Trim(RsGral.Fields("TelsPagos").Value)
                txtTelefonosPagos.Tag = Trim(RsGral.Fields("TelsPagos").Value)
                txtCtasBancarias.Text = Trim(RsGral.Fields("CuentasBancarias").Value)
                txtCtasBancarias.Tag = Trim(RsGral.Fields("CuentasBancarias").Value)
                If Trim(RsGral.Fields("Tipo").Value) = "P" Then ' Es un Proveedor
                    optProveedor.Checked = True
                    optProveedor.Tag = True
                    optAcreedor.Checked = False
                    optAcreedor.Tag = False
                Else
                    optProveedor.Checked = False
                    optProveedor.Tag = False
                    optAcreedor.Checked = True
                    optAcreedor.Tag = True
                End If
                If RsGral.Fields("Nacional").Value = True Then
                    optNacional.Checked = True
                    optNacional.Tag = True
                    optExtranjero.Checked = False
                    optExtranjero.Tag = False
                Else
                    optNacional.Checked = False
                    optNacional.Tag = False
                    optExtranjero.Checked = True
                    optExtranjero.Tag = True
                End If
                If Trim(RsGral.Fields("Servicio").Value) = "P" Then
                    optPersonal.Checked = True
                    optPersonal.Tag = True
                    optEmpresa.Checked = False
                    optEmpresa.Tag = False
                Else
                    optPersonal.Checked = False
                    optPersonal.Tag = False
                    optEmpresa.Checked = True
                    optEmpresa.Tag = True
                End If
                If RsGral.Fields("AgenciaAduanal").Value = True Then
                    chkAgenciaAduanal.CheckState = System.Windows.Forms.CheckState.Checked
                    chkAgenciaAduanal.Tag = System.Windows.Forms.CheckState.Checked
                Else
                    chkAgenciaAduanal.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkAgenciaAduanal.Tag = System.Windows.Forms.CheckState.Unchecked
                End If
                rtbObservaciones.Text = Trim(RsGral.Fields("Observaciones").Value)
                rtbObservaciones.Tag = Trim(RsGral.Fields("Observaciones").Value)
            Else
                MsjNoExiste("El ProvAcreed", gstrNombCortoEmpresa)
                Limpiar()
            End If

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

            txtCodProvAcreed.Text = ""
            Nuevo()
            mblnNuevo = True
            mblnCambiosEnCodigo = False
            txtCodProvAcreed.Focus()
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
            If Trim(txtNombre.Text) <> Trim(txtNombre.Tag) Then Exit Function
            If Trim(txtDomicilio.Text) <> Trim(txtDomicilio.Tag) Then Exit Function
            If Trim(txtLocalidad.Text) <> Trim(txtLocalidad.Tag) Then Exit Function
            If Trim(txtCodigoPostal.Text) <> Trim(txtCodigoPostal.Tag) Then Exit Function
            If Trim(txtPais.Text) <> Trim(txtPais.Tag) Then Exit Function
            If Trim(txtTelefonos.Text) <> Trim(txtTelefonos.Tag) Then Exit Function
            If Trim(txtRFC.Text) <> Trim(txtRFC.Tag) Then Exit Function
            If Trim(txtEmail.Text) <> Trim(txtEmail.Tag) Then Exit Function
            If Trim(txtTAXID.Text) <> Trim(txtTAXID.Tag) Then Exit Function
            If Val(txtDiasCredito.Text) <> Val(txtDiasCredito.Tag) Then Exit Function
            If Val(txtDescVolumen.Text) <> Val(txtDescVolumen.Tag) Then Exit Function
            If Val(txtDescFinanciero.Text) <> Val(txtDescFinanciero.Tag) Then Exit Function
            If Trim(txtContactoVentas.Text) <> Trim(txtContactoVentas.Tag) Then Exit Function
            If Trim(txtTelefonosVentas.Text) <> Trim(txtTelefonosVentas.Tag) Then Exit Function
            If Trim(txtContactoPagos.Text) <> Trim(txtContactoPagos.Tag) Then Exit Function
            If Trim(txtTelefonosPagos.Text) <> Trim(txtTelefonosPagos.Tag) Then Exit Function
            If Trim(txtCtasBancarias.Text) <> Trim(txtCtasBancarias.Tag) Then Exit Function
            If optProveedor.Checked <> CBool(optProveedor.Tag) Then Exit Function
            If optAcreedor.Checked <> CBool(optAcreedor.Tag) Then Exit Function
            If optNacional.Checked <> CBool(optNacional.Tag) Then Exit Function
            If optExtranjero.Checked <> CBool(optExtranjero.Tag) Then Exit Function
            If optPersonal.Checked <> CBool(optPersonal.Tag) Then Exit Function
            If optEmpresa.Checked <> CBool(optEmpresa.Tag) Then Exit Function
            If chkAgenciaAduanal.CheckState <> CDbl(chkAgenciaAduanal.Tag) Then Exit Function
            If rtbObservaciones.Text <> rtbObservaciones.Tag Then Exit Function
            Cambios = False

            Exit Function
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Function ValidaDatos() As Object
        'Esta Función Valida que todos los datos en el Formulario se introduzcan, para poder realizar la Alta del registro
        'On Error GoTo MErr
        Try
            'ValidaDatos = False -- No es necesario especificarlo, ya que la funcion se inicializa con falso
            If Len(Trim(txtNombre.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Nombre", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtNombre.Focus()
                Exit Function
            End If
            If optProveedor.Checked = False And optAcreedor.Checked = False Then
                MsgBox(C_msgFALTADATO & "Tipo", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.optProveedor.Focus()
                Exit Function
            End If
            If optNacional.Checked = False And optExtranjero.Checked = False Then
                MsgBox(C_msgFALTADATO & "Nacionalidad", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.optNacional.Focus()
                Exit Function
            End If
            If optPersonal.Checked = False And optEmpresa.Checked = False Then
                MsgBox(C_msgFALTADATO & "Tipo de Servicio", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.optEmpresa.Focus()
                Exit Function
            End If

            If Len(Trim(txtDomicilio.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Domicilio", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtDomicilio.Focus()
                Exit Function
            End If
            If Len(Trim(txtLocalidad.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Localidad", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtLocalidad.Focus()
                Exit Function
            End If
            If Len(Trim(txtPais.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "País", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtPais.Focus()
                Exit Function
            End If
            If Len(Trim(txtCodigoPostal.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Código Postal", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtCodigoPostal.Focus()
                Exit Function
            End If
            If Len(Trim(txtRFC.Text)) = 0 And optNacional.Checked = True Then
                MsgBox(C_msgFALTADATO & "R.F.C.", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtRFC.Focus()
                Exit Function
            End If
            If Len(Trim(txtTelefonos.Text)) = 0 Then
                MsgBox(C_msgFALTADATO & "Teléfono", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtTelefonos.Focus()
                Exit Function
            End If
            If Len(Trim(txtTAXID.Text)) = 0 And optExtranjero.Checked = True Then
                MsgBox(C_msgFALTADATO & "TaxId", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtTAXID.Focus()
                Exit Function
            End If
            'Validar que le Numero de dias no sea mayor de 255
            If CDbl(Numerico(txtDiasCredito.Text)) > 255 Then
                MsgBox("El Número de Días de Crédito debe estar entre 0 y 255" & vbNewLine & "Verifique Profavor...", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
                Me.txtDiasCredito.Focus()
                Exit Function
            End If
            If CDec(txtDescFinanciero.Text) > 100 Then
                MsgBox("El % de Descuento Financiero debe ser menor o igual a 100" & vbNewLine & "Verifique Porfavor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtDescFinanciero.Focus()
                Exit Function
            End If
            If CDec(txtDescVolumen.Text) > 100 Then
                MsgBox("El % de Descuento por Volumen debe ser menor o igual a 100" & vbNewLine & "Verifique Porfavor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                txtDescVolumen.Focus()
                Exit Function
            End If
            ValidaDatos = True
            Exit Function
            'MErr:
        Catch ex As Exception
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Private Sub frmCorpoAbcProvAcreed_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoAbcProvAcreed_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmCorpoAbcProvAcreed_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        'Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Me.Top = 0
        'Me.Left = TwipsToPixelsX(3000)
        InicializaVariables()
        Nuevo()
    End Sub

    Private Sub frmCorpoAbcProvAcreed_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmCorpoAbcProvAcreed_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        '    KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoAbcProvAcreed_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Trim(Me.Tag) = "" Then 'Si el formulario no fue llamado desde CxP sale de manera normal
        '    If Not mblnSALIR Then
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
        '        mblnSALIR = False
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

    Private Sub frmCorpoAbcProvAcreed_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        If Trim(Me.Tag) = "FRMCXPAGREGARPAGO" Then
            'frmCXPAgregarPago.Enabled = True
            'frmCXPAgregarPago.dbcProveedor.Focus()
        End If
        Me.Tag = ""

        ' Me = Nothing
    End Sub

    Private Sub txtCodigoPostal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigoPostal.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodigoPostal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigoPostal.Enter
        SelTextoTxt(txtCodigoPostal)
        Pon_Tool()
    End Sub

    Private Sub txtCodigoPostal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigoPostal.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodProvAcreed_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodProvAcreed.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodProvAcreed_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodProvAcreed.Enter
        strControlActual = UCase("txtCodProvAcreed")
        SelTextoTxt(txtCodProvAcreed)
        Pon_Tool()
    End Sub

    Private Sub txtCodProvAcreed_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodProvAcreed.KeyDown
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
                        Nuevo() 'Lipiar tambien el contenido de todos los controles
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        txtCodProvAcreed.Focus()
                        KeyCode = 0
                        Exit Sub
                End Select
            End If
        End If
    End Sub

    Private Sub txtCodProvAcreed_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodProvAcreed.KeyPress
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
                        txtCodProvAcreed.Focus()
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

    Private Sub txtCodProvAcreed_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodProvAcreed.Leave
        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If Val(Trim(txtCodProvAcreed.Text)) = 0 Then txtCodProvAcreed.Text = "000"
        If mblnCambiosEnCodigo = True And CDbl(Numerico(txtCodProvAcreed.Text)) <> 0 Then 'si hubo cambios en el codigo hace la consulta para llenar los datos
            LlenaDatos()
        End If
    End Sub

    Private Sub txtContactoPagos_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtContactoPagos.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtContactoPagos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtContactoPagos.Enter
        SelTextoTxt(txtContactoPagos)
        Pon_Tool()
    End Sub

    Private Sub txtContactoPagos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtContactoPagos.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtContactoVentas_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtContactoVentas.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtContactoVentas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtContactoVentas.Enter
        SelTextoTxt(txtContactoVentas)
        Pon_Tool()
    End Sub

    Private Sub txtContactoVentas_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtContactoVentas.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCtasBancarias_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCtasBancarias.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCtasBancarias_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCtasBancarias.Enter
        SelTextoTxt(txtCtasBancarias)
        Pon_Tool()
    End Sub

    Private Sub txtCtasBancarias_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCtasBancarias.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDescFinanciero_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescFinanciero.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDescFinanciero_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescFinanciero.Enter
        SelTextoTxt(txtDescFinanciero)
        Pon_Tool()
    End Sub

    Private Sub txtDescFinanciero_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDescFinanciero.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.MskCantidad(txtDescFinanciero.Text, KeyAscii, 3, 2, (txtDescFinanciero.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDescFinanciero_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescFinanciero.Leave
        If ActiveControl.Text <> Me.Text Then Exit Sub
        txtDescFinanciero.Text = Format(CDec(txtDescFinanciero.Text), "0.00")
        If CDec(txtDescFinanciero.Text) > 100 Then
            MsgBox("El % de Descuento Financiero debe ser menor o igual a 100" & vbNewLine & "Verifique Porfavor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtDescFinanciero.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub txtDescVolumen_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescVolumen.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDescVolumen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescVolumen.Enter
        SelTextoTxt(txtDescVolumen)
        Pon_Tool()
    End Sub

    Private Sub txtDescVolumen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDescVolumen.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.MskCantidad(txtDescVolumen.Text, KeyAscii, 3, 2, (txtDescVolumen.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDescVolumen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescVolumen.Leave
        If ActiveControl.Text <> Me.Text Then Exit Sub
        txtDescVolumen.Text = Format(CDec(Numerico(txtDescVolumen.Text)), "0.00")
        If CDec(txtDescVolumen.Text) > 100 Then
            MsgBox("El % de Descuento por Volumen debe ser menor o igual a 100" & vbNewLine & "Verifique Porfavor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            txtDescVolumen.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub txtDiasCredito_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDiasCredito.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDiasCredito_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDiasCredito.Enter
        SelTextoTxt(txtDiasCredito)
        Pon_Tool()
    End Sub

    Private Sub txtDiasCredito_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDiasCredito.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.MskCantidad(txtDiasCredito.Text, KeyAscii, 8, 0, (txtDiasCredito.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDiasCredito_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDiasCredito.Leave
        If ActiveControl.Text <> Me.Text Then Exit Sub
        'Validar que el número proprocionado este entre 0 y 255.
        If CDbl(Numerico(txtDiasCredito.Text)) > 255 Then
            MsgBox("El Número de Días de Crédito debe estar entre 0 y 255" & vbNewLine & "Verifique Profavor...", MsgBoxStyle.Exclamation, gstrNombCortoEmpresa)
            Me.txtDiasCredito.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub txtDomicilio_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtDomicilio_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDomicilio.Enter
        SelTextoTxt(txtDomicilio)
        Pon_Tool()
    End Sub

    Private Sub txtDomicilio_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDomicilio.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtEmail_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEmail.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtEmail_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEmail.Enter
        SelTextoTxt(txtEmail)
        Pon_Tool()
    End Sub

    Private Sub txtEmail_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtEmail.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If Shift = 0 Then
            KeyCode = 97
        End If
    End Sub

    Private Sub txtEmail_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEmail.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "_.-@")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtLocalidad_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocalidad.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtLocalidad_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocalidad.Enter
        SelTextoTxt(txtLocalidad)
        Pon_Tool()
    End Sub

    Private Sub txtLocalidad_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLocalidad.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtNombre_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtNombre_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.Enter
        strControlActual = UCase("txtNombre")
        SelTextoTxt(txtNombre)
        Pon_Tool()
    End Sub

    Private Sub TxtNombre_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNombre.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPais_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPais.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtPais_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPais.Enter
        SelTextoTxt(txtPais)
        Pon_Tool()
    End Sub

    Private Sub txtPais_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPais.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtRFC_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRFC.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtRFC_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRFC.Enter
        SelTextoTxt(txtRFC)
        Pon_Tool()
    End Sub


    'Private Sub txtRFC_KeyDown(KeyCode As Integer, Shift As Integer)
    '    KeyCode = ModEstandar.Valida_RFC(txtRFC, KeyCode, Len(txtRFC) + 1)
    'End Sub

    Private Sub txtRFC_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRFC.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Back Then GoTo EventExitSub

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        KeyAscii = ModEstandar.Valida_RFC(txtRFC.Text, KeyAscii, Len(txtRFC.Text) + 1)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTAXID_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTAXID.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtTAXID_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTAXID.Enter
        SelTextoTxt(txtTAXID)
        Pon_Tool()
    End Sub

    Private Sub txtTAXID_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTAXID.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTelefonos_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTelefonos.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtTelefonos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTelefonos.Enter
        SelTextoTxt(txtTelefonos)
        Pon_Tool()
    End Sub

    Private Sub txtTelefonos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTelefonos.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTelefonosPagos_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTelefonosPagos.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtTelefonosPagos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTelefonosPagos.Enter
        SelTextoTxt(txtTelefonosPagos)
        Pon_Tool()
    End Sub

    Private Sub txtTelefonosPagos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTelefonosPagos.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)

        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTelefonosVentas_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTelefonosVentas.TextChanged
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtTelefonosVentas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTelefonosVentas.Enter
        SelTextoTxt(txtTelefonosVentas)
        Pon_Tool()
    End Sub

    Private Sub optExtranjero_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optExtranjero.CheckedChanged
        If eventSender.Checked Then
            Select Case optExtranjero.Checked
                Case False
                    txtTAXID.Text = ""
                    txtTAXID.Enabled = False
                    'lblProvAcreed(11).Enabled = False 'Etiqueta de TaxId
                    txtRFC.Enabled = True
                    lblProvAcreed(18).Enabled = True 'Etiqueta de RFC
                Case True
                    txtRFC.Text = ""
                    txtRFC.Enabled = False
                    'lblProvAcreed(18).Enabled = False 'Etiqueta de RFC
                    txtTAXID.Enabled = True
                    'lblProvAcreed(11).Enabled = True 'Etiqueta de TaxId
            End Select
        End If
    End Sub

    Private Sub optNacional_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optNacional.CheckedChanged
        If eventSender.Checked Then
            Select Case optNacional.Checked
                Case True
                    txtTAXID.Text = ""
                    txtTAXID.Enabled = False
                    'lblProvAcreed(11).Enabled = False 'Etiqueta de TaxId
                    txtRFC.Enabled = True
                    'lblProvAcreed(18).Enabled = True 'Etiqueta de RFC
                Case False
                    txtRFC.Text = ""
                    txtRFC.Enabled = False
                    'lblProvAcreed(18).Enabled = False 'Etiqueta de RFC
                    txtTAXID.Enabled = True
                    'lblProvAcreed(11).Enabled = True 'Etiqueta de TaxId
            End Select
        End If
    End Sub

    Private Sub txtTelefonosVentas_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTelefonosVentas.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'UPGRADE_WARNING: Couldn't resolve default property of object ModEstandar.gp_CampoMayusculas(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.txtRFC = New System.Windows.Forms.TextBox()
        Me.txtPais = New System.Windows.Forms.TextBox()
        Me.txtCodigoPostal = New System.Windows.Forms.TextBox()
        Me.txtTelefonos = New System.Windows.Forms.TextBox()
        Me.txtDomicilio = New System.Windows.Forms.TextBox()
        Me.txtLocalidad = New System.Windows.Forms.TextBox()
        Me.txtContactoPagos = New System.Windows.Forms.TextBox()
        Me.txtCtasBancarias = New System.Windows.Forms.TextBox()
        Me.txtTelefonosPagos = New System.Windows.Forms.TextBox()
        Me.txtTelefonosVentas = New System.Windows.Forms.TextBox()
        Me.txtDescVolumen = New System.Windows.Forms.TextBox()
        Me.txtDiasCredito = New System.Windows.Forms.TextBox()
        Me.txtContactoVentas = New System.Windows.Forms.TextBox()
        Me.txtDescFinanciero = New System.Windows.Forms.TextBox()
        Me.txtTAXID = New System.Windows.Forms.TextBox()
        Me.optEmpresa = New System.Windows.Forms.RadioButton()
        Me.optPersonal = New System.Windows.Forms.RadioButton()
        Me.chkAgenciaAduanal = New System.Windows.Forms.CheckBox()
        Me.optAcreedor = New System.Windows.Forms.RadioButton()
        Me.optProveedor = New System.Windows.Forms.RadioButton()
        Me.optExtranjero = New System.Windows.Forms.RadioButton()
        Me.optNacional = New System.Windows.Forms.RadioButton()
        Me.chkMostrarTodos = New System.Windows.Forms.CheckBox()
        Me.txtCodProvAcreed = New System.Windows.Forms.TextBox()
        Me.Frame8 = New System.Windows.Forms.GroupBox()
        Me.Frame9 = New System.Windows.Forms.GroupBox()
        Me._lblProvAcreed_18 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_9 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_7 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_5 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_4 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_1 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_3 = New System.Windows.Forms.Label()
        Me.Frame10 = New System.Windows.Forms.GroupBox()
        Me.rtbObservaciones = New System.Windows.Forms.RichTextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_17 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_16 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_10 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_6 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_15 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_14 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_13 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_12 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_11 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.fraServicio = New System.Windows.Forms.Panel()
        Me.fraTipo = New System.Windows.Forms.Panel()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.Frame7 = New System.Windows.Forms.GroupBox()
        Me.fraNacional = New System.Windows.Forms.Panel()
        Me.Frame6 = New System.Windows.Forms.GroupBox()
        Me._lblProvAcreed_2 = New System.Windows.Forms.Label()
        Me._lblProvAcreed_0 = New System.Windows.Forms.Label()
        Me.lblProvAcreed = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.Frame8.SuspendLayout()
        Me.Frame10.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.fraServicio.SuspendLayout()
        Me.fraTipo.SuspendLayout()
        Me.fraNacional.SuspendLayout()
        CType(Me.lblProvAcreed, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtNombre
        '
        Me.txtNombre.AcceptsReturn = True
        Me.txtNombre.BackColor = System.Drawing.SystemColors.Window
        Me.txtNombre.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombre.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombre.Location = New System.Drawing.Point(69, 34)
        Me.txtNombre.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNombre.MaxLength = 50
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombre.Size = New System.Drawing.Size(238, 20)
        Me.txtNombre.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtNombre, "Nombre del Prov/Acreed")
        '
        'txtEmail
        '
        Me.txtEmail.AcceptsReturn = True
        Me.txtEmail.BackColor = System.Drawing.SystemColors.Window
        Me.txtEmail.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEmail.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtEmail.Location = New System.Drawing.Point(58, 120)
        Me.txtEmail.Margin = New System.Windows.Forms.Padding(2)
        Me.txtEmail.MaxLength = 50
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEmail.Size = New System.Drawing.Size(278, 20)
        Me.txtEmail.TabIndex = 29
        Me.ToolTip1.SetToolTip(Me.txtEmail, "E-mail del Prov/Acreed")
        '
        'txtRFC
        '
        Me.txtRFC.AcceptsReturn = True
        Me.txtRFC.BackColor = System.Drawing.SystemColors.Window
        Me.txtRFC.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRFC.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRFC.Location = New System.Drawing.Point(264, 63)
        Me.txtRFC.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRFC.MaxLength = 15
        Me.txtRFC.Name = "txtRFC"
        Me.txtRFC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRFC.Size = New System.Drawing.Size(126, 20)
        Me.txtRFC.TabIndex = 24
        Me.ToolTip1.SetToolTip(Me.txtRFC, "RFC del Prov/Acreed")
        '
        'txtPais
        '
        Me.txtPais.AcceptsReturn = True
        Me.txtPais.BackColor = System.Drawing.SystemColors.Window
        Me.txtPais.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPais.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPais.Location = New System.Drawing.Point(263, 40)
        Me.txtPais.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPais.MaxLength = 20
        Me.txtPais.Name = "txtPais"
        Me.txtPais.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPais.Size = New System.Drawing.Size(133, 20)
        Me.txtPais.TabIndex = 20
        Me.ToolTip1.SetToolTip(Me.txtPais, "País")
        '
        'txtCodigoPostal
        '
        Me.txtCodigoPostal.AcceptsReturn = True
        Me.txtCodigoPostal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodigoPostal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodigoPostal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodigoPostal.Location = New System.Drawing.Point(81, 64)
        Me.txtCodigoPostal.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodigoPostal.MaxLength = 10
        Me.txtCodigoPostal.Name = "txtCodigoPostal"
        Me.txtCodigoPostal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodigoPostal.Size = New System.Drawing.Size(126, 20)
        Me.txtCodigoPostal.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.txtCodigoPostal, "Código Postal")
        '
        'txtTelefonos
        '
        Me.txtTelefonos.AcceptsReturn = True
        Me.txtTelefonos.BackColor = System.Drawing.SystemColors.Window
        Me.txtTelefonos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTelefonos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTelefonos.Location = New System.Drawing.Point(73, 95)
        Me.txtTelefonos.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTelefonos.MaxLength = 50
        Me.txtTelefonos.Name = "txtTelefonos"
        Me.txtTelefonos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTelefonos.Size = New System.Drawing.Size(263, 20)
        Me.txtTelefonos.TabIndex = 27
        Me.ToolTip1.SetToolTip(Me.txtTelefonos, "Teléfonos del Prov/Acreed")
        '
        'txtDomicilio
        '
        Me.txtDomicilio.AcceptsReturn = True
        Me.txtDomicilio.BackColor = System.Drawing.SystemColors.Window
        Me.txtDomicilio.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDomicilio.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDomicilio.Location = New System.Drawing.Point(66, 19)
        Me.txtDomicilio.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDomicilio.MaxLength = 150
        Me.txtDomicilio.Name = "txtDomicilio"
        Me.txtDomicilio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDomicilio.Size = New System.Drawing.Size(330, 20)
        Me.txtDomicilio.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.txtDomicilio, "Domicilio del Prov/Acreed")
        '
        'txtLocalidad
        '
        Me.txtLocalidad.AcceptsReturn = True
        Me.txtLocalidad.BackColor = System.Drawing.SystemColors.Window
        Me.txtLocalidad.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLocalidad.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLocalidad.Location = New System.Drawing.Point(66, 41)
        Me.txtLocalidad.Margin = New System.Windows.Forms.Padding(2)
        Me.txtLocalidad.MaxLength = 20
        Me.txtLocalidad.Name = "txtLocalidad"
        Me.txtLocalidad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLocalidad.Size = New System.Drawing.Size(141, 20)
        Me.txtLocalidad.TabIndex = 18
        Me.ToolTip1.SetToolTip(Me.txtLocalidad, "Nombre de la Localidad")
        '
        'txtContactoPagos
        '
        Me.txtContactoPagos.AcceptsReturn = True
        Me.txtContactoPagos.BackColor = System.Drawing.SystemColors.Window
        Me.txtContactoPagos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtContactoPagos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtContactoPagos.Location = New System.Drawing.Point(114, 115)
        Me.txtContactoPagos.Margin = New System.Windows.Forms.Padding(2)
        Me.txtContactoPagos.MaxLength = 40
        Me.txtContactoPagos.Name = "txtContactoPagos"
        Me.txtContactoPagos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtContactoPagos.Size = New System.Drawing.Size(290, 20)
        Me.txtContactoPagos.TabIndex = 44
        Me.ToolTip1.SetToolTip(Me.txtContactoPagos, "Contacto en Pagos")
        '
        'txtCtasBancarias
        '
        Me.txtCtasBancarias.AcceptsReturn = True
        Me.txtCtasBancarias.BackColor = System.Drawing.SystemColors.Window
        Me.txtCtasBancarias.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCtasBancarias.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCtasBancarias.Location = New System.Drawing.Point(113, 159)
        Me.txtCtasBancarias.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCtasBancarias.MaxLength = 50
        Me.txtCtasBancarias.Name = "txtCtasBancarias"
        Me.txtCtasBancarias.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCtasBancarias.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtCtasBancarias.Size = New System.Drawing.Size(290, 20)
        Me.txtCtasBancarias.TabIndex = 48
        Me.ToolTip1.SetToolTip(Me.txtCtasBancarias, "Cuentas Bancarias")
        '
        'txtTelefonosPagos
        '
        Me.txtTelefonosPagos.AcceptsReturn = True
        Me.txtTelefonosPagos.BackColor = System.Drawing.SystemColors.Window
        Me.txtTelefonosPagos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTelefonosPagos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTelefonosPagos.Location = New System.Drawing.Point(116, 137)
        Me.txtTelefonosPagos.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTelefonosPagos.MaxLength = 50
        Me.txtTelefonosPagos.Name = "txtTelefonosPagos"
        Me.txtTelefonosPagos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTelefonosPagos.Size = New System.Drawing.Size(290, 20)
        Me.txtTelefonosPagos.TabIndex = 46
        Me.ToolTip1.SetToolTip(Me.txtTelefonosPagos, "Teléfonos de Pagos")
        '
        'txtTelefonosVentas
        '
        Me.txtTelefonosVentas.AcceptsReturn = True
        Me.txtTelefonosVentas.BackColor = System.Drawing.SystemColors.Window
        Me.txtTelefonosVentas.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTelefonosVentas.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTelefonosVentas.Location = New System.Drawing.Point(116, 91)
        Me.txtTelefonosVentas.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTelefonosVentas.MaxLength = 50
        Me.txtTelefonosVentas.Name = "txtTelefonosVentas"
        Me.txtTelefonosVentas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTelefonosVentas.Size = New System.Drawing.Size(290, 20)
        Me.txtTelefonosVentas.TabIndex = 42
        Me.ToolTip1.SetToolTip(Me.txtTelefonosVentas, "Teléfonos de  Ventas")
        '
        'txtDescVolumen
        '
        Me.txtDescVolumen.AcceptsReturn = True
        Me.txtDescVolumen.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescVolumen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescVolumen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescVolumen.Location = New System.Drawing.Point(371, 21)
        Me.txtDescVolumen.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescVolumen.MaxLength = 8
        Me.txtDescVolumen.Name = "txtDescVolumen"
        Me.txtDescVolumen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescVolumen.Size = New System.Drawing.Size(50, 20)
        Me.txtDescVolumen.TabIndex = 34
        Me.txtDescVolumen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDescVolumen, "Descuento Por Volumen")
        '
        'txtDiasCredito
        '
        Me.txtDiasCredito.AcceptsReturn = True
        Me.txtDiasCredito.BackColor = System.Drawing.SystemColors.Window
        Me.txtDiasCredito.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDiasCredito.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDiasCredito.Location = New System.Drawing.Point(96, 21)
        Me.txtDiasCredito.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDiasCredito.MaxLength = 3
        Me.txtDiasCredito.Name = "txtDiasCredito"
        Me.txtDiasCredito.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDiasCredito.Size = New System.Drawing.Size(50, 20)
        Me.txtDiasCredito.TabIndex = 32
        Me.txtDiasCredito.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDiasCredito, "Días de Crédito")
        '
        'txtContactoVentas
        '
        Me.txtContactoVentas.AcceptsReturn = True
        Me.txtContactoVentas.BackColor = System.Drawing.SystemColors.Window
        Me.txtContactoVentas.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtContactoVentas.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtContactoVentas.Location = New System.Drawing.Point(114, 67)
        Me.txtContactoVentas.Margin = New System.Windows.Forms.Padding(2)
        Me.txtContactoVentas.MaxLength = 40
        Me.txtContactoVentas.Name = "txtContactoVentas"
        Me.txtContactoVentas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtContactoVentas.Size = New System.Drawing.Size(290, 20)
        Me.txtContactoVentas.TabIndex = 40
        Me.ToolTip1.SetToolTip(Me.txtContactoVentas, "Contacto en Ventas")
        '
        'txtDescFinanciero
        '
        Me.txtDescFinanciero.AcceptsReturn = True
        Me.txtDescFinanciero.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescFinanciero.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescFinanciero.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescFinanciero.Location = New System.Drawing.Point(123, 43)
        Me.txtDescFinanciero.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescFinanciero.MaxLength = 8
        Me.txtDescFinanciero.Name = "txtDescFinanciero"
        Me.txtDescFinanciero.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescFinanciero.Size = New System.Drawing.Size(50, 20)
        Me.txtDescFinanciero.TabIndex = 36
        Me.txtDescFinanciero.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDescFinanciero, "Descuento Financiero")
        '
        'txtTAXID
        '
        Me.txtTAXID.AcceptsReturn = True
        Me.txtTAXID.BackColor = System.Drawing.SystemColors.Window
        Me.txtTAXID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTAXID.Enabled = False
        Me.txtTAXID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTAXID.Location = New System.Drawing.Point(248, 43)
        Me.txtTAXID.Margin = New System.Windows.Forms.Padding(2)
        Me.txtTAXID.MaxLength = 20
        Me.txtTAXID.Name = "txtTAXID"
        Me.txtTAXID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTAXID.Size = New System.Drawing.Size(142, 20)
        Me.txtTAXID.TabIndex = 37
        Me.ToolTip1.SetToolTip(Me.txtTAXID, "TAXID")
        '
        'optEmpresa
        '
        Me.optEmpresa.BackColor = System.Drawing.Color.Silver
        Me.optEmpresa.Checked = True
        Me.optEmpresa.Cursor = System.Windows.Forms.Cursors.Default
        Me.optEmpresa.ForeColor = System.Drawing.SystemColors.WindowText
        Me.optEmpresa.Location = New System.Drawing.Point(6, 0)
        Me.optEmpresa.Margin = New System.Windows.Forms.Padding(2)
        Me.optEmpresa.Name = "optEmpresa"
        Me.optEmpresa.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optEmpresa.Size = New System.Drawing.Size(107, 20)
        Me.optEmpresa.TabIndex = 11
        Me.optEmpresa.TabStop = True
        Me.optEmpresa.Text = "De la Empresa"
        Me.ToolTip1.SetToolTip(Me.optEmpresa, "Prov/Acreed de la Empresa")
        Me.optEmpresa.UseVisualStyleBackColor = False
        '
        'optPersonal
        '
        Me.optPersonal.BackColor = System.Drawing.Color.Silver
        Me.optPersonal.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPersonal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.optPersonal.Location = New System.Drawing.Point(6, 20)
        Me.optPersonal.Margin = New System.Windows.Forms.Padding(2)
        Me.optPersonal.Name = "optPersonal"
        Me.optPersonal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPersonal.Size = New System.Drawing.Size(74, 20)
        Me.optPersonal.TabIndex = 12
        Me.optPersonal.TabStop = True
        Me.optPersonal.Text = "Personal"
        Me.ToolTip1.SetToolTip(Me.optPersonal, "Prov/Acreed Personal")
        Me.optPersonal.UseVisualStyleBackColor = False
        '
        'chkAgenciaAduanal
        '
        Me.chkAgenciaAduanal.BackColor = System.Drawing.Color.Silver
        Me.chkAgenciaAduanal.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAgenciaAduanal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.chkAgenciaAduanal.Location = New System.Drawing.Point(368, 11)
        Me.chkAgenciaAduanal.Margin = New System.Windows.Forms.Padding(2)
        Me.chkAgenciaAduanal.Name = "chkAgenciaAduanal"
        Me.chkAgenciaAduanal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAgenciaAduanal.Size = New System.Drawing.Size(85, 41)
        Me.chkAgenciaAduanal.TabIndex = 13
        Me.chkAgenciaAduanal.Text = "Es Agencia Aduanal"
        Me.ToolTip1.SetToolTip(Me.chkAgenciaAduanal, "Es Agencia Aduanal")
        Me.chkAgenciaAduanal.UseVisualStyleBackColor = False
        '
        'optAcreedor
        '
        Me.optAcreedor.BackColor = System.Drawing.Color.Silver
        Me.optAcreedor.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAcreedor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.optAcreedor.Location = New System.Drawing.Point(6, 20)
        Me.optAcreedor.Margin = New System.Windows.Forms.Padding(2)
        Me.optAcreedor.Name = "optAcreedor"
        Me.optAcreedor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAcreedor.Size = New System.Drawing.Size(82, 20)
        Me.optAcreedor.TabIndex = 6
        Me.optAcreedor.TabStop = True
        Me.optAcreedor.Text = "Acreedor"
        Me.ToolTip1.SetToolTip(Me.optAcreedor, "Acreedor")
        Me.optAcreedor.UseVisualStyleBackColor = False
        '
        'optProveedor
        '
        Me.optProveedor.BackColor = System.Drawing.Color.Silver
        Me.optProveedor.Checked = True
        Me.optProveedor.Cursor = System.Windows.Forms.Cursors.Default
        Me.optProveedor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.optProveedor.Location = New System.Drawing.Point(6, 0)
        Me.optProveedor.Margin = New System.Windows.Forms.Padding(2)
        Me.optProveedor.Name = "optProveedor"
        Me.optProveedor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optProveedor.Size = New System.Drawing.Size(82, 20)
        Me.optProveedor.TabIndex = 5
        Me.optProveedor.TabStop = True
        Me.optProveedor.Text = "Proveedor"
        Me.ToolTip1.SetToolTip(Me.optProveedor, "Proveedor")
        Me.optProveedor.UseVisualStyleBackColor = False
        '
        'optExtranjero
        '
        Me.optExtranjero.BackColor = System.Drawing.Color.Silver
        Me.optExtranjero.Cursor = System.Windows.Forms.Cursors.Default
        Me.optExtranjero.ForeColor = System.Drawing.SystemColors.WindowText
        Me.optExtranjero.Location = New System.Drawing.Point(6, 20)
        Me.optExtranjero.Margin = New System.Windows.Forms.Padding(2)
        Me.optExtranjero.Name = "optExtranjero"
        Me.optExtranjero.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optExtranjero.Size = New System.Drawing.Size(77, 20)
        Me.optExtranjero.TabIndex = 9
        Me.optExtranjero.TabStop = True
        Me.optExtranjero.Text = "Extranjero"
        Me.ToolTip1.SetToolTip(Me.optExtranjero, "Prov/Acreed Extranjero")
        Me.optExtranjero.UseVisualStyleBackColor = False
        '
        'optNacional
        '
        Me.optNacional.BackColor = System.Drawing.Color.Silver
        Me.optNacional.Checked = True
        Me.optNacional.Cursor = System.Windows.Forms.Cursors.Default
        Me.optNacional.ForeColor = System.Drawing.SystemColors.WindowText
        Me.optNacional.Location = New System.Drawing.Point(6, 0)
        Me.optNacional.Margin = New System.Windows.Forms.Padding(2)
        Me.optNacional.Name = "optNacional"
        Me.optNacional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optNacional.Size = New System.Drawing.Size(77, 20)
        Me.optNacional.TabIndex = 8
        Me.optNacional.TabStop = True
        Me.optNacional.Text = "Nacional"
        Me.ToolTip1.SetToolTip(Me.optNacional, "Prov/Acreed Nacional")
        Me.optNacional.UseVisualStyleBackColor = False
        '
        'chkMostrarTodos
        '
        Me.chkMostrarTodos.BackColor = System.Drawing.Color.Silver
        Me.chkMostrarTodos.Checked = True
        Me.chkMostrarTodos.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMostrarTodos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMostrarTodos.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkMostrarTodos.Location = New System.Drawing.Point(320, 10)
        Me.chkMostrarTodos.Margin = New System.Windows.Forms.Padding(2)
        Me.chkMostrarTodos.Name = "chkMostrarTodos"
        Me.chkMostrarTodos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMostrarTodos.Size = New System.Drawing.Size(148, 44)
        Me.chkMostrarTodos.TabIndex = 53
        Me.chkMostrarTodos.Text = "Mostrar todos los Prov/Acreed"
        Me.chkMostrarTodos.UseVisualStyleBackColor = False
        '
        'txtCodProvAcreed
        '
        Me.txtCodProvAcreed.AcceptsReturn = True
        Me.txtCodProvAcreed.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodProvAcreed.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodProvAcreed.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodProvAcreed.Location = New System.Drawing.Point(69, 10)
        Me.txtCodProvAcreed.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodProvAcreed.MaxLength = 3
        Me.txtCodProvAcreed.Name = "txtCodProvAcreed"
        Me.txtCodProvAcreed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodProvAcreed.Size = New System.Drawing.Size(44, 20)
        Me.txtCodProvAcreed.TabIndex = 1
        '
        'Frame8
        '
        Me.Frame8.BackColor = System.Drawing.Color.Silver
        Me.Frame8.Controls.Add(Me.txtEmail)
        Me.Frame8.Controls.Add(Me.txtRFC)
        Me.Frame8.Controls.Add(Me.Frame9)
        Me.Frame8.Controls.Add(Me.txtPais)
        Me.Frame8.Controls.Add(Me.txtCodigoPostal)
        Me.Frame8.Controls.Add(Me.txtTelefonos)
        Me.Frame8.Controls.Add(Me.txtDomicilio)
        Me.Frame8.Controls.Add(Me.txtLocalidad)
        Me.Frame8.Controls.Add(Me._lblProvAcreed_18)
        Me.Frame8.Controls.Add(Me._lblProvAcreed_9)
        Me.Frame8.Controls.Add(Me._lblProvAcreed_7)
        Me.Frame8.Controls.Add(Me._lblProvAcreed_5)
        Me.Frame8.Controls.Add(Me._lblProvAcreed_4)
        Me.Frame8.Controls.Add(Me._lblProvAcreed_1)
        Me.Frame8.Controls.Add(Me._lblProvAcreed_3)
        Me.Frame8.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame8.Location = New System.Drawing.Point(14, 121)
        Me.Frame8.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame8.Name = "Frame8"
        Me.Frame8.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame8.Size = New System.Drawing.Size(453, 145)
        Me.Frame8.TabIndex = 14
        Me.Frame8.TabStop = False
        Me.Frame8.Text = " Información &Básica "
        '
        'Frame9
        '
        Me.Frame9.BackColor = System.Drawing.SystemColors.Control
        Me.Frame9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame9.Location = New System.Drawing.Point(4, 87)
        Me.Frame9.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame9.Name = "Frame9"
        Me.Frame9.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame9.Size = New System.Drawing.Size(386, 2)
        Me.Frame9.TabIndex = 25
        Me.Frame9.TabStop = False
        '
        '_lblProvAcreed_18
        '
        Me._lblProvAcreed_18.AutoSize = True
        Me._lblProvAcreed_18.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_18.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_18.Location = New System.Drawing.Point(232, 67)
        Me._lblProvAcreed_18.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_18.Name = "_lblProvAcreed_18"
        Me._lblProvAcreed_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_18.Size = New System.Drawing.Size(34, 13)
        Me._lblProvAcreed_18.TabIndex = 23
        Me._lblProvAcreed_18.Text = "RFC :"
        Me._lblProvAcreed_18.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblProvAcreed_9
        '
        Me._lblProvAcreed_9.AutoSize = True
        Me._lblProvAcreed_9.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_9.Location = New System.Drawing.Point(12, 120)
        Me._lblProvAcreed_9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_9.Name = "_lblProvAcreed_9"
        Me._lblProvAcreed_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_9.Size = New System.Drawing.Size(41, 13)
        Me._lblProvAcreed_9.TabIndex = 28
        Me._lblProvAcreed_9.Text = "E-mail: "
        '
        '_lblProvAcreed_7
        '
        Me._lblProvAcreed_7.AutoSize = True
        Me._lblProvAcreed_7.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_7.Location = New System.Drawing.Point(226, 46)
        Me._lblProvAcreed_7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_7.Name = "_lblProvAcreed_7"
        Me._lblProvAcreed_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_7.Size = New System.Drawing.Size(35, 13)
        Me._lblProvAcreed_7.TabIndex = 19
        Me._lblProvAcreed_7.Text = "País :"
        Me._lblProvAcreed_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblProvAcreed_5
        '
        Me._lblProvAcreed_5.AutoSize = True
        Me._lblProvAcreed_5.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_5.Location = New System.Drawing.Point(9, 64)
        Me._lblProvAcreed_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_5.Name = "_lblProvAcreed_5"
        Me._lblProvAcreed_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_5.Size = New System.Drawing.Size(75, 13)
        Me._lblProvAcreed_5.TabIndex = 21
        Me._lblProvAcreed_5.Text = "Código Postal:"
        '
        '_lblProvAcreed_4
        '
        Me._lblProvAcreed_4.AutoSize = True
        Me._lblProvAcreed_4.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_4.Location = New System.Drawing.Point(12, 98)
        Me._lblProvAcreed_4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_4.Name = "_lblProvAcreed_4"
        Me._lblProvAcreed_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_4.Size = New System.Drawing.Size(57, 13)
        Me._lblProvAcreed_4.TabIndex = 26
        Me._lblProvAcreed_4.Text = "Teléfonos:"
        '
        '_lblProvAcreed_1
        '
        Me._lblProvAcreed_1.AutoSize = True
        Me._lblProvAcreed_1.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_1.Location = New System.Drawing.Point(9, 43)
        Me._lblProvAcreed_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_1.Name = "_lblProvAcreed_1"
        Me._lblProvAcreed_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_1.Size = New System.Drawing.Size(56, 13)
        Me._lblProvAcreed_1.TabIndex = 17
        Me._lblProvAcreed_1.Text = "Localidad:"
        '
        '_lblProvAcreed_3
        '
        Me._lblProvAcreed_3.AutoSize = True
        Me._lblProvAcreed_3.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_3.Location = New System.Drawing.Point(8, 20)
        Me._lblProvAcreed_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_3.Name = "_lblProvAcreed_3"
        Me._lblProvAcreed_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_3.Size = New System.Drawing.Size(52, 13)
        Me._lblProvAcreed_3.TabIndex = 15
        Me._lblProvAcreed_3.Text = "Domicilio:"
        '
        'Frame10
        '
        Me.Frame10.BackColor = System.Drawing.Color.Silver
        Me.Frame10.Controls.Add(Me.rtbObservaciones)
        Me.Frame10.Controls.Add(Me.txtContactoPagos)
        Me.Frame10.Controls.Add(Me.txtCtasBancarias)
        Me.Frame10.Controls.Add(Me.txtTelefonosPagos)
        Me.Frame10.Controls.Add(Me.txtTelefonosVentas)
        Me.Frame10.Controls.Add(Me.txtDescVolumen)
        Me.Frame10.Controls.Add(Me.txtDiasCredito)
        Me.Frame10.Controls.Add(Me.txtContactoVentas)
        Me.Frame10.Controls.Add(Me.txtDescFinanciero)
        Me.Frame10.Controls.Add(Me.txtTAXID)
        Me.Frame10.Controls.Add(Me.Label1)
        Me.Frame10.Controls.Add(Me._lblProvAcreed_17)
        Me.Frame10.Controls.Add(Me._lblProvAcreed_16)
        Me.Frame10.Controls.Add(Me._lblProvAcreed_10)
        Me.Frame10.Controls.Add(Me._lblProvAcreed_6)
        Me.Frame10.Controls.Add(Me._lblProvAcreed_15)
        Me.Frame10.Controls.Add(Me._lblProvAcreed_14)
        Me.Frame10.Controls.Add(Me._lblProvAcreed_13)
        Me.Frame10.Controls.Add(Me._lblProvAcreed_12)
        Me.Frame10.Controls.Add(Me._lblProvAcreed_11)
        Me.Frame10.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame10.Location = New System.Drawing.Point(14, 273)
        Me.Frame10.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame10.Name = "Frame10"
        Me.Frame10.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame10.Size = New System.Drawing.Size(454, 260)
        Me.Frame10.TabIndex = 30
        Me.Frame10.TabStop = False
        Me.Frame10.Text = " Información &Adicional"
        '
        'rtbObservaciones
        '
        Me.rtbObservaciones.Location = New System.Drawing.Point(11, 195)
        Me.rtbObservaciones.Margin = New System.Windows.Forms.Padding(2)
        Me.rtbObservaciones.Name = "rtbObservaciones"
        Me.rtbObservaciones.Size = New System.Drawing.Size(377, 53)
        Me.rtbObservaciones.TabIndex = 55
        Me.rtbObservaciones.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Silver
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 179)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 54
        Me.Label1.Text = "Observaciones:"
        '
        '_lblProvAcreed_17
        '
        Me._lblProvAcreed_17.AutoSize = True
        Me._lblProvAcreed_17.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_17.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_17.Location = New System.Drawing.Point(10, 159)
        Me._lblProvAcreed_17.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_17.Name = "_lblProvAcreed_17"
        Me._lblProvAcreed_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_17.Size = New System.Drawing.Size(99, 13)
        Me._lblProvAcreed_17.TabIndex = 47
        Me._lblProvAcreed_17.Text = "Cuentas Bancarias:"
        '
        '_lblProvAcreed_16
        '
        Me._lblProvAcreed_16.AutoSize = True
        Me._lblProvAcreed_16.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_16.Location = New System.Drawing.Point(10, 137)
        Me._lblProvAcreed_16.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_16.Name = "_lblProvAcreed_16"
        Me._lblProvAcreed_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_16.Size = New System.Drawing.Size(105, 13)
        Me._lblProvAcreed_16.TabIndex = 45
        Me._lblProvAcreed_16.Text = "Teléfonos de Pagos:"
        '
        '_lblProvAcreed_10
        '
        Me._lblProvAcreed_10.AutoSize = True
        Me._lblProvAcreed_10.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_10.Location = New System.Drawing.Point(10, 115)
        Me._lblProvAcreed_10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_10.Name = "_lblProvAcreed_10"
        Me._lblProvAcreed_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_10.Size = New System.Drawing.Size(101, 13)
        Me._lblProvAcreed_10.TabIndex = 43
        Me._lblProvAcreed_10.Text = "Contacto en Pagos:"
        '
        '_lblProvAcreed_6
        '
        Me._lblProvAcreed_6.AutoSize = True
        Me._lblProvAcreed_6.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_6.Location = New System.Drawing.Point(10, 91)
        Me._lblProvAcreed_6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_6.Name = "_lblProvAcreed_6"
        Me._lblProvAcreed_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_6.Size = New System.Drawing.Size(108, 13)
        Me._lblProvAcreed_6.TabIndex = 41
        Me._lblProvAcreed_6.Text = "Teléfonos de Ventas:"
        '
        '_lblProvAcreed_15
        '
        Me._lblProvAcreed_15.AutoSize = True
        Me._lblProvAcreed_15.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_15.Location = New System.Drawing.Point(243, 24)
        Me._lblProvAcreed_15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_15.Name = "_lblProvAcreed_15"
        Me._lblProvAcreed_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_15.Size = New System.Drawing.Size(124, 13)
        Me._lblProvAcreed_15.TabIndex = 33
        Me._lblProvAcreed_15.Text = "Descuento por Volumen:"
        '
        '_lblProvAcreed_14
        '
        Me._lblProvAcreed_14.AutoSize = True
        Me._lblProvAcreed_14.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_14.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_14.Location = New System.Drawing.Point(10, 24)
        Me._lblProvAcreed_14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_14.Name = "_lblProvAcreed_14"
        Me._lblProvAcreed_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_14.Size = New System.Drawing.Size(83, 13)
        Me._lblProvAcreed_14.TabIndex = 31
        Me._lblProvAcreed_14.Text = "Días de crédito:"
        '
        '_lblProvAcreed_13
        '
        Me._lblProvAcreed_13.AutoSize = True
        Me._lblProvAcreed_13.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_13.Location = New System.Drawing.Point(10, 69)
        Me._lblProvAcreed_13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_13.Name = "_lblProvAcreed_13"
        Me._lblProvAcreed_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_13.Size = New System.Drawing.Size(104, 13)
        Me._lblProvAcreed_13.TabIndex = 39
        Me._lblProvAcreed_13.Text = "Contacto en Ventas:"
        '
        '_lblProvAcreed_12
        '
        Me._lblProvAcreed_12.AutoSize = True
        Me._lblProvAcreed_12.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_12.Location = New System.Drawing.Point(10, 46)
        Me._lblProvAcreed_12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_12.Name = "_lblProvAcreed_12"
        Me._lblProvAcreed_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_12.Size = New System.Drawing.Size(114, 13)
        Me._lblProvAcreed_12.TabIndex = 35
        Me._lblProvAcreed_12.Text = "Descuento Financiero:"
        '
        '_lblProvAcreed_11
        '
        Me._lblProvAcreed_11.AutoSize = True
        Me._lblProvAcreed_11.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_11.Enabled = False
        Me._lblProvAcreed_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_11.Location = New System.Drawing.Point(213, 46)
        Me._lblProvAcreed_11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_11.Name = "_lblProvAcreed_11"
        Me._lblProvAcreed_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_11.Size = New System.Drawing.Size(39, 13)
        Me._lblProvAcreed_11.TabIndex = 38
        Me._lblProvAcreed_11.Text = "TaxID:"
        Me._lblProvAcreed_11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.Color.Silver
        Me.Frame1.Controls.Add(Me.fraServicio)
        Me.Frame1.Controls.Add(Me.chkAgenciaAduanal)
        Me.Frame1.Controls.Add(Me.fraTipo)
        Me.Frame1.Controls.Add(Me.Frame7)
        Me.Frame1.Controls.Add(Me.fraNacional)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(14, 58)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(454, 59)
        Me.Frame1.TabIndex = 4
        Me.Frame1.TabStop = False
        Me.Frame1.Text = " Clasi&ficación "
        '
        'fraServicio
        '
        Me.fraServicio.BackColor = System.Drawing.Color.Silver
        Me.fraServicio.Controls.Add(Me.optEmpresa)
        Me.fraServicio.Controls.Add(Me.optPersonal)
        Me.fraServicio.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraServicio.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.fraServicio.Location = New System.Drawing.Point(223, 14)
        Me.fraServicio.Margin = New System.Windows.Forms.Padding(2)
        Me.fraServicio.Name = "fraServicio"
        Me.fraServicio.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraServicio.Size = New System.Drawing.Size(141, 44)
        Me.fraServicio.TabIndex = 52
        '
        'fraTipo
        '
        Me.fraTipo.BackColor = System.Drawing.Color.Silver
        Me.fraTipo.Controls.Add(Me.optAcreedor)
        Me.fraTipo.Controls.Add(Me.optProveedor)
        Me.fraTipo.Controls.Add(Me.Frame5)
        Me.fraTipo.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraTipo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTipo.Location = New System.Drawing.Point(30, 14)
        Me.fraTipo.Margin = New System.Windows.Forms.Padding(2)
        Me.fraTipo.Name = "fraTipo"
        Me.fraTipo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTipo.Size = New System.Drawing.Size(85, 44)
        Me.fraTipo.TabIndex = 51
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame5.Location = New System.Drawing.Point(66, 0)
        Me.Frame5.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(2, 38)
        Me.Frame5.TabIndex = 7
        Me.Frame5.TabStop = False
        '
        'Frame7
        '
        Me.Frame7.BackColor = System.Drawing.SystemColors.Control
        Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame7.Location = New System.Drawing.Point(252, 13)
        Me.Frame7.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame7.Name = "Frame7"
        Me.Frame7.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame7.Size = New System.Drawing.Size(2, 38)
        Me.Frame7.TabIndex = 50
        Me.Frame7.TabStop = False
        '
        'fraNacional
        '
        Me.fraNacional.BackColor = System.Drawing.Color.Silver
        Me.fraNacional.Controls.Add(Me.optExtranjero)
        Me.fraNacional.Controls.Add(Me.optNacional)
        Me.fraNacional.Controls.Add(Me.Frame6)
        Me.fraNacional.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraNacional.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.fraNacional.Location = New System.Drawing.Point(123, 13)
        Me.fraNacional.Margin = New System.Windows.Forms.Padding(2)
        Me.fraNacional.Name = "fraNacional"
        Me.fraNacional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraNacional.Size = New System.Drawing.Size(95, 44)
        Me.fraNacional.TabIndex = 49
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame6.Location = New System.Drawing.Point(66, 0)
        Me.Frame6.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(2, 38)
        Me.Frame6.TabIndex = 10
        Me.Frame6.TabStop = False
        '
        '_lblProvAcreed_2
        '
        Me._lblProvAcreed_2.AutoSize = True
        Me._lblProvAcreed_2.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_2.Location = New System.Drawing.Point(18, 34)
        Me._lblProvAcreed_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_2.Name = "_lblProvAcreed_2"
        Me._lblProvAcreed_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_2.Size = New System.Drawing.Size(47, 13)
        Me._lblProvAcreed_2.TabIndex = 2
        Me._lblProvAcreed_2.Text = "Nombre:"
        '
        '_lblProvAcreed_0
        '
        Me._lblProvAcreed_0.AutoSize = True
        Me._lblProvAcreed_0.BackColor = System.Drawing.Color.Silver
        Me._lblProvAcreed_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblProvAcreed_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblProvAcreed_0.Location = New System.Drawing.Point(18, 14)
        Me._lblProvAcreed_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblProvAcreed_0.Name = "_lblProvAcreed_0"
        Me._lblProvAcreed_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblProvAcreed_0.Size = New System.Drawing.Size(43, 13)
        Me._lblProvAcreed_0.TabIndex = 0
        Me._lblProvAcreed_0.Text = "Código:"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Frame8)
        Me.Panel1.Controls.Add(Me.Frame10)
        Me.Panel1.Controls.Add(Me.chkMostrarTodos)
        Me.Panel1.Controls.Add(Me._lblProvAcreed_0)
        Me.Panel1.Controls.Add(Me.txtCodProvAcreed)
        Me.Panel1.Controls.Add(Me._lblProvAcreed_2)
        Me.Panel1.Controls.Add(Me.txtNombre)
        Me.Panel1.Controls.Add(Me.Frame1)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(481, 622)
        Me.Panel1.TabIndex = 54
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(14, 537)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(454, 74)
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
        'frmCorpoAbcProvAcreed
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(505, 647)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(215, 83)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoAbcProvAcreed"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Proveedores y Acreedores"
        Me.Frame8.ResumeLayout(False)
        Me.Frame8.PerformLayout()
        Me.Frame10.ResumeLayout(False)
        Me.Frame10.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.fraServicio.ResumeLayout(False)
        Me.fraTipo.ResumeLayout(False)
        Me.fraNacional.ResumeLayout(False)
        CType(Me.lblProvAcreed, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
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