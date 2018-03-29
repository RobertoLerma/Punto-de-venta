'**********************************************************************************************************************'
'*PROGRAMA: ABC BANCOS JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class FrmAbcBancos

    Inherits System.Windows.Forms.Form
    'Programa: ABC de Bancos
    'Autor: Rosaura Torres López
    'Fecha de Creación: 08/Mayo/2003

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkSucursal As System.Windows.Forms.CheckBox
    Public WithEvents chkBancoInterno As System.Windows.Forms.CheckBox
    Public WithEvents txtCodBanco As System.Windows.Forms.TextBox
    Public WithEvents txtDescripcion As System.Windows.Forms.TextBox
    Public WithEvents _lblBancos_1 As System.Windows.Forms.Label
    Public WithEvents _lblBancos_0 As System.Windows.Forms.Label
    Public WithEvents fraGeneral As System.Windows.Forms.GroupBox
    Public WithEvents lblBancos As Microsoft.VisualBasic.Compatibility.VB6.LabelArray



    'Estas Variables se declaran de manera local, para evitar conflictos al estar usando
    'la misma variable en distintos modulos, que pueden afectar el valor que hayan tomado en un form. distinto al actual
    Dim mblnNuevo As Boolean 'Para Controlar si un registro es Nuevo o se trata de una consulta
    Dim mblnCambiosEnCodigo As Boolean 'Para Controlar si se han efectuado cambios en el código
    Dim mblnSALIR As Boolean 'se usa para cuando un usuario presiona escape en el primer control de formulario
    Dim bytInterno As Byte 'Para saber si un Banco es Interno
    Dim bytSucursal As Byte 'Para Saber si es Sucursal
    Public WithEvents Panel3 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents btnBuscar As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents Panel1 As Panel
    Public mblnBancoPrincipal As Boolean

    Public strControlActual As String 'Nombre del control actual
    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraGeneral = New System.Windows.Forms.GroupBox()
        Me.chkSucursal = New System.Windows.Forms.CheckBox()
        Me.chkBancoInterno = New System.Windows.Forms.CheckBox()
        Me.txtCodBanco = New System.Windows.Forms.TextBox()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me._lblBancos_1 = New System.Windows.Forms.Label()
        Me._lblBancos_0 = New System.Windows.Forms.Label()
        Me.lblBancos = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.fraGeneral.SuspendLayout()
        CType(Me.lblBancos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraGeneral
        '
        Me.fraGeneral.BackColor = System.Drawing.Color.Silver
        Me.fraGeneral.Controls.Add(Me.chkSucursal)
        Me.fraGeneral.Controls.Add(Me.chkBancoInterno)
        Me.fraGeneral.Controls.Add(Me.txtCodBanco)
        Me.fraGeneral.Controls.Add(Me.txtDescripcion)
        Me.fraGeneral.Controls.Add(Me._lblBancos_1)
        Me.fraGeneral.Controls.Add(Me._lblBancos_0)
        Me.fraGeneral.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGeneral.Location = New System.Drawing.Point(15, 15)
        Me.fraGeneral.Margin = New System.Windows.Forms.Padding(2)
        Me.fraGeneral.Name = "fraGeneral"
        Me.fraGeneral.Padding = New System.Windows.Forms.Padding(2)
        Me.fraGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGeneral.Size = New System.Drawing.Size(324, 147)
        Me.fraGeneral.TabIndex = 4
        Me.fraGeneral.TabStop = False
        Me.ToolTip1.SetToolTip(Me.fraGeneral, "Descripción")
        '
        'chkSucursal
        '
        Me.chkSucursal.BackColor = System.Drawing.Color.Silver
        Me.chkSucursal.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSucursal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSucursal.Location = New System.Drawing.Point(12, 106)
        Me.chkSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSucursal.Name = "chkSucursal"
        Me.chkSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSucursal.Size = New System.Drawing.Size(88, 17)
        Me.chkSucursal.TabIndex = 3
        Me.chkSucursal.Text = "Sucursal"
        Me.chkSucursal.UseVisualStyleBackColor = False
        '
        'chkBancoInterno
        '
        Me.chkBancoInterno.BackColor = System.Drawing.Color.Silver
        Me.chkBancoInterno.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBancoInterno.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBancoInterno.Location = New System.Drawing.Point(12, 76)
        Me.chkBancoInterno.Margin = New System.Windows.Forms.Padding(2)
        Me.chkBancoInterno.Name = "chkBancoInterno"
        Me.chkBancoInterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBancoInterno.Size = New System.Drawing.Size(101, 17)
        Me.chkBancoInterno.TabIndex = 2
        Me.chkBancoInterno.Text = "Banco Interno"
        Me.chkBancoInterno.UseVisualStyleBackColor = False
        '
        'txtCodBanco
        '
        Me.txtCodBanco.AcceptsReturn = True
        Me.txtCodBanco.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodBanco.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodBanco.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodBanco.Location = New System.Drawing.Point(60, 23)
        Me.txtCodBanco.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCodBanco.MaxLength = 3
        Me.txtCodBanco.Name = "txtCodBanco"
        Me.txtCodBanco.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodBanco.Size = New System.Drawing.Size(54, 20)
        Me.txtCodBanco.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtCodBanco, "Código del Banco")
        '
        'txtDescripcion
        '
        Me.txtDescripcion.AcceptsReturn = True
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescripcion.Location = New System.Drawing.Point(82, 48)
        Me.txtDescripcion.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDescripcion.MaxLength = 40
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(200, 20)
        Me.txtDescripcion.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtDescripcion, "Descripción del Banco")
        '
        '_lblBancos_1
        '
        Me._lblBancos_1.AutoSize = True
        Me._lblBancos_1.BackColor = System.Drawing.Color.Silver
        Me._lblBancos_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBancos_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblBancos_1.Location = New System.Drawing.Point(14, 48)
        Me._lblBancos_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblBancos_1.Name = "_lblBancos_1"
        Me._lblBancos_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBancos_1.Size = New System.Drawing.Size(66, 13)
        Me._lblBancos_1.TabIndex = 6
        Me._lblBancos_1.Text = "Descripción:"
        '
        '_lblBancos_0
        '
        Me._lblBancos_0.AutoSize = True
        Me._lblBancos_0.BackColor = System.Drawing.Color.Silver
        Me._lblBancos_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBancos_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblBancos_0.Location = New System.Drawing.Point(14, 23)
        Me._lblBancos_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblBancos_0.Name = "_lblBancos_0"
        Me._lblBancos_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBancos_0.Size = New System.Drawing.Size(43, 13)
        Me._lblBancos_0.TabIndex = 5
        Me._lblBancos_0.Text = "Código:"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(15, 167)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(324, 74)
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
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.fraGeneral)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(354, 259)
        Me.Panel1.TabIndex = 72
        '
        'FrmAbcBancos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(379, 285)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(177, 160)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "FrmAbcBancos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Bancos"
        Me.fraGeneral.ResumeLayout(False)
        Me.fraGeneral.PerformLayout()
        CType(Me.lblBancos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub



    Sub InicializaVariables()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        bytInterno = 0
        bytSucursal = 0
        mblnBancoPrincipal = False
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
            'strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
            strTag = UCase("FRMCORPOABCBANCOS" & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

            Select Case strControlActual
                Case "TXTCODBANCO"
                    strCaptionForm = "Consulta de Bancos"
                    gStrSql = "SELECT RIGHT('000'+LTRIM(Codbanco),3) AS CODIGO,Descbanco AS DESCRIPCION FROM Catbancos  ORDER BY CodBanco"
                Case "TXTDESCRIPCION"
                    strCaptionForm = "Consulta de Bancos"
                    gStrSql = "SELECT Descbanco AS DESCRIPCION, RIGHT('000'+LTRIM(Codbanco),3) AS CODIGO FROM Catbancos WHERE Descbanco LIKE '" & Trim(txtDescripcion.Text) & "%' ORDER BY Descripcion"
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

            ''Carga el formulario de consulta
            Dim FrmConsultas As FrmConsultas = New FrmConsultas()
            ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

            With FrmConsultas.Flexdet
                Select Case strControlActual
                    Case "TXTCODBANCO"
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
        'On Error GoTo MErr
        Try
            'Screen.MousePointer = vbHourglass Esto se manejará hasta antes de iniciar la transacción

            gStrSql = "SELECT DescBanco FROM CatBancos WHERE CodBanco=" & Val(txtCodBanco.Text)

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

            ModStoredProcedures.PR_IMECatBancos(Trim(txtCodBanco.Text), Trim(txtDescripcion.Text), IIf(chkBancoInterno.CheckState = 1, "1", "0"), IIf(chkSucursal.CheckState = 1, "1", "0"), C_ELIMINACION, CStr(0))
            Cmd.Execute()
            MsgBox("El Banco ha sido eliminado correctamente con el Código: " & txtCodBanco.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
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
            'Si no se realizaron cambios, entonces no se guardara nada
            'Si el Código del Banco es "", entonces no se validará nada, solamente se saldrá del proc.
            'And Trim(txtCodBanco) = ""
            If Cambios() = False Then
                Limpiar()
                Exit Function
            End If

            'Validar si todos los datos fueron proporcionados para ser guardados
            If ValidaDatos() = False Then
                Exit Function
            End If

            If Val(txtCodBanco.Text) = 0 Then
                mblnNuevo = True
            End If

            'Se inicia la Transacción aquí, porque en este momento se hara la inserción de los datos.
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Cnn.BeginTrans()

            If mblnNuevo = True Then 'Se realizará una insercion
                ModStoredProcedures.PR_IMECatBancos(Trim(txtCodBanco.Text), Trim(txtDescripcion.Text), IIf(chkBancoInterno.CheckState = 1, "1", "0"), IIf(chkSucursal.CheckState = 1, "1", "0"), C_INSERCION, CStr(0))
                Cmd.Execute()
                txtCodBanco.Text = Format(Cmd.Parameters("ID").Value, "000")
            Else ' Se realizará una Modificación
                ModStoredProcedures.PR_IMECatBancos(Trim(txtCodBanco.Text), Trim(txtDescripcion.Text), IIf(chkBancoInterno.CheckState = 1, "1", "0"), IIf(chkSucursal.CheckState = 1, "1", "0"), C_MODIFICACION, CStr(0))
                Cmd.Execute()
            End If
            Cnn.CommitTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            'Por cuestiones de estética el cambio al puntero del mouse se hace antes de iniciar la transacción y al finalizar la misma.

            If mblnNuevo Then
                MsgBox("El Banco ha sido grabado correctamente con el Código: " & txtCodBanco.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
            Else
                MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
            End If
            'Dejar el Procedimiento Nuevo, sirve para que al usar limpiar,. no pregunte si se desea guardar cambios en el codigo
            Nuevo()
            InicializaVariables()
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
            txtCodBanco.Enabled = True
            txtCodBanco.Text = ""
            txtDescripcion.Text = ""
            txtDescripcion.Tag = ""
            chkBancoInterno.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkSucursal.CheckState = System.Windows.Forms.CheckState.Unchecked
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
            'Screen.MousePointer = vbHourglass
            If Val(txtCodBanco.Text) = 0 Then
                Nuevo()
                'ModEstandar.AvanzarTab Me
                Exit Sub
            End If

            'txtCodBanco.Text = Format(txtCodBanco.Text, "000")

            For i = 1 To 3 - (txtCodBanco.TextLength)
                txtCodBanco.Text = String.Concat("0" + txtCodBanco.Text)
            Next i

            gStrSql = "SELECT * FROM  CatBancos WHERE CodBanco= '" & txtCodBanco.Text & "'"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                txtDescripcion.Text = Trim(RsGral.Fields("DescBanco").Value)
                txtDescripcion.Tag = Trim(RsGral.Fields("DescBanco").Value)
                If RsGral.Fields("ControlInterno").Value Then
                    chkBancoInterno.CheckState = System.Windows.Forms.CheckState.Checked
                    bytInterno = 1
                Else
                    chkBancoInterno.CheckState = System.Windows.Forms.CheckState.Unchecked
                    bytInterno = 0
                End If
                If RsGral.Fields("Sucursal").Value Then
                    chkSucursal.CheckState = System.Windows.Forms.CheckState.Checked
                    bytSucursal = 1
                Else
                    chkSucursal.CheckState = System.Windows.Forms.CheckState.Unchecked
                    bytSucursal = 0
                End If
                If RsGral.Fields("ControlInterno").Value And Not RsGral.Fields("Sucursal").Value Then
                    mblnBancoPrincipal = True
                End If
            Else
                MsjNoExiste("El Banco", gstrNombCortoEmpresa)
                Limpiar()
            End If

            'txtCodBanco.Enabled = False
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

            txtCodBanco.Text = ""
            Nuevo()
            InicializaVariables()
            mblnNuevo = True
            mblnCambiosEnCodigo = False
            txtCodBanco.Focus()
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
        On Error GoTo MErr
        '    Screen.MousePointer = vbHourglass
        Cambios = True

        '    If Trim(txtCodBanco) <> Trim(txtCodBanco.Tag) Then Exit Function

        If Trim(txtDescripcion.Text) <> Trim(txtDescripcion.Tag) Then Exit Function
        If chkBancoInterno.CheckState <> bytInterno Then Exit Function
        If chkSucursal.CheckState <> bytSucursal Then Exit Function
        Cambios = False
        '    Screen.MousePointer = vbDefault
        Exit Function
MErr:
        '    Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then ModEstandar.MostrarError()
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
            If chkBancoInterno.CheckState = 1 And chkSucursal.CheckState = 0 Then
                If Not BuscaBancoPrincipal() Then
                    Exit Function
                End If
            End If
            ValidaDatos = True
            '    Screen.MousePointer = vbDefault
            Exit Function
            'MErr:
        Catch ex As Exception
            '    Screen.MousePointer = vbDefault
            If Err.Number <> 0 Then ModEstandar.MostrarError()
        End Try
    End Function

    Function BuscaBancoPrincipal() As Boolean
        gStrSql = "SELECT * FROM CatBancos WHERE ControlInterno = 1 AND Sucursal = 0"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            BuscaBancoPrincipal = True
        ElseIf RsGral.RecordCount = 1 And Not mblnBancoPrincipal Then
            MsgBox("Ya Existe un Banco Principal Interno, No se Puede Registrar Este Banco ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            BuscaBancoPrincipal = False
        ElseIf RsGral.RecordCount > 1 And Not mblnBancoPrincipal Then
            MsgBox("Existe mas de un Banco Principal Interno No se Puede Registrar Este Banco ...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            BuscaBancoPrincipal = False
        ElseIf mblnBancoPrincipal Then
            BuscaBancoPrincipal = True
        End If
    End Function

    Private Sub chkBancoInterno_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkBancoInterno.Enter
        Pon_Tool()
    End Sub

    Private Sub chkSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSucursal.Enter
        Pon_Tool()
    End Sub

    Private Sub frmCorpoAbcBancos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoAbcBancos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmCorpoAbcBancos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        InicializaVariables()
    End Sub

    Private Sub frmCorpoAbcBancos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmCorpoAbcBancos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoAbcBancos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmCorpoAbcBancos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
    End Sub

    Private Sub txtCodBanco_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodBanco.TextChanged
        If mblnNuevo = False Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodBanco_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodBanco.Enter
        strControlActual = UCase("txtCodBanco")
        SelTextoTxt(txtCodBanco)
        Pon_Tool()
    End Sub

    Private Sub txtCodBanco_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodBanco.KeyDown
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
                        txtCodBanco.Focus()
                        KeyCode = 0
                        Exit Sub
                End Select
            End If
        End If
    End Sub

    Private Sub txtCodBanco_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodBanco.KeyPress
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
                        txtCodBanco.Focus()
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

    Private Sub txtCodBanco_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodBanco.Leave
        'If ActiveControl.Text <> Me.Text Then
        '    Exit Sub
        'End If
        If Val(Trim(txtCodBanco.Text)) = 0 Then txtCodBanco.Text = "000"
        If mblnCambiosEnCodigo = True And CDbl(Numerico(txtCodBanco.Text)) <> 0 Then 'si hubo cambios en el codigo hace la consulta para llenar los datos
            LlenaDatos()
        End If
    End Sub

    Private Sub txtDescripcion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.TextChanged
        mblnCambiosEnCodigo = True
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

