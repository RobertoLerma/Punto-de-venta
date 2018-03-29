'**********************************************************************************************************************'
'*PROGRAMA: ABC DE GRUPOS JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmCorpoABCGrupos
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ' Programa :                ABC de Grupos
    ' Autor :                   Paimí
    ' Fecha de Inicio:          8 de Mayo de 2003
    ' Fecha de Finalización:    8 de Mayo de 2003
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtPorcTasa As System.Windows.Forms.TextBox
    Public WithEvents txtImporte As System.Windows.Forms.TextBox
    Public WithEvents txtDescGrupo As System.Windows.Forms.TextBox
    Public WithEvents txtCodGrupo As System.Windows.Forms.TextBox
    Public WithEvents _lblGrupo_3 As System.Windows.Forms.Label
    Public WithEvents _lblGrupo_2 As System.Windows.Forms.Label
    Public WithEvents _lblGrupo_1 As System.Windows.Forms.Label
    Public WithEvents _lblGrupo_0 As System.Windows.Forms.Label
    Public WithEvents fraGrupos As System.Windows.Forms.GroupBox
    Public WithEvents lblGrupo As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Dim mblnSalir As Boolean 'Controla la salida con ESCAPE

    Dim mblnNuevo As Boolean
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btnSalir As Button
    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnGuardar As Button
    Friend WithEvents btnLimpiar As Button
    Friend WithEvents btnEliminar As Button
    Dim mblnCambiosEnCodigo As Boolean


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtPorcTasa = New System.Windows.Forms.TextBox()
        Me.txtImporte = New System.Windows.Forms.TextBox()
        Me.txtDescGrupo = New System.Windows.Forms.TextBox()
        Me.txtCodGrupo = New System.Windows.Forms.TextBox()
        Me.fraGrupos = New System.Windows.Forms.GroupBox()
        Me._lblGrupo_3 = New System.Windows.Forms.Label()
        Me._lblGrupo_2 = New System.Windows.Forms.Label()
        Me._lblGrupo_1 = New System.Windows.Forms.Label()
        Me._lblGrupo_0 = New System.Windows.Forms.Label()
        Me.lblGrupo = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.fraGrupos.SuspendLayout()
        CType(Me.lblGrupo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtPorcTasa
        '
        Me.txtPorcTasa.AcceptsReturn = True
        Me.txtPorcTasa.BackColor = System.Drawing.SystemColors.Window
        Me.txtPorcTasa.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPorcTasa.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPorcTasa.Location = New System.Drawing.Point(208, 120)
        Me.txtPorcTasa.MaxLength = 0
        Me.txtPorcTasa.Name = "txtPorcTasa"
        Me.txtPorcTasa.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPorcTasa.Size = New System.Drawing.Size(105, 20)
        Me.txtPorcTasa.TabIndex = 8
        Me.txtPorcTasa.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtPorcTasa, "Porcentaje de Impuesto Especial a Productos Suntuarios")
        '
        'txtImporte
        '
        Me.txtImporte.AcceptsReturn = True
        Me.txtImporte.BackColor = System.Drawing.SystemColors.Window
        Me.txtImporte.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImporte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImporte.Location = New System.Drawing.Point(208, 88)
        Me.txtImporte.MaxLength = 0
        Me.txtImporte.Name = "txtImporte"
        Me.txtImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImporte.Size = New System.Drawing.Size(105, 20)
        Me.txtImporte.TabIndex = 6
        Me.txtImporte.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtImporte, "Importe a partir del cual se aplicará el IEPS")
        '
        'txtDescGrupo
        '
        Me.txtDescGrupo.AcceptsReturn = True
        Me.txtDescGrupo.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescGrupo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescGrupo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescGrupo.Location = New System.Drawing.Point(88, 56)
        Me.txtDescGrupo.MaxLength = 0
        Me.txtDescGrupo.Name = "txtDescGrupo"
        Me.txtDescGrupo.ReadOnly = True
        Me.txtDescGrupo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescGrupo.Size = New System.Drawing.Size(305, 20)
        Me.txtDescGrupo.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtDescGrupo, "Descripción del Grupo de Artículos")
        '
        'txtCodGrupo
        '
        Me.txtCodGrupo.AcceptsReturn = True
        Me.txtCodGrupo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodGrupo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodGrupo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodGrupo.Location = New System.Drawing.Point(88, 24)
        Me.txtCodGrupo.MaxLength = 0
        Me.txtCodGrupo.Name = "txtCodGrupo"
        Me.txtCodGrupo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodGrupo.Size = New System.Drawing.Size(41, 20)
        Me.txtCodGrupo.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtCodGrupo, "Clave del Grupo de Artículos")
        '
        'fraGrupos
        '
        Me.fraGrupos.BackColor = System.Drawing.Color.Silver
        Me.fraGrupos.Controls.Add(Me.txtPorcTasa)
        Me.fraGrupos.Controls.Add(Me.txtImporte)
        Me.fraGrupos.Controls.Add(Me.txtDescGrupo)
        Me.fraGrupos.Controls.Add(Me.txtCodGrupo)
        Me.fraGrupos.Controls.Add(Me._lblGrupo_3)
        Me.fraGrupos.Controls.Add(Me._lblGrupo_2)
        Me.fraGrupos.Controls.Add(Me._lblGrupo_1)
        Me.fraGrupos.Controls.Add(Me._lblGrupo_0)
        Me.fraGrupos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGrupos.Location = New System.Drawing.Point(10, 14)
        Me.fraGrupos.Name = "fraGrupos"
        Me.fraGrupos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGrupos.Size = New System.Drawing.Size(409, 161)
        Me.fraGrupos.TabIndex = 0
        Me.fraGrupos.TabStop = False
        '
        '_lblGrupo_3
        '
        Me._lblGrupo_3.AutoSize = True
        Me._lblGrupo_3.BackColor = System.Drawing.Color.Silver
        Me._lblGrupo_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblGrupo_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblGrupo_3.Location = New System.Drawing.Point(151, 124)
        Me._lblGrupo_3.Name = "_lblGrupo_3"
        Me._lblGrupo_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblGrupo_3.Size = New System.Drawing.Size(48, 13)
        Me._lblGrupo_3.TabIndex = 7
        Me._lblGrupo_3.Text = "IEPS (%)"
        Me._lblGrupo_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblGrupo_2
        '
        Me._lblGrupo_2.AutoSize = True
        Me._lblGrupo_2.BackColor = System.Drawing.Color.Silver
        Me._lblGrupo_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblGrupo_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblGrupo_2.Location = New System.Drawing.Point(40, 92)
        Me._lblGrupo_2.Name = "_lblGrupo_2"
        Me._lblGrupo_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblGrupo_2.Size = New System.Drawing.Size(159, 13)
        Me._lblGrupo_2.TabIndex = 5
        Me._lblGrupo_2.Text = "Importe para aplicación de IEPS"
        '
        '_lblGrupo_1
        '
        Me._lblGrupo_1.AutoSize = True
        Me._lblGrupo_1.BackColor = System.Drawing.Color.Silver
        Me._lblGrupo_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblGrupo_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblGrupo_1.Location = New System.Drawing.Point(16, 60)
        Me._lblGrupo_1.Name = "_lblGrupo_1"
        Me._lblGrupo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblGrupo_1.Size = New System.Drawing.Size(63, 13)
        Me._lblGrupo_1.TabIndex = 3
        Me._lblGrupo_1.Text = "Descripción"
        '
        '_lblGrupo_0
        '
        Me._lblGrupo_0.AutoSize = True
        Me._lblGrupo_0.BackColor = System.Drawing.Color.Silver
        Me._lblGrupo_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblGrupo_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblGrupo_0.Location = New System.Drawing.Point(16, 24)
        Me._lblGrupo_0.Name = "_lblGrupo_0"
        Me._lblGrupo_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblGrupo_0.Size = New System.Drawing.Size(36, 13)
        Me._lblGrupo_0.TabIndex = 1
        Me._lblGrupo_0.Text = "Grupo"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.fraGrupos)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(429, 268)
        Me.Panel1.TabIndex = 69
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(10, 181)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(409, 74)
        Me.Panel3.TabIndex = 72
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
        'frmCorpoABCGrupos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(453, 292)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoABCGrupos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Grupos de Artículos"
        Me.fraGrupos.ResumeLayout(False)
        Me.fraGrupos.PerformLayout()
        CType(Me.lblGrupo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub


    'Este codigo es de Prueba y no se deberá tomar en cuenta como plantilla
    'hasta que Patricia determine la nueva froma de consultas que se utilizará
    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendrá el string del tag que se le mandara al fromulario de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim strControlActual As String 'Nombre del control actual

        strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        strCaptionForm = "Consulta de Grupos de Artículos"
        Select Case strControlActual
            Case "TXTCODGRUPO"
                gStrSql = "SELECT RIGHT('00'+LTRIM(CodGrupo),2) AS CODIGO, DescGrupo AS DESCRIPCION FROM CatGrupos ORDER BY CodGrupo"
            Case "TXTDESCGRUPO"
                gStrSql = "SELECT DescGrupo AS DESCRIPCION, RIGHT('00'+LTRIM(CodGrupo),2) AS CODIGO FROM CatGrupos WHERE DescGrupo LIKE '" & Trim(txtDescGrupo.Text) & "%' ORDER BY DescGrupo"
            Case Else
                'Sale de este sub para que no ejecute ninguna opción
                Exit Sub
        End Select

        strSQL = gStrSql 'Se hace uso de una variable temporal para el query

        'Si hubo cambios y es una modificacion entonces preguntará si desea grabar los cambios
        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se cargue la consulta
                Case MsgBoxResult.Cancel 'Cancela la consulta
                    Exit Sub
            End Select
        End If

        gStrSql = strSQL 'Se regresa el valor de la variable temporal a la variable original

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsjNoExiste(C_msgSINDATOS, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If

        ''Carga el formulario de consulta
        'Load(FrmConsultas)
        'ModVariables.frmConsultas.Show()
        ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODGRUPO"
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                Case "TXTDESCGRUPO"
                    .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
                    .set_ColWidth(1, 0, 900) 'Columna del Código
            End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub LlenaDatos()
        On Error GoTo Merr
        If CDbl(ModEstandar.Numerico((Me.txtCodGrupo.Text))) = 0 Then
            Nuevo()
            ModEstandar.AvanzarTab(Me)
            Exit Sub
        End If
        Me.txtCodGrupo.Text = Format(Me.txtCodGrupo.Text, "00")
        gStrSql = "select * from CatGrupos where codGrupo =" & ModEstandar.Numerico((Me.txtCodGrupo.Text))
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount <> 0 Then
            Me.txtDescGrupo.Text = Trim(RsGral.Fields("DescGrupo").Value)
            Me.txtDescGrupo.Tag = Me.txtDescGrupo.Text
            Me.txtImporte.Text = Format(RsGral.Fields("Importe").Value, "###,###,##0.00")
            Me.txtImporte.Tag = Me.txtImporte.Text
            Me.txtPorcTasa.Text = Format(RsGral.Fields("PorcTasa").Value, "##0.00")
            Me.txtPorcTasa.Tag = Me.txtPorcTasa.Text
        Else
            MsjNoExiste("El Grupo de Artículos", gstrNombCortoEmpresa)
            Limpiar()
        End If
        mblnCambiosEnCodigo = False
        mblnNuevo = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub Eliminar()
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        gStrSql = "select codGrupo from CatGrupos where codGrupo = '" & Trim(Me.txtCodGrupo.Text) & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            MsgBox("Proporcione un código válido para eliminar el Grupo de Artículos", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If
        'Preguntar si desea borrar el registro
        If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) = MsgBoxResult.No Then
            Exit Sub
        End If
        Cnn.BeginTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        blnTransaction = True
        ModStoredProcedures.PR_IMECatGrupos(Trim(Me.txtCodGrupo.Text), Trim(Me.txtDescGrupo.Text), Str(CDbl(ModEstandar.Numerico((Me.txtImporte.Text)))), Str(CDbl(ModEstandar.Numerico((Me.txtPorcTasa.Text)))), C_ELIMINACION, CStr(0))
        Cmd.Execute()
        Cnn.CommitTrans()
        blnTransaction = False
        Limpiar()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Sub Nuevo()
        On Error GoTo Merr
        Me.txtDescGrupo.Text = ""
        Me.txtDescGrupo.Tag = Me.txtDescGrupo.Text
        Me.txtImporte.Text = "0.00"
        Me.txtImporte.Tag = Me.txtImporte.Text
        Me.txtPorcTasa.Text = "0.00"
        Me.txtPorcTasa.Tag = Me.txtPorcTasa.Text
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function Cambios() As Boolean
        Select Case True
            Case Trim(Me.txtDescGrupo.Text) <> Trim(Me.txtDescGrupo.Tag)
                Cambios = True
            Case ModEstandar.Numerico((Me.txtImporte.Text)) <> ModEstandar.Numerico((Me.txtImporte.Tag))
                Cambios = True
            Case ModEstandar.Numerico((Me.txtPorcTasa.Text)) <> ModEstandar.Numerico((Me.txtPorcTasa.Tag))
                Cambios = True
            Case Else
                Cambios = False
        End Select
    End Function

    Public Function ValidaDatos() As Boolean
        Select Case True
            Case Len(Me.txtDescGrupo.Text) = 0
                MsgBox(C_msgFALTADATO & "Descripción de Grupo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Me.txtDescGrupo.Focus()
                ValidaDatos = False
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Public Function Guardar() As Boolean
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        'Valida si todos los datos han sido llenados correctamnte para poder ser guardados
        If Not Cambios() Then
            Limpiar()
            Exit Function
        End If
        If Not ValidaDatos() Then
            mblnNuevo = True
            Exit Function
        End If
        Cnn.BeginTrans()
        blnTransaction = True
        If mblnNuevo Then
            ModStoredProcedures.PR_IMECatGrupos(Trim(Me.txtCodGrupo.Text), Trim(Me.txtDescGrupo.Text), Str(CDbl(ModEstandar.Numerico((Me.txtImporte.Text)))), Str(CDbl(ModEstandar.Numerico((Me.txtPorcTasa.Text)))), C_INSERCION, CStr(0))
            Cmd.Execute()
            Me.txtCodGrupo.Text = Format(Cmd.Parameters("ID").Value, "000")
        Else
            ModStoredProcedures.PR_IMECatGrupos(Trim(Me.txtCodGrupo.Text), Trim(Me.txtDescGrupo.Text), Str(CDbl(ModEstandar.Numerico((Me.txtImporte.Text)))), Str(CDbl(ModEstandar.Numerico((Me.txtPorcTasa.Text)))), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        Cnn.CommitTrans()
        blnTransaction = False
        If mblnNuevo Then
            MsgBox("El Grupo de Artículos ha sido grabado correctamente con el código " & Me.txtCodGrupo.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        End If
        Nuevo()
        Guardar = True
        Limpiar()
Merr:
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub Limpiar()
        On Error Resume Next
        'Validar si hubo cambios que desee guardar
        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                Case MsgBoxResult.Cancel 'Cancela la acción de limpiar pantalla
                    Exit Sub
            End Select
        End If
        Me.txtCodGrupo.Text = ""
        Nuevo()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        Me.txtCodGrupo.Focus()
    End Sub
    Private Sub frmCorpoABCGrupos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoABCGrupos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCorpoABCGrupos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If Trim(UCase(Me.ActiveControl.Name)) = "TXTCODGRUPO" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCorpoABCGrupos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayúscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoABCGrupos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        mblnNuevo = True
        mblnCambiosEnCodigo = False
    End Sub

    Private Sub frmCorpoABCGrupos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ''Si desea cerrar la forma y esta se encuentra minimizada, esta se restaura
        'If Not mblnSalir Then
        '    ModEstandar.RestaurarForma(Me, False)
        '    If Cambios() And Not (mblnNuevo) Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes
        '                If Not (Guardar()) Then
        '                    Cancel = 1
        '                End If
        '            Case MsgBoxResult.No 'No hace nada y permite que se cierre el formulario
        '                Cancel = 0
        '            Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin Guardar
        '                Cancel = 1
        '        End Select
        '    End If
        'Else 'Se quiere salir con escape
        '    mblnSalir = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Me.txtCodGrupo.Focus()
        '            ModEstandar.SelTxt()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoABCGrupos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
    End Sub

    Private Sub txtCodGrupo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodGrupo.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodGrupo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodGrupo.Enter
        SelTextoTxt((Me.txtCodGrupo))
        Pon_Tool()
    End Sub

    Private Sub txtCodGrupo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodGrupo.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'Pregunta sólo en caso de que existan cambios en la clave (esto es, cuando se teclea una clave diferente a la actual)
        If Cambios() And KeyCode = System.Windows.Forms.Keys.Delete Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el registro
                    If Not Guardar() Then
                        KeyCode = 0
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se borre el contenido del text
                Case MsgBoxResult.Cancel
                    KeyCode = 0
                    Me.txtCodGrupo.Focus()
            End Select
        End If
    End Sub

    Private Sub txtCodGrupo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodGrupo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back Then
            KeyAscii = 0
        Else
            'Pregunta sólo si ha habido cambios
            If Cambios() And Not mblnNuevo Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        If Not Guardar() Then
                            KeyAscii = 0
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        KeyAscii = 0
                        Me.txtCodGrupo.Focus()
                End Select
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodGrupo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodGrupo.Leave
        If ActiveControl.Text = Me.Text Then
            If mblnCambiosEnCodigo = True Then 'Si hubo cambios en el código hace la consulta
                LlenaDatos()
            End If
        End If
    End Sub

    Private Sub txtDescGrupo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescGrupo.Enter
        SelTextoTxt((Me.txtDescGrupo))
        Pon_Tool()
    End Sub

    Private Sub txtImporte_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Enter
        SelTextoTxt((Me.txtImporte))
        Pon_Tool()
    End Sub

    Private Sub txtImporte_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtImporte.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                Me.txtImporte.Text = Format(Numerico((Me.txtImporte.Text)), "###,###,##0.00")
        End Select
    End Sub

    Private Sub txtImporte_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImporte.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        gp_CampoNumerico(KeyAscii, ".")
        KeyAscii = ModEstandar.MskCantidad((txtImporte.Text), KeyAscii, 10, 2, (txtImporte.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtImporte_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtImporte.Leave
        Me.txtImporte.Text = Format(Me.txtImporte.Text, "###,###,##0.00")
    End Sub

    Private Sub txtPorcTasa_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcTasa.Enter
        SelTextoTxt((Me.txtPorcTasa))
        Pon_Tool()
    End Sub

    Private Sub txtPorcTasa_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPorcTasa.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                Me.txtPorcTasa.Text = Format(Numerico((Me.txtPorcTasa.Text)), "##0.00")
        End Select
    End Sub

    Private Sub txtPorcTasa_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPorcTasa.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        '''Para que en la columna de porcentage
        '''no deje capturar caracteres sino solo numeros
        gp_CampoNumerico(KeyAscii, ".")
        KeyAscii = ModEstandar.MskCantidad((txtPorcTasa.Text), KeyAscii, 5, 2, (txtPorcTasa.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPorcTasa_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPorcTasa.Leave
        Me.txtPorcTasa.Text = Format(Me.txtPorcTasa.Text, "##0.00")
    End Sub


    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub
End Class