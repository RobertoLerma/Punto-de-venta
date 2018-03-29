Option Explicit On
Option Strict Off
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmABCFunciones
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtDescFuncion As System.Windows.Forms.TextBox
    Public WithEvents txtCodFuncion As System.Windows.Forms.TextBox
    Public WithEvents txtForma As System.Windows.Forms.TextBox
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents _Label2_2 As System.Windows.Forms.Label
    Public WithEvents fraFuncion As System.Windows.Forms.GroupBox
    Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    ' Programa :                ABC de Funciones
    ' Autor :                   Paimí
    ' Fecha de Inicio:          22 de Mayo de 2003
    ' Fecha de Finalización:


    Public mintCodModulo As Integer

    Dim mblnSalir As Boolean 'Controla la salida con ESCAPE

    Dim mblnNuevo As Boolean
    Friend WithEvents Panel1 As Panel
    Public WithEvents btnSalir As Button
    Public WithEvents btnNuevo As Button
    Dim mblnCambiosEnCodigo As Boolean


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtDescFuncion = New System.Windows.Forms.TextBox()
        Me.txtCodFuncion = New System.Windows.Forms.TextBox()
        Me.txtForma = New System.Windows.Forms.TextBox()
        Me.fraFuncion = New System.Windows.Forms.GroupBox()
        Me._Label2_1 = New System.Windows.Forms.Label()
        Me._Label2_0 = New System.Windows.Forms.Label()
        Me._Label2_2 = New System.Windows.Forms.Label()
        Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.fraFuncion.SuspendLayout()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtDescFuncion
        '
        Me.txtDescFuncion.AcceptsReturn = True
        Me.txtDescFuncion.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescFuncion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescFuncion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescFuncion.Location = New System.Drawing.Point(96, 56)
        Me.txtDescFuncion.MaxLength = 50
        Me.txtDescFuncion.Name = "txtDescFuncion"
        Me.txtDescFuncion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescFuncion.Size = New System.Drawing.Size(281, 20)
        Me.txtDescFuncion.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtDescFuncion, "Descripción de la Función")
        '
        'txtCodFuncion
        '
        Me.txtCodFuncion.AcceptsReturn = True
        Me.txtCodFuncion.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodFuncion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodFuncion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodFuncion.Location = New System.Drawing.Point(96, 24)
        Me.txtCodFuncion.MaxLength = 3
        Me.txtCodFuncion.Name = "txtCodFuncion"
        Me.txtCodFuncion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodFuncion.Size = New System.Drawing.Size(41, 20)
        Me.txtCodFuncion.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtCodFuncion, "Código de Función < ENTER = Nuevo >")
        '
        'txtForma
        '
        Me.txtForma.AcceptsReturn = True
        Me.txtForma.BackColor = System.Drawing.SystemColors.Window
        Me.txtForma.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtForma.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtForma.Location = New System.Drawing.Point(96, 88)
        Me.txtForma.MaxLength = 50
        Me.txtForma.Name = "txtForma"
        Me.txtForma.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtForma.Size = New System.Drawing.Size(281, 20)
        Me.txtForma.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtForma, "Nombre del Formulario")
        '
        'fraFuncion
        '
        Me.fraFuncion.BackColor = System.Drawing.Color.Silver
        Me.fraFuncion.Controls.Add(Me.txtDescFuncion)
        Me.fraFuncion.Controls.Add(Me.txtCodFuncion)
        Me.fraFuncion.Controls.Add(Me.txtForma)
        Me.fraFuncion.Controls.Add(Me._Label2_1)
        Me.fraFuncion.Controls.Add(Me._Label2_0)
        Me.fraFuncion.Controls.Add(Me._Label2_2)
        Me.fraFuncion.ForeColor = System.Drawing.Color.Black
        Me.fraFuncion.Location = New System.Drawing.Point(11, 10)
        Me.fraFuncion.Name = "fraFuncion"
        Me.fraFuncion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraFuncion.Size = New System.Drawing.Size(400, 121)
        Me.fraFuncion.TabIndex = 0
        Me.fraFuncion.TabStop = False
        Me.fraFuncion.Text = "Información General"
        '
        '_Label2_1
        '
        Me._Label2_1.AutoSize = True
        Me._Label2_1.BackColor = System.Drawing.Color.Silver
        Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_1.ForeColor = System.Drawing.Color.Black
        Me._Label2_1.Location = New System.Drawing.Point(21, 60)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_1.Size = New System.Drawing.Size(63, 13)
        Me._Label2_1.TabIndex = 3
        Me._Label2_1.Text = "Descripción"
        '
        '_Label2_0
        '
        Me._Label2_0.AutoSize = True
        Me._Label2_0.BackColor = System.Drawing.Color.Silver
        Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_0.ForeColor = System.Drawing.Color.Black
        Me._Label2_0.Location = New System.Drawing.Point(21, 28)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_0.Size = New System.Drawing.Size(40, 13)
        Me._Label2_0.TabIndex = 1
        Me._Label2_0.Text = "Código"
        '
        '_Label2_2
        '
        Me._Label2_2.AutoSize = True
        Me._Label2_2.BackColor = System.Drawing.Color.Silver
        Me._Label2_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_2.ForeColor = System.Drawing.Color.Black
        Me._Label2_2.Location = New System.Drawing.Point(21, 92)
        Me._Label2_2.Name = "_Label2_2"
        Me._Label2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_2.Size = New System.Drawing.Size(44, 13)
        Me._Label2_2.TabIndex = 5
        Me._Label2_2.Text = "Nombre"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.btnSalir)
        Me.Panel1.Controls.Add(Me.btnNuevo)
        Me.Panel1.Controls.Add(Me.fraFuncion)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(421, 191)
        Me.Panel1.TabIndex = 1
        '
        'btnSalir
        '
        Me.btnSalir.BackColor = System.Drawing.SystemColors.Control
        Me.btnSalir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSalir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSalir.Location = New System.Drawing.Point(222, 141)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSalir.Size = New System.Drawing.Size(109, 36)
        Me.btnSalir.TabIndex = 72
        Me.btnSalir.Text = "&Salir"
        Me.btnSalir.UseVisualStyleBackColor = False
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(107, 141)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 71
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'frmABCFunciones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(442, 215)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(314, 166)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmABCFunciones"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Funciones"
        Me.fraFuncion.ResumeLayout(False)
        Me.fraFuncion.PerformLayout()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
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

        If (txtCodFuncion.Text = "") Then
            strControlActual = UCase(txtCodFuncion.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
            strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
        ElseIf (txtDescFuncion.Text = "") Then
            strControlActual = UCase(txtDescFuncion.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
            strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
        End If

        strCaptionForm = "Consulta de Funciones"
        Select Case strControlActual
            Case "TXTCODFUNCION"
                gStrSql = "SELECT RIGHT('000'+LTRIM(CodFuncion),3) AS CODIGO, DescFuncion AS DESCRIPCION FROM CatFunciones WHERE CodModulo = " & mintCodModulo & " ORDER BY CodFuncion"
            Case "TXTDESCFUNCION"
                gStrSql = "SELECT DescFuncion AS DESCRIPCION, RIGHT('000'+LTRIM(CodFuncion),3) AS CODIGO FROM CatFunciones WHERE CodModulo = " & mintCodModulo & " ORDER BY descFuncion"
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

        'Carga el formulario de consulta
        'Load(FrmConsultas)
        Call ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODFUNCION"
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                Case "TXTDESCFUNCION"
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
        If CDbl(ModEstandar.Numerico((Me.txtCodFuncion.Text))) = 0 Then
            Nuevo()
            ModEstandar.AvanzarTab(Me)
            Exit Sub
        End If

        'Me.txtCodFuncion.Text = Format(Me.txtCodFuncion.Text, "000")

        For I = 0 To 2 - txtCodFuncion.TextLength
            txtCodFuncion.Text = String.Concat("0" + txtCodFuncion.Text)
        Next I

        gStrSql = "select * from CatFunciones where codModulo =" & mintCodModulo & " and codFuncion = " & ModEstandar.Numerico((Me.txtCodFuncion.Text))
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount <> 0 Then
            Me.txtDescFuncion.Text = Trim(RsGral.Fields("DescFuncion").Value)
            Me.txtDescFuncion.Tag = Me.txtDescFuncion.Text
            Me.txtForma.Text = Trim(RsGral.Fields("Forma").Value)
            Me.txtForma.Tag = Me.txtForma.Text
        Else
            MsjNoExiste("La Función solicitada en el Módulo", gstrNombCortoEmpresa)
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
        gStrSql = "select codFuncion from CatFunciones where codModulo = " & mintCodModulo & " and codFuncion = " & ModEstandar.Numerico((Me.txtCodFuncion.Text))
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            MsgBox("Proporcione un código válido para eliminar la Función", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
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
        ModStoredProcedures.PR_IMECatFunciones(Str(mintCodModulo), Trim(Me.txtCodFuncion.Text), Trim(Me.txtDescFuncion.Text), Trim(Me.txtForma.Text), C_ELIMINACION, CStr(0))
        Cmd.Execute()
        Cnn.CommitTrans()
        blnTransaction = False
        frmABCModulos.Encabezado()
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
        Me.txtDescFuncion.Text = ""
        Me.txtDescFuncion.Tag = Me.txtDescFuncion.Text
        Me.txtForma.Text = ""
        Me.txtForma.Tag = Me.txtForma.Text
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function Cambios() As Boolean
        Select Case True
            Case Trim(Me.txtDescFuncion.Text) <> Trim(Me.txtDescFuncion.Tag)
                Cambios = True
            Case Trim(Me.txtForma.Text) <> Trim(Me.txtForma.Tag)
                Cambios = True
            Case Else
                Cambios = False
        End Select
    End Function

    Public Function ValidaDatos() As Boolean
        Select Case True
            Case mintCodModulo = 0
                MsgBox(C_msgFALTADATO & "un Módulo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.txtCodFuncion.Focus()
            Case Len(Trim(Me.txtDescFuncion.Text)) = 0
                MsgBox(C_msgFALTADATO & "Descripción de Grupo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Me.txtDescFuncion.Focus()
                ValidaDatos = False
            Case Len(Trim(Me.txtForma.Text)) = 0
                MsgBox(C_msgFALTADATO & "Nombre de la Forma", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Me.txtForma.Focus()
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
            ModStoredProcedures.PR_IMECatFunciones(Str(mintCodModulo), Trim(Me.txtCodFuncion.Text), Trim(Me.txtDescFuncion.Text), Trim(Me.txtForma.Text), C_INSERCION, CStr(0))
            Cmd.Execute()
            Me.txtCodFuncion.Text = Format(Cmd.Parameters("ID").Value, "000")
        Else
            ModStoredProcedures.PR_IMECatFunciones(Str(mintCodModulo), Trim(Me.txtCodFuncion.Text), Trim(Me.txtDescFuncion.Text), Trim(Me.txtForma.Text), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        Cnn.CommitTrans()
        blnTransaction = False
        If mblnNuevo Then
            MsgBox("La Función ha sido grabada correctamente con el código " & Me.txtCodFuncion.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        End If
        frmABCModulos.Encabezado()
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
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                    If Not Guardar() Then
                        Exit Sub
                    End If
                Case MsgBoxResult.Cancel 'Cancela la acción de limpiar pantalla
                    Exit Sub
            End Select
        End If
        Me.txtCodFuncion.Text = ""
        Nuevo()
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        Me.txtCodFuncion.Focus()
    End Sub

    Private Sub frmABCFunciones_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmABCFunciones_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmABCFunciones_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If Trim(UCase(Me.ActiveControl.Name)) = "TXTCODFUNCION" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmABCFunciones_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayúscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmABCFunciones_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        mblnNuevo = True
        mblnCambiosEnCodigo = False
    End Sub

    Private Sub frmABCFunciones_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmABCFunciones_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
    End Sub

    Private Sub txtCodFuncion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodFuncion.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodFuncion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodFuncion.Enter
        SelTextoTxt((Me.txtCodFuncion))
        Pon_Tool()
    End Sub

    Private Sub txtCodFuncion_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodFuncion.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
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
                    Me.txtCodFuncion.Focus()
            End Select
        End If
    End Sub

    Private Sub txtCodFuncion_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodFuncion.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
                        Me.txtCodFuncion.Focus()
                End Select
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodFuncion_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodFuncion.Leave
        'If ActiveControl.Text = Me.Text Then
        'If mblnCambiosEnCodigo = True Then 'Si hubo cambios en el código hace la consulta
        If (txtCodFuncion.Text <> "") Then
            LlenaDatos()
        End If
        'End If
        'End If
    End Sub

    Private Sub txtDescFuncion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescFuncion.Enter
        SelTextoTxt((Me.txtDescFuncion))
        Pon_Tool()
    End Sub

    Private Sub txtForma_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtForma.Enter
        SelTextoTxt((Me.txtForma))
        Pon_Tool()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub
End Class