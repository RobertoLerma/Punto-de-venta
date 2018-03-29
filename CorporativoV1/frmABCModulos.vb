Option Explicit On
Option Strict Off
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmABCModulos

    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents btnABCFunciones As System.Windows.Forms.Button
    Public WithEvents txtDescModulo As System.Windows.Forms.TextBox
    Public WithEvents txtCodModulo As System.Windows.Forms.TextBox
    Public WithEvents mshFlex As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents _lblModulo_1 As System.Windows.Forms.Label
    Public WithEvents _lblModulo_0 As System.Windows.Forms.Label
    Public WithEvents _fraModulo_1 As System.Windows.Forms.GroupBox
    Public WithEvents fraModulo As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblModulo As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    ' Programa :                ABC de Módulos del Sistema
    ' Autor :                   Paimí
    ' Fecha de Inicio:          23 de Mayo de 2003
    ' Fecha de Finalización:


    Const C_RENENCABEZADO As Integer = 0
    Const C_COLDESCRIPCION As Integer = 0
    Const C_COLFORMULARIO As Integer = 1
    Const C_ColCODIGO As Integer = 2

    Dim rsLocal As ADODB.Recordset

    Dim mblnSalir As Boolean 'Controla la salida con ESCAPE

    Dim mblnNuevo As Boolean
    Public WithEvents btnGuardar As Button
    Public WithEvents btnNuevo As Button
    Public WithEvents btnBuscar As Button
    Dim mblnCambiosenCodigo As Boolean
    Public strControlActual As String 'Nombre del control actual

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmABCModulos))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtDescModulo = New System.Windows.Forms.TextBox()
        Me.txtCodModulo = New System.Windows.Forms.TextBox()
        Me._fraModulo_1 = New System.Windows.Forms.GroupBox()
        Me.btnABCFunciones = New System.Windows.Forms.Button()
        Me.mshFlex = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._lblModulo_1 = New System.Windows.Forms.Label()
        Me._lblModulo_0 = New System.Windows.Forms.Label()
        Me.fraModulo = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblModulo = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me._fraModulo_1.SuspendLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraModulo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblModulo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtDescModulo
        '
        Me.txtDescModulo.AcceptsReturn = True
        Me.txtDescModulo.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescModulo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescModulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescModulo.Location = New System.Drawing.Point(72, 60)
        Me.txtDescModulo.MaxLength = 0
        Me.txtDescModulo.Name = "txtDescModulo"
        Me.txtDescModulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescModulo.Size = New System.Drawing.Size(321, 20)
        Me.txtDescModulo.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtDescModulo, "Descripción del Módulo")
        '
        'txtCodModulo
        '
        Me.txtCodModulo.AcceptsReturn = True
        Me.txtCodModulo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodModulo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodModulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodModulo.Location = New System.Drawing.Point(72, 24)
        Me.txtCodModulo.MaxLength = 0
        Me.txtCodModulo.Name = "txtCodModulo"
        Me.txtCodModulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodModulo.Size = New System.Drawing.Size(49, 20)
        Me.txtCodModulo.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtCodModulo, "Código del Módulo")
        '
        '_fraModulo_1
        '
        Me._fraModulo_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraModulo_1.Controls.Add(Me.btnABCFunciones)
        Me._fraModulo_1.Controls.Add(Me.txtDescModulo)
        Me._fraModulo_1.Controls.Add(Me.txtCodModulo)
        Me._fraModulo_1.Controls.Add(Me.mshFlex)
        Me._fraModulo_1.Controls.Add(Me._lblModulo_1)
        Me._fraModulo_1.Controls.Add(Me._lblModulo_0)
        Me._fraModulo_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraModulo.SetIndex(Me._fraModulo_1, CType(1, Short))
        Me._fraModulo_1.Location = New System.Drawing.Point(8, 8)
        Me._fraModulo_1.Name = "_fraModulo_1"
        Me._fraModulo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraModulo_1.Size = New System.Drawing.Size(537, 320)
        Me._fraModulo_1.TabIndex = 0
        Me._fraModulo_1.TabStop = False
        '
        'btnABCFunciones
        '
        Me.btnABCFunciones.BackColor = System.Drawing.SystemColors.Control
        Me.btnABCFunciones.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnABCFunciones.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnABCFunciones.Location = New System.Drawing.Point(395, 272)
        Me.btnABCFunciones.Name = "btnABCFunciones"
        Me.btnABCFunciones.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnABCFunciones.Size = New System.Drawing.Size(121, 38)
        Me.btnABCFunciones.TabIndex = 6
        Me.btnABCFunciones.Text = "ABC a F&unciones"
        Me.btnABCFunciones.UseVisualStyleBackColor = False
        '
        'mshFlex
        '
        Me.mshFlex.DataSource = Nothing
        Me.mshFlex.Location = New System.Drawing.Point(16, 96)
        Me.mshFlex.Name = "mshFlex"
        Me.mshFlex.OcxState = CType(resources.GetObject("mshFlex.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mshFlex.Size = New System.Drawing.Size(500, 161)
        Me.mshFlex.TabIndex = 5
        '
        '_lblModulo_1
        '
        Me._lblModulo_1.AutoSize = True
        Me._lblModulo_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblModulo_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblModulo_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblModulo.SetIndex(Me._lblModulo_1, CType(1, Short))
        Me._lblModulo_1.Location = New System.Drawing.Point(24, 64)
        Me._lblModulo_1.Name = "_lblModulo_1"
        Me._lblModulo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblModulo_1.Size = New System.Drawing.Size(42, 13)
        Me._lblModulo_1.TabIndex = 3
        Me._lblModulo_1.Text = "Módulo"
        '
        '_lblModulo_0
        '
        Me._lblModulo_0.AutoSize = True
        Me._lblModulo_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblModulo_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblModulo_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblModulo.SetIndex(Me._lblModulo_0, CType(0, Short))
        Me._lblModulo_0.Location = New System.Drawing.Point(24, 28)
        Me._lblModulo_0.Name = "_lblModulo_0"
        Me._lblModulo_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblModulo_0.Size = New System.Drawing.Size(40, 13)
        Me._lblModulo_0.TabIndex = 1
        Me._lblModulo_0.Text = "Código"
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(8, 342)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(108, 39)
        Me.btnGuardar.TabIndex = 9
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(236, 342)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(108, 39)
        Me.btnNuevo.TabIndex = 10
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.BackColor = System.Drawing.SystemColors.Control
        Me.btnBuscar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnBuscar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBuscar.Location = New System.Drawing.Point(122, 342)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnBuscar.Size = New System.Drawing.Size(108, 39)
        Me.btnBuscar.TabIndex = 11
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmABCModulos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(553, 393)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me._fraModulo_1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmABCModulos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC de Módulos y Funciones"
        Me._fraModulo_1.ResumeLayout(False)
        Me._fraModulo_1.PerformLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraModulo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblModulo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub



    Sub Buscar()
        On Error GoTo MErr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendrá el string del tag que se le mandará al fromulario de consultas
        Dim strCaptionForm As String 'Titulo que mostrará el formulario de consultas


        'strControlActual = UCase(btnBuscar.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag será el nombre del formulario + el nombre del control
        strCaptionForm = "Consulta de Módulos del Sistema"

        Select Case strControlActual
            Case "TXTCODMODULO"
                gStrSql = "SELECT RIGHT('00'+LTRIM(CodModulo),2) AS CODIGO, DescModulo AS DESCRIPCION FROM CatModulos ORDER BY CodModulo"
            Case "TXTDESCMODULO"
                gStrSql = "SELECT DescModulo AS DESCRIPCION, RIGHT('00'+LTRIM(CodModulo),2) AS CODIGO FROM CatModulos WHERE DescModulo LIKE '" & Trim(Me.txtDescModulo.Text) & "%' ORDER BY DescModulo"
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
        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 5700, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODMODULO"
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4800) 'Columna de la Descripción
                Case "TXTDESCMODULO"
                    .set_ColWidth(0, 0, 4800) 'Columna de la Descripción
                    .set_ColWidth(1, 0, 900) 'Columna del Código
            End Select
        End With
        FrmConsultas.ShowDialog()
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub LimpiarFlex()
        On Error Resume Next
        Dim i As Object
        'Me.dbcGrupo.text = ""
        'Me.dbcGrupo.Tag = ""
        'Pone el enfoque en la última línea disponible para dar de alta una descripción más
        With mshFlex
            .Clear()
            .set_TextMatrix(C_RENENCABEZADO, C_COLDESCRIPCION, "Función")
            .set_TextMatrix(C_RENENCABEZADO, C_COLFORMULARIO, "Formulario")
            .set_TextMatrix(C_RENENCABEZADO, C_ColCODIGO, "Código")
            'Colocar los textos de los encabezados centrados
            .Row = C_RENENCABEZADO
            For i = 0 To (.get_Cols() - 1) Step 1
                .Col = i
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next i
        End With
    End Sub

    Public Sub ScrollGrid()
        'Procedimiento que pone el enfoque en el primer renglón vacío del Grid
        Dim i As Integer
        Dim nCont As Integer 'Cuenta los renglones que están ocupados (que no están vacíos)
        Dim nRen As Integer
        'Aparecen 7 renglones disponibles en el Grid
        'Si son menos de siete registros ocupados, no se utiliza el .TopRow
        'Pero, si son 7 ó más registros, el .TopRow manda el enfoque al primer renglón vacío
        'después de los renglones ocupados
        nRen = 7 'El máximo de renglones que aparece en el grid (Además del encabezado)
        nCont = 0
        With Me.mshFlex
            For i = 1 To .Rows
                If Trim(.get_TextMatrix(i, C_COLDESCRIPCION)) <> "" Then
                    nCont = nCont + 1
                Else
                    Exit For
                End If
            Next i
            If nCont < 7 Then
                'Hay menos de 7 registros
                .Row = nCont + 1
                .Col = C_COLDESCRIPCION
            Else
                'Hay 7 ó más registros, hay que recorrer el grid
                .TopRow = (nCont - nRen) + 2
                .Row = nCont + 1
                .Col = C_COLDESCRIPCION
            End If
        End With
    End Sub

    Function BuscarFlex() As Boolean
        On Error GoTo MErr
        gStrSql = "select codFuncion from CatFunciones where codModulo = " & ModEstandar.Numerico((Me.txtCodModulo.Text)) & " and codFuncion = " & ModEstandar.Numerico(Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_ColCODIGO))
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount = 0 Then
            BuscarFlex = False
        Else
            BuscarFlex = True
        End If
MErr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub EliminarFlex()
        On Error GoTo MErr
        Dim blnTransaction As Boolean
        Dim TopRowAnterior As Object
        Dim RowAnterior As Integer
        TopRowAnterior = Me.mshFlex.TopRow
        RowAnterior = Me.mshFlex.Row

        If Me.mshFlex.get_TextMatrix(mshFlex.Row, C_ColCODIGO) <> "" And BuscarFlex() Then
            gStrSql = "select codFuncion from CatFunciones where codModulo = " & ModEstandar.Numerico((Me.txtCodModulo.Text)) & " and codFuncion = " & ModEstandar.Numerico(Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_ColCODIGO))
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_SELECT_DATOS"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            rsLocal = Cmd.Execute

            If rsLocal.RecordCount > 0 Then
                'Pregunta por la integridad referencial
                If Referencia("SELECT Forma FROM Accesos WHERE LTRIM(RTRIM(Forma)) = '" & Trim(Me.mshFlex.get_TextMatrix(mshFlex.Row, C_COLFORMULARIO)) & "'") Then
                    MsgBox("No puede eliminar este formulario porque forma parte de la tabla Accesos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Sub
                End If
                If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) = MsgBoxResult.Yes Then
                    Cnn.BeginTrans()
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    blnTransaction = True
                    ModStoredProcedures.PR_IMECatFunciones(Trim(Me.txtCodModulo.Text), Trim(Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_ColCODIGO)), Trim(""), Trim(""), C_ELIMINACION, CStr(0))
                    Cmd.Execute()
                    Cnn.CommitTrans()
                    blnTransaction = False
                Else
                    Exit Sub
                End If
            End If
        Else
            Exit Sub
        End If
        Encabezado()
        Me.mshFlex.TopRow = TopRowAnterior
        Me.mshFlex.Row = RowAnterior
        Me.mshFlex.Col = C_COLDESCRIPCION
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
MErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
        If blnTransaction Then Cnn.RollbackTrans()
    End Sub


    Sub Encabezado()
        Dim LnContador As Integer
        Dim lStrSql As String
        Dim i As Integer

        With mshFlex
            '.FocusRect = flexFocusHeavy 'flexFocusLight 'flexFocusNone
            '.WordWrap = False
            .FixedRows = 1
            .FixedCols = 0
            .set_Cols(0, 3)

            .set_ColWidth(C_COLDESCRIPCION, 0, 3570)
            .set_ColWidth(C_COLFORMULARIO, 0, 3570)
            .set_ColWidth(C_ColCODIGO, 0, 1)

            .set_TextMatrix(0, C_COLDESCRIPCION, "Función")
            .set_TextMatrix(0, C_COLFORMULARIO, "Formulario")
            .set_TextMatrix(0, C_ColCODIGO, "Código")

            'Colocar los textos de los encabezados centrados
            .Row = C_RENENCABEZADO
            For LnContador = 0 To (.get_Cols() - 1) Step 1
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next LnContador


            'Hacer la consulta con la cual llenar el grid
            lStrSql = "SELECT * FROM CatFunciones WHERE codModulo = " & ModEstandar.Numerico((Me.txtCodModulo.Text))
            ModEstandar.BorraCmd()
            Cmd.CommandText = "Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, lStrSql))
            rsLocal = Cmd.Execute

            LimpiarFlex()
            'Obtiene el último registro o renglón
            If rsLocal.RecordCount > 0 Then
                If rsLocal.RecordCount + 2 < 11 Then
                    .Rows = 11
                Else
                    .Rows = rsLocal.RecordCount + 2
                End If
                rsLocal.MoveFirst()
                For i = 1 To rsLocal.RecordCount
                    .set_TextMatrix(i, C_COLDESCRIPCION, Trim(rsLocal.Fields("DescFuncion").Value))
                    .set_TextMatrix(i, C_COLFORMULARIO, Trim(rsLocal.Fields("Forma").Value))
                    .set_TextMatrix(i, C_ColCODIGO, Trim(rsLocal.Fields("CodFuncion").Value))
                    rsLocal.MoveNext()
                Next i
                .TopRow = 1
                .Row = 1
                .Col = C_COLDESCRIPCION
            Else
                .Rows = 11
                .Row = 1
                .Col = C_COLDESCRIPCION
            End If
        End With
    End Sub

    Public Sub LlenaDatos()
        On Error GoTo MErr
        If CDbl(ModEstandar.Numerico((Me.txtCodModulo.Text))) = 0 Then
            Nuevo()
            ModEstandar.AvanzarTab(Me)
            Exit Sub
        End If
        'Me.txtCodModulo.Text = Format(Me.txtCodModulo.Text, "00")

        For i = 1 To 2 - txtCodModulo.TextLength
            txtCodModulo.Text = String.Concat("0", txtCodModulo.Text)
        Next i

        gStrSql = "select * from CatModulos where codModulo =" & ModEstandar.Numerico((Me.txtCodModulo.Text))
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount <> 0 Then
            Me.txtDescModulo.Text = Trim(RsGral.Fields("DescModulo").Value)
            Me.txtDescModulo.Tag = Me.txtDescModulo.Text
            Encabezado()
        Else
            MsjNoExiste("el Módulo solicitado", gstrNombCortoEmpresa)
            Limpiar()
        End If
        mblnCambiosenCodigo = False
        mblnNuevo = False
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub Eliminar()
        On Error GoTo MErr
        Dim blnTransaction As Boolean
        gStrSql = "select codModulo from CatModulos where codModulo = " & Trim(Me.txtCodModulo.Text)
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            MsgBox("Proporcione un código válido para eliminar.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If

        'Preguntar por la Integridad Referencial
        If Referencia("select codModulo from CatFunciones where codModulo = " & Numerico((Me.txtCodModulo.Text))) Then
            MsgBox("No puede eliminar este módulo porque existen Funciones que dependen de él", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If
        If Referencia("select codModulo from Accesos where codModulo = " & Numerico((Me.txtCodModulo.Text))) Then
            MsgBox("No puede eliminar este módulo porque existen dependencias en la tabla de Accesos", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Exit Sub
        End If
        'Preguntar si desea borrar el registro
        If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
            Exit Sub
        End If
        Cnn.BeginTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        blnTransaction = True
        ModStoredProcedures.PR_IMECatModulos(Str(CDbl(Numerico((Me.txtCodModulo.Text)))), Trim(Me.txtDescModulo.Text), C_ELIMINACION, CStr(0))
        Cmd.Execute()
        Cnn.CommitTrans()
        blnTransaction = False
        Limpiar()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
MErr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Sub Nuevo()
        On Error GoTo MErr
        Me.txtDescModulo.Text = ""
        Me.txtDescModulo.Tag = Me.txtDescModulo.Text
        LimpiarFlex()
MErr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function Cambios() As Boolean
        Select Case True
            Case Trim(Me.txtDescModulo.Text) <> Trim(Me.txtDescModulo.Tag)
                Cambios = True
            Case Else
                Cambios = False
        End Select
    End Function

    Public Function ValidaDatos() As Boolean
        Select Case True
            Case Len(Me.txtDescModulo.Text) = 0
                MsgBox(C_msgFALTADATO & "Descripción de Módulo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                Me.txtDescModulo.Focus()
                ValidaDatos = False
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Public Function Guardar() As Boolean
        On Error GoTo MErr
        Dim blnTransaction As Boolean
        'Valida si todos los datos han sido llenados correctamnte para poder ser guardados
        If Not ValidaDatos() Then
            mblnNuevo = True
            Exit Function
        End If
        If Not Cambios() Then
            Limpiar()
            Exit Function
        End If
        Cnn.BeginTrans()
        blnTransaction = True
        If mblnNuevo Then
            ModStoredProcedures.PR_IMECatModulos(Trim(Me.txtCodModulo.Text), Trim(Me.txtDescModulo.Text), C_INSERCION, CStr(0))
            Cmd.Execute()
            Me.txtCodModulo.Text = Format(Cmd.Parameters("ID").Value, "00")
        Else
            ModStoredProcedures.PR_IMECatModulos(Trim(Me.txtCodModulo.Text), Trim(Me.txtDescModulo.Text), C_MODIFICACION, CStr(0))
            Cmd.Execute()
        End If
        Cnn.CommitTrans()
        blnTransaction = False
        If mblnNuevo Then
            MsgBox("El Módulo ha sido grabado correctamente con el código " & Me.txtCodModulo.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        End If
        Guardar = True
        Nuevo()
        Limpiar()
MErr:
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
        Me.txtCodModulo.Text = ""
        Nuevo()
        mblnNuevo = True
        mblnCambiosenCodigo = False
        Me.txtCodModulo.Focus()
    End Sub

    Private Sub btnABCFunciones_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnABCFunciones.Click
        Me.Enabled = False
        frmABCFunciones.Text = "Módulo : " & Trim(Me.txtDescModulo.Text) & " [ " & Trim(Me.txtCodModulo.Text) & " ]"
        frmABCFunciones.mintCodModulo = CInt(ModEstandar.Numerico((Me.txtCodModulo.Text)))
        frmABCFunciones.Show()
        Me.Enabled = True
    End Sub

    Private Sub frmABCModulos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmABCModulos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmABCModulos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Delete
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    If Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_COLDESCRIPCION) <> "" Then
                        Call EliminarFlex()
                    End If
                End If
            Case System.Windows.Forms.Keys.Escape
                If Trim(UCase(Me.ActiveControl.Name)) = "TXTCODMODULO" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmABCModulos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayúscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmABCModulos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        mblnNuevo = True
        mblnCambiosenCodigo = False
        Encabezado()
    End Sub

    Private Sub frmABCModulos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        '            Me.txtCodModulo.Focus()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmABCModulos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
    End Sub

    Private Sub mshFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.Enter
        Pon_Tool()
        Me.mshFlex.Row = 1
        Me.mshFlex.Col = 0
        Me.mshFlex.TopRow = 1
    End Sub

    Private Sub txtCodModulo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodModulo.TextChanged
        If Not mblnNuevo Then
            Nuevo()
            mblnNuevo = True
        End If
        mblnCambiosenCodigo = True
    End Sub

    Private Sub txtCodModulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodModulo.Enter
        strControlActual = UCase("txtCodModulo")
        SelTextoTxt((Me.txtCodModulo))
        Pon_Tool()
    End Sub

    Private Sub txtCodModulo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodModulo.KeyDown
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
                    Me.txtCodModulo.Focus()
            End Select
        End If
    End Sub

    Private Sub txtCodModulo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodModulo.KeyPress
        'Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        'If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back Then
        '    KeyAscii = 0
        'Else
        '    'Pregunta sólo si ha habido cambios
        '    If Cambios() And Not mblnNuevo Then
        '        Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes
        '                If Not Guardar() Then
        '                    KeyAscii = 0
        '                End If
        '            Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
        '            Case MsgBoxResult.Cancel 'Cancela la captura
        '                KeyAscii = 0
        '                Me.txtCodModulo.Focus()
        '        End Select
        '    End If
        'End If
        'eventArgs.KeyChar = Chr(KeyAscii)
        'If KeyAscii = 0 Then
        '    eventArgs.Handled = True
        'End If
    End Sub

    Private Sub txtCodModulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodModulo.Leave
        'If ActiveControl.Text = Me.Text Then
        If mblnCambiosenCodigo = True Then 'Si hubo cambios en el código hace la consulta
            LlenaDatos()
        End If
        'End If
    End Sub

    Private Sub txtDescModulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescModulo.Enter
        strControlActual = UCase("txtDescModulo")
        SelTextoTxt((Me.txtDescModulo))
        Pon_Tool()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub
End Class