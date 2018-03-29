Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCorpoABCDescuentosVendExternos

    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public WithEvents _sstGrupos_TabPage0 As System.Windows.Forms.TabPage
    Public WithEvents _sstGrupos_TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents _sstGrupos_TabPage2 As System.Windows.Forms.TabPage
    Public WithEvents sstGrupos As System.Windows.Forms.TabControl
    Public WithEvents flexJoyeria As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents flexRelojeria As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents flexVarios As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid

    Public WithEvents cmdInicializar As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtFlexJoyeria As System.Windows.Forms.TextBox
    Public WithEvents txtFlexRelojeria As System.Windows.Forms.TextBox
    Public WithEvents txtFlexVarios As System.Windows.Forms.TextBox
    Public WithEvents cmdCargarFamilias As System.Windows.Forms.Button
    Public WithEvents btnNuevo As Button
    Public WithEvents btnGuardar As Button
    Public WithEvents cmdCargarMarcas As System.Windows.Forms.Button

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCorpoABCDescuentosVendExternos))
        Me.sstGrupos = New System.Windows.Forms.TabControl()
        Me._sstGrupos_TabPage0 = New System.Windows.Forms.TabPage()
        Me.txtFlexJoyeria = New System.Windows.Forms.TextBox()
        Me.flexJoyeria = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._sstGrupos_TabPage1 = New System.Windows.Forms.TabPage()
        Me.txtFlexRelojeria = New System.Windows.Forms.TextBox()
        Me.flexRelojeria = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.cmdCargarMarcas = New System.Windows.Forms.Button()
        Me._sstGrupos_TabPage2 = New System.Windows.Forms.TabPage()
        Me.txtFlexVarios = New System.Windows.Forms.TextBox()
        Me.flexVarios = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.cmdCargarFamilias = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdInicializar = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.sstGrupos.SuspendLayout()
        Me._sstGrupos_TabPage0.SuspendLayout()
        CType(Me.flexJoyeria, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._sstGrupos_TabPage1.SuspendLayout()
        CType(Me.flexRelojeria, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._sstGrupos_TabPage2.SuspendLayout()
        CType(Me.flexVarios, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'sstGrupos
        '
        Me.sstGrupos.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.sstGrupos.Controls.Add(Me._sstGrupos_TabPage0)
        Me.sstGrupos.Controls.Add(Me._sstGrupos_TabPage1)
        Me.sstGrupos.Controls.Add(Me._sstGrupos_TabPage2)
        Me.sstGrupos.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.sstGrupos.ItemSize = New System.Drawing.Size(42, 18)
        Me.sstGrupos.Location = New System.Drawing.Point(12, 12)
        Me.sstGrupos.Name = "sstGrupos"
        Me.sstGrupos.SelectedIndex = 0
        Me.sstGrupos.Size = New System.Drawing.Size(530, 255)
        Me.sstGrupos.TabIndex = 9
        '
        '_sstGrupos_TabPage0
        '
        Me._sstGrupos_TabPage0.Controls.Add(Me.txtFlexJoyeria)
        Me._sstGrupos_TabPage0.Controls.Add(Me.flexJoyeria)
        Me._sstGrupos_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._sstGrupos_TabPage0.Name = "_sstGrupos_TabPage0"
        Me._sstGrupos_TabPage0.Size = New System.Drawing.Size(522, 229)
        Me._sstGrupos_TabPage0.TabIndex = 0
        Me._sstGrupos_TabPage0.Text = "Joyería"
        '
        'txtFlexJoyeria
        '
        Me.txtFlexJoyeria.AcceptsReturn = True
        Me.txtFlexJoyeria.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlexJoyeria.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlexJoyeria.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlexJoyeria.Location = New System.Drawing.Point(14, 35)
        Me.txtFlexJoyeria.MaxLength = 0
        Me.txtFlexJoyeria.Name = "txtFlexJoyeria"
        Me.txtFlexJoyeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlexJoyeria.Size = New System.Drawing.Size(64, 22)
        Me.txtFlexJoyeria.TabIndex = 7
        Me.txtFlexJoyeria.Visible = False
        '
        'flexJoyeria
        '
        Me.flexJoyeria.DataSource = Nothing
        Me.flexJoyeria.Location = New System.Drawing.Point(12, 13)
        Me.flexJoyeria.Name = "flexJoyeria"
        Me.flexJoyeria.OcxState = CType(resources.GetObject("flexJoyeria.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexJoyeria.Size = New System.Drawing.Size(491, 199)
        Me.flexJoyeria.TabIndex = 10
        '
        '_sstGrupos_TabPage1
        '
        Me._sstGrupos_TabPage1.Controls.Add(Me.txtFlexRelojeria)
        Me._sstGrupos_TabPage1.Controls.Add(Me.flexRelojeria)
        Me._sstGrupos_TabPage1.Controls.Add(Me.cmdCargarMarcas)
        Me._sstGrupos_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._sstGrupos_TabPage1.Name = "_sstGrupos_TabPage1"
        Me._sstGrupos_TabPage1.Size = New System.Drawing.Size(522, 229)
        Me._sstGrupos_TabPage1.TabIndex = 1
        Me._sstGrupos_TabPage1.Text = "Relojería"
        '
        'txtFlexRelojeria
        '
        Me.txtFlexRelojeria.AcceptsReturn = True
        Me.txtFlexRelojeria.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlexRelojeria.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlexRelojeria.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlexRelojeria.Location = New System.Drawing.Point(14, 30)
        Me.txtFlexRelojeria.MaxLength = 0
        Me.txtFlexRelojeria.Name = "txtFlexRelojeria"
        Me.txtFlexRelojeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlexRelojeria.Size = New System.Drawing.Size(64, 22)
        Me.txtFlexRelojeria.TabIndex = 8
        Me.txtFlexRelojeria.Visible = False
        '
        'flexRelojeria
        '
        Me.flexRelojeria.DataSource = Nothing
        Me.flexRelojeria.Location = New System.Drawing.Point(14, 10)
        Me.flexRelojeria.Name = "flexRelojeria"
        Me.flexRelojeria.OcxState = CType(resources.GetObject("flexRelojeria.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexRelojeria.Size = New System.Drawing.Size(494, 170)
        Me.flexRelojeria.TabIndex = 11
        '
        'cmdCargarMarcas
        '
        Me.cmdCargarMarcas.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCargarMarcas.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCargarMarcas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCargarMarcas.Location = New System.Drawing.Point(16, 190)
        Me.cmdCargarMarcas.Name = "cmdCargarMarcas"
        Me.cmdCargarMarcas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCargarMarcas.Size = New System.Drawing.Size(140, 25)
        Me.cmdCargarMarcas.TabIndex = 3
        Me.cmdCargarMarcas.Text = "Cargar todas las Marcas"
        Me.ToolTip1.SetToolTip(Me.cmdCargarMarcas, "Carga Todas las Marcas")
        Me.cmdCargarMarcas.UseVisualStyleBackColor = False
        '
        '_sstGrupos_TabPage2
        '
        Me._sstGrupos_TabPage2.Controls.Add(Me.txtFlexVarios)
        Me._sstGrupos_TabPage2.Controls.Add(Me.flexVarios)
        Me._sstGrupos_TabPage2.Controls.Add(Me.cmdCargarFamilias)
        Me._sstGrupos_TabPage2.Location = New System.Drawing.Point(4, 22)
        Me._sstGrupos_TabPage2.Name = "_sstGrupos_TabPage2"
        Me._sstGrupos_TabPage2.Size = New System.Drawing.Size(522, 229)
        Me._sstGrupos_TabPage2.TabIndex = 2
        Me._sstGrupos_TabPage2.Text = "Varios"
        '
        'txtFlexVarios
        '
        Me.txtFlexVarios.AcceptsReturn = True
        Me.txtFlexVarios.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlexVarios.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlexVarios.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlexVarios.Location = New System.Drawing.Point(16, 50)
        Me.txtFlexVarios.MaxLength = 0
        Me.txtFlexVarios.Name = "txtFlexVarios"
        Me.txtFlexVarios.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlexVarios.Size = New System.Drawing.Size(64, 22)
        Me.txtFlexVarios.TabIndex = 9
        Me.txtFlexVarios.Visible = False
        '
        'flexVarios
        '
        Me.flexVarios.DataSource = Nothing
        Me.flexVarios.Location = New System.Drawing.Point(14, 30)
        Me.flexVarios.Name = "flexVarios"
        Me.flexVarios.OcxState = CType(resources.GetObject("flexVarios.OcxState"), System.Windows.Forms.AxHost.State)
        Me.flexVarios.Size = New System.Drawing.Size(494, 152)
        Me.flexVarios.TabIndex = 12
        '
        'cmdCargarFamilias
        '
        Me.cmdCargarFamilias.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCargarFamilias.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCargarFamilias.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCargarFamilias.Location = New System.Drawing.Point(16, 190)
        Me.cmdCargarFamilias.Name = "cmdCargarFamilias"
        Me.cmdCargarFamilias.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCargarFamilias.Size = New System.Drawing.Size(140, 25)
        Me.cmdCargarFamilias.TabIndex = 5
        Me.cmdCargarFamilias.Text = "Cargar todas las Familias"
        Me.ToolTip1.SetToolTip(Me.cmdCargarFamilias, "Carga Todas las Familias")
        Me.cmdCargarFamilias.UseVisualStyleBackColor = False
        '
        'cmdInicializar
        '
        Me.cmdInicializar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdInicializar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdInicializar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdInicializar.Location = New System.Drawing.Point(255, 284)
        Me.cmdInicializar.Name = "cmdInicializar"
        Me.cmdInicializar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdInicializar.Size = New System.Drawing.Size(77, 25)
        Me.cmdInicializar.TabIndex = 6
        Me.cmdInicializar.Text = "&Borrar todo"
        Me.ToolTip1.SetToolTip(Me.cmdInicializar, "Inicializa Importes")
        Me.cmdInicializar.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(16, 284)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(169, 21)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Supr Eliminar Renglón"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(134, 342)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 97
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.SystemColors.Control
        Me.btnGuardar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGuardar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGuardar.Location = New System.Drawing.Point(19, 342)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGuardar.Size = New System.Drawing.Size(109, 36)
        Me.btnGuardar.TabIndex = 96
        Me.btnGuardar.Text = "&Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'frmCorpoABCDescuentosVendExternos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(549, 388)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.cmdInicializar)
        Me.Controls.Add(Me.sstGrupos)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(298, 150)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoABCDescuentosVendExternos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Abc a Descuentos de Vendedores Externos"
        Me.sstGrupos.ResumeLayout(False)
        Me._sstGrupos_TabPage0.ResumeLayout(False)
        CType(Me.flexJoyeria, System.ComponentModel.ISupportInitialize).EndInit()
        Me._sstGrupos_TabPage1.ResumeLayout(False)
        CType(Me.flexRelojeria, System.ComponentModel.ISupportInitialize).EndInit()
        Me._sstGrupos_TabPage2.ResumeLayout(False)
        CType(Me.flexVarios, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             ABC DESCUENTOS A VENDEDORES EXTERNOS                                                         *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      VIERNES 10 DE OCTUBRE DE 2003                                                                *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Dim mblnSALIR As Boolean
    Dim Edita As Boolean
    Dim valida As Boolean
    Dim FueraClick As Boolean
    Dim EditVarios As Boolean
    Dim EditRelojeria As Boolean
    Dim I As Integer

    Function Guardar() As Boolean
        Dim NumPartida As Integer
        On Error GoTo Err_Renamed
        Dim blnTransaccion As Boolean
        If ValidaDatos() = False Then
            Exit Function
        End If
        Cnn.BeginTrans()
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnTransaccion = True
        'Guardar los Datos de Joyeria
        'Primero Inicializamos los Datos de Joyeria
        ModStoredProcedures.PR_IMECatDesctosVExternos(CStr(gCODJOYERIA), "0", "0", "0", "0", "0", "0", C_ELIMINACION, CStr(0))
        Cmd.Execute()
        'Guardamos los Datos del Grid de Joyeria
        NumPartida = 1
        With flexJoyeria
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Trim(.get_TextMatrix(I, 2)) <> "" Then
                    ModStoredProcedures.PR_IMECatDesctosVExternos(CStr(gCODJOYERIA), CStr(NumPartida), VB6.Format(.get_TextMatrix(I, 0), "#####0.00"), VB6.Format(.get_TextMatrix(I, 1), "#####0.00"), "0", "0", VB6.Format(.get_TextMatrix(I, 2), "#####0.00"), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    NumPartida = NumPartida + 1
                End If
            Next
        End With
        'Guardar los datos de Relojeria
        'Inicializamos los Datos de Relojeria
        ModStoredProcedures.PR_IMECatDesctosVExternos(CStr(gCODRELOJERIA), "0", "0", "0", "0", "0", "0", C_ELIMINACION, CStr(0))
        Cmd.Execute()
        'Guardamos los Datos del Grid de Relojeria
        NumPartida = 1
        With flexRelojeria
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                    ModStoredProcedures.PR_IMECatDesctosVExternos(CStr(gCODRELOJERIA), CStr(NumPartida), "0", "0", .get_TextMatrix(I, 2), "0", VB6.Format(.get_TextMatrix(I, 1), "#####0.00"), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    NumPartida = NumPartida + 1
                End If
            Next
        End With
        'Guardar los Datos de Varios
        'Inicializamos los Datos de Varios
        ModStoredProcedures.PR_IMECatDesctosVExternos(CStr(gCODVARIOS), "0", "0", "0", "0", "0", "0", C_ELIMINACION, CStr(0))
        Cmd.Execute()
        'Guardamos los Datos del Grid de Varios
        NumPartida = 1
        With flexVarios
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                    ModStoredProcedures.PR_IMECatDesctosVExternos(CStr(gCODVARIOS), CStr(NumPartida), "0", "0", "0", .get_TextMatrix(I, 2), VB6.Format(.get_TextMatrix(I, 1), "#####0.00"), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    NumPartida = NumPartida + 1
                End If
            Next
        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        MsgBox("¡¡¡Los datos se han guardado exitosamente!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        Me.Close()
Err_Renamed:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Private Sub CargaInicial()
        Dim EstaVacio As Boolean
        On Error GoTo Err_Renamed
        'Primero Checamos si hay Informacion en Joyeria
        gStrSql = "SELECT * FROM CatDesctosVExternos WHERE CodGrupo = " & gCODJOYERIA
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            Exit Sub
        Else
            'Carga los Datos del Grupo Joyeria
            gStrSql = "SELECT NumPartida,ISNULL(ImporteIni,0) AS ImpIni,ISNULL(ImporteFin,0) AS ImpFin,ISNULL(PorcDescto,0) AS Porcentaje " & "FROM CatDesctosVExternos WHERE CodGrupo = " & gCODJOYERIA & " ORDER BY Numpartida"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                With flexJoyeria
                    .Row = 1
                    Do While Not RsGral.EOF
                        .set_TextMatrix(.Row, 0, VB6.Format(RsGral.Fields("ImpIni").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, 1, VB6.Format(RsGral.Fields("ImpFin").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, 2, VB6.Format(RsGral.Fields("Porcentaje").Value, "###,##0.00"))
                        RsGral.MoveNext()
                        If .Row = .Rows - 1 And Not RsGral.EOF Then
                            .Rows = .Rows + 1
                            .set_TextMatrix(.Rows - 1, 1, "0.00")
                        End If
                        If Not RsGral.EOF Then
                            .Row = .Row + 1
                        End If
                    Loop
                    .Col = 0
                    .Row = 1
                End With
            End If
        End If
        'Carga los Datos del Grupo Relojeria
        gStrSql = "SELECT * FROM CatDesctosVExternos WHERE CodGrupo = " & gCODRELOJERIA
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            Exit Sub
        Else
            gStrSql = "SELECT CM.DESCMARCA,CD.CODMARCA,CD.PORCDESCTO,CD.NUMPARTIDA " & "FROM CATDESCTOSVEXTERNOS CD INNER JOIN CATMARCAS CM " & "ON CD.CODMARCA = CM.CODMARCA " & "WHERE CD.CODGRUPO = " & gCODRELOJERIA & " " & "ORDER BY CD.NUMPARTIDA"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                With flexRelojeria
                    .Row = 1
                    Do While Not RsGral.EOF
                        .set_TextMatrix(.Row, 0, Trim(RsGral.Fields("DescMarca").Value))
                        .Col = 0
                        .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter
                        .set_TextMatrix(.Row, 1, VB6.Format(RsGral.Fields("PorcDescto").Value, "###,##0.00"))
                        .Col = 1
                        .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
                        .set_TextMatrix(.Row, 2, RsGral.Fields("CodMArca").Value)
                        RsGral.MoveNext()
                        If .Row = .Rows - 1 And Not RsGral.EOF Then
                            .Rows = .Rows + 1
                        End If
                        If Not RsGral.EOF Then
                            .Row = .Row + 1
                        End If
                    Loop
                    .Rows = RsGral.RecordCount + 1
                    .Col = 0
                    .Row = 1
                End With
            End If
        End If
        'Carga los Datos del Grupo Varios
        gStrSql = "SELECT * FROM CatDesctosVExternos WHERE CodGrupo = " & gCODVARIOS
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount = 0 Then
            Exit Sub
        Else
            gStrSql = "SELECT CF.DESCFAMILIA,CD.CODFAMILIA,CD.PORCDESCTO,CD.NUMPARTIDA " & "FROM CATDESCTOSVEXTERNOS CD INNER JOIN CATFAMILIAS CF " & "ON CD.CODFAMILIA = CF.CODFAMILIA AND CD.CODGRUPO = CF.CODGRUPO " & "WHERE CD.CodGrupo = " & gCODVARIOS & " " & "ORDER BY CD.NUMPARTIDA"
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            RsGral = Cmd.Execute
            If RsGral.RecordCount > 0 Then
                With flexVarios
                    .Row = 1
                    Do While Not RsGral.EOF
                        .Col = 0
                        .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter
                        .set_TextMatrix(.Row, 0, Trim(RsGral.Fields("DescFamilia").Value))
                        .Col = 1
                        .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
                        .set_TextMatrix(.Row, 1, VB6.Format(RsGral.Fields("PorcDescto").Value, "###,##0.00"))
                        .set_TextMatrix(.Row, 2, RsGral.Fields("CodFamilia").Value)
                        RsGral.MoveNext()
                        If .Row = .Rows - 1 And Not RsGral.EOF Then
                            .Rows = .Rows + 1
                        End If
                        If Not RsGral.EOF Then
                            .Row = .Row + 1
                        End If
                    Loop
                    .Rows = RsGral.RecordCount + 1
                    .Col = 0
                    .Row = 1
                End With
            End If
        End If
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub CambiarFormatoTxtenCaptura()
        If sstGrupos.SelectedIndex = 0 Then
            With txtFlexJoyeria
                Select Case flexJoyeria.Col
                    Case 0 'Cantidad Inicial
                        .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                        'UPGRADE_WARNING: TextBox property txtFlexJoyeria.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                        .MaxLength = 15
                    Case 1 'Cantidad Final
                        .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                        'UPGRADE_WARNING: TextBox property txtFlexJoyeria.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                        .MaxLength = 15
                    Case 2 'Porcentaje de Descuento
                        .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                        'UPGRADE_WARNING: TextBox property txtFlexJoyeria.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                        .MaxLength = 5
                End Select
            End With
        ElseIf sstGrupos.SelectedIndex = 1 Then
            With txtFlexRelojeria
                .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                'UPGRADE_WARNING: TextBox property txtFlexRelojeria.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                .MaxLength = 5
            End With
        ElseIf sstGrupos.SelectedIndex = 2 Then
            With txtFlexVarios
                .TextAlign = System.Windows.Forms.HorizontalAlignment.Right
                'UPGRADE_WARNING: TextBox property txtFlexVarios.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                .MaxLength = 5
            End With
        End If
    End Sub

    Sub EncabezadoJoyeria()
        With flexJoyeria
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 1750)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Importe Inicial"
            .Col = 1
            .set_ColWidth(1, 0, 1750)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Importe Final"
            .Col = 2
            .set_ColWidth(2, 0, 1000)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "% Desc"
            .Rows = 11
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub EncabezadoRelojeria()
        With flexRelojeria
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 3500)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Marca"
            .Col = 1
            .set_ColWidth(1, 0, 1000)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "% Desc"
            .Col = 2
            .set_ColWidth(2, 0, 0)
            .Rows = 11
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub EncabezadoVarios()
        With flexVarios
            .Col = 0
            .Row = 0
            .set_ColWidth(0, 0, 3500)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "Familia"
            .Col = 1
            .set_ColWidth(1, 0, 1000)
            .CellAlignment = 5
            .CellFontBold = True
            .Text = "% Desc"
            .Col = 2
            .set_ColWidth(2, 0, 0)
            .Rows = 11
            .Col = 0
            .Row = 1
        End With
    End Sub

    Sub Limpiar()
        Nuevo()
        sstGrupos.Focus()
    End Sub

    Sub Nuevo()
        flexJoyeria.Clear()
        flexRelojeria.Clear()
        flexVarios.Clear()
        EncabezadoJoyeria()
        EncabezadoRelojeria()
        EncabezadoVarios()
        sstGrupos.SelectedIndex = 0
        FueraClick = False
        Edita = False
        mblnSALIR = False
        EditVarios = False
        EditRelojeria = False
        valida = False
    End Sub

    Private Sub EliminaRenglon()
        Select Case sstGrupos.SelectedIndex
            Case 0
                flexJoyeria.RemoveItem((flexJoyeria.Row))
                flexJoyeria.Rows = flexJoyeria.Rows + 1
            Case 1
                flexRelojeria.RemoveItem((flexRelojeria.Row))
                flexRelojeria.Rows = flexRelojeria.Rows + 1
            Case 2
                flexVarios.RemoveItem((flexVarios.Row))
                flexVarios.Rows = flexVarios.Rows + 1
        End Select
    End Sub

    Function EstaVaciaJoyeria() As Boolean
        With flexJoyeria
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" And Numerico(.get_TextMatrix(I, 2)) <> "" Then
                    EstaVaciaJoyeria = False
                    Exit Function
                End If
            Next
            MsgBox("El grupo joyería no tiene informacion de descuentos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            EstaVaciaJoyeria = True
        End With
    End Function

    Function EstaVaciaRelojeria() As Boolean
        With flexRelojeria
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                    EstaVaciaRelojeria = False
                    Exit Function
                End If
            Next
            MsgBox("El grupo relojería no tiene informacion de descuentos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            EstaVaciaRelojeria = True
        End With
    End Function

    Function EstaVacioVarios() As Boolean
        With flexVarios
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                    EstaVacioVarios = False
                    Exit Function
                End If
            Next
            MsgBox("El grupo varios no tiene informacion de descuentos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            EstaVacioVarios = True
        End With
    End Function

    Function ChecaPorcentajes() As Boolean
        With flexJoyeria
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" And Trim(.get_TextMatrix(I, 1)) <> "" Then
                    If CDbl(Numerico(.get_TextMatrix(I, 2))) = 0 Then
                        MsgBox("Existen porcentajes en cero en el grupo joyería, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        ChecaPorcentajes = False
                        sstGrupos.SelectedIndex = 0
                        flexJoyeria.Row = I
                        flexJoyeria.Col = 2
                        If flexJoyeria.Row > 6 Then
                            flexJoyeria.TopRow = flexJoyeria.Row
                        End If
                        flexJoyeria.Focus()
                        Exit Function
                    End If
                End If
            Next
        End With
        With flexRelojeria
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" Then
                    If CDbl(Numerico(.get_TextMatrix(I, 1))) = 0 Then
                        MsgBox("Existen porcentajes en cero en el grupo relojería, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        ChecaPorcentajes = False
                        sstGrupos.SelectedIndex = 1
                        flexRelojeria.Row = I
                        flexRelojeria.Col = 1
                        If flexRelojeria.Row > 5 Then
                            flexRelojeria.TopRow = flexRelojeria.Row
                        End If
                        flexRelojeria.Focus()
                        Exit Function
                    End If
                End If
            Next
        End With
        With flexVarios
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, 0)) <> "" Then
                    If CDbl(Numerico(.get_TextMatrix(I, 1))) = 0 Then
                        MsgBox("Existen porcentajes en cero en el grupo varios, Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        ChecaPorcentajes = False
                        sstGrupos.SelectedIndex = 2
                        flexVarios.Row = I
                        flexVarios.Col = 1
                        If flexVarios.Row > 5 Then
                            flexVarios.TopRow = flexVarios.Row
                        End If
                        flexVarios.Focus()
                        Exit Function
                    End If
                End If
            Next
        End With
        ChecaPorcentajes = True
    End Function

    Function ValidaDatos() As Boolean
        ValidaDatos = False
        'Validar el Grid de Joyeria
        If EstaVaciaJoyeria() Then
            sstGrupos.SelectedIndex = 0
            flexJoyeria.Focus()
            Exit Function
        End If
        If EstaVaciaRelojeria() Then
            sstGrupos.SelectedIndex = 1
            flexRelojeria.Focus()
            Exit Function
        End If
        If EstaVacioVarios() Then
            sstGrupos.SelectedIndex = 2
            flexVarios.Focus()
            Exit Function
        End If
        If Not ChecaPorcentajes() Then
            Exit Function
        End If
        '        With flexJoyeria
        '            For I = 1 To .Rows - 1
        '                If (Trim(.TextMatrix(I, 0)) <> "" And Trim(.TextMatrix(I, 1)) = "") Or _
        ''                   (Trim(.TextMatrix(I, 0)) = "" And Trim(.TextMatrix(I, 1)) <> "") Then
        '                    MsgBox "No se ha capturado toda la información para el grupo joyería" & vbNewLine & _
        ''                    "Favor de verificar...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
        '                    sstGrupos.Tab = 0
        '                    flexJoyeria.SetFocus
        '                    Exit Function
        '                End If
        '            Next
        '        End With
        ValidaDatos = True
    End Function

    Private Sub cmdCargarFamilias_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCargarFamilias.Click
        On Error GoTo Err_Renamed
        gStrSql = "SELECT DescFamilia,CodFamilia FROM CatFamilias " & "WHERE CodGrupo = " & gCODVARIOS & " ORDER BY DescFamilia"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            With flexVarios
                .Row = 1
                Do While Not RsGral.EOF
                    .set_TextMatrix(.Row, 0, Trim(RsGral.Fields("DescFamilia").Value))
                    .Col = 0
                    .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter
                    .set_TextMatrix(.Row, 1, VB6.Format(0, "###,##0.00"))
                    .Col = 1
                    .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
                    .set_TextMatrix(.Row, 2, RsGral.Fields("CodFamilia").Value)
                    RsGral.MoveNext()
                    If .Row = .Rows - 1 And Not RsGral.EOF Then
                        .Rows = .Rows + 1
                    End If
                    If Not RsGral.EOF Then
                        .Row = .Row + 1
                    End If
                Loop
                .Rows = RsGral.RecordCount + 1
            End With
        End If
        flexVarios.Row = 1
        flexVarios.Col = 0
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub cmdCargarFamilias_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCargarFamilias.Enter
        Pon_Tool()
    End Sub

    Private Sub cmdCargarFamilias_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmdCargarFamilias.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            flexVarios.Focus()
        End If
    End Sub

    Private Sub cmdCargarMarcas_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCargarMarcas.Click
        On Error GoTo Err_Renamed
        gStrSql = "SELECT DescMarca,CodMarca FROM CatMarcas " & "WHERE CodGrupo = " & gCODRELOJERIA & " ORDER BY DescMarca"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            With flexRelojeria
                .Row = 1
                Do While Not RsGral.EOF
                    .set_TextMatrix(.Row, 0, Trim(RsGral.Fields("DescMarca").Value))
                    .Col = 0
                    .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter
                    .set_TextMatrix(.Row, 1, VB6.Format(0, "###,##0.00"))
                    .Col = 1
                    .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
                    .set_TextMatrix(.Row, 2, RsGral.Fields("CodMArca").Value)
                    RsGral.MoveNext()
                    If .Row = .Rows - 1 And Not RsGral.EOF Then
                        .Rows = .Rows + 1
                    End If
                    If Not RsGral.EOF Then
                        .Row = .Row + 1
                    End If
                Loop
                .Rows = RsGral.RecordCount + 1
            End With
        End If
        flexRelojeria.Row = 1
        flexRelojeria.Col = 0
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub cmdCargarMarcas_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCargarMarcas.Enter
        Pon_Tool()
    End Sub

    Private Sub cmdCargarMarcas_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmdCargarMarcas.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            flexRelojeria.Focus()
        End If
    End Sub

    Private Sub cmdCargarMarcas_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCargarMarcas.Leave
        'UPGRADE_ISSUE: Control TabIndex could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.TabIndex > cmdCargarMarcas.TabIndex And sstGrupos.SelectedIndex = 1 Then
            cmdInicializar.Focus()
        End If
    End Sub

    Private Sub cmdInicializar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInicializar.Click
        If sstGrupos.SelectedIndex = 0 Then
            flexJoyeria.Clear()
            EncabezadoJoyeria()
        ElseIf sstGrupos.SelectedIndex = 1 Then
            flexRelojeria.Clear()
            EncabezadoRelojeria()
        ElseIf sstGrupos.SelectedIndex = 2 Then
            flexVarios.Clear()
            EncabezadoVarios()
        End If
    End Sub

    Private Sub cmdInicializar_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInicializar.Enter
        Pon_Tool()
    End Sub

    Private Sub cmdInicializar_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmdInicializar.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape And sstGrupos.SelectedIndex = 0 Then
            flexJoyeria.Focus()
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape And sstGrupos.SelectedIndex = 1 Then
            cmdCargarMarcas.Focus()
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape And sstGrupos.SelectedIndex = 2 Then
            cmdCargarFamilias.Focus()
        End If
    End Sub

    Private Sub flexJoyeria_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexJoyeria.DblClick
        flexJoyeria_KeyPressEvent(flexJoyeria, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub flexJoyeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexJoyeria.Enter
        Pon_Tool()
    End Sub

    Private Sub flexJoyeria_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexJoyeria.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete Then
            EliminaRenglon()
        End If
    End Sub

    Private Sub flexJoyeria_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexJoyeria.KeyPressEvent
        Dim lonR, lonI As Integer
        'KeyAscii = vbKeyReturn
        If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then
            'Verifica si se puede capturar la fila
            If flexJoyeria.Row > 1 Then
                If flexJoyeria.get_TextMatrix(flexJoyeria.Row - 1, 0) <> "" Then
                    For lonR = 1 To flexJoyeria.Row - 1 Step 1
                        For lonI = 0 To 2 Step 1
                            If flexJoyeria.get_TextMatrix(lonR, lonI) = "" Then
                                'MsgBox "Hace falta información en la captura", vbExclamation, cNomEmp
                                flexJoyeria.Row = lonR
                                flexJoyeria.Col = lonI
                                If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                                CambiarFormatoTxtenCaptura()
                                MSHFlexGridEdit(flexJoyeria, txtFlexJoyeria, eventArgs.keyAscii)
                                If Len(Trim(txtFlexJoyeria.Text)) = 1 Then
                                    System.Windows.Forms.SendKeys.Send("{right}")
                                End If
                                'Edita = True
                                Exit Sub
                            End If
                        Next lonI
                    Next lonR
                Else
                    'Edita = False
                    cmdInicializar.Focus()
                    Exit Sub
                End If
            End If
            'Edita el campo sólo si es Editable
            If flexJoyeria.get_TextMatrix(flexJoyeria.Row - 1, 1) = "0.00" Then
                'Edita = False
                cmdInicializar.Focus()
                Exit Sub
            End If
            If flexJoyeria.Row >= 1 And flexJoyeria.Col < 3 Then
                If flexJoyeria.Col = 1 And flexJoyeria.get_TextMatrix(flexJoyeria.Row, 0) = "" Then
                    MsgBox("Se requiere importe inicial...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    flexJoyeria.Col = 0
                    flexJoyeria.Focus()
                    'Edita = True
                    Exit Sub
                ElseIf flexJoyeria.Col = 2 And Trim(flexJoyeria.get_TextMatrix(flexJoyeria.Row, 0)) = "" Then
                    MsgBox("Se requiere importe inicial...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    flexJoyeria.Col = 0
                    flexJoyeria.Focus()
                    'Edita = True
                    Exit Sub
                ElseIf flexJoyeria.Col = 2 And Trim(flexJoyeria.get_TextMatrix(flexJoyeria.Row, 1)) = "" And Trim(flexJoyeria.get_TextMatrix(flexJoyeria.Row, 0)) <> "" Then
                    '''MsgBox "Primero debe teclear el importe final...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
                    flexJoyeria.set_TextMatrix(flexJoyeria.Row, 1, "0")
                    '''flexJoyeria.SetFocus
                    Edita = True
                End If
                If flexJoyeria.Row = flexJoyeria.Rows - 1 Then
                    flexJoyeria.Rows = flexJoyeria.Rows + 1
                    flexJoyeria.set_TextMatrix(flexJoyeria.Rows - 1, 2, "0.00")
                End If
                If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                CambiarFormatoTxtenCaptura()
                MSHFlexGridEdit(flexJoyeria, txtFlexJoyeria, eventArgs.keyAscii)
                If Len(Trim(txtFlexJoyeria.Text)) = 1 Then
                    System.Windows.Forms.SendKeys.Send("{right}")
                End If
                'Edita = True
            End If
        ElseIf eventArgs.keyAscii = System.Windows.Forms.Keys.Escape Then
            sstGrupos.SelectedIndex = 0
            sstGrupos.Focus()
        End If
    End Sub

    Private Sub flexRelojeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexRelojeria.Enter
        Pon_Tool()
    End Sub

    Private Sub flexRelojeria_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexRelojeria.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Return Then
            flexRelojeria_KeyPressEvent(flexRelojeria, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
            If Not Edita Then
                cmdCargarMarcas.Focus()
            Else
                Edita = False
            End If
        ElseIf eventArgs.keyCode = System.Windows.Forms.Keys.Escape Then
            sstGrupos.SelectedIndex = 1
            sstGrupos.Focus()
        ElseIf eventArgs.keyCode = System.Windows.Forms.Keys.Delete Then
            EliminaRenglon()
        End If
    End Sub

    Private Sub flexRelojeria_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexRelojeria.KeyPressEvent
        If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape And flexRelojeria.Col = 1 Then
            If Trim(flexRelojeria.get_TextMatrix(flexRelojeria.Row, 0)) <> "" Then
                If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                CambiarFormatoTxtenCaptura()
                MSHFlexGridEdit(flexRelojeria, txtFlexRelojeria, eventArgs.keyAscii)
                If Len(Trim(txtFlexRelojeria.Text)) = 1 Then
                    System.Windows.Forms.SendKeys.Send("{right}")
                End If
                Edita = True
                Exit Sub
            Else
                Edita = False
                Exit Sub
            End If
        End If
    End Sub

    Private Sub flexVarios_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles flexVarios.Enter
        Pon_Tool()
    End Sub

    Private Sub flexVarios_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles flexVarios.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Return Then
            flexVarios_KeyPressEvent(flexVarios, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
            If Not Edita Then
                cmdCargarFamilias.Focus()
            Else
                Edita = False
            End If
        End If
        If eventArgs.keyCode = System.Windows.Forms.Keys.Escape Then
            sstGrupos.SelectedIndex = 2
            sstGrupos.Focus()
        End If
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete Then
            EliminaRenglon()
        End If
    End Sub

    Private Sub flexVarios_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles flexVarios.KeyPressEvent
        If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape And flexVarios.Col = 1 Then
            If Trim(flexVarios.get_TextMatrix(flexVarios.Row, 0)) <> "" Then
                If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                CambiarFormatoTxtenCaptura()
                MSHFlexGridEdit(flexVarios, txtFlexVarios, eventArgs.keyAscii)
                If Len(Trim(txtFlexVarios.Text)) = 1 Then
                    System.Windows.Forms.SendKeys.Send("{right}")
                End If
                Edita = True
                Exit Sub
            Else
                Edita = False
                Exit Sub
            End If
        End If
    End Sub

    'UPGRADE_WARNING: Form event frmCorpoABCDescuentosVendExternos.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmCorpoABCDescuentosVendExternos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        'UPGRADE_WARNING: Form method frmCorpoABCDescuentosVendExternos.ZOrder has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        Me.BringToFront()
    End Sub

    'UPGRADE_WARNING: Form event frmCorpoABCDescuentosVendExternos.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmCorpoABCDescuentosVendExternos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCorpoABCDescuentosVendExternos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        Nuevo()
        FueraClick = True
        CargaInicial()
    End Sub

    Private Sub frmCorpoABCDescuentosVendExternos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        'Si se cierra el formulario y existio algun cambio en el registro se
        'informa al usuario del cabio y si desea guardar el registro, ya sea
        'que sea nuevo o un registro modificado
        If Not mblnSALIR Then
            '        If Cambios = True And mblnNuevo = False Then
            '            Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
            '                Case vbYes: 'Guardar el registro
            '                    If Guardar = False Then
            '                        Cancel = 1
            '                    End If
            '                Case vbNo: 'No hace nada y permite el cierre del formulario
            '                Case vbCancel: 'Cancela el cierre del formulario sin guardar
            '                    Cancel = 1
            '            End Select
            '        End If
        Else
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSALIR = False
                    Cancel = 1
                    sstGrupos.Focus()
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoABCDescuentosVendExternos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
    End Sub

    Private Sub sstGrupos_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sstGrupos.SelectedIndexChanged
        Static PreviousTab As Short = sstGrupos.SelectedIndex()
        If FueraClick Then
            FueraClick = False
            Exit Sub
        End If
        If txtFlexJoyeria.Visible Then
            If (flexJoyeria.Col = 0 Or flexJoyeria.Col = 1 Or flexJoyeria.Col = 2) And CDbl(Numerico(txtFlexJoyeria.Text)) = 0 Then
                txtFlexJoyeria.Focus()
                FueraClick = True
                sstGrupos.SelectedIndex = 0
                Exit Sub
            End If
        End If
        If txtFlexRelojeria.Visible Then
            If flexRelojeria.Col = 1 And CDbl(Numerico(txtFlexRelojeria.Text)) = 0 Then
                FueraClick = True
                sstGrupos.SelectedIndex = 1
                txtFlexRelojeria.Focus()
                Exit Sub
            End If
        End If
        If txtFlexVarios.Visible Then
            If flexVarios.Col = 1 And CDbl(Numerico(txtFlexVarios.Text)) = 0 Then
                FueraClick = True
                sstGrupos.SelectedIndex = 2
                txtFlexVarios.Focus()
                Exit Sub
            End If
        End If
        If sstGrupos.SelectedIndex = 0 Then
            ToolTip1.SetToolTip(sstGrupos, "Grupo Joyería")
            flexJoyeria.Focus()
        ElseIf sstGrupos.SelectedIndex = 1 Then
            ToolTip1.SetToolTip(sstGrupos, "Grupo Relojería")
            flexRelojeria.Focus()
        ElseIf sstGrupos.SelectedIndex = 2 Then
            ToolTip1.SetToolTip(sstGrupos, "Grupo Varios")
            flexVarios.Focus()
        End If
        PreviousTab = sstGrupos.SelectedIndex()
    End Sub

    Private Sub sstGrupos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sstGrupos.Enter
        If sstGrupos.SelectedIndex = 0 Then
            ToolTip1.SetToolTip(sstGrupos, "Grupo Joyería")
        ElseIf sstGrupos.SelectedIndex = 1 Then
            ToolTip1.SetToolTip(sstGrupos, "Grupo Relojería")
        ElseIf sstGrupos.SelectedIndex = 2 Then
            ToolTip1.SetToolTip(sstGrupos, "Grupo Varios")
        End If
        Pon_Tool()
    End Sub

    Private Sub sstGrupos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles sstGrupos.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If sstGrupos.SelectedIndex = 0 Then
                flexJoyeria.Focus()
            ElseIf sstGrupos.SelectedIndex = 1 Then
                flexRelojeria.Focus()
            ElseIf sstGrupos.SelectedIndex = 2 Then
                flexVarios.Focus()
            End If
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSALIR = True
            Me.Close()
        End If
    End Sub

    Private Sub sstGrupos_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles sstGrupos.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Left Or KeyCode = System.Windows.Forms.Keys.Right Then
            If sstGrupos.SelectedIndex = 0 Then
                ToolTip1.SetToolTip(sstGrupos, "Grupo Joyería")
            ElseIf sstGrupos.SelectedIndex = 1 Then
                ToolTip1.SetToolTip(sstGrupos, "Grupo Relojería")
            ElseIf sstGrupos.SelectedIndex = 2 Then
                ToolTip1.SetToolTip(sstGrupos, "Grupo Varios")
            End If
            Pon_Tool()
        End If
    End Sub

    Private Sub sstGrupos_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sstGrupos.Leave
        If sstGrupos.SelectedIndex = 1 Then
            flexRelojeria.Focus()
        ElseIf sstGrupos.SelectedIndex = 2 Then
            flexVarios.Focus()
        End If
    End Sub

    Private Sub txtFlexJoyeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlexJoyeria.Enter
        SelTextoTxt(txtFlexJoyeria)
        If flexJoyeria.Col = 0 Then
            ToolTip1.SetToolTip(txtFlexJoyeria, "Teclee la Cantidad Inicial.")
        ElseIf flexJoyeria.Col = 1 Then
            ToolTip1.SetToolTip(txtFlexJoyeria, "Teclee la Cantidad Final.")
        ElseIf flexJoyeria.Col = 2 Then
            ToolTip1.SetToolTip(txtFlexJoyeria, "Teclee el Porcentaje de Descuento.")
        End If
        Pon_Tool()
    End Sub

    Private Sub txtFlexJoyeria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlexJoyeria.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        valida = False
        If KeyCode = System.Windows.Forms.Keys.Return Then
            With flexJoyeria
                If .Row = 1 And .Col = 0 Then
                    If CDbl(Numerico(txtFlexJoyeria.Text)) = 0 Then
                        MsgBox("El importe inicial de la primera partida no puede ser cero" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
                        txtFlexJoyeria.Text = ""
                        If .get_TextMatrix(.Row, 2) = "" Then
                            .set_TextMatrix(.Row, 2, "")
                        End If
                        txtFlexJoyeria.Visible = False
                        valida = False
                        .Focus()
                        Exit Sub
                    End If
                    If Trim(.get_TextMatrix(.Row, 1)) <> "0.00" And Trim(.get_TextMatrix(.Row, 1)) <> "" Then
                        If CDec(Numerico(txtFlexJoyeria.Text)) >= CDec(Numerico(.get_TextMatrix(.Row, 1))) Then
                            MsgBox("La cantidad inicial no puede ser mayor o igual que la cantidad final" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            txtFlexJoyeria.Text = ""
                            valida = False
                            txtFlexJoyeria.Focus()
                            Exit Sub
                        End If
                    End If
                    .Text = VB6.Format(CDec(Numerico(txtFlexJoyeria.Text)), "###,##0.00")
                    If CDbl(Numerico(.get_TextMatrix(.Row, 2))) = 0 Then
                        .set_TextMatrix(.Row, 1, "0.00")
                        .set_TextMatrix(.Row, 2, "0.00")
                    End If
                    txtFlexJoyeria.Text = ""
                    txtFlexJoyeria.Visible = False
                    .Col = 1
                    valida = True
                    Exit Sub
                ElseIf .Row = 1 And .Col = 1 Then
                    If CDec(Numerico(txtFlexJoyeria.Text)) <= CDec(Numerico(.get_TextMatrix(.Row, 0))) Then
                        If CDec(Numerico(.get_TextMatrix(.Row, 0))) > 0 And CDec(Numerico(txtFlexJoyeria.Text)) = 0 And CDec(Numerico(.get_TextMatrix(.Row + 1, 0))) = 0 Then
                            .Text = VB6.Format(CDec(Numerico(txtFlexJoyeria.Text)), "###,##0.00")
                            txtFlexJoyeria.Text = ""
                            txtFlexJoyeria.Visible = False
                            .Col = 2
                            valida = True
                            Exit Sub
                        Else
                            MsgBox("El importe final no puede ser menor al del importe inicial del siguiente rango" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            .set_TextMatrix(.Row, .Col, "0.00")
                            txtFlexJoyeria.Text = ""
                            txtFlexJoyeria.Visible = False
                            valida = False
                            .Focus()
                            Exit Sub
                        End If
                        MsgBox("La cantidad final no puede ser menor o igual que la cantidad inicial" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        txtFlexJoyeria.Text = ""
                        valida = False
                        txtFlexJoyeria.Focus()
                        Exit Sub
                    ElseIf Trim(.get_TextMatrix(.Row + 1, 0)) <> "" Then
                        If CDec(Numerico(txtFlexJoyeria.Text)) >= CDec(Numerico(.get_TextMatrix(.Row + 1, 0))) Then
                            MsgBox("El importe final no puede ser mayor o igual que el importe inicial" & vbNewLine & "del siguiente rango, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            txtFlexJoyeria.Text = ""
                            valida = False
                            txtFlexJoyeria.Focus()
                            Exit Sub
                        Else
                            .Text = VB6.Format(CDec(Numerico(txtFlexJoyeria.Text)), "###,##0.00")
                            txtFlexJoyeria.Text = ""
                            txtFlexJoyeria.Visible = False
                            .Col = 2
                            valida = True
                            Exit Sub
                        End If
                    Else
                        .Text = VB6.Format(CDec(Numerico(txtFlexJoyeria.Text)), "###,##0.00")
                        txtFlexJoyeria.Text = ""
                        txtFlexJoyeria.Visible = False
                        .Col = 2
                        valida = True
                        Exit Sub
                    End If
                ElseIf .Row = 1 And .Col = 2 Then
                    If CDbl(Numerico(txtFlexJoyeria.Text)) = 0 Then
                        txtFlexJoyeria.Text = VB6.Format(Numerico(txtFlexJoyeria.Text), "###,##0.00")
                        MsgBox("El porcentaje debe ser mayor que cero, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .set_TextMatrix(.Row, 2, "0.00")
                        txtFlexJoyeria.Text = ""
                        txtFlexJoyeria.Visible = False
                        valida = False
                        .Focus()
                        Exit Sub
                    End If
                    .Text = VB6.Format(CDec(Numerico(txtFlexJoyeria.Text)), "###,##0.00")
                    txtFlexJoyeria.Text = ""
                    txtFlexJoyeria.Visible = False
                    .Col = 0
                    '                If .Row = .Rows - 2 Then
                    '                    .Rows = .Rows + 1
                    '                    .TextMatrix(.Rows - 1, 2) = "0.00"
                    '                End If
                    .Row = .Row + 1
                    valida = True
                    Exit Sub
                ElseIf .Row > 1 And .Col = 0 Then
                    If CDec(Numerico(txtFlexJoyeria.Text)) <= CDec(Numerico(.get_TextMatrix(.Row - 1, 1))) Then
                        MsgBox("La cantidad inicial del nuevo rango no puede ser menor o igual;" & vbNewLine & "que la cantidad final del rango anterior" & vbNewLine & vbNewLine & "Favor de verificar ...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        txtFlexJoyeria.Text = ""
                        txtFlexJoyeria.Visible = False
                        .Focus()
                        valida = False
                        Exit Sub
                    End If
                    If Trim(.get_TextMatrix(.Row, 1)) <> "0.00" And Trim(.get_TextMatrix(.Row, 1)) <> "" Then
                        If CDec(Numerico(txtFlexJoyeria.Text)) >= CDec(Numerico(.get_TextMatrix(.Row, 1))) Then
                            MsgBox("La cantidad inicial no puede ser mayor o igual que la cantidad final" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            txtFlexJoyeria.Text = ""
                            valida = False
                            txtFlexJoyeria.Focus()
                            Exit Sub
                        End If
                    End If
                    .Text = VB6.Format(CDec(Numerico(txtFlexJoyeria.Text)), "###,##0.00")
                    If CDbl(Numerico(.get_TextMatrix(.Row, 1))) = 0 Then .set_TextMatrix(.Row, 1, "0.00")
                    If CDbl(Numerico(.get_TextMatrix(.Row, 2))) = 0 Then .set_TextMatrix(.Row, 2, "0.00")
                    txtFlexJoyeria.Text = ""
                    txtFlexJoyeria.Visible = False
                    .Col = 1
                    valida = True
                    Exit Sub
                ElseIf .Row > 1 And .Col = 1 Then
                    If CDec(Numerico(txtFlexJoyeria.Text)) <= CDec(Numerico(.get_TextMatrix(.Row, 0))) Then
                        If CDec(Numerico(.get_TextMatrix(.Row, 0))) > 0 And CDec(Numerico(txtFlexJoyeria.Text)) = 0 And CDec(Numerico(.get_TextMatrix(.Row + 1, 0))) = 0 Then
                            .Text = VB6.Format(CDec(Numerico(txtFlexJoyeria.Text)), "###,##0.00")
                            txtFlexJoyeria.Text = ""
                            txtFlexJoyeria.Visible = False
                            .Col = 2
                            valida = True
                            Exit Sub
                        Else
                            MsgBox("El importe final no puede ser menor al del importe inicial del siguiente rango" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            txtFlexJoyeria.Text = ""
                            valida = False
                            txtFlexJoyeria.Focus()
                            Exit Sub
                        End If
                        MsgBox("La cantidad final no puede ser menor o igual que la cantidad inicial" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        txtFlexJoyeria.Text = ""
                        valida = False
                        txtFlexJoyeria.Focus()
                        Exit Sub
                    ElseIf Trim(.get_TextMatrix(.Row + 1, 0)) <> "" Then
                        If CDec(Numerico(txtFlexJoyeria.Text)) >= CDec(Numerico(.get_TextMatrix(.Row + 1, 0))) Then
                            MsgBox("El importe final no puede ser mayor o igual que el importe inicial" & vbNewLine & "del siguiente rango, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            txtFlexJoyeria.Text = ""
                            valida = False
                            txtFlexJoyeria.Focus()
                            Exit Sub
                        Else
                            .Text = VB6.Format(CDec(Numerico(txtFlexJoyeria.Text)), "###,##0.00")
                            txtFlexJoyeria.Text = ""
                            txtFlexJoyeria.Visible = False
                            .Col = 2
                            valida = True
                            Exit Sub
                        End If
                    Else
                        .Text = VB6.Format(CDec(Numerico(txtFlexJoyeria.Text)), "###,##0.00")
                        txtFlexJoyeria.Text = ""
                        txtFlexJoyeria.Visible = False
                        .Col = 2
                        valida = True
                        Exit Sub
                    End If
                ElseIf .Row > 1 And .Col = 2 Then
                    If CDbl(Numerico(txtFlexJoyeria.Text)) = 0 Then
                        MsgBox("El porcentaje debe ser mayor que cero, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .set_TextMatrix(.Row, 2, "")
                        txtFlexJoyeria.Visible = False
                        '''txtFlexJoyeria = ""
                        valida = False
                        .Focus()
                        Exit Sub
                    End If
                    .Text = VB6.Format(CDec(Numerico(txtFlexJoyeria.Text)), "###,##0.00")
                    txtFlexJoyeria.Text = ""
                    txtFlexJoyeria.Visible = False
                    .Col = 0
                    '                If .Row = .Rows - 2 Then
                    '                    .Rows = .Rows + 1
                    '                    .TextMatrix(.Rows - 1, 2) = "0.00"
                    '                End If
                    .Row = .Row + 1
                    valida = True
                    Exit Sub
                End If
            End With
        End If
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            'If CCur(Numerico(txtFlexJoyeria)) = 0 Or CCur(Numerico(flexJoyeria.TextMatrix(flexJoyeria.Row, flexJoyeria.Col))) <> 0 Then
            txtFlexJoyeria.Visible = False
            flexJoyeria.Focus()
            'Else
            '    txtFlexJoyeria_LostFocus
            'End If
        End If
    End Sub

    Private Sub txtFlexJoyeria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlexJoyeria.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        With flexJoyeria
            Select Case .Col
                Case 0, 1
                    ModEstandar.MskCantidad(txtFlexJoyeria.Text, KeyAscii, 12, 2, (txtFlexJoyeria.SelectionStart))
                Case 2
                    ModEstandar.MskCantidad(txtFlexJoyeria.Text, KeyAscii, 2, 2, (txtFlexJoyeria.SelectionStart))
            End Select
        End With
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlexJoyeria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlexJoyeria.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        If Not txtFlexJoyeria.Visible Then Exit Sub
        txtFlexJoyeria_Validating(txtFlexJoyeria, New System.ComponentModel.CancelEventArgs(True))
        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "sstGrupos" Then
            sstGrupos_SelectedIndexChanged(sstGrupos, New System.EventArgs())
            Exit Sub
        End If
        If Not valida Then Exit Sub
        If flexJoyeria.Col = 2 And Trim(flexJoyeria.get_TextMatrix(flexJoyeria.Row, 2)) = "" Then
            flexJoyeria.set_TextMatrix(flexJoyeria.Row, 2, "0.00")
        End If
        txtFlexJoyeria.Visible = False
    End Sub

    Private Sub txtFlexJoyeria_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtFlexJoyeria.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtFlexJoyeria_KeyDown(txtFlexJoyeria, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Return Or 0 * &H10000))
        If Not valida Then
            Cancel = True
        Else
            Cancel = False
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtFlexRelojeria_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlexRelojeria.Enter
        SelTextoTxt(txtFlexRelojeria)
        ToolTip1.SetToolTip(txtFlexRelojeria, "Teclee el Porcentaje de Descuento")
    End Sub

    Private Sub txtFlexRelojeria_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlexRelojeria.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        valida = False
        If KeyCode = System.Windows.Forms.Keys.Return Then
            With flexRelojeria
                If CDbl(Numerico(txtFlexRelojeria.Text)) = 0 Then
                    txtFlexRelojeria.Text = VB6.Format(txtFlexRelojeria.Text, "###,##0.00")
                    MsgBox("El porcentaje de descuento no puede ser cero" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    txtFlexRelojeria.Text = ""
                    txtFlexRelojeria.Visible = False
                    valida = False
                    .set_TextMatrix(.Row, .Col, "0.00")
                    .Focus()
                    Exit Sub
                End If
                .Text = VB6.Format(txtFlexRelojeria.Text, "###,##0.00")
                txtFlexRelojeria.Text = ""
                txtFlexRelojeria.Visible = False
                valida = True
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                End If
            End With
        End If
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            txtFlexRelojeria.Visible = False
            flexRelojeria.Focus()
            'txtFlexRelojeria_LostFocus
            EditRelojeria = False
        End If
    End Sub

    Private Sub txtFlexRelojeria_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlexRelojeria.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        ModEstandar.MskCantidad(txtFlexRelojeria.Text, KeyAscii, 2, 2, (txtFlexRelojeria.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlexRelojeria_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlexRelojeria.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        If Not txtFlexRelojeria.Visible Then Exit Sub
        txtFlexRelojeria_Validating(txtFlexRelojeria, New System.ComponentModel.CancelEventArgs(True))
        If EditRelojeria Then
            EditRelojeria = False
            Exit Sub
        End If
        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "sstGrupos" Then
            sstGrupos_SelectedIndexChanged(sstGrupos, New System.EventArgs())
            Exit Sub
        End If
        If Not valida Then
            EditRelojeria = True
            Exit Sub
        End If
        If Trim(flexRelojeria.get_TextMatrix(flexRelojeria.Row, 1)) = "" Then
            flexRelojeria.set_TextMatrix(flexRelojeria.Row, 1, "0.00")
        End If
        txtFlexRelojeria.Visible = False
    End Sub

    Private Sub txtFlexRelojeria_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtFlexRelojeria.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If EditRelojeria Then
            Cancel = True
            txtFlexRelojeria.Focus()
            GoTo EventExitSub
        End If
        txtFlexRelojeria_KeyDown(txtFlexRelojeria, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Return Or 0 * &H10000))
        If Not valida Then
            Cancel = True
        Else
            Cancel = False
        End If
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtFlexVarios_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlexVarios.Enter
        SelTextoTxt(txtFlexVarios)
        ToolTip1.SetToolTip(txtFlexVarios, "Teclee el Porcentaje de Descuento")
    End Sub

    Private Sub txtFlexVarios_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlexVarios.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            With flexVarios
                If CDbl(Numerico(txtFlexVarios.Text)) = 0 Then
                    txtFlexVarios.Text = VB6.Format(Numerico(txtFlexVarios.Text), "###,##0.00")
                    MsgBox("El porcentaje de descuento no puede ser cero" & vbNewLine & "Favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    txtFlexVarios.Text = ""
                    txtFlexVarios.Visible = False
                    .set_TextMatrix(.Row, .Col, "0.00")
                    valida = False
                    .Focus()
                    Exit Sub
                End If
                .Text = VB6.Format(txtFlexVarios.Text, "###,##0.00")
                txtFlexVarios.Text = ""
                txtFlexVarios.Visible = False
                valida = True
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                End If
            End With
        End If
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            txtFlexVarios.Visible = False
            flexVarios.Focus()
            'txtFlexVarios_LostFocus
            EditVarios = False
        End If
    End Sub

    Private Sub txtFlexVarios_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlexVarios.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        ModEstandar.MskCantidad(txtFlexVarios.Text, KeyAscii, 2, 2, (txtFlexVarios.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlexVarios_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlexVarios.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        If Not txtFlexVarios.Visible Then Exit Sub
        txtFlexVarios_Validating(txtFlexVarios, New System.ComponentModel.CancelEventArgs(True))
        If EditVarios Then
            EditVarios = False
            Exit Sub
        End If
        'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name = "sstGrupos" Then
            sstGrupos_SelectedIndexChanged(sstGrupos, New System.EventArgs())
            Exit Sub
        End If
        If Not valida Then
            EditVarios = True
            Exit Sub
        End If
        If Trim(flexVarios.get_TextMatrix(flexVarios.Row, 1)) = "" Then
            flexVarios.set_TextMatrix(flexVarios.Row, 1, "0.00")
        End If
        txtFlexVarios.Visible = False
    End Sub

    Private Sub txtFlexVarios_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtFlexVarios.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If EditVarios Then
            Cancel = True
            txtFlexVarios.Focus()
            GoTo EventExitSub
        End If
        txtFlexVarios_KeyDown(txtFlexVarios, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Return Or 0 * &H10000))
        If Not valida Then
            Cancel = True
        Else
            Cancel = False
        End If
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub
End Class