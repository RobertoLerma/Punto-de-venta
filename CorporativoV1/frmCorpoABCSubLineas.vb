Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCorpoABCSubLineas
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ' Programa :                ABC de SubLíneas de Joyería
    ' Autor :                   Paimí
    ' Fecha de Inicio :         14 de Mayo de 2003
    ' Fecha de Finalización :
    ' Nota :

    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents dbcDescFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcDescLinea As System.Windows.Forms.ComboBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents mshFlex As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents lblGrupo As System.Windows.Forms.Label
    Public WithEvents _lblSubLinea_2 As System.Windows.Forms.Label
    Public WithEvents _lblSubLinea_1 As System.Windows.Forms.Label
    Public WithEvents _lblSubLinea_0 As System.Windows.Forms.Label
    Public WithEvents fraMarcas As System.Windows.Forms.GroupBox
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btnSalir As Button
    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnGuardar As Button
    Friend WithEvents btnLimpiar As Button
    Friend WithEvents btnEliminar As Button
    Public WithEvents lblSubLinea As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCorpoABCSubLineas))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.fraMarcas = New System.Windows.Forms.GroupBox()
        Me.dbcDescFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcDescLinea = New System.Windows.Forms.ComboBox()
        Me.mshFlex = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.lblGrupo = New System.Windows.Forms.Label()
        Me._lblSubLinea_2 = New System.Windows.Forms.Label()
        Me._lblSubLinea_1 = New System.Windows.Forms.Label()
        Me._lblSubLinea_0 = New System.Windows.Forms.Label()
        Me.lblSubLinea = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.fraMarcas.SuspendLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSubLinea, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(73, 154)
        Me.txtFlex.MaxLength = 50
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(70, 20)
        Me.txtFlex.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.txtFlex, "Digite la SubLínea de Joyería")
        Me.txtFlex.Visible = False
        '
        'fraMarcas
        '
        Me.fraMarcas.BackColor = System.Drawing.SystemColors.Control
        Me.fraMarcas.Controls.Add(Me.dbcDescFamilia)
        Me.fraMarcas.Controls.Add(Me.dbcDescLinea)
        Me.fraMarcas.Controls.Add(Me.txtFlex)
        Me.fraMarcas.Controls.Add(Me.mshFlex)
        Me.fraMarcas.Controls.Add(Me.lblGrupo)
        Me.fraMarcas.Controls.Add(Me._lblSubLinea_2)
        Me.fraMarcas.Controls.Add(Me._lblSubLinea_1)
        Me.fraMarcas.Controls.Add(Me._lblSubLinea_0)
        Me.fraMarcas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMarcas.Location = New System.Drawing.Point(8, 4)
        Me.fraMarcas.Name = "fraMarcas"
        Me.fraMarcas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMarcas.Size = New System.Drawing.Size(393, 295)
        Me.fraMarcas.TabIndex = 0
        Me.fraMarcas.TabStop = False
        '
        'dbcDescFamilia
        '
        Me.dbcDescFamilia.Location = New System.Drawing.Point(82, 56)
        Me.dbcDescFamilia.Name = "dbcDescFamilia"
        Me.dbcDescFamilia.Size = New System.Drawing.Size(297, 21)
        Me.dbcDescFamilia.TabIndex = 3
        '
        'dbcDescLinea
        '
        Me.dbcDescLinea.Location = New System.Drawing.Point(82, 88)
        Me.dbcDescLinea.Name = "dbcDescLinea"
        Me.dbcDescLinea.Size = New System.Drawing.Size(297, 21)
        Me.dbcDescLinea.TabIndex = 5
        '
        'mshFlex
        '
        Me.mshFlex.DataSource = Nothing
        Me.mshFlex.Location = New System.Drawing.Point(12, 135)
        Me.mshFlex.Name = "mshFlex"
        Me.mshFlex.OcxState = CType(resources.GetObject("mshFlex.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mshFlex.Size = New System.Drawing.Size(384, 144)
        Me.mshFlex.TabIndex = 6
        '
        'lblGrupo
        '
        Me.lblGrupo.BackColor = System.Drawing.SystemColors.Window
        Me.lblGrupo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrupo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblGrupo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGrupo.Location = New System.Drawing.Point(82, 24)
        Me.lblGrupo.Name = "lblGrupo"
        Me.lblGrupo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblGrupo.Size = New System.Drawing.Size(297, 21)
        Me.lblGrupo.TabIndex = 8
        '
        '_lblSubLinea_2
        '
        Me._lblSubLinea_2.AutoSize = True
        Me._lblSubLinea_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblSubLinea_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSubLinea_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSubLinea_2.Location = New System.Drawing.Point(12, 92)
        Me._lblSubLinea_2.Name = "_lblSubLinea_2"
        Me._lblSubLinea_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSubLinea_2.Size = New System.Drawing.Size(35, 13)
        Me._lblSubLinea_2.TabIndex = 4
        Me._lblSubLinea_2.Text = "Línea"
        '
        '_lblSubLinea_1
        '
        Me._lblSubLinea_1.AutoSize = True
        Me._lblSubLinea_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblSubLinea_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSubLinea_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSubLinea_1.Location = New System.Drawing.Point(12, 60)
        Me._lblSubLinea_1.Name = "_lblSubLinea_1"
        Me._lblSubLinea_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSubLinea_1.Size = New System.Drawing.Size(39, 13)
        Me._lblSubLinea_1.TabIndex = 2
        Me._lblSubLinea_1.Text = "Familia"
        '
        '_lblSubLinea_0
        '
        Me._lblSubLinea_0.AutoSize = True
        Me._lblSubLinea_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblSubLinea_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblSubLinea_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblSubLinea_0.Location = New System.Drawing.Point(12, 28)
        Me._lblSubLinea_0.Name = "_lblSubLinea_0"
        Me._lblSubLinea_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblSubLinea_0.Size = New System.Drawing.Size(36, 13)
        Me._lblSubLinea_0.TabIndex = 1
        Me._lblSubLinea_0.Text = "Grupo"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(8, 305)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(377, 74)
        Me.Panel3.TabIndex = 69
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
        'frmCorpoABCSubLineas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(409, 390)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.fraMarcas)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(258, 147)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoABCSubLineas"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Sub - Líneas de Joyería"
        Me.fraMarcas.ResumeLayout(False)
        Me.fraMarcas.PerformLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSubLinea, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub



    Const C_RENENCABEZADO As Integer = 0

    Const C_ColDESCRIPCION As Integer = 0
    Const C_ColDESCCORTA As Integer = 1
    Const C_ColDESCRIPCIONTAG As Integer = 2
    Const C_COLCODIGO As Integer = 3
    Const C_COLSTATUS As Integer = 4
    Const C_ColDESCCORTATAG As Integer = 5
    Const C_COLDEPEND As Integer = 6
    Const C_COLSUBLORIGINAL As Integer = 7

    Dim rsLocal As ADODB.Recordset

    Dim mblnSalir As Boolean 'Controla la salida con ESC
    Dim mblnEscape As Boolean
    Dim mblnCambiosEnCodigo1, mblnCambiosEnCodigo2 As Object
    Dim mblnCambiosEnCodigo3 As Boolean
    Dim mblnNuevo As Boolean
    Dim mintDepend As Integer

    'Variables para manejar el combo de Familia
    Dim Tecla1 As Integer
    Dim Tecla2 As Integer
    Dim Tecla3 As Integer
    Dim mblnFueraChange As Boolean
    Public mintCodFamilia As Integer
    Public mintCodLinea As Integer
    Dim I As Integer
    Dim mblnGuardar As Boolean

    Sub BuscarGrupo()
        gStrSql = "select DescGrupo from CatGrupos where codGrupo = " & gCODJOYERIA
        ModEstandar.BorraCmd()
        Cmd.CommandText = "Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount <= 0 Then
            mblnFueraChange = True
            'mintCodGrupo = 0
            lblGrupo.Text = ""
            lblGrupo.Tag = ""
            mblnFueraChange = False
        Else
            Me.lblGrupo.Text = Trim(RsGral.Fields("DescGrupo").Value)
        End If
    End Sub

    Public Sub ScrollGrid()
        'Procedimiento que pone el enfoque en el primer renglón vacío del Grid
        Dim I As Integer
        Dim nCont As Integer 'Cuenta los renglones que están ocupados (que no están vacíos)
        Dim nRen As Integer
        'Aparecen 9 renglones disponibles en el Grid
        'Si son menos de siete registros ocupados, no se utiliza el .TopRow
        'Pero, si son 9 ó más registros, el .TopRow manda el enfoque al primer renglón vacío
        'después de los renglones ocupados
        nRen = 9 'El máximo de renglones que aparece en el grid (Además del encabezado)
        nCont = 0
        With Me.mshFlex
            For I = 1 To .Rows
                If Trim(.get_TextMatrix(I, C_ColDESCRIPCION)) <> "" Then
                    nCont = nCont + 1
                Else
                    Exit For
                End If
            Next I
            If nCont < 9 Then
                'Hay menos de 9 registros
                .Row = nCont + 1
                .Col = C_ColDESCRIPCION
            Else
                'Hay 9 ó más registros, hay que recorrer el grid
                .TopRow = (nCont - nRen) + 2
                .Row = nCont + 1
                .Col = C_ColDESCRIPCION
            End If
        End With
    End Sub

    Public Sub LimpiarFlex()
        On Error Resume Next
        Dim I As Object
        'Pone el enfoque en la última línea disponible para dar de alta una descripción más
        With mshFlex
            .Clear()
            .set_TextMatrix(C_RENENCABEZADO, C_ColDESCRIPCION, "SubLíneas de Joyería")
            .set_TextMatrix(C_RENENCABEZADO, C_ColDESCCORTA, "Desc Corta")
            .set_TextMatrix(C_RENENCABEZADO, C_ColDESCRIPCIONTAG, "DescripcionTag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODIGO, "Código")
            .set_TextMatrix(C_RENENCABEZADO, C_COLSTATUS, "STATUS")
            .set_TextMatrix(C_RENENCABEZADO, C_ColDESCCORTATAG, "DescCortaTag")
            'Colocar los textos de los encabezados centrados
            .Row = C_RENENCABEZADO
            For I = 0 To (.get_Cols() - 1) Step 1
                .Col = I
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next I
        End With
    End Sub

    Public Sub Nuevo()
        mintCodFamilia = 0
        dbcDescFamilia.Text = ""
        dbcDescFamilia.Tag = ""
        mintCodLinea = 0
        dbcDescLinea.Text = ""
        dbcDescLinea.Tag = ""
        mintDepend = 0
        mblnGuardar = False
        LimpiarFlex()
    End Sub

    Public Sub Limpiar()
        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar los registros
                    If Not Guardar() Then
                        mblnNuevo = True
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No guarda los cambios y permite que se limpie el contenido
                    mblnNuevo = True
                Case MsgBoxResult.Cancel 'No hace nada
                    mblnNuevo = True
                    Exit Sub
            End Select
        End If
        Nuevo()
        mblnNuevo = True
        mblnCambiosEnCodigo1 = False
        mblnCambiosEnCodigo2 = False
        mblnCambiosEnCodigo3 = False
        dbcDescFamilia.Focus()
    End Sub

    Public Function ValidaDatos() As Boolean
        On Error Resume Next
        Dim I As Object
        If mintCodFamilia = 0 Then
            MsgBox("Debe especificar la Familia del Artículo.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            ValidaDatos = False
            Exit Function
        End If
        If mintCodLinea = 0 Then
            MsgBox("Debe especificar la Línea a la que pertenece la SubLínea.")
            ValidaDatos = False
            Exit Function
        End If
        With mshFlex
            mintDepend = 0
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_ColDESCRIPCION)) = "" And Trim(.get_TextMatrix(I, C_ColDESCCORTA)) = "" Then Exit For
                If Trim(.get_TextMatrix(I, C_ColDESCRIPCION)) = "" Then
                    MsgBox("Debe especificar la descripción de la SubLínea del Artículo, o borrar el registro", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    .Col = 0
                    .Row = I
                    .Focus()
                    ValidaDatos = False
                    Exit Function
                ElseIf Trim(.get_TextMatrix(I, C_ColDESCCORTA)) = "" Then
                    MsgBox("Debe especificar una descripción corta, o borrar el registro", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    .Col = C_ColDESCCORTA
                    .Row = I
                    .Focus()
                    ValidaDatos = False
                    Exit Function
                Else
                    ValidaDatos = True
                End If
                If .get_TextMatrix(I, C_COLDEPEND) = "S" Then mintDepend = mintDepend + 1
            Next I
        End With
    End Function

    Public Sub Eliminar()
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        Dim TopRowAnterior As Object
        Dim RowAnterior As Integer
        TopRowAnterior = Me.mshFlex.TopRow
        RowAnterior = Me.mshFlex.Row
        If Me.mshFlex.get_TextMatrix(mshFlex.Row, C_COLCODIGO) <> "" And BuscarFlex() Then
            'Preguntar si la columna Status es diferente de ""
            If Me.mshFlex.get_TextMatrix(mshFlex.Row, C_COLSTATUS) <> "" Then
                If Referencia("Select * From CatArticulos Where CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintCodFamilia & " and CodLinea = " & mintCodLinea & " and CodSubLinea = " & CInt(Numerico(Me.mshFlex.get_TextMatrix(mshFlex.Row, C_COLCODIGO)))) Then
                    MsgBox("No es posible eliminar esta SubLínea" & vbNewLine & "debido a que está asociada" & vbNewLine & "con algunos artículos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                    Exit Sub
                End If
                If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

                Cnn.BeginTrans()
                blnTransaction = True
                ModStoredProcedures.PR_IMECatSubLineas(Str(gCODJOYERIA), Str(mintCodFamilia), Str(mintCodLinea), Trim(Me.mshFlex.get_TextMatrix(mshFlex.Row, C_COLCODIGO)), Trim(mshFlex.get_TextMatrix(mshFlex.Row, C_ColDESCRIPCION)), Trim(mshFlex.get_TextMatrix(mshFlex.Row, C_ColDESCCORTA)), C_ELIMINACION, CStr(0))
                Cmd.Execute()
                Cnn.CommitTrans()
                blnTransaction = False
                mshFlex.RemoveItem(mshFlex.Row)
                mshFlex.Rows = mshFlex.Rows + 1
            End If
        Else
            '''no esta dado de alta en el catalogo por lo tanto es nuevo
            With mshFlex
                If Trim(.get_TextMatrix(.Row, C_COLSTATUS)) = "N" Then
                    .RemoveItem(.Row)
                    .Rows = .Rows + 1
                End If
            End With
        End If
        mshFlex.TopRow = TopRowAnterior
        mshFlex.Row = RowAnterior
        mshFlex.Col = C_ColDESCRIPCION

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
        If blnTransaction Then Cnn.RollbackTrans()
    End Sub

    Function BuscarFlex() As Boolean
        On Error GoTo Merr
        gStrSql = "select * from CatSubLineas where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintCodFamilia & " and codLinea = " & mintCodLinea & "and CodSubLinea = " & ModEstandar.Numerico(Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_COLCODIGO))
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
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function Cambios() As Boolean
        On Error Resume Next
        Dim I As Object
        With mshFlex
            For I = 1 To .Rows
                If IsNothing(.get_TextMatrix(I, C_COLSTATUS)) Then
                    Exit For
                End If
                If .get_TextMatrix(I, C_COLSTATUS) = C_ELIMINADO Or .get_TextMatrix(I, C_COLSTATUS) = C_ACTIVO Then
                    'No hace nada
                ElseIf (Trim(.get_TextMatrix(I, C_ColDESCRIPCION)) <> Trim(.get_TextMatrix(I, C_ColDESCRIPCIONTAG)) And (.get_TextMatrix(I, C_COLCODIGO) <> "")) Or (Trim(.get_TextMatrix(I, C_ColDESCCORTA)) <> Trim(.get_TextMatrix(I, C_ColDESCCORTATAG)) And (.get_TextMatrix(I, C_COLCODIGO) <> "")) Then
                    .set_TextMatrix(I, C_COLSTATUS, C_MODIFICADO)
                    Cambios = True
                ElseIf Trim(.get_TextMatrix(I, C_ColDESCRIPCION)) <> Trim(.get_TextMatrix(I, C_ColDESCRIPCIONTAG)) And (.get_TextMatrix(I, C_COLCODIGO) = "") Then
                    .set_TextMatrix(I, C_COLSTATUS, C_NUEVO)
                    Cambios = True
                End If
            Next I
        End With
    End Function

    Public Function Guardar() As Boolean
        On Error GoTo Merr
        Dim nNuevos, nModif As Object
        Dim nBorrados As Integer
        Dim blnTransaction As Boolean
        Dim I As Object
        Dim nPosicion As Integer

        mblnGuardar = True
        nNuevos = 0
        nModif = 0
        nBorrados = 0
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Return Or 0 * &H10000))
        txtFlex.Visible = False

        If Not ValidaDatos() Then
            Exit Function
        End If
        If Not Cambios() Then
            Exit Function
        End If
        If mintDepend >= 1 Then
            MsgBox("Existen artículos clasificados con" & vbNewLine & "los datos que se modificaron" & vbNewLine & vbNewLine & "Estos serán reclasificados pero" & vbNewLine & "su descripción no será alterada" & vbNewLine & vbNewLine & "", MsgBoxStyle.Information, "AVISO")
        End If

        I = 0
        blnTransaction = True
        With mshFlex
            Cnn.BeginTrans()
            For I = 1 To (.Rows)
                If Trim(.get_TextMatrix(I, C_COLSTATUS)) = "" Then Exit For
                With mshFlex
                    Select Case .get_TextMatrix(I, C_COLSTATUS)
                        Case C_MODIFICADO
                            ModStoredProcedures.PR_IMECatSubLineas(Str(gCODJOYERIA), Str(mintCodFamilia), Str(mintCodLinea), .get_TextMatrix(I, C_COLCODIGO), Trim(.get_TextMatrix(I, C_ColDESCRIPCION)), Trim(.get_TextMatrix(I, C_ColDESCCORTA)), C_MODIFICACION, CStr(0))
                            Cmd.Execute()
                            .set_TextMatrix(I, C_ColDESCRIPCIONTAG, .get_TextMatrix(I, C_ColDESCRIPCION))
                            .set_TextMatrix(I, C_ColDESCCORTATAG, .get_TextMatrix(I, C_ColDESCCORTA))
                            .set_TextMatrix(I, C_COLSTATUS, C_ACTIVO)
                            nModif = nModif + 1
                            nPosicion = I
                        Case C_NUEVO
                            ModStoredProcedures.PR_IMECatSubLineas(Str(gCODJOYERIA), Str(mintCodFamilia), Str(mintCodLinea), .get_TextMatrix(I, C_COLCODIGO), Trim(.get_TextMatrix(I, C_ColDESCRIPCION)), Trim(.get_TextMatrix(I, C_ColDESCCORTA)), C_INSERCION, CStr(0))
                            Cmd.Execute()
                            .set_TextMatrix(I, C_ColDESCRIPCIONTAG, .get_TextMatrix(I, C_ColDESCRIPCION))
                            .set_TextMatrix(I, C_ColDESCCORTATAG, .get_TextMatrix(I, C_ColDESCCORTA))
                            .set_TextMatrix(I, C_COLCODIGO, Cmd.Parameters("ID").Value)
                            .set_TextMatrix(I, C_COLSTATUS, C_ACTIVO)
                            .Rows = .Rows + 1
                            nNuevos = nNuevos + 1
                            nPosicion = I
                        Case C_ELIMINADO
                            ModStoredProcedures.PR_IMECatSubLineas(Str(gCODJOYERIA), Str(mintCodFamilia), Str(mintCodLinea), .get_TextMatrix(I, C_COLCODIGO), Trim(.get_TextMatrix(I, C_ColDESCRIPCION)), Trim(.get_TextMatrix(I, C_ColDESCCORTA)), C_ELIMINACION, CStr(0))
                            Cmd.Execute()
                    End Select
                End With
            Next I
            Cnn.CommitTrans()

            If Trim(Me.Tag) = "FRMCXPJOYERIA" Then
                With frmCXPJoyeria
                    .mblnFueraChange = True
                    .dbcSubLinea.Text = Trim(Me.mshFlex.get_TextMatrix(nPosicion, C_ColDESCRIPCION))
                    .dbcSubLinea.Tag = .dbcSubLinea.Text
                    .mintCodSubLinea = CInt(Numerico(Me.mshFlex.get_TextMatrix(nPosicion, C_COLCODIGO)))
                    .mblnFueraChange = False
                End With
                Guardar = True
                Me.Close()
                Exit Function
            End If
        End With
        blnTransaction = False
        MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        mshFlex.TopRow = 1
        mshFlex.Row = 1
        mshFlex.Col = 0
        mblnGuardar = False
        Guardar = True

        Nuevo()
        dbcDescFamilia.Focus()

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
        If blnTransaction Then Cnn.RollbackTrans()
    End Function

    Sub LlenaDatos()
        On Error GoTo Merr
        Dim I As Integer
        Dim nRow As Integer
        If Me.Tag = "FRMCXPJOYERIA" Then
            mblnFueraChange = True
            Me.lblGrupo.Text = C_JOYERIA
            Me.lblGrupo.Tag = C_JOYERIA
            Me.dbcDescFamilia.Text = Trim(frmCXPJoyeria.dbcFamilia.Text)
            Me.dbcDescFamilia.Tag = Me.dbcDescFamilia.Text
            Me.dbcDescFamilia.Text = True
            Me.dbcDescLinea.Text = Trim(frmCXPJoyeria.dbcLinea.Text)
            Me.dbcDescLinea.Tag = Me.dbcDescLinea.Text
            Me.dbcDescLinea.Text = True
            mblnFueraChange = False
        End If
        With mshFlex
            gStrSql = "SELECT LTrim(RTrim(DescSubLinea)) AS DescSubLinea, DescCorta" & ", LTrim(RTrim(DescSubLinea)) AS DescripcionTag" & ", CodSubLinea, '" & C_ACTIVO & "' AS Estatus, '' AS Depend, LTrim(RTrim(DescSubLinea)) as SubLOriginal " & "FROM CatSubLineas " & "WHERE CodGrupo = " & gCODJOYERIA & " AND CodFamilia = " & mintCodFamilia & " AND CodLinea = " & mintCodLinea & " Order by DescSubLinea "
            nRow = .Row
            .Clear()
            ModEstandar.BorraCmd()
            Cmd.CommandText = "Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
            rsLocal = Cmd.Execute
            If rsLocal.RecordCount > 0 Then
                .Recordset = rsLocal
                Encabezado()
                If rsLocal.RecordCount < 8 Then
                    .Rows = 11
                Else
                    .Rows = (rsLocal.RecordCount - 7) + 11
                End If
            Else
                Encabezado()
            End If
            .TopRow = 1
            .Row = 1
            .Col = C_ColDESCRIPCION
            If Trim(Me.Tag) <> "" Then
                ScrollGrid()
                .Focus()
            End If
        End With
        mblnNuevo = False
        mblnCambiosEnCodigo1 = False
        mblnCambiosEnCodigo2 = False
        mblnCambiosEnCodigo3 = False

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Encabezado()
        Dim LnContador As Integer

        With mshFlex
            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusHeavy 'flexFocusLight 'flexFocusNone
            .WordWrap = False
            .FixedRows = 1
            .FixedCols = 0
            .set_Cols(0, 8)

            .set_ColWidth(C_ColDESCRIPCION, 0, 4100)
            .set_ColWidth(C_ColDESCCORTA, 0, 1100)
            .set_ColWidth(C_ColDESCRIPCIONTAG, 0, 1)
            .set_ColWidth(C_COLCODIGO, 0, 1)
            .set_ColWidth(C_COLSTATUS, 0, 1)
            .set_ColWidth(C_ColDESCCORTATAG, 0, 1)
            .set_ColWidth(C_COLDEPEND, 0, 1)
            .set_ColWidth(C_COLSUBLORIGINAL, 0, 1)

            .set_TextMatrix(0, C_ColDESCRIPCION, "SubLíneas de Joyería")
            .set_TextMatrix(0, C_ColDESCCORTA, "Desc Corta")
            .set_TextMatrix(0, C_ColDESCRIPCIONTAG, "DescripcionTag")
            .set_TextMatrix(0, C_COLCODIGO, "Código")
            .set_TextMatrix(0, C_COLSTATUS, "STATUS")
            .set_TextMatrix(0, C_ColDESCRIPCIONTAG, "DescCortaTag")
            .set_TextMatrix(0, C_COLDEPEND, "DEPEND")
            .set_TextMatrix(0, C_COLSUBLORIGINAL, "SUBLORIGINAL")

            'Colocar los textos de los encabezados centrados
            .Row = C_RENENCABEZADO
            For LnContador = 0 To (.get_Cols() - 1) Step 1
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next LnContador

            'Obtiene el último registro o renglón
            If rsLocal.RecordCount > 0 Then
                If rsLocal.RecordCount + 2 < 11 Then
                    .Rows = 11
                Else
                    .Rows = rsLocal.RecordCount + 2
                End If
            Else
                .Rows = 11
                .Row = 1
                .Col = C_ColDESCRIPCION
            End If
        End With
    End Sub

    Private Sub dbcDescFamilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcDescFamilia.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codFamilia, rtrim(ltrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia LIKE '" & Trim(Me.dbcDescFamilia.Text) & "%' Order by DescFamilia "
        ModDCombo.DCChange(lStrSql, Tecla2, dbcDescFamilia)

        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox("¿Desea guardar los cambios?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    If Guardar() Then
                        mblnNuevo = True
                        mblnFueraChange = True
                        Me.dbcDescLinea.Text = ""
                        Me.dbcDescLinea.Tag = ""
                        mintCodLinea = 0
                        mblnFueraChange = False
                        LimpiarFlex()
                        Me.dbcDescFamilia.Focus()
                        ModEstandar.SelTxt()
                    End If
                    Call dbcDescFamilia_Enter(dbcDescFamilia, New System.EventArgs())
                Case MsgBoxResult.No
                    mblnNuevo = True
                    mblnFueraChange = True
                    Me.dbcDescLinea.Text = ""
                    Me.dbcDescLinea.Tag = ""
                    mintCodLinea = 0
                    mblnFueraChange = False
                    LimpiarFlex()
                    Me.dbcDescFamilia.Focus()
                    ModEstandar.SelTxt()
                Case MsgBoxResult.Cancel
            End Select
        End If
        If Me.dbcDescFamilia.Text = "" Then
            mblnFueraChange = True
            Me.dbcDescLinea.Text = ""
            Me.dbcDescLinea.Tag = ""
            mintCodLinea = 0
            mblnFueraChange = False
            LimpiarFlex()
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcDescFamilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcDescFamilia.Enter
        Pon_Tool()
        gStrSql = "SELECT codFamilia, rtrim(ltrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " ORDER BY descFamilia "
        ModDCombo.DCGotFocus(gStrSql, dbcDescFamilia)
    End Sub

    Private Sub dbcDescFamilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcDescFamilia.KeyDown
        Dim Aux As String
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                mblnSalir = True
                Me.Close()
                eventSender.KeyCode = 0
            Case System.Windows.Forms.Keys.Return
                Aux = Trim(Me.dbcDescFamilia.Text)
                'If Me.dbcDescFamilia.SelectedItem <> 0 Then
                '    dbcDescFamilia_Leave(dbcDescFamilia, New System.EventArgs())
                'End If
                Me.dbcDescFamilia.Text = Aux
                Exit Sub
            Case System.Windows.Forms.Keys.Tab
                Aux = Trim(Me.dbcDescFamilia.Text)
                'If Me.dbcDescFamilia.SelectedItem <> 0 Then
                '    Me.dbcDescFamilia.Text = Me.dbcDescFamilia.SelectedItem
                '    dbcDescFamilia_Leave(dbcDescFamilia, New System.EventArgs())
                'End If
                Me.dbcDescFamilia.Text = Aux
                Exit Sub
        End Select
        Tecla2 = eventArgs.KeyCode
    End Sub

    Private Sub dbcDescFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcDescFamilia.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codFamilia, rtrim(ltrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia LIKE '" & Trim(Me.dbcDescFamilia.Text) & "%'"
        Aux = mintCodFamilia
        mintCodFamilia = 0
        ModDCombo.DCLostFocus(dbcDescFamilia, gStrSql, mintCodFamilia)
        If Aux <> mintCodFamilia Or mintCodFamilia = 0 Then
            mblnCambiosEnCodigo2 = True
            mblnFueraChange = True
            Me.dbcDescLinea.Text = ""
            Me.dbcDescLinea.Tag = ""
            mintCodLinea = 0
            Call LimpiarFlex()
            mblnFueraChange = False
        End If
    End Sub

    Private Sub dbcDescFamilia_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcDescFamilia.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcDescFamilia.Text)
        'If Me.dbcDescFamilia.SelectedItem <> "" Then
        'dbcDescFamilia_Leave(dbcDescFamilia, New System.EventArgs())
        'End If
        Me.dbcDescFamilia.Text = Aux
    End Sub

    Private Sub dbcDescLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcDescLinea.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String
        If mblnFueraChange Then Exit Sub
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcDescLinea.Name Then Exit Sub
        lStrSql = "SELECT codLinea, rtrim(ltrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintCodFamilia & " and descLinea LIKE '" & Trim(Me.dbcDescLinea.Text) & "%' Order by DescLinea "
        ModDCombo.DCChange(lStrSql, Tecla3, dbcDescLinea)

        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox("¿Desea guardar los cambios?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    If Guardar() Then
                        mblnNuevo = True
                        LimpiarFlex()
                        Me.dbcDescLinea.Focus()
                        ModEstandar.SelTxt()
                    End If
                    Call dbcDescLinea_Enter(dbcDescLinea, New System.EventArgs())
                Case MsgBoxResult.No
                    mblnNuevo = True
                    LimpiarFlex()
                    Me.dbcDescLinea.Focus()
                    ModEstandar.SelTxt()
                Case MsgBoxResult.Cancel
            End Select
        End If
        If Me.dbcDescLinea.Text = "" Then
            LimpiarFlex()
        End If
        If dbcDescLinea.SelectedItem <> "" Then
            Call dbcDescLinea_Leave(dbcDescLinea, New System.EventArgs())
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcDescLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcDescLinea.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcDescLinea.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT codLinea, rtrim(ltrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintCodFamilia & "ORDER BY descLinea "
        ModDCombo.DCGotFocus(gStrSql, dbcDescLinea)
    End Sub

    Private Sub dbcDescLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcDescLinea.KeyDown
        Dim Aux As String
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                mblnEscape = True
                ModEstandar.RetrocederTab(Me)
                eventSender.KeyCode = 0
            Case System.Windows.Forms.Keys.Return
                Aux = Trim(Me.dbcDescLinea.Text)
                'If Me.dbcDescLinea.SelectedItem <> 0 Then
                '    dbcDescLinea_Leave(dbcDescLinea, New System.EventArgs())
                'End If
                Me.dbcDescLinea.Text = Aux
                Exit Sub
            Case System.Windows.Forms.Keys.Tab
                Aux = Trim(Me.dbcDescLinea.Text)
                'If Me.dbcDescLinea.SelectedItem <> 0 Then
                '    Me.dbcDescLinea.Text = Me.dbcDescLinea.SelectedItem
                '    dbcDescLinea_Leave(dbcDescLinea, New System.EventArgs())
                'End If
                Me.dbcDescLinea.Text = Aux
                Exit Sub
        End Select
        Tecla3 = eventArgs.KeyCode
    End Sub

    Private Sub dbcDescLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcDescLinea.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codLinea, rtrim(ltrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintCodFamilia & " and descLinea LIKE '" & Trim(Me.dbcDescLinea.Text) & "%'"
        Aux = mintCodLinea
        mintCodLinea = 0
        ModDCombo.DCLostFocus(dbcDescLinea, gStrSql, mintCodLinea)

        If mintCodLinea <> Aux Then
            mblnCambiosEnCodigo3 = True
            If Not mblnEscape Then
            Else
                If Not Cambios() Then
                    LlenaDatos()
                End If
            End If
        Else
            mblnCambiosEnCodigo3 = False
        End If
        If Not mblnEscape Then
            If Not Cambios() Then
                LlenaDatos()
            End If
        End If
        mblnEscape = False
    End Sub

    Private Sub dbcDescLinea_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcDescLinea.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcDescLinea.Text)
        'If Me.dbcDescLinea.SelectedItem <> 0 Then
        ' dbcDescLinea_Leave(dbcDescLinea, New System.EventArgs())
        'End If
        Me.dbcDescLinea.Text = Aux
    End Sub

    Private Sub frmCorpoABCSubLineas_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoABCSubLineas_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCorpoABCSubLineas_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                'If UCase(Me.ActiveControl.Name) <> "TXTFLEX" Then ModEstandar.AvanzarTab Me
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    Me.dbcDescLinea.Focus()
                End If
            Case System.Windows.Forms.Keys.Delete
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    If Me.mshFlex.get_TextMatrix(mshFlex.Row, C_ColDESCRIPCION) <> "" Then
                        Call Eliminar()
                    End If
                End If
        End Select
    End Sub

    Private Sub frmCorpoABCSubLineas_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoABCSubLineas_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        'mintCodGrupo = 0
        mintCodFamilia = 0
        mintCodLinea = 0
        BuscarGrupo()
        LlenaDatos()
    End Sub

    Private Sub frmCorpoABCSubLineas_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Trim(Me.Tag) = "" Then
        '    If Not mblnSalir Then
        '        'Si desea cerrar la forma y ésta se encuentra minimizada, se debe restaurar
        '        ModEstandar.RestaurarForma(Me, False)
        '        If Cambios() Then
        '            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
        '                Case MsgBoxResult.Yes
        '                    If Not Guardar() Then 'Si falla el guardar, no cierra la forma
        '                        Cancel = 1
        '                    Else
        '                        mblnNuevo = True
        '                        Cancel = 0
        '                    End If
        '                Case MsgBoxResult.No 'No hace nada y permite que se cierre el formulario
        '                    mblnNuevo = True
        '                    Cancel = 0
        '                Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
        '                    Cancel = 1
        '            End Select
        '        End If
        '    Else 'Se quiere salir con escape
        '        mblnSalir = False
        '        Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '            Case MsgBoxResult.Yes 'Sale del Formulario
        '                Cancel = 0
        '            Case MsgBoxResult.No 'No sale del formulario
        '                Me.dbcDescFamilia.Focus()
        '                ModEstandar.SelTxt()
        '                Cancel = 1
        '        End Select
        '    End If
        'Else
        '    Cancel = 0
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoABCSubLineas_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Select Case Me.Tag
            Case "FRMCXPJOYERIA"
                frmCXPJoyeria.Enabled = True
                frmCXPJoyeria.dbcSubLinea.Focus()
        End Select
        Me.Tag = ""
        'Me = Nothing
    End Sub

    Private Sub mshFlex_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.DblClick
        mshFlex_KeyPressEvent(mshFlex, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub mshFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.Enter
        Pon_Tool()
    End Sub

    Private Sub mshFlex_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles mshFlex.KeyDownEvent
        If mintCodFamilia = 0 Or mintCodLinea = 0 Then
            eventArgs.keyCode = 0
        End If
        With Me.mshFlex
            Select Case eventArgs.keyCode
                Case System.Windows.Forms.Keys.Left
                    .Col = C_ColDESCRIPCION
                Case System.Windows.Forms.Keys.Right
                    .Col = C_ColDESCCORTA
                Case System.Windows.Forms.Keys.Down
                    .Col = C_ColDESCRIPCION
            End Select
        End With

    End Sub

    Private Sub mshFlex_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles mshFlex.KeyPressEvent

        If mintCodFamilia = 0 Or mintCodLinea = 0 Then
            eventArgs.keyAscii = 0
        End If
        With mshFlex
            '''si ya se capturo algo entonces se edita el grid
            '''ya sea con numeros, letras o enter
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then
                If (.Col = C_ColDESCRIPCION Or .Col = C_ColDESCCORTA) Then
                    Select Case .Col
                        Case C_ColDESCCORTA
                            txtFlex.MaxLength = 3
                        Case C_ColDESCRIPCION
                            txtFlex.MaxLength = 50
                    End Select
                    '''en esta parte se validará si es el rengón, columna que le
                    '''corresponde editarse
                    If (.Row > 1) Then
                        '''de tal modo que si el renglón es mayor que 1
                        '''y si un renglón antes del renglón actual está vacío,
                        '''el renglón actual no se editará
                        If Trim(.get_TextMatrix(.Row - 1, C_ColDESCRIPCION)) = "" Or Trim(.get_TextMatrix(.Row - 1, C_ColDESCCORTA)) = "" Then
                            .Focus()
                            Exit Sub
                        End If

                    End If
                    ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, eventArgs.keyAscii)
                    If Len(Trim(txtFlex.Text)) = 1 Then
                        'System.Windows.Forms.SendKeys.Send("{Right}")
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub mshFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.Leave
        mshFlex.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub

    Private Sub txtFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Enter
        SelTextoTxt(txtFlex)
        Pon_Tool()
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub

        With mshFlex
            Select Case .Col
                Case C_ColDESCRIPCION
                    txtFlex.MaxLength = 50
                Case C_ColDESCCORTA
                    txtFlex.MaxLength = 3
            End Select

            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    'txtFlex.Visible = False
                    'txtFlex.Text = ""
                    'mshFlex.SetFocus

                Case System.Windows.Forms.Keys.Return
                    '''If Trim(txtFlex.text) = "" Then Exit Sub
                    Select Case .Col
                        Case C_ColDESCRIPCION
                            If Trim(txtFlex.Text) = "" And Trim(.get_TextMatrix(.Row, C_ColDESCRIPCION)) = "" Then
                                txtFlex.Visible = False
                                mshFlex.Focus()
                                Exit Sub
                            End If
                            If Trim(txtFlex.Text) <> "" Then .set_TextMatrix(.Row, C_ColDESCRIPCION, Trim(txtFlex.Text))
                            .set_TextMatrix(.Row, C_COLDEPEND, IIf(ReferenciaSubL(.get_TextMatrix(.Row, C_COLCODIGO)), "S", "N"))
                            If ArticuloRepetidoenGrid(Trim(.get_TextMatrix(.Row, C_ColDESCRIPCION)), "A") = True Then
                                mshFlex.set_TextMatrix(mshFlex.Row, C_ColDESCRIPCION, Trim(txtFlex.Text))
                                txtFlex.Visible = False
                                MsgBox("Existe un artículo capturado con la misma descripción" & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                                LimpiaDatosArticulo(C_ColDESCRIPCION)
                                txtFlex.Visible = False
                                .Col = C_ColDESCRIPCION
                                mshFlex.Focus()
                                Exit Sub
                            End If
                            If .get_TextMatrix(.Row, C_COLSTATUS) = "" Then
                                .set_TextMatrix(.Row, C_COLSTATUS, C_NUEVO)
                            Else
                                .set_TextMatrix(.Row, C_COLSTATUS, C_MODIFICADO)
                            End If

                            mblnNuevo = False
                            .Col = C_ColDESCCORTA
                            ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, KeyCode)
                            SelTextoTxt(txtFlex)

                        Case C_ColDESCCORTA
                            If Trim(txtFlex.Text) = "" And Trim(.get_TextMatrix(.Row, C_ColDESCCORTA)) = "" Then
                                txtFlex.Visible = False
                                mshFlex.Focus()
                                Exit Sub
                            End If
                            .set_TextMatrix(.Row, C_ColDESCCORTA, Trim(txtFlex.Text))
                            .set_TextMatrix(.Row, C_COLSTATUS, "M")
                            txtFlex.Visible = False
                            If ArticuloRepetidoenGrid(Trim(.get_TextMatrix(.Row, C_ColDESCCORTA)), "B") = True Then
                                mshFlex.set_TextMatrix(mshFlex.Row, C_ColDESCCORTA, Trim(txtFlex.Text))
                                MsgBox("Existe un artículo capturado con la misma descripción corta" & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                                LimpiaDatosArticulo(C_ColDESCCORTA)
                                txtFlex.Visible = False
                                .Col = C_ColDESCCORTA
                                mshFlex.Focus()
                                Exit Sub
                            End If
                            If (.Row = .Rows - 1) Then
                                .Rows = .Rows + 1
                                '''ScrollGrid
                            Else
                                .Row = .Row + 1
                            End If
                            ScrollGrid()
                            mblnNuevo = False
                            '''.Row = .Row + 1
                            .Col = C_ColDESCRIPCION
                            If Not mblnGuardar Then
                                If .Enabled Then .Focus()
                            End If
                            txtFlex.Text = ""
                            txtFlex.Visible = False
                            mshFlex.Focus()
                    End Select
            End Select

        End With
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        txtFlex.Visible = False
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Function ArticuloRepetidoenGrid(ByRef lDESC As String, ByRef Tipo As String) As Boolean
        Dim UnaVez As Integer
        '''A -->  Descirpcion Normal
        '''B -->  Descirpcion Corta
        ArticuloRepetidoenGrid = False
        UnaVez = 0
        If Tipo = "A" Then
            If Trim(lDESC) <> "" Then
                'Descripcion
                With mshFlex
                    For I = 1 To .Rows - 1
                        If UCase(Trim(.get_TextMatrix(I, C_ColDESCRIPCION))) = lDESC Then
                            UnaVez = UnaVez + 1
                            If UnaVez > 1 Then
                                ArticuloRepetidoenGrid = True
                                Exit For
                            End If
                        End If
                    Next
                End With
            End If
        End If
        If Tipo = "B" Then
            If Trim(lDESC) <> "" Then
                'DecCorta
                With mshFlex
                    For I = 1 To .Rows - 1
                        If Trim(.get_TextMatrix(.Row, C_ColDESCCORTA)) <> "" Then
                            If UCase(Trim(.get_TextMatrix(I, C_ColDESCCORTA))) = lDESC Then
                                UnaVez = UnaVez + 1
                                If UnaVez > 1 Then ArticuloRepetidoenGrid = True
                            End If
                        End If
                    Next
                End With
            End If
        End If
        Exit Function
    End Function

    Sub LimpiaDatosArticulo(ByRef lColumna As Integer)
        On Error GoTo Merr
        'Este Procedimiento Limpialos Campos Correspondientes a un Artículo, cuando se cambie de Articulo, que se limpien los datos
        With Me.mshFlex
            For I = lColumna To 7
                .set_TextMatrix(.Row, I, "")
            Next
            txtFlex.Text = ""
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub


    Public Function ReferenciaSubL(ByRef lSubLinea As String) As Boolean
        On Error GoTo Merr
        Dim rsLocal As ADODB.Recordset
        Dim lSql As String

        If Trim(lSubLinea) = "" Then Exit Function

        ReferenciaSubL = False
        lSql = "Select * from CatArticulos(Nolock) Where CodGrupo = " & gCODJOYERIA & " And CodFamilia = " & mintCodFamilia & " And CodLinea = " & mintCodLinea & " And CodSubLinea = " & lSubLinea
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then ReferenciaSubL = True

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function


    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        Eliminar()
    End Sub


End Class