'**********************************************************************************************************************'
'*PROGRAMA: ABC DE MARCA JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmCorpoABCMarca
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ' Programa :                ABC de Marcas de Relojería
    ' Autor :                   Paimí
    ' Fecha de Inicio:          12 de Mayo de 2003
    ' Fecha de Finalización:
    ' Nota:                     Si este cambia, debe cambiar también el de Familias y viceversa
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents mshFlex As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents lblGrupo As System.Windows.Forms.Label
    Public WithEvents _lblFamilias_2 As System.Windows.Forms.Label
    Public WithEvents fraMarcas As System.Windows.Forms.GroupBox
    Public WithEvents lblFamilias As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Const C_RENENCABEZADO As Integer = 0

    Const C_ColDESCRIPCION As Integer = 0
    Const C_ColDESCRIPCIONTAG As Integer = 1
    Const C_COLCODIGO As Integer = 2
    Const C_COLSTATUS As Integer = 3
    Const C_COLDEPEND As Integer = 4
    Const C_COLMARORIGINAL As Integer = 5

    Dim tecla As Integer
    Dim mblnFueraChange As Boolean
    Dim mintDepend As Integer
    Dim mblnCambiosEnCodigo As Boolean
    Dim mblnNuevo As Boolean
    Dim mblnSalir As Boolean
    Dim mblnGuardar As Boolean

    Dim rsLocal As ADODB.Recordset
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btnSalir As Button
    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnGuardar As Button
    Friend WithEvents btnLimpiar As Button
    Friend WithEvents btnEliminar As Button
    Dim I As Integer


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCorpoABCMarca))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.fraMarcas = New System.Windows.Forms.GroupBox()
        Me.mshFlex = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.lblGrupo = New System.Windows.Forms.Label()
        Me._lblFamilias_2 = New System.Windows.Forms.Label()
        Me.lblFamilias = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.fraMarcas.SuspendLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFamilias, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(40, 160)
        Me.txtFlex.MaxLength = 50
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(65, 20)
        Me.txtFlex.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtFlex, "Digite la Marca de Relojería")
        Me.txtFlex.Visible = False
        '
        'fraMarcas
        '
        Me.fraMarcas.BackColor = System.Drawing.Color.Silver
        Me.fraMarcas.Controls.Add(Me.txtFlex)
        Me.fraMarcas.Controls.Add(Me.mshFlex)
        Me.fraMarcas.Controls.Add(Me.lblGrupo)
        Me.fraMarcas.Controls.Add(Me._lblFamilias_2)
        Me.fraMarcas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMarcas.Location = New System.Drawing.Point(16, 12)
        Me.fraMarcas.Name = "fraMarcas"
        Me.fraMarcas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMarcas.Size = New System.Drawing.Size(393, 313)
        Me.fraMarcas.TabIndex = 0
        Me.fraMarcas.TabStop = False
        '
        'mshFlex
        '
        Me.mshFlex.DataSource = Nothing
        Me.mshFlex.Location = New System.Drawing.Point(16, 96)
        Me.mshFlex.Name = "mshFlex"
        Me.mshFlex.OcxState = CType(resources.GetObject("mshFlex.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mshFlex.Size = New System.Drawing.Size(359, 195)
        Me.mshFlex.TabIndex = 3
        '
        'lblGrupo
        '
        Me.lblGrupo.BackColor = System.Drawing.SystemColors.Window
        Me.lblGrupo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrupo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblGrupo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGrupo.Location = New System.Drawing.Point(70, 40)
        Me.lblGrupo.Name = "lblGrupo"
        Me.lblGrupo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblGrupo.Size = New System.Drawing.Size(305, 21)
        Me.lblGrupo.TabIndex = 4
        '
        '_lblFamilias_2
        '
        Me._lblFamilias_2.AutoSize = True
        Me._lblFamilias_2.BackColor = System.Drawing.Color.Silver
        Me._lblFamilias_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFamilias_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblFamilias_2.Location = New System.Drawing.Point(24, 44)
        Me._lblFamilias_2.Name = "_lblFamilias_2"
        Me._lblFamilias_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFamilias_2.Size = New System.Drawing.Size(36, 13)
        Me._lblFamilias_2.TabIndex = 1
        Me._lblFamilias_2.Text = "Grupo"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.fraMarcas)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(421, 417)
        Me.Panel1.TabIndex = 1
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(16, 329)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(393, 74)
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
        'frmCorpoABCMarca
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(446, 442)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(269, 157)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoABCMarca"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Marcas de Relojería"
        Me.fraMarcas.ResumeLayout(False)
        Me.fraMarcas.PerformLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFamilias, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Public Sub ScrollGrid()
        'Procedimiento que pone el enfoque en el primer renglón vacío del Grid
        Dim I As Integer
        Dim nCont As Integer 'Cuenta los renglones que están ocupados (que no están vacíos)
        Dim nRen As Integer
        'Aparecen 7 renglones disponibles en el Grid
        'Si son menos de siete registros ocupados, no se utiliza el .TopRow
        'Pero, si son 7 ó más registros, el .TopRow manda el enfoque al primer renglón vacío
        'después de los renglones ocupados
        nRen = 7 'El máximo de renglones que aparece en el grid (Además del encabezado)
        nCont = 0
        With Me.mshFlex
            For I = 1 To .Rows
                If Trim(.get_TextMatrix(I, C_ColDESCRIPCION)) <> "" Then
                    nCont = nCont + 1
                Else
                    Exit For
                End If
            Next I
            If nCont < 7 Then
                'Hay menos de 7 registros
                .Row = nCont + 1
                .Col = C_ColDESCRIPCION
            Else
                'Hay 7 ó más registros, hay que recorrer el grid
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
            .set_TextMatrix(C_RENENCABEZADO, C_ColDESCRIPCION, "Marcas")
            .set_TextMatrix(C_RENENCABEZADO, C_ColDESCRIPCIONTAG, "DescripcionTag")
            .set_TextMatrix(C_RENENCABEZADO, C_COLCODIGO, "Código")
            .set_TextMatrix(C_RENENCABEZADO, C_COLSTATUS, "STATUS")
            'Colocar los textos de los encabezados centrados
            .Row = C_RENENCABEZADO
            For I = 0 To (.get_Cols() - 1) Step 1
                .Col = I
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next I
        End With
    End Sub

    Sub BuscarGrupo()
        gStrSql = "select DescGrupo from CatGrupos where codGrupo = " & gCODRELOJERIA
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
            lblGrupo.Text = Trim(RsGral.Fields("DescGrupo").Value)
        End If
    End Sub

    Public Sub Nuevo()
        mblnFueraChange = True
        lblGrupo.Text = ""
        lblGrupo.Tag = ""
        mblnFueraChange = False
        mintDepend = 0
        mblnGuardar = False
        LimpiarFlex()
    End Sub

    'Public Sub Limpiar()
    '    If Cambios() And Not mblnNuevo Then
    '        Select Case MsgBox(C_msgGUARDAR, vbQuestion + vbYesNoCancel, gstrNombCortoEmpresa)
    '            Case vbYes: 'Guardar los registros
    '                If Not Guardar() Then
    '                    mblnNuevo = True
    '                    Exit Sub
    '                End If
    '            Case vbNo: 'No guarda los cambios y permite que se limpie el contenido
    '                mblnNuevo = True
    '            Case vbCancel: 'No hace nada
    '                mblnNuevo = True
    '                Exit Sub
    '        End Select
    '    End If
    '    Nuevo
    '    mblnCambiosEnCodigo = False
    '    mblnNuevo = True
    'End Sub

    Public Function ValidaDatos() As Boolean
        On Error Resume Next
        Dim I As Object
        BuscarGrupo()
        '    If mintCodGrupo = 0 Then
        '        MsgBox "Debe especificar el Grupo al que pertenecerá la Marca", vbInformation + vbOKOnly, gstrNombCortoEmpresa
        '        mblnNuevo = True
        '        Limpiar
        '        ValidaDatos = False
        '        Exit Function
        '    End If
        With mshFlex
            mintDepend = 0
            For I = 1 To .Rows - 1
                If IsNothing(.get_TextMatrix(I, C_COLSTATUS)) Then
                    Exit For
                End If
                If .get_TextMatrix(I, C_ColDESCRIPCION) = "" Then
                    MsgBox("Debe especificar la Marca, o borrar el registro", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    .Col = 0
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
                If Referencia("Select * From CatModelos Where CodGrupo = " & gCODRELOJERIA & " And CodMarca = " & mshFlex.get_TextMatrix(mshFlex.Row, C_COLCODIGO)) Then
                    MsgBox("No es posible eliminar esta Marca" & vbNewLine & "debido a que tiene Modelos asociados a ella", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                    Exit Sub
                End If
                If Referencia("Select * from CatArticulos Where CodGrupo = " & gCODRELOJERIA & " And CodMarca = " & mshFlex.get_TextMatrix(mshFlex.Row, C_COLCODIGO)) Then
                    MsgBox("No es posible eliminar esta Marca" & vbNewLine & "debido a que está asociada" & vbNewLine & "con algunos artículos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    Exit Sub
                End If
                If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
                    Exit Sub
                End If
                Cnn.BeginTrans()
                blnTransaction = True
                ModStoredProcedures.PR_IMECatMarcas(Str(gCODRELOJERIA), Trim(Me.mshFlex.get_TextMatrix(mshFlex.Row, C_COLCODIGO)), Trim(mshFlex.get_TextMatrix(mshFlex.Row, C_ColDESCRIPCION)), C_ELIMINACION, CStr(0))
                Cmd.Execute()
                Cnn.CommitTrans()
                blnTransaction = False
            End If
        End If
        LlenaDatos()
        Me.mshFlex.TopRow = TopRowAnterior
        Me.mshFlex.Row = RowAnterior
        Me.mshFlex.Col = C_ColDESCRIPCION
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
        If blnTransaction Then Cnn.RollbackTrans()
    End Sub

    Function BuscarFlex() As Boolean
        On Error GoTo Merr
        gStrSql = "select * from CatMarcas where codGrupo = " & gCODRELOJERIA & " and codMarca = " & ModEstandar.Numerico(Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_COLCODIGO))
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
                ElseIf Trim(.get_TextMatrix(I, C_ColDESCRIPCION)) <> Trim(.get_TextMatrix(I, C_ColDESCRIPCIONTAG)) And (.get_TextMatrix(I, C_COLCODIGO) <> "") Then
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
        If Not ValidaDatos() Then
            Exit Function
        End If
        If Not Cambios() Then
            Exit Function
        End If
        I = 0
        If mintDepend >= 1 Then
            MsgBox("Existen artículos clasificados con" & vbNewLine & "los datos que se modificaron" & vbNewLine & vbNewLine & "Estos serán reclasificados pero" & vbNewLine & "su descripción no será alterada" & vbNewLine & vbNewLine & "", MsgBoxStyle.Information, "AVISO")
        End If

        With mshFlex
            Cnn.BeginTrans()
            For I = 1 To (.Rows)
                If Trim(.get_TextMatrix(I, C_COLSTATUS)) = "" Then Exit For
                blnTransaction = True
                With mshFlex
                    Select Case .get_TextMatrix(I, C_COLSTATUS)
                        Case C_MODIFICADO
                            ModStoredProcedures.PR_IMECatMarcas(Str(gCODRELOJERIA), .get_TextMatrix(I, C_COLCODIGO), Trim(.get_TextMatrix(I, C_ColDESCRIPCION)), C_MODIFICACION, CStr(0))
                            Cmd.Execute()
                            .set_TextMatrix(I, C_ColDESCRIPCIONTAG, .get_TextMatrix(I, C_ColDESCRIPCION))
                            .set_TextMatrix(I, C_COLSTATUS, C_ACTIVO)
                            nModif = nModif + 1
                            nPosicion = I
                        Case C_NUEVO
                            ModStoredProcedures.PR_IMECatMarcas(Str(gCODRELOJERIA), .get_TextMatrix(I, C_COLCODIGO), Trim(.get_TextMatrix(I, C_ColDESCRIPCION)), C_INSERCION, CStr(0))
                            Cmd.Execute()
                            .set_TextMatrix(I, C_ColDESCRIPCIONTAG, .get_TextMatrix(I, C_ColDESCRIPCION))
                            .set_TextMatrix(I, C_COLCODIGO, Cmd.Parameters("ID").Value)
                            .set_TextMatrix(I, C_COLSTATUS, C_ACTIVO)
                            .Rows = .Rows + 1
                            nNuevos = nNuevos + 1
                            nPosicion = I
                        Case C_ELIMINADO
                            ModStoredProcedures.PR_IMECatMarcas(Str(gCODRELOJERIA), .get_TextMatrix(I, C_COLCODIGO), Trim(.get_TextMatrix(I, C_ColDESCRIPCION)), C_ELIMINACION, CStr(0))
                            Cmd.Execute()
                    End Select
                End With
                blnTransaction = False
            Next I
            Cnn.CommitTrans()
            If Trim(Me.Tag) = "FRMCXPRELOJERIA" Then
                With frmCXPRelojeria
                    .mblnFueraChange = True
                    .dbcMarca.Text = Trim(Me.mshFlex.get_TextMatrix(nPosicion, C_ColDESCRIPCION))
                    .dbcMarca.Tag = .dbcMarca.Text
                    .mintCodMarca = CInt(Trim(Me.mshFlex.get_TextMatrix(nPosicion, C_COLCODIGO)))
                    '''se inicializa el dato dependiente de la marca
                    .mintCodModelo = 0
                    .dbcModelo.Text = ""
                    .mblnFueraChange = False
                End With
                Me.Close()
                Exit Function
            End If
        End With
        MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        mblnGuardar = False
        Guardar = True

        Nuevo()
        LlenaDatos()
        mshFlex.Focus()

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
        If blnTransaction Then Cnn.RollbackTrans()
    End Function

    Sub LlenaDatos()
        On Error GoTo Merr
        Dim I As Integer
        Dim nRow As Integer
        If Trim(Me.Tag) = "FRMCXPRELOJERIA" Then
            mblnFueraChange = True
            Me.lblGrupo.Text = Trim(C_RELOJERIA)
            Me.lblGrupo.Tag = Me.lblGrupo.Text
            mblnFueraChange = False
        End If
        With Me.mshFlex
            BuscarGrupo()
            gStrSql = "select LTrim(RTrim(DescMarca)) as DescMarca, DescMarca as DescripcionTag, CodMarca, '" & C_ACTIVO & "' as Estatus, '' as Depend, LTrim(RTrim(DescMarca)) as MarOriginal From CatMarcas Where CodGrupo = " & gCODRELOJERIA & " Order by DescMarca "
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
            .Row = 1
            .Col = C_ColDESCRIPCION
            .TopRow = 1
            If Trim(Me.Tag) <> "" Then
                ScrollGrid()
                .Focus()
            End If
        End With
        mblnNuevo = False
        mblnCambiosEnCodigo = False
        System.Windows.Forms.Application.DoEvents()

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub Encabezado()
        Dim LnContador As Integer
        With mshFlex
            .set_Cols(0, 6)
            .set_ColWidth(C_ColDESCRIPCION, 0, 5070)
            .set_ColWidth(C_ColDESCRIPCIONTAG, 0, 1)
            .set_ColWidth(C_COLCODIGO, 0, 1)
            .set_ColWidth(C_COLSTATUS, 0, 1)
            .set_ColWidth(C_COLDEPEND, 0, 1)
            .set_ColWidth(C_COLMARORIGINAL, 0, 1)

            .set_TextMatrix(0, C_ColDESCRIPCION, "Marcas")
            .set_TextMatrix(0, C_ColDESCRIPCIONTAG, "DescripcionTag")
            .set_TextMatrix(0, C_COLCODIGO, "Código")
            .set_TextMatrix(0, C_COLSTATUS, "STATUS")
            .set_TextMatrix(0, C_COLDEPEND, "DEPEND")
            .set_TextMatrix(0, C_COLMARORIGINAL, "MARORIGINAL")

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

    Private Sub frmCorpoABCMarca_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoABCMarca_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCorpoABCMarca_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                'mshFlex_KeyPress vbKeyReturn
            Case System.Windows.Forms.Keys.Escape
                'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    mblnSalir = True
                    Me.Close()
                    KeyCode = 0
                End If
            Case System.Windows.Forms.Keys.Delete
                'UPGRADE_ISSUE: Control Name could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    If Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_ColDESCRIPCION) <> "" Then
                        Call Eliminar()
                    End If
                End If
        End Select
    End Sub

    Private Sub frmCorpoABCMarca_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoABCMarca_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
        LlenaDatos()
    End Sub

    Private Sub frmCorpoABCMarca_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If Trim(Me.Tag) = "" Then
            If Not mblnSalir Then
                'Si desea cerrar la forma y ésta se encuentra minimizada, se debe restaurar
                ModEstandar.RestaurarForma(Me, False)
                If Cambios() Then
                    Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                        Case MsgBoxResult.Yes
                            If Not Guardar() Then 'Si falla el guardar, no cierra la forma
                                Cancel = 1
                            Else
                                mblnNuevo = True
                                Cancel = 0
                            End If
                        Case MsgBoxResult.No 'No hace nada y permite que se cierre el formulario
                            mblnNuevo = True
                            Cancel = 0
                        Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
                            Cancel = 1
                    End Select
                End If
            Else 'Se quiere salir con escape
                mblnSalir = False
                Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes 'Sale del Formulario
                        Cancel = 0
                    Case MsgBoxResult.No 'No sale del formulario
                        Me.mshFlex.Focus()
                        Cancel = 1
                End Select
            End If
        Else
            Cancel = 0
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoABCMarca_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Select Case Me.Tag
            Case "FRMCXPRELOJERIA"
                frmCXPRelojeria.Enabled = True
                frmCXPRelojeria.dbcMarca.Focus()
        End Select
        Me.Tag = ""
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub mshFlex_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.DblClick
        mshFlex_KeyPressEvent(mshFlex, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub mshFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.Enter
        Pon_Tool()
    End Sub

    Private Sub mshFlex_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles mshFlex.KeyDownEvent
        With Me.mshFlex
            Select Case eventArgs.keyCode
                Case System.Windows.Forms.Keys.Left
                    .Col = C_ColDESCRIPCION
                Case System.Windows.Forms.Keys.Right
                    .Col = C_ColDESCRIPCION
                Case System.Windows.Forms.Keys.Down
                    .Col = C_ColDESCRIPCION
            End Select
        End With
    End Sub

    Private Sub mshFlex_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles mshFlex.KeyPressEvent
        With mshFlex
            '''si ya se capturo algo entonces se edita el grid
            '''ya sea con numeros, letras o enter
            'If KeyAscii = 13 Then
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then
                If (.Col = C_ColDESCRIPCION) Then
                    '''en esta parte se validará si es el rengón, columna que le
                    '''corresponde editarse
                    If (.Row > 1) Then
                        '''de tal modo que si el renglón es mayor que 1
                        '''y si un renglón antes del renglón actual está vacío,
                        '''el renglón actual no se editará
                        If Trim(.get_TextMatrix(.Row - 1, C_ColDESCRIPCION)) = "" Then
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
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        With mshFlex
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    'txtFlex.Visible = False
                    'txtFlex.Text = ""
                    'mshFlex.Focus()
                Case System.Windows.Forms.Keys.Return
                    'If Trim(txtFlex.Text) = "" Then
                    '    Exit Sub
                    'End If
                    .set_TextMatrix(.Row, C_ColDESCRIPCION, Trim(txtFlex.Text))
                    If ArticuloRepetidoenGrid(Trim(.get_TextMatrix(.Row, C_ColDESCRIPCION)), "A") = True Then
                        MsgBox("Existe un artículo capturado con la misma descripción" & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                        LimpiaDatosArticulo(C_ColDESCRIPCION)
                        txtFlex.Visible = False
                        .Col = C_ColDESCRIPCION
                        mshFlex.Focus()
                        Exit Sub
                    End If
                    .set_TextMatrix(.Row, C_COLDEPEND, IIf(ReferenciaMar(.get_TextMatrix(.Row, C_COLCODIGO)), "S", "N"))
                    If .get_TextMatrix(.Row, C_COLSTATUS) = "" Then
                        .set_TextMatrix(.Row, C_COLSTATUS, C_NUEVO)
                        mblnNuevo = False
                        .Rows = .Rows + 1
                        ScrollGrid()
                    ElseIf .get_TextMatrix(.Row, C_COLSTATUS) <> "" Then
                        .set_TextMatrix(.Row, C_COLSTATUS, C_MODIFICADO)
                        mblnNuevo = False
                        .Row = .Row + 1
                    End If
                    If Not mblnGuardar Then
                        If .Enabled Then .Focus()
                    End If
                    txtFlex.Text = ""
                    txtFlex.Visible = False
            End Select
        End With
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Public Function ReferenciaMar(ByRef lMarca As String) As Boolean
        On Error GoTo Merr
        Dim rsLocal As ADODB.Recordset
        Dim lSql As String

        If Trim(lMarca) = "" Then Exit Function

        ReferenciaMar = False
        lSql = "Select * from CatArticulos(Nolock) Where CodGrupo = " & gCODRELOJERIA & " And CodMarca = " & lMarca
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then ReferenciaMar = True

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

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
                        If Trim(.get_TextMatrix(I, C_ColDESCRIPCION)) = "" Then Exit For
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
        Exit Function
    End Function

    Sub LimpiaDatosArticulo(ByRef lColumna As Integer)
        On Error GoTo Merr
        'Este Procedimiento Limpialos Campos Correspondientes a un Artículo, cuando se cambie de Articulo, que se limpien los datos
        With Me.mshFlex
            For I = lColumna To 5
                .set_TextMatrix(.Row, I, "")
            Next
            txtFlex.Text = ""
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub
End Class