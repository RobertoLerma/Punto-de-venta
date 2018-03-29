'**********************************************************************************************************************'
'*PROGRAMA: ABC DE LINEAS JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmCorpoABCLineas
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ' Programa :                ABC de Líneas de Artículos
    ' Autor :                   Paimí
    ' Fecha de Inicio:          13 de Mayo de 2003
    ' Fecha de Finalización:
    ' Nota:                     Si este cambia, debe cambiar también el de Modelos y viceversa
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents dbcGrupo As System.Windows.Forms.ComboBox
    Public WithEvents dbcDescFamilia As System.Windows.Forms.ComboBox
    Public WithEvents txtFlex As System.Windows.Forms.TextBox
    Public WithEvents mshFlex As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents _lblModelo_0 As System.Windows.Forms.Label
    Public WithEvents _lblModelo_1 As System.Windows.Forms.Label
    Public WithEvents fraMarcas As System.Windows.Forms.GroupBox
    Public WithEvents lblModelo As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Const C_RENENCABEZADO As Integer = 0
    Const C_ColDESCRIPCION As Integer = 0
    Const C_ColDESCRIPCIONTAG As Integer = 1
    Const C_COLCODIGO As Integer = 2
    Const C_COLSTATUS As Integer = 3
    Const C_COLDEPEND As Integer = 4
    Const C_COLLINORIGINAL As Integer = 5

    Dim rsLocal As ADODB.Recordset

    Dim mblnSalir As Boolean
    Dim mblnEscape As Boolean
    Dim mblnCambiosEnCodigo1 As Object
    Dim mblnCambiosEnCodigo2 As Boolean
    Dim mblnNuevo As Boolean
    Dim mintDepend As Integer

    'Variables para manejar el combo de Familia
    Dim Tecla1 As Integer
    Dim Tecla2 As Integer
    Dim mblnFueraChange As Boolean
    Public mintCodFamilia As Integer
    Public mintCodGrupo As Integer
    Dim I As Integer
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btnSalir As Button
    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnGuardar As Button
    Friend WithEvents btnLimpiar As Button
    Friend WithEvents btnEliminar As Button
    Dim mblnGuardar As Boolean


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCorpoABCLineas))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.fraMarcas = New System.Windows.Forms.GroupBox()
        Me.dbcGrupo = New System.Windows.Forms.ComboBox()
        Me.dbcDescFamilia = New System.Windows.Forms.ComboBox()
        Me.mshFlex = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._lblModelo_0 = New System.Windows.Forms.Label()
        Me._lblModelo_1 = New System.Windows.Forms.Label()
        Me.lblModelo = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.fraMarcas.SuspendLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblModelo, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.txtFlex.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtFlex, "Digite la Línea de Artículos")
        Me.txtFlex.Visible = False
        '
        'fraMarcas
        '
        Me.fraMarcas.BackColor = System.Drawing.Color.Silver
        Me.fraMarcas.Controls.Add(Me.dbcGrupo)
        Me.fraMarcas.Controls.Add(Me.dbcDescFamilia)
        Me.fraMarcas.Controls.Add(Me.txtFlex)
        Me.fraMarcas.Controls.Add(Me.mshFlex)
        Me.fraMarcas.Controls.Add(Me._lblModelo_0)
        Me.fraMarcas.Controls.Add(Me._lblModelo_1)
        Me.fraMarcas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMarcas.Location = New System.Drawing.Point(13, 14)
        Me.fraMarcas.Name = "fraMarcas"
        Me.fraMarcas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMarcas.Size = New System.Drawing.Size(393, 305)
        Me.fraMarcas.TabIndex = 0
        Me.fraMarcas.TabStop = False
        '
        'dbcGrupo
        '
        Me.dbcGrupo.Location = New System.Drawing.Point(80, 24)
        Me.dbcGrupo.Name = "dbcGrupo"
        Me.dbcGrupo.Size = New System.Drawing.Size(297, 21)
        Me.dbcGrupo.TabIndex = 2
        '
        'dbcDescFamilia
        '
        Me.dbcDescFamilia.Location = New System.Drawing.Point(80, 56)
        Me.dbcDescFamilia.Name = "dbcDescFamilia"
        Me.dbcDescFamilia.Size = New System.Drawing.Size(297, 21)
        Me.dbcDescFamilia.TabIndex = 4
        '
        'mshFlex
        '
        Me.mshFlex.DataSource = Nothing
        Me.mshFlex.Location = New System.Drawing.Point(19, 96)
        Me.mshFlex.Name = "mshFlex"
        Me.mshFlex.OcxState = CType(resources.GetObject("mshFlex.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mshFlex.Size = New System.Drawing.Size(359, 196)
        Me.mshFlex.TabIndex = 5
        '
        '_lblModelo_0
        '
        Me._lblModelo_0.AutoSize = True
        Me._lblModelo_0.BackColor = System.Drawing.Color.Silver
        Me._lblModelo_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblModelo_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblModelo_0.Location = New System.Drawing.Point(24, 28)
        Me._lblModelo_0.Name = "_lblModelo_0"
        Me._lblModelo_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblModelo_0.Size = New System.Drawing.Size(36, 13)
        Me._lblModelo_0.TabIndex = 1
        Me._lblModelo_0.Text = "Grupo"
        '
        '_lblModelo_1
        '
        Me._lblModelo_1.AutoSize = True
        Me._lblModelo_1.BackColor = System.Drawing.Color.Silver
        Me._lblModelo_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblModelo_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblModelo_1.Location = New System.Drawing.Point(24, 60)
        Me._lblModelo_1.Name = "_lblModelo_1"
        Me._lblModelo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblModelo_1.Size = New System.Drawing.Size(39, 13)
        Me._lblModelo_1.TabIndex = 3
        Me._lblModelo_1.Text = "Familia"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.fraMarcas)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(420, 410)
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
        Me.Panel3.Location = New System.Drawing.Point(13, 322)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(393, 74)
        Me.Panel3.TabIndex = 73
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
        'frmCorpoABCLineas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(444, 435)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(348, 115)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoABCLineas"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Líneas de Artículos"
        Me.fraMarcas.ResumeLayout(False)
        Me.fraMarcas.PerformLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblModelo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

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
            .set_TextMatrix(C_RENENCABEZADO, C_ColDESCRIPCION, "Líneas de Artículos")
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

    Public Sub Nuevo()
        If Not mblnNuevo Then
            mintCodGrupo = 0
            dbcGrupo.Text = ""
            dbcGrupo.Tag = ""
        End If
        mintCodFamilia = 0
        dbcDescFamilia.Text = ""
        dbcDescFamilia.Tag = ""
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
        Me.dbcGrupo.Focus()
    End Sub

    Public Function ValidaDatos() As Boolean
        On Error Resume Next
        Dim I As Object
        If mintCodGrupo = 0 Then
            MsgBox("Debe especificar el Grupo al que pertenece la Familia del Artículo.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            mblnNuevo = True
            Limpiar()
            ValidaDatos = False
            Exit Function
        End If
        If mintCodFamilia = 0 Then
            MsgBox("Debe especificar la Familia del Artículo.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            ValidaDatos = False
            Exit Function
        End If
        With mshFlex
            mintDepend = 0
            For I = 1 To .Rows - 1
                If IsNothing(.get_TextMatrix(I, C_COLSTATUS)) Then
                    Exit For
                End If
                If .get_TextMatrix(I, C_ColDESCRIPCION) = "" Then
                    MsgBox("Debe especificar la Línea del Artículo, o borrar el registro", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
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
                If Referencia("Select CodSubLinea From CatSubLineas where codGrupo = " & mintCodGrupo & " and codFamilia = " & mintCodFamilia & " and CodLinea = " & CShort(Numerico(mshFlex.get_TextMatrix(mshFlex.Row, C_COLCODIGO)))) Then
                    MsgBox("No es posible eliminar esta Línea" & vbNewLine & "debido a que tiene SubLíneas asociadas a ella", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                    Exit Sub
                End If
                If Referencia("Select codArticulo From CatArticulos Where codGrupo = " & mintCodGrupo & " and codFamilia = " & mintCodFamilia & " and codLinea = " & Numerico(Me.mshFlex.get_TextMatrix(mshFlex.Row, C_COLCODIGO))) Then
                    MsgBox("No es posible eliminar esta Línea" & vbNewLine & "debido a que está asociada" & vbNewLine & "con algunos artículos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                    Exit Sub
                End If

                If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA) <> MsgBoxResult.Yes Then
                    Exit Sub
                End If
                Cnn.BeginTrans()
                blnTransaction = True
                ModStoredProcedures.PR_IMECatLineas(Str(mintCodGrupo), Str(mintCodFamilia), Trim(Me.mshFlex.get_TextMatrix(mshFlex.Row, C_COLCODIGO)), Trim(mshFlex.get_TextMatrix(mshFlex.Row, C_ColDESCRIPCION)), C_ELIMINACION, CStr(0))
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
        'mshFlex.TopRow = 1
        'mshFlex.Row = 1
        mshFlex.Col = C_ColDESCRIPCION

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
        If blnTransaction Then Cnn.RollbackTrans()
    End Sub

    Function BuscarFlex() As Boolean
        On Error GoTo Merr
        gStrSql = "select * from CatLineas where codGrupo = " & mintCodGrupo & " and codFamilia = " & mintCodFamilia & " and codLinea = " & ModEstandar.Numerico(Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_COLCODIGO))
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

        'If Not ValidaDatos() Then
        '    Limpiar()
        '    Exit Function
        'End If

        If Not Cambios() Then
            Limpiar()
            Exit Function
        End If
        If mintDepend >= 1 Then
            MsgBox("Existen artículos clasificados con" & vbNewLine & "los datos que se modificaron" & vbNewLine & vbNewLine & "Estos serán reclasificados pero" & vbNewLine & "su descripción no será alterada" & vbNewLine & vbNewLine & "", MsgBoxStyle.Information, "AVISO")
        End If

        I = 0
        With mshFlex
            Cnn.BeginTrans()
            For I = 1 To (.Rows)
                If IsNothing(.get_TextMatrix(I, C_COLSTATUS)) Then
                    Exit For
                End If
                blnTransaction = True
                With mshFlex
                    Select Case .get_TextMatrix(I, C_COLSTATUS)
                        Case C_MODIFICADO
                            ModStoredProcedures.PR_IMECatLineas(Str(mintCodGrupo), Str(mintCodFamilia), .get_TextMatrix(I, C_COLCODIGO), Trim(.get_TextMatrix(I, C_ColDESCRIPCION)), C_MODIFICACION, CStr(0))
                            Cmd.Execute()
                            .set_TextMatrix(I, C_ColDESCRIPCIONTAG, .get_TextMatrix(I, C_ColDESCRIPCION))
                            .set_TextMatrix(I, C_COLSTATUS, C_ACTIVO)
                            nModif = nModif + 1
                            nPosicion = I
                        Case C_NUEVO
                            ModStoredProcedures.PR_IMECatLineas(Str(mintCodGrupo), Str(mintCodFamilia), .get_TextMatrix(I, C_COLCODIGO), Trim(.get_TextMatrix(I, C_ColDESCRIPCION)), C_INSERCION, CStr(0))
                            Cmd.Execute()
                            .set_TextMatrix(I, C_ColDESCRIPCIONTAG, .get_TextMatrix(I, C_ColDESCRIPCION))
                            .set_TextMatrix(I, C_COLCODIGO, Cmd.Parameters("ID").Value)
                            .set_TextMatrix(I, C_COLSTATUS, C_ACTIVO)
                            .Rows = .Rows + 1
                            nNuevos = nNuevos + 1
                            nPosicion = I
                        Case C_ELIMINADO
                            ModStoredProcedures.PR_IMECatLineas(Str(mintCodGrupo), Str(mintCodFamilia), .get_TextMatrix(I, C_COLCODIGO), Trim(.get_TextMatrix(I, C_ColDESCRIPCION)), C_ELIMINACION, CStr(0))
                            Cmd.Execute()
                    End Select
                End With
                blnTransaction = False
            Next I
            Cnn.CommitTrans()

            'If Trim(Me.Tag) = "FRMCXPJOYERIA" Then
            '    With frmCXPJoyeria
            '        .mblnFueraChange = True
            '        .dbcLinea.CtlText = Trim(Me.mshFlex.get_TextMatrix(nPosicion, C_ColDESCRIPCION))
            '        .dbcLinea.Tag = .dbcLinea.CtlText
            '        .mintCodLinea = CInt(Numerico(Me.mshFlex.get_TextMatrix(nPosicion, C_COLCODIGO)))
            '        .mblnFueraChange = False
            '    End With
            '    Me.Close()
            '    Exit Function
            'ElseIf Trim(Me.Tag) = "FRMCXPVARIOS" Then
            '    With frmCXPVarios
            '        .mblnFueraChange = True
            '        .dbcLinea.CtlText = Trim(Me.mshFlex.get_TextMatrix(nPosicion, C_ColDESCRIPCION))
            '        .dbcLinea.Tag = .dbcLinea.CtlText
            '        .mintCodLinea = CInt(Numerico(Me.mshFlex.get_TextMatrix(nPosicion, C_COLCODIGO)))
            '        .mblnFueraChange = False
            '    End With
            '    Me.Close()
            '    Exit Function
            'End If
        End With
        MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        mshFlex.Col = 0
        mshFlex.Row = 1
        mblnGuardar = False
        Guardar = True
        Nuevo()
        dbcGrupo.Focus()

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
            Me.dbcGrupo.Text = C_JOYERIA
            Me.dbcGrupo.Tag = C_JOYERIA
            Me.dbcGrupo.Text = True
            'Me.dbcDescFamilia.Text = Trim(frmCXPJoyeria.dbcFamilia.Text)
            Me.dbcDescFamilia.Tag = Me.dbcDescFamilia.Text
            Me.dbcDescFamilia.Text = True
            mblnFueraChange = False
        ElseIf Me.Tag = "FRMCXPVARIOS" Then
            mblnFueraChange = True
            Me.dbcGrupo.Text = C_VARIOS
            Me.dbcGrupo.Tag = C_VARIOS
            Me.dbcGrupo.Text = True
            'Me.dbcDescFamilia.Text = Trim(frmCXPVarios.dbcFamilia.Text)
            Me.dbcDescFamilia.Tag = Me.dbcDescFamilia.Text
            Me.dbcDescFamilia.Text = True
            mblnFueraChange = False
        End If
        With Me.mshFlex
            gStrSql = "select LTrim(RTrim(DescLinea)) as DescLinea, DescLinea as DescripcionTag, CodLinea, '" & C_ACTIVO & "' as Estatus, '' as Depend,  LTrim(RTrim(DescLinea)) as LinOriginal From CatLineas Where CodGrupo = " & mintCodGrupo & " and CodFamilia = " & mintCodFamilia & " Order by DescLinea "
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
            If Trim(Me.Tag) <> "" Then
                ScrollGrid()
                .Focus()
            End If
        End With
        mblnNuevo = False
        mblnCambiosEnCodigo1 = False
        mblnCambiosEnCodigo2 = False
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Sub Encabezado()
        Dim LnContador As Integer

        With mshFlex
            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusHeavy 'flexFocusLight 'flexFocusNone
            .WordWrap = False
            .FixedRows = 1
            .FixedCols = 0
            .set_Cols(0, 6)

            .set_ColWidth(C_ColDESCRIPCION, 0, 5070)
            .set_ColWidth(C_ColDESCRIPCIONTAG, 0, 1)
            .set_ColWidth(C_COLCODIGO, 0, 1)
            .set_ColWidth(C_COLSTATUS, 0, 1)
            .set_ColWidth(C_COLDEPEND, 0, 1)
            .set_ColWidth(C_COLLINORIGINAL, 0, 1)

            .set_TextMatrix(0, C_ColDESCRIPCION, "Líneas de Artículos")
            .set_TextMatrix(0, C_ColDESCRIPCIONTAG, "DescripcionTag")
            .set_TextMatrix(0, C_COLCODIGO, "Código")
            .set_TextMatrix(0, C_COLSTATUS, "STATUS")
            .set_TextMatrix(0, C_COLDEPEND, "DEPEND")
            .set_TextMatrix(0, C_COLLINORIGINAL, "LINORIGINAL")

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

    Private Sub dbcDescFamilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcDescFamilia.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codFamilia, rtrim(ltrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & mintCodGrupo & " and descFamilia LIKE '" & Trim(Me.dbcDescFamilia.Text) & "%' Order by DescFamilia "
        ModDCombo.DCChange(lStrSql, Tecla2, dbcDescFamilia)

        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox("¿Desea guardar los cambios?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    If Guardar() Then
                        mblnNuevo = True
                        Limpiar()
                    End If
                    Call dbcDescFamilia_Enter(dbcDescFamilia, New System.EventArgs())
                Case MsgBoxResult.No
                    mblnNuevo = True
                    Limpiar()
                Case MsgBoxResult.Cancel
            End Select
        End If
        If Me.dbcDescFamilia.Text = "" Then
            LimpiarFlex()
        End If
        If dbcDescFamilia.SelectedItem <> "" Then
            Call dbcDescFamilia_Leave(dbcDescFamilia, New System.EventArgs())
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcDescFamilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcDescFamilia.Enter
        Pon_Tool()
        gStrSql = "SELECT codFamilia, rtrim(ltrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & mintCodGrupo & " ORDER BY descFamilia "
        ModDCombo.DCGotFocus(gStrSql, dbcDescFamilia)
    End Sub

    Private Sub dbcDescFamilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcDescFamilia.KeyDown
        Dim Aux As String
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                mblnEscape = True
                ModEstandar.RetrocederTab(Me)
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
        gStrSql = "SELECT codFamilia, rtrim(ltrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & mintCodGrupo & " and descFamilia LIKE '" & Trim(Me.dbcDescFamilia.Text) & "%'"
        Aux = mintCodFamilia
        mintCodFamilia = 0
        ModDCombo.DCLostFocus(dbcDescFamilia, gStrSql, mintCodFamilia)
        If mintCodFamilia <> Aux Then
            mblnCambiosEnCodigo2 = True
            If Not mblnEscape Then
            Else
                If Not Cambios() Then
                    LlenaDatos()
                End If
            End If
        Else
            mblnCambiosEnCodigo2 = False
        End If
        If Not mblnEscape Then
            If Not Cambios() Then
                LlenaDatos()
            End If
        End If
        mblnEscape = False
    End Sub

    Private Sub dbcDescFamilia_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcDescFamilia.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcDescFamilia.Text)
        'If Me.dbcDescFamilia.SelectedItem <> 0 Then
        'dbcDescFamilia_Leave(dbcDescFamilia, New System.EventArgs())
        'End If
        Me.dbcDescFamilia.Text = Aux
    End Sub

    Private Sub dbcGrupo_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcGrupo.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        lStrSql = "SELECT codGrupo, ltrim(rtrim(descGrupo)) as descGrupo FROM catGrupos WHERE  codGrupo <> " & gCODRELOJERIA & " and descGrupo LIKE '" & Trim(Me.dbcGrupo.Text) & "%'"
        ModDCombo.DCChange(lStrSql, Tecla1, dbcGrupo)

        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox("¿Desea guardar los cambios?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    If Guardar() Then
                        mblnNuevo = True
                        mblnFueraChange = True
                        Me.dbcDescFamilia.Text = ""
                        Me.dbcDescFamilia.Tag = ""
                        mintCodFamilia = 0
                        mblnFueraChange = False
                        LimpiarFlex()
                        Me.dbcGrupo.Focus()
                        ModEstandar.SelTxt()
                    End If
                    Call dbcGrupo_Enter(dbcGrupo, New System.EventArgs())
                Case MsgBoxResult.No
                    mblnNuevo = True
                    mblnFueraChange = True
                    Me.dbcDescFamilia.Text = ""
                    Me.dbcDescFamilia.Tag = ""
                    mintCodFamilia = 0
                    mblnFueraChange = False
                    LimpiarFlex()
                    Me.dbcGrupo.Focus()
                    ModEstandar.SelTxt()
                Case MsgBoxResult.Cancel
            End Select
        End If

        If Me.dbcGrupo.Text = "" Then
            mblnFueraChange = True
            Me.dbcDescFamilia.Text = ""
            Me.dbcDescFamilia.Tag = ""
            mintCodFamilia = 0
            mblnFueraChange = False
            LimpiarFlex()
        End If

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcGrupo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcGrupo.Enter
        Pon_Tool()
        gStrSql = "SELECT codGrupo, ltrim(rtrim(descGrupo)) as descGrupo FROM catGrupos WHERE codGrupo <> " & gCODRELOJERIA & " ORDER BY DescGrupo"
        ModDCombo.DCGotFocus(gStrSql, dbcGrupo)
    End Sub

    Private Sub dbcGrupo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcGrupo.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            eventSender.KeyCode = 0
        End If
        Tecla1 = eventArgs.KeyCode
    End Sub

    Private Sub dbcGrupo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcGrupo.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codGrupo, ltrim(rtrim(descGrupo)) as descGrupo FROM catGrupos WHERE  codGrupo <> " & gCODRELOJERIA & " and descGrupo LIKE '" & Trim(Me.dbcGrupo.Text) & "%'"
        Aux = mintCodGrupo
        mintCodGrupo = 0
        ModDCombo.DCLostFocus(dbcGrupo, gStrSql, mintCodGrupo)
        If mintCodGrupo <> Aux Then
            mblnFueraChange = True
            Me.dbcDescFamilia.Text = ""
            Me.dbcDescFamilia.Tag = ""
            mintCodFamilia = 0
            LimpiarFlex()
            mblnFueraChange = False
        End If
    End Sub

    Private Sub frmCorpoABCLineas_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoABCLineas_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCorpoABCLineas_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    Me.dbcDescFamilia.Focus()
                End If
            Case System.Windows.Forms.Keys.Delete
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    If Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_ColDESCRIPCION) <> "" Then
                        Eliminar()
                    End If
                End If
        End Select
    End Sub

    Private Sub frmCorpoABCLineas_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoABCLineas_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        mintCodGrupo = 0
        mintCodFamilia = 0
        LlenaDatos()
    End Sub

    Private Sub frmCorpoABCLineas_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Trim(Me.Tag) = "" Then 'Si el formulario NO fue llamado desde CxP, sale de la manera habitual
        '    If Not mblnSalir Then
        '        'Si desea cerrar la forma y ésta se encuentra minimizada, se debe restaurar
        '        ModEstandar.RestaurarForma(Me, False)
        '        If Cambios() Then ' And Not mblnNuevo
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
        '                Me.dbcGrupo.Focus()
        '                ModEstandar.SelTxt()
        '                Cancel = 1
        '        End Select
        '    End If
        'Else
        '    Cancel = 0
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoABCLineas_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Select Case Me.Tag
            Case "FRMCXPJOYERIA"
                'frmCXPJoyeria.Enabled = True
                'frmCXPJoyeria.dbcLinea.Focus()
            Case "FRMCXPVARIOS"
                'frmCXPVarios.Enabled = True
                'frmCXPVarios.dbcLinea.Focus()
        End Select
        Me.Tag = ""
        ' Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub mshFlex_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.DblClick
        mshFlex_KeyPressEvent(mshFlex, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent((System.Windows.Forms.Keys.Return)))
    End Sub

    Private Sub mshFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.Enter
        Pon_Tool()
    End Sub

    Private Sub mshFlex_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles mshFlex.KeyDownEvent
        If mintCodGrupo = 0 Or mintCodFamilia = 0 Then
            eventArgs.keyCode = 0
        End If
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
        If mintCodGrupo = 0 Or mintCodFamilia = 0 Then
            eventArgs.keyAscii = 0
        End If
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
                    If Trim(txtFlex.Text) = "" Then
                        Exit Sub
                    End If
                    .set_TextMatrix(.Row, C_ColDESCRIPCION, Trim(txtFlex.Text))
                    If ArticuloRepetidoenGrid(Trim(.get_TextMatrix(.Row, C_ColDESCRIPCION)), "A") = True Then
                        MsgBox("Existe un artículo capturado con la misma descripción" & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
                        LimpiaDatosArticulo(C_ColDESCRIPCION)
                        txtFlex.Visible = False
                        .Col = C_ColDESCRIPCION
                        mshFlex.Focus()
                        Exit Sub
                    End If
                    .set_TextMatrix(.Row, C_COLDEPEND, IIf(ReferenciaLin(.get_TextMatrix(.Row, C_COLCODIGO)), "S", "N"))
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

    Public Function ReferenciaLin(ByRef lLinea As String) As Boolean
        On Error GoTo Merr
        Dim rsLocal As ADODB.Recordset
        Dim lSql As String

        If Trim(lLinea) = "" Then Exit Function

        ReferenciaLin = False
        lSql = "Select * from CatArticulos(Nolock) Where CodGrupo = " & mintCodGrupo & " And CodFamilia = " & mintCodFamilia & " And CodLinea = " & lLinea
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then ReferenciaLin = True

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

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        Eliminar()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub
End Class