'**********************************************************************************************************************'
'*PROGRAMA: ABC DE COMISIONES JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmCorpoABCComisiones

    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public WithEvents ToolTip1 As ToolTip
    Public WithEvents mshFlex As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents txtFlex As TextBox
    Friend WithEvents dtpMes As DateTimePicker
    Friend WithEvents lblComision As Label


    Const C_MOD As String = "M"
    Const C_NVO As String = "N"

    Const C_RENENCABEZADO As Integer = 0

    Const C_COLMES As Integer = 0
    Const C_COLPORCCOMISION As Integer = 1
    Const C_COLMESANIO As Integer = 2
    Const C_COLMESANIOTAG As Integer = 3
    Const C_COLPORCCOMISIONTAG As Integer = 4
    Const C_COLESTATUS As Integer = 5

    Dim mblnTecleoFecha As Boolean
    Dim msglTiempoCambio As Single
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btnSalir As Button
    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnGuardar As Button
    Friend WithEvents btnLimpiar As Button
    Friend WithEvents btnEliminar As Button
    Dim mblnLoad As Boolean



    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCorpoABCComisiones))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.mshFlex = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.txtFlex = New System.Windows.Forms.TextBox()
        Me.dtpMes = New System.Windows.Forms.DateTimePicker()
        Me.lblComision = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'mshFlex
        '
        Me.mshFlex.DataSource = Nothing
        Me.mshFlex.Location = New System.Drawing.Point(16, 28)
        Me.mshFlex.Name = "mshFlex"
        Me.mshFlex.OcxState = CType(resources.GetObject("mshFlex.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mshFlex.Size = New System.Drawing.Size(223, 230)
        Me.mshFlex.TabIndex = 2
        '
        'txtFlex
        '
        Me.txtFlex.AcceptsReturn = True
        Me.txtFlex.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlex.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlex.Location = New System.Drawing.Point(101, 317)
        Me.txtFlex.MaxLength = 0
        Me.txtFlex.Name = "txtFlex"
        Me.txtFlex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlex.Size = New System.Drawing.Size(81, 20)
        Me.txtFlex.TabIndex = 3
        Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtFlex.Visible = False
        '
        'dtpMes
        '
        Me.dtpMes.Location = New System.Drawing.Point(37, 281)
        Me.dtpMes.Name = "dtpMes"
        Me.dtpMes.Size = New System.Drawing.Size(193, 20)
        Me.dtpMes.TabIndex = 4
        '
        'lblComision
        '
        Me.lblComision.AutoSize = True
        Me.lblComision.Location = New System.Drawing.Point(13, 12)
        Me.lblComision.Name = "lblComision"
        Me.lblComision.Size = New System.Drawing.Size(89, 13)
        Me.lblComision.TabIndex = 5
        Me.lblComision.Text = "Comisión por mes"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(284, 454)
        Me.Panel1.TabIndex = 6
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(12, 369)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(260, 74)
        Me.Panel3.TabIndex = 72
        '
        'btnSalir
        '
        Me.btnSalir.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.salir
        Me.btnSalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSalir.Location = New System.Drawing.Point(203, 14)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(50, 42)
        Me.btnSalir.TabIndex = 70
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.buscar
        Me.btnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBuscar.Location = New System.Drawing.Point(155, 14)
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
        Me.btnGuardar.Location = New System.Drawing.Point(6, 14)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(50, 42)
        Me.btnGuardar.TabIndex = 64
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.nuevo
        Me.btnLimpiar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnLimpiar.Location = New System.Drawing.Point(105, 14)
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
        Me.btnEliminar.Location = New System.Drawing.Point(56, 14)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(50, 42)
        Me.btnEliminar.TabIndex = 65
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Silver
        Me.Panel2.Controls.Add(Me.lblComision)
        Me.Panel2.Controls.Add(Me.dtpMes)
        Me.Panel2.Controls.Add(Me.mshFlex)
        Me.Panel2.Controls.Add(Me.txtFlex)
        Me.Panel2.Location = New System.Drawing.Point(12, 13)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(260, 350)
        Me.Panel2.TabIndex = 0
        '
        'frmCorpoABCComisiones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(308, 479)
        Me.Controls.Add(Me.Panel1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmCorpoABCComisiones"
        Me.Text = "frmCorpoABCComisiones"
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
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
                If Trim(.get_TextMatrix(I, C_COLMES)) <> "" Then
                    nCont = nCont + 1
                Else
                    Exit For
                End If
            Next I
            If nCont < 7 Then
                'Hay menos de 7 registros
                .Row = nCont + 1
                .Col = C_COLMES
            Else
                'Hay 7 ó más registros, hay que recorrer el grid
                .TopRow = (nCont - nRen) + 2
                .Row = nCont + 1
                .Col = C_COLMES
            End If
        End With
    End Sub

    Public Sub Encabezado()
        On Error GoTo Merr
        Dim LnContador As Integer
        Dim I As Integer

        With Me.mshFlex
            If Not mblnLoad Then
                .Rows = 2
                .Rows = 12
                .set_Cols(0, 6)
                .RemoveItem((1))
                Exit Sub
            End If
            .set_Cols(0, 6)
            .Clear()

            .set_ColWidth(C_COLMES, 0, 1500)
            .set_ColAlignment(C_COLMES, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColWidth(C_COLPORCCOMISION, 0, 1250)
            .set_ColAlignment(C_COLPORCCOMISION, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColWidth(C_COLMESANIO, 0, 0)
            .set_ColWidth(C_COLMESANIOTAG, 0, 0)
            .set_ColWidth(C_COLPORCCOMISIONTAG, 0, 0)
            .set_ColWidth(C_COLESTATUS, 0, 0)

            .set_TextMatrix(C_RENENCABEZADO, C_COLMES, "Mes - Año")
            .set_TextMatrix(C_RENENCABEZADO, C_COLPORCCOMISION, "Comisión (%)")

            .Row = C_RENENCABEZADO
            .Col = C_COLMES
            .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
            .CellFontBold = True
            .Col = C_COLPORCCOMISION
            .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
            .CellFontBold = True
            .Rows = 16
            .Col = 0
            .Row = 2
            '''.TopRow = 1
        End With
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub LlenaDatos()
        On Error GoTo Merr
        Dim I As Integer
        gStrSql = "Select FechaPeriodo, PorcComision from CatComisionXVendedor order by FechaPeriodo"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            With Me.mshFlex
                If RsGral.RecordCount > 8 Then
                    .Rows = .Rows + 4
                Else
                    .Rows = 16
                End If
                RsGral.MoveFirst()
                For I = 1 To RsGral.RecordCount
                    .set_TextMatrix(I, C_COLMES, Format(RsGral.Fields("FechaPeriodo").Value, "MMM - yyyy"))
                    .set_TextMatrix(I, C_COLPORCCOMISION, Format(RsGral.Fields("PorcComision").Value, "##0.0"))
                    .set_TextMatrix(I, C_COLMESANIO, Format(RsGral.Fields("FechaPeriodo").Value, C_FORMATFECHAMOSTRAR))
                    .set_TextMatrix(I, C_COLMESANIOTAG, .get_TextMatrix(I, C_COLMESANIO))
                    .set_TextMatrix(I, C_COLPORCCOMISIONTAG, .get_TextMatrix(I, C_COLPORCCOMISION))
                    .set_TextMatrix(I, C_COLESTATUS, "")
                    RsGral.MoveNext()
                Next I
                .Rows = RsGral.RecordCount + 2
                .ScrollBars = MSHierarchicalFlexGridLib.ScrollBarsSettings.flexScrollBarVertical
            End With
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function Cambios() As Boolean
        On Error Resume Next
        Dim I As Integer
        With Me.mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLMES)) = "" Then
                    Exit For
                End If
                Select Case True
                    Case Trim(.get_TextMatrix(I, C_COLMESANIO)) <> Trim(.get_TextMatrix(I, C_COLMESANIOTAG))
                        If Trim(.get_TextMatrix(I, C_COLESTATUS)) <> C_NVO Then
                            .set_TextMatrix(I, C_COLESTATUS, C_MOD)
                        End If
                        Cambios = True
                    Case CShort(.get_TextMatrix(I, C_COLPORCCOMISION)) <> CShort(.get_TextMatrix(I, C_COLPORCCOMISIONTAG))
                        If Trim(.get_TextMatrix(I, C_COLESTATUS)) <> C_NVO Then
                            .set_TextMatrix(I, C_COLESTATUS, C_MOD)
                        End If
                        Cambios = True
                End Select
            Next I
        End With
    End Function

    Public Function ValidaDatos() As Boolean
        On Error Resume Next
        Dim I As Integer
        With Me.mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLMES)) = "" Then
                    Exit For
                End If
                Select Case True

                End Select
            Next I
        End With
        ValidaDatos = True
    End Function

    Public Sub Limpiar()
        ScrollGrid()
    End Sub

    Public Function Guardar() As Boolean
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        Dim I As Integer

        If Not Cambios() Then
            Exit Function
        End If
        Cnn.BeginTrans()
        blnTransaction = True
        With Me.mshFlex
            For I = 1 To .Rows - 1
                If Trim(.get_TextMatrix(I, C_COLMES)) = "" Then Exit For

                If Trim(.get_TextMatrix(I, C_COLESTATUS)) = C_NVO Then
                    ModStoredProcedures.PR_IMECatComisiones(Format(CDate(.get_TextMatrix(I, C_COLMESANIO)), C_FORMATFECHAGUARDAR), Trim(.get_TextMatrix(I, C_COLPORCCOMISION)), C_INSERCION, CStr(0))
                    Cmd.Execute()
                    .set_TextMatrix(I, C_COLMESANIOTAG, .get_TextMatrix(I, C_COLMESANIO))
                    .set_TextMatrix(I, C_COLPORCCOMISIONTAG, .get_TextMatrix(I, C_COLPORCCOMISION))
                    .set_TextMatrix(I, C_COLESTATUS, "")
                ElseIf Trim(.get_TextMatrix(I, C_COLESTATUS)) = C_MOD Then
                    ModStoredProcedures.PR_IMECatComisiones(Format(CDate(.get_TextMatrix(I, C_COLMESANIO)), C_FORMATFECHAGUARDAR), Trim(.get_TextMatrix(I, C_COLPORCCOMISION)), C_MODIFICACION, CStr(0))
                    Cmd.Execute()
                    .set_TextMatrix(I, C_COLMESANIOTAG, .get_TextMatrix(I, C_COLMESANIO))
                    .set_TextMatrix(I, C_COLPORCCOMISIONTAG, .get_TextMatrix(I, C_COLPORCCOMISION))
                    .set_TextMatrix(I, C_COLESTATUS, "")
                End If
            Next I
        End With
        Cnn.CommitTrans()
        blnTransaction = False
        MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        LlenaDatos()
        mshFlex.Col = 0
        mshFlex.Row = 1
        mshFlex.TopRow = 1
        mshFlex.Focus()
        Guardar = True
Merr:
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function Eliminar() As Boolean
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        With Me.mshFlex
            'Si es un registro que aún no ha sido guardado, sólo quita la línea del grid y añade una en blanco
            If Trim(.get_TextMatrix(.Row, C_COLMES)) <> "" Then
                If Trim(.get_TextMatrix(.Row, C_COLMESANIOTAG)) = "" Or Trim(.get_TextMatrix(.Row, C_COLESTATUS)) = C_NVO Then
                    If MsgBox(C_msgBORRAR, MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) = MsgBoxResult.Yes Then
                        .RemoveItem(.Row)
                    End If
                    .Focus()
                    Exit Function
                End If
            Else
                MsgBox("Debe seleccionar un registro válido para borrarlo", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                .Focus()
                Exit Function
            End If
            If MsgBox(C_msgBORRAR, MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) = MsgBoxResult.Yes Then
                Cnn.BeginTrans()
                blnTransaction = True

                ModStoredProcedures.PR_IMECatComisiones(Format(CDate(.get_TextMatrix(.Row, C_COLMESANIO)), C_FORMATFECHAGUARDAR), Trim(.get_TextMatrix(.Row, C_COLPORCCOMISION)), C_ELIMINACION, CStr(0))
                Cmd.Execute()

                Cnn.CommitTrans()
                blnTransaction = False
                .RemoveItem(.Row)
                MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                .Rows = .Rows + 1
                .Focus()
                Eliminar = True
            Else
                .Focus()
            End If
        End With

Merr:
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Function

    Private Sub dtpMes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpMes.Enter
        Dim x As Integer
        x = 0
    End Sub

    Private Sub dtpMes_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpMes.KeyDown
        Dim cFechaTmp As String
        Dim I As Integer
        With Me.mshFlex
            Select Case eventArgs.KeyCode
                Case System.Windows.Forms.Keys.Escape
                    Me.dtpMes.Value = Format(Today, "MMM/yyyy")
                    Me.dtpMes.Visible = False
                Case System.Windows.Forms.Keys.Return
                    If mblnTecleoFecha Then
                        dtpMes.Refresh()
                        '''Do While (Timer - msglTiempoCambio) <= 2.1
                        '''Loop
                        mblnTecleoFecha = False
                    End If
                    System.Windows.Forms.Application.DoEvents()
                    'Buscar si ya existe el registro que quiere añadir
                    cFechaTmp = Format(dtpMes.Value, "MMM/yyyy")
                    For I = 1 To .Rows - 1
                        If Trim(cFechaTmp) = Format(Trim(.get_TextMatrix(I, C_COLMES)), "MMM/yyyy") Then
                            MsgBox("El período que quiere introducir ya existe", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            Me.dtpMes.Focus()
                            Exit Sub
                        End If
                    Next I
                    'Si no existe, es un registro nuevo, y debe cambiar el estatus en el grid a C_NVO
                    .set_TextMatrix(.Row, C_COLMES, Trim(cFechaTmp))
                    .set_TextMatrix(.Row, C_COLPORCCOMISION, "0")
                    .set_TextMatrix(.Row, C_COLMESANIO, Format(Me.dtpMes.Value, C_FORMATFECHAMOSTRAR))
                    .set_TextMatrix(.Row, C_COLMESANIOTAG, "")
                    .set_TextMatrix(.Row, C_COLPORCCOMISIONTAG, "")
                    .set_TextMatrix(.Row, C_COLESTATUS, C_NVO)
                    Me.dtpMes.Value = Format(Today, "MMM/yyyy")
                    Me.dtpMes.Visible = False
                    .Col = C_COLPORCCOMISION
                    .Rows = .Rows + 1
                    .Focus()
            End Select
        End With
    End Sub

    Private Sub dtpMes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpMes.KeyPress
        mblnTecleoFecha = True
        'msglTiempoCambio = Timer()
    End Sub

    Private Sub dtpMes_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpMes.Leave
        dtpMes_KeyDown(dtpMes, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape))
    End Sub

    Private Sub frmCorpoABCComisiones_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCorpoABCComisiones_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCorpoABCComisiones_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    Me.Close()
                End If
            Case System.Windows.Forms.Keys.Delete
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    If Me.mshFlex.get_TextMatrix(Me.mshFlex.Row, C_COLMES) <> "" Then
                        Call Eliminar()
                    End If
                End If
        End Select
    End Sub

    Private Sub frmCorpoABCComisiones_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoABCComisiones_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        mblnLoad = True
        Encabezado()
        LlenaDatos()
        mblnLoad = False
        dtpMes.Visible = False
    End Sub

    Private Sub frmCorpoABCComisiones_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        ModEstandar.RestaurarForma(Me, False)
        If Cambios() Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    If Not (Guardar()) Then
                        Cancel = 1
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se cierre el formulario
                    Cancel = 0
                Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin Guardar
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoABCComisiones_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub mshFlex_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.DblClick
        mshFlex_KeyPressEvent(mshFlex, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub mshFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.Enter
        Pon_Tool()
    End Sub

    Private Sub mshFlex_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles mshFlex.KeyPressEvent
        Dim nCol As Integer
        Dim nRow As Integer
        With mshFlex
            '''If KeyAscii = 13 Then
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then
                nCol = .Col
                nRow = .Row
                'Si el grid ya tiene un valor en período, no debe editarlo debido a que estará editando la clave principal
                If .Col = C_COLPORCCOMISION Then
                    If Trim(.get_TextMatrix(.Row, C_COLMES)) = "" Then
                        MsgBox("Primero debe introducir un período válido", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        .Col = C_COLMES
                        .Focus()
                        Exit Sub
                    End If
                End If
                Select Case .Col
                    Case C_COLMES
                        dtpMes.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(mshFlex.Left) - 25 + mshFlex.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(mshFlex.Top) - 25 + mshFlex.CellTop), VB6.TwipsToPixelsX(mshFlex.CellWidth + 15), 0, System.Windows.Forms.BoundsSpecified.X Or System.Windows.Forms.BoundsSpecified.Y Or System.Windows.Forms.BoundsSpecified.Width)
                        If .get_TextMatrix(.Row, C_COLMES) <> "" Then
                            dtpMes.Value = Format(CDate("01/" & Trim(.get_TextMatrix(.Row, C_COLMES))), "MMM/yyyy")
                        Else
                            dtpMes.Value = Format(Today, "MMM/yyyy")
                        End If
                        dtpMes.Visible = True
                        '''SendKeys "{Right}", False
                        dtpMes.Focus()
                    Case C_COLPORCCOMISION
                        txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                        txtFlex.BackColor = .CellBackColor
                        eventArgs.keyAscii = ModEstandar.MskCantidad((Me.txtFlex.Text), eventArgs.keyAscii, 3, 1, (Me.txtFlex.SelectionStart))
                        'ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, eventArgs.keyAscii)
                        txtFlex.SelectionStart = Len(txtFlex.Text)
                End Select
            ElseIf eventArgs.keyAscii = 27 Then
            Else
                If (.Row > 1) Then
                    '''de tal modo que si el renglón es mayor que 1
                    '''y si un renglón antes del renglón actual está vacío,
                    '''el renglón actual no se editará
                    If Trim(.get_TextMatrix(.Row - 1, C_COLMES)) = "" Then
                        .Focus()
                        Exit Sub
                    End If
                End If

                Select Case .Col
                    Case C_COLMES
                        'Sólo debe editarse con <ENTER>
                        dtpMes.Focus()
                        '''.SetFocus
                        Exit Sub
                    Case C_COLPORCCOMISION
                        If Trim(.get_TextMatrix(.Row, C_COLMES)) <> "" Then
                            Me.txtFlex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right 'Alinear a la derecha
                            Me.txtFlex.BackColor = .CellBackColor

                            eventArgs.keyAscii = ModEstandar.MskCantidad((Me.txtFlex.Text), eventArgs.keyAscii, 3, 1, (Me.txtFlex.SelectionStart))

                            ModEstandar.MSHFlexGridEdit(mshFlex, txtFlex, eventArgs.keyAscii)
                            If Len(Me.txtFlex.Text) <> 1 Then
                                ModEstandar.SelTextoTxt((Me.txtFlex))
                            End If
                        Else
                            .Focus()
                        End If
                End Select
            End If
        End With
    End Sub

    Private Sub txtFlex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFlex.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Dim nCol As Object
        Dim nRen As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        With mshFlex
            nCol = .Col
            nRen = .Row
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    txtFlex.Text = ""
                    txtFlex.Visible = False
                Case System.Windows.Forms.Keys.Return
                    'Validar que se haya tecleado un porcentaje válido
                    If CShort(Numerico((Me.txtFlex.Text))) >= 0 And CShort(Numerico((Me.txtFlex.Text))) <= 99 Then
                        .set_TextMatrix(.Row, C_COLPORCCOMISION, Format(txtFlex.Text, "##0.0"))
                        If Trim(.get_TextMatrix(.Row, C_COLESTATUS)) = "" Then .set_TextMatrix(.Row, C_COLESTATUS, C_MOD)
                    Else
                        MsgBox("El porcentaje indicado no es válido", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        Me.txtFlex.Text = CStr(0)
                        Me.txtFlex.Focus()
                        ModEstandar.SelTxt()
                        Exit Sub
                    End If
                    Me.txtFlex.Text = ""
                    Me.txtFlex.Visible = False
                    .Col = C_COLMES
                    If .Row <= .Rows Then
                        .Row = .Row + 1
                    End If
                    .Focus()
            End Select
        End With
    End Sub

    Private Sub txtFlex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlex.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Me.txtFlex.Text = Format(Numerico((Me.txtFlex.Text)), "###,###,##0.0")
        End If
        KeyAscii = ModEstandar.MskCantidad((Me.txtFlex.Text), KeyAscii, 3, 1, (Me.txtFlex.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlex_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlex.Leave
        txtFlex_KeyDown(txtFlex, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub


End Class