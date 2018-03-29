Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.Compatibility
Imports ADODB
Public Class frmCXPConsultaOrden
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents _chkTipoConsulta_3 As System.Windows.Forms.CheckBox
    Public WithEvents _chkTipoConsulta_2 As System.Windows.Forms.CheckBox
    Public WithEvents _chkTipoConsulta_1 As System.Windows.Forms.CheckBox
    Public WithEvents _chkTipoConsulta_0 As System.Windows.Forms.CheckBox
    Public WithEvents _fraTipoConsulta_0 As System.Windows.Forms.GroupBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents mshFlex As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents _lblConsulta_0 As System.Windows.Forms.Label
    Public WithEvents chkTipoConsulta As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
    Public WithEvents fraTipoConsulta As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblConsulta As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Const C_RENENCABEZADO As Integer = 0

    Const C_COLFOLIO As Integer = 0
    Const C_COLFECHAORDEN As Integer = 1
    Const C_COLESTATUS As Integer = 2
    Const C_COLIMPORTE As Integer = 3

    Public nPROV As Integer
    Public cFORM As String
    Public lDESC As Boolean 'Para indicar si se van a desplegar en orden Descendente

    Dim mblnSalir As Boolean

    Dim mblnFueraChange As Boolean
    Dim Tecla As Integer
    Dim mintCodProveedor As Integer

    Public Function BuscaNombreProveedor(ByRef Codigo As Integer) As String
        On Error GoTo MErr
        gStrSql = "SELECT DescProvAcreed FROM CatProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' and codProvAcreed = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            BuscaNombreProveedor = Trim(RsGral.Fields("DescProvACreed").Value)
        Else
            BuscaNombreProveedor = ""
        End If
MErr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Sub LlenaDatos()
        On Error GoTo MErr
        Dim i As Integer
        Dim lStrSql As String
        Dim rsLocal As ADODB.Recordset
        Dim cWHERE As String 'Variable donde se va concatenando el where

        lDESC = True
        If nPROV > 0 Then
            mblnFueraChange = True
            mintCodProveedor = nPROV
            Me.dbcProveedor.Text = BuscaNombreProveedor(mintCodProveedor)
            Me.dbcProveedor.Tag = Me.dbcProveedor.Text
            mblnFueraChange = False
        End If

        lStrSql = "SELECT FolioOrdenCompra, FechaOrdenCompra, dbo.EstatusStr(Estatus) as Estatus, Total FROM OrdenesCompra"
        cWHERE = ""

        Select Case True
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(3).CheckState = System.Windows.Forms.CheckState.Checked
                'Activas - Terminadas - Canceladas - Registradas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STVIGENTE & "' or Estatus = '" & C_STGENERADA & "' or Estatus = '" & C_STCANCELADA & "' or Estatus = '" & C_STREGISTRADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(3).CheckState = System.Windows.Forms.CheckState.Checked
                'Activas - Terminadas - Registradas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STVIGENTE & "' or Estatus = '" & C_STGENERADA & "' or Estatus = '" & C_STREGISTRADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(3).CheckState = System.Windows.Forms.CheckState.Checked
                'Activas - Canceladas - Registradas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STVIGENTE & "' or Estatus = '" & C_STCANCELADA & "' or Estatus = '" & C_STREGISTRADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(3).CheckState = System.Windows.Forms.CheckState.Checked
                'Terminadas - Canceladas - Registradas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STGENERADA & "' or Estatus = '" & C_STCANCELADA & "' or Estatus = '" & C_STREGISTRADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(3).CheckState = System.Windows.Forms.CheckState.Checked
                'Activas - Registradas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STVIGENTE & "' or Estatus = '" & C_STREGISTRADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(3).CheckState = System.Windows.Forms.CheckState.Checked
                'Terminadas - Registradas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STGENERADA & "' or Estatus = '" & C_STREGISTRADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(3).CheckState = System.Windows.Forms.CheckState.Checked
                'Canceladas - Registradas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STCANCELADA & "' or Estatus = '" & C_STREGISTRADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(3).CheckState = System.Windows.Forms.CheckState.Checked
                'Registradas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STREGISTRADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Checked
                'Activas - Terminadas - Canceladas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STVIGENTE & "' or Estatus = '" & C_STGENERADA & "' or Estatus = '" & C_STCANCELADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                'Activas - Terminadas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STVIGENTE & "' or Estatus = '" & C_STGENERADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                'Activas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STVIGENTE & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                'Ninguna
                Call Encabezado()
                Exit Sub
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                'Terminadas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STGENERADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Checked
                'Terminadas - Canceladas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STGENERADA & "' or Estatus = '" & C_STCANCELADA & "')"
            Case Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Checked
                'Canceladas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STCANCELADA & "')"
            Case Else
                'Me.chkTipoConsulta(0).Value = vbChecked And Me.chkTipoConsulta(1).Value = vbUnchecked And Me.chkTipoConsulta(2).Value = vbChecked
                'Canceladas - Activas
                cWHERE = cWHERE & " WHERE (Estatus = '" & C_STCANCELADA & "' or Estatus = '" & C_STVIGENTE & "')"
        End Select

        If mintCodProveedor <> 0 Then
            cWHERE = cWHERE & " and CodProvAcreed = " & mintCodProveedor
        End If

        If Trim(Me.Tag) <> "" Then
            cWHERE = cWHERE & " and FolioOrdenCompra LIKE '" & Trim(Me.Tag) & "%'"
        End If

        If lDESC Then
            cWHERE = cWHERE & " ORDER BY FolioOrdenCompra DESC "
        End If

        lStrSql = lStrSql & cWHERE


        Call Encabezado()

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            If rsLocal.RecordCount < 11 Then
                Me.mshFlex.Rows = 11
            Else
                Me.mshFlex.Rows = rsLocal.RecordCount + 3
            End If
            'Llena el Grid
            With Me.mshFlex
                rsLocal.MoveFirst()
                For i = 1 To rsLocal.RecordCount
                    .set_TextMatrix(i, C_COLFOLIO, Trim(rsLocal.Fields("FolioOrdenCompra").Value))
                    .set_TextMatrix(i, C_COLFECHAORDEN, VB6.Format(rsLocal.Fields("FechaOrdenCompra").Value, "dd/MMM/yyyy"))
                    .set_TextMatrix(i, C_COLESTATUS, Trim(rsLocal.Fields("Estatus").Value))
                    .set_TextMatrix(i, C_COLIMPORTE, VB6.Format(rsLocal.Fields("Total").Value, "###,###,##0.00"))
                    rsLocal.MoveNext()
                Next i
            End With
        End If
MErr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Sub Encabezado()
        Dim LnContador As Integer

        With mshFlex
            .Clear()
            .Height = VB6.TwipsToPixelsY(2630)
            .set_ColWidth(C_COLFOLIO, 0, 2400)
            .set_ColWidth(C_COLFECHAORDEN, 0, 1230)
            .set_ColWidth(C_COLESTATUS, 0, 1360)
            .set_ColWidth(C_COLIMPORTE, 0, 1850)

            .set_TextMatrix(C_RENENCABEZADO, C_COLFOLIO, "Folio de Orden")
            .set_TextMatrix(C_RENENCABEZADO, C_COLFECHAORDEN, "Fecha")
            .set_TextMatrix(C_RENENCABEZADO, C_COLESTATUS, "Estatus")
            .set_TextMatrix(C_RENENCABEZADO, C_COLIMPORTE, "Importe")

            'Colocar los textos de los encabezados centrados
            .Row = C_RENENCABEZADO
            For LnContador = 0 To (.get_Cols() - 1) Step 1
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = False
            Next LnContador

            .Rows = 11

            .Col = C_COLFOLIO
            .Row = 1

        End With
    End Sub

    Private Sub chkTipoConsulta_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTipoConsulta.CheckStateChanged
        Dim Index As Integer = chkTipoConsulta.GetIndex(eventSender)
        LlenaDatos()
    End Sub

    Private Sub chkTipoConsulta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTipoConsulta.Enter
        Dim Index As Integer = chkTipoConsulta.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo MErr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ModDCombo.DCChange(lStrSql, Tecla, dbcProveedor)

        If Me.dbcProveedor.Text = "" Then
            LlenaDatos()
        End If

MErr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        Pon_Tool()
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' ORDER BY descProvAcreed"
        ModDCombo.DCGotFocus(gStrSql, dbcProveedor)
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcProveedor.KeyDown
        Dim Aux As String
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                mblnSalir = True
                Me.Close()
                eventSender.KeyCode = 0
            Case System.Windows.Forms.Keys.Return
                Aux = Trim(Me.dbcProveedor.Text)
                If Me.dbcProveedor.SelectedItem <> 0 Then
                    dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
                End If
                Me.dbcProveedor.Text = Aux
                Exit Sub
            Case System.Windows.Forms.Keys.Tab
                Aux = Trim(Me.dbcProveedor.Text)
                If Me.dbcProveedor.SelectedItem <> 0 Then
                    dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
                End If
                Me.dbcProveedor.Text = Aux
                Exit Sub
        End Select
        Tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
        Dim i As Integer
        Dim Aux As Integer
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        Aux = mintCodProveedor
        mintCodProveedor = 0
        ModDCombo.DCLostFocus(dbcProveedor, gStrSql, mintCodProveedor)
        If Aux <> mintCodProveedor Then
            LlenaDatos()
        End If
    End Sub

    Private Sub dbcProveedor_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcProveedor.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcProveedor.Text)
        'If Me.dbcProveedor.SelectedItem <> 0 Then
        '    dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        'End If
        Me.dbcProveedor.Text = Aux
    End Sub

    Private Sub frmCXPConsultaOrden_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Me.ZOrder
    End Sub

    Private Sub frmCXPConsultaOrden_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If UCase(Me.ActiveControl.Name) <> "MSHFLEX" Then
                    ModEstandar.AvanzarTab(Me)
                    If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                        Me.mshFlex.Col = 0
                        Me.mshFlex.Row = 1
                        Me.mshFlex.Focus()
                    End If
                Else
                    ModEstandar.AvanzarTab(Me)
                End If
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "MSHFLEX" Then
                    If Me.chkTipoConsulta(0).CheckState = System.Windows.Forms.CheckState.Checked Then
                        Me.chkTipoConsulta(0).Focus()
                    ElseIf Me.chkTipoConsulta(1).CheckState = System.Windows.Forms.CheckState.Checked Then
                        Me.chkTipoConsulta(1).Focus()
                    ElseIf Me.chkTipoConsulta(2).CheckState = System.Windows.Forms.CheckState.Checked Then
                        Me.chkTipoConsulta(2).Focus()
                    Else
                        Me.chkTipoConsulta(2).Focus()
                    End If
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCXPConsultaOrden_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPConsultaOrden_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Encabezado()
        LlenaDatos()
    End Sub

    Private Sub frmCXPConsultaOrden_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub mshFlex_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.DblClick
        mshFlex_KeyPressEvent(mshFlex, New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub mshFlex_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mshFlex.Enter
        Pon_Tool()
        '    Me.mshFlex.Row = 1
        '    Me.mshFlex.Col = 0
        '    Me.mshFlex.TopRow = 1
    End Sub

    Private Sub mshFlex_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles mshFlex.KeyPressEvent
        With Me.mshFlex
            If eventArgs.keyAscii = 13 Then
                If Trim(.get_TextMatrix(.Row, C_COLFOLIO)) <> "" Then
                    Select Case Me.cFORM
                        Case "FRMCXPORDENCOMPRA"
                            frmCXPOrdenCompra.txtFolio.Text = Trim(.get_TextMatrix(.Row, C_COLFOLIO))
                            frmCXPOrdenCompra.LlenaDatos()
                            Me.Close()
                            frmCXPOrdenCompra.txtFolio.Focus()
                        Case "FRMCXPREGFACTCOMPRAS"
                            'frmCXPRegFactCompras.txtFolio.Text = Trim(.get_TextMatrix(.Row, C_COLFOLIO))
                            'frmCXPRegFactCompras.LlenaDatosOrdenCompra()
                            'Me.Close()
                            'frmCXPRegFactCompras.txtFolio.Focus()
                    End Select
                End If
            End If
        End With
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCXPConsultaOrden))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._chkTipoConsulta_3 = New System.Windows.Forms.CheckBox()
        Me._chkTipoConsulta_2 = New System.Windows.Forms.CheckBox()
        Me._chkTipoConsulta_1 = New System.Windows.Forms.CheckBox()
        Me._chkTipoConsulta_0 = New System.Windows.Forms.CheckBox()
        Me._fraTipoConsulta_0 = New System.Windows.Forms.GroupBox()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me.mshFlex = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._lblConsulta_0 = New System.Windows.Forms.Label()
        Me.chkTipoConsulta = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.fraTipoConsulta = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblConsulta = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._fraTipoConsulta_0.SuspendLayout()
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkTipoConsulta, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraTipoConsulta, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblConsulta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_chkTipoConsulta_3
        '
        Me._chkTipoConsulta_3.BackColor = System.Drawing.SystemColors.Control
        Me._chkTipoConsulta_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkTipoConsulta_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTipoConsulta.SetIndex(Me._chkTipoConsulta_3, CType(3, Short))
        Me._chkTipoConsulta_3.Location = New System.Drawing.Point(312, 16)
        Me._chkTipoConsulta_3.Name = "_chkTipoConsulta_3"
        Me._chkTipoConsulta_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkTipoConsulta_3.Size = New System.Drawing.Size(95, 33)
        Me._chkTipoConsulta_3.TabIndex = 6
        Me._chkTipoConsulta_3.Text = "Registradas (con factura)"
        Me.ToolTip1.SetToolTip(Me._chkTipoConsulta_3, "Mostrar Órdenes Registradas en Factura")
        Me._chkTipoConsulta_3.UseVisualStyleBackColor = False
        '
        '_chkTipoConsulta_2
        '
        Me._chkTipoConsulta_2.BackColor = System.Drawing.SystemColors.Control
        Me._chkTipoConsulta_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkTipoConsulta_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTipoConsulta.SetIndex(Me._chkTipoConsulta_2, CType(2, Short))
        Me._chkTipoConsulta_2.Location = New System.Drawing.Point(216, 16)
        Me._chkTipoConsulta_2.Name = "_chkTipoConsulta_2"
        Me._chkTipoConsulta_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkTipoConsulta_2.Size = New System.Drawing.Size(90, 33)
        Me._chkTipoConsulta_2.TabIndex = 5
        Me._chkTipoConsulta_2.Text = "Canceladas"
        Me.ToolTip1.SetToolTip(Me._chkTipoConsulta_2, "Mostrar Órdenes Canceladas")
        Me._chkTipoConsulta_2.UseVisualStyleBackColor = False
        '
        '_chkTipoConsulta_1
        '
        Me._chkTipoConsulta_1.BackColor = System.Drawing.SystemColors.Control
        Me._chkTipoConsulta_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkTipoConsulta_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTipoConsulta.SetIndex(Me._chkTipoConsulta_1, CType(1, Short))
        Me._chkTipoConsulta_1.Location = New System.Drawing.Point(120, 16)
        Me._chkTipoConsulta_1.Name = "_chkTipoConsulta_1"
        Me._chkTipoConsulta_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkTipoConsulta_1.Size = New System.Drawing.Size(90, 33)
        Me._chkTipoConsulta_1.TabIndex = 4
        Me._chkTipoConsulta_1.Text = "Generadas (Recibidas)"
        Me.ToolTip1.SetToolTip(Me._chkTipoConsulta_1, "Mostrar Órdenes Terminadas")
        Me._chkTipoConsulta_1.UseVisualStyleBackColor = False
        '
        '_chkTipoConsulta_0
        '
        Me._chkTipoConsulta_0.BackColor = System.Drawing.SystemColors.Control
        Me._chkTipoConsulta_0.Checked = True
        Me._chkTipoConsulta_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkTipoConsulta_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkTipoConsulta_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTipoConsulta.SetIndex(Me._chkTipoConsulta_0, CType(0, Short))
        Me._chkTipoConsulta_0.Location = New System.Drawing.Point(24, 16)
        Me._chkTipoConsulta_0.Name = "_chkTipoConsulta_0"
        Me._chkTipoConsulta_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkTipoConsulta_0.Size = New System.Drawing.Size(90, 33)
        Me._chkTipoConsulta_0.TabIndex = 3
        Me._chkTipoConsulta_0.Text = "Vigentes (por recibir)"
        Me.ToolTip1.SetToolTip(Me._chkTipoConsulta_0, "Mostrar Órdenes Vigentes")
        Me._chkTipoConsulta_0.UseVisualStyleBackColor = False
        '
        '_fraTipoConsulta_0
        '
        Me._fraTipoConsulta_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraTipoConsulta_0.Controls.Add(Me._chkTipoConsulta_3)
        Me._fraTipoConsulta_0.Controls.Add(Me._chkTipoConsulta_2)
        Me._fraTipoConsulta_0.Controls.Add(Me._chkTipoConsulta_1)
        Me._fraTipoConsulta_0.Controls.Add(Me._chkTipoConsulta_0)
        Me._fraTipoConsulta_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTipoConsulta.SetIndex(Me._fraTipoConsulta_0, CType(0, Short))
        Me._fraTipoConsulta_0.Location = New System.Drawing.Point(80, 48)
        Me._fraTipoConsulta_0.Name = "_fraTipoConsulta_0"
        Me._fraTipoConsulta_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraTipoConsulta_0.Size = New System.Drawing.Size(409, 65)
        Me._fraTipoConsulta_0.TabIndex = 2
        Me._fraTipoConsulta_0.TabStop = False
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(80, 20)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(409, 21)
        Me.dbcProveedor.TabIndex = 1
        '
        'mshFlex
        '
        Me.mshFlex.DataSource = Nothing
        Me.mshFlex.Location = New System.Drawing.Point(8, 120)
        Me.mshFlex.Name = "mshFlex"
        Me.mshFlex.OcxState = CType(resources.GetObject("mshFlex.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mshFlex.Size = New System.Drawing.Size(479, 175)
        Me.mshFlex.TabIndex = 7
        '
        '_lblConsulta_0
        '
        Me._lblConsulta_0.AutoSize = True
        Me._lblConsulta_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblConsulta_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblConsulta_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblConsulta.SetIndex(Me._lblConsulta_0, CType(0, Short))
        Me._lblConsulta_0.Location = New System.Drawing.Point(16, 24)
        Me._lblConsulta_0.Name = "_lblConsulta_0"
        Me._lblConsulta_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblConsulta_0.Size = New System.Drawing.Size(56, 13)
        Me._lblConsulta_0.TabIndex = 0
        Me._lblConsulta_0.Text = "Proveedor"
        '
        'chkTipoConsulta
        '
        '
        'frmCXPConsultaOrden
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(498, 344)
        Me.Controls.Add(Me._fraTipoConsulta_0)
        Me.Controls.Add(Me.dbcProveedor)
        Me.Controls.Add(Me.mshFlex)
        Me.Controls.Add(Me._lblConsulta_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 29)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCXPConsultaOrden"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Consulta de Ódenes de Compra"
        Me._fraTipoConsulta_0.ResumeLayout(False)
        CType(Me.mshFlex, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkTipoConsulta, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraTipoConsulta, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblConsulta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

End Class