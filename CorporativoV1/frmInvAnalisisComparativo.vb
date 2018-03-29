Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmInvAnalisisComparativo
    Inherits System.Windows.Forms.Form

    Public isload As Boolean = False

    Private components As System.ComponentModel.IContainer
    'Programa : Análisis Comparativo
    'Elaboró: Rosaura Torres López.
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents msgArticulos As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents txtDesArticulo As System.Windows.Forms.Label
    Public WithEvents lblOrigen As System.Windows.Forms.Label
    Public WithEvents _Label_2 As System.Windows.Forms.Label
    Public WithEvents lblAlmacen As System.Windows.Forms.Label
    Public WithEvents _Label_7 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents chkOrdenarporGrupo As System.Windows.Forms.CheckBox
    Public WithEvents optImpTodo As System.Windows.Forms.RadioButton
    Public WithEvents optImpSoloDiferencias As System.Windows.Forms.RadioButton
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents chkOrdenarCodAnt As System.Windows.Forms.CheckBox
    Public WithEvents chkVizSubNivel As System.Windows.Forms.CheckBox
    Public WithEvents CmbRefrescar As System.Windows.Forms.Button
    Public WithEvents TVArticulos As System.Windows.Forms.TreeView
    Public WithEvents Imagenes As System.Windows.Forms.ImageList
    Public WithEvents _Label_12 As System.Windows.Forms.Label
    Public WithEvents _Label_13 As System.Windows.Forms.Label
    Public WithEvents _Label_14 As System.Windows.Forms.Label
    Public WithEvents _Label_15 As System.Windows.Forms.Label
    Public WithEvents _Label_16 As System.Windows.Forms.Label
    Public WithEvents _Label_17 As System.Windows.Forms.Label
    Public WithEvents _Label_18 As System.Windows.Forms.Label
    Public WithEvents _Label_19 As System.Windows.Forms.Label
    Public WithEvents frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Dim Sql As String
    Dim rec As ADODB.Recordset
    Dim Nodo As System.Windows.Forms.TreeNode ''esta variable es para insertar nodos al treeview
    '''estas variables son contadores para los ciclos
    Dim pry As Integer '''proyectos
    Dim mdl As Integer '''modulos
    Dim prc As Integer '''procesos
    Dim prg As Integer '''programas
    Dim mblnSalir As Boolean
    Dim mblnFueraChange As Boolean
    Dim tecla As Integer
    Dim I As Integer
    Dim NodosExp() As Integer
    Dim mintCodSucursal As Integer
    Dim FueraChange As Boolean

    Const C_TODAS As String = "[ Todas ... ]"
    Const C_COLCODIGO As Integer = 0
    Const C_COLDESCRIPCION As Integer = 1
    Const C_ColCODIGOANT As Integer = 2
    Const C_COLUNIDAD As Integer = 3
    Const C_ColTEORICO As Integer = 4
    Const C_ColFISICO As Integer = 5
    Const C_ColAJUSTE As Integer = 6
    Const C_ColGRUPO As Integer = 14
    Const C_ColFAMILIA As Integer = 7
    Const C_ColLINEA As Integer = 8
    Const C_ColSUBLINEA As Integer = 9
    Const C_ColMARCA As Integer = 10
    Const C_ColMODELO As Integer = 11
    Const C_ColCOSTOUNITARIO As Integer = 12
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Const C_ColIMPORTE As Integer = 13

    Sub Nuevo()
        On Error GoTo Merr
        FueraChange = False
        msgArticulos.Clear()
        Encabezado()
        MostrarDatosAlmacen()

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub chkOrdenarCodAnt_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkOrdenarCodAnt.CheckStateChanged
        DatosGenerales()
    End Sub

    Private Sub chkOrdenarCodAnt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkOrdenarCodAnt.Enter
        Pon_Tool()
    End Sub

    Private Sub chkOrdenarporGrupo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkOrdenarporGrupo.Enter
        Pon_Tool()
    End Sub

    Private Sub chkVizSubNivel_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkVizSubNivel.CheckStateChanged
        If chkVizSubNivel.CheckState = System.Windows.Forms.CheckState.Checked Then
            TVArticulos.Enabled = True
            Refrescar()
            Nuevo()
        Else
            DatosGenerales()
            TVArticulos.Enabled = False
            Refrescar()
        End If
    End Sub

    Private Sub chkVizSubNivel_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkVizSubNivel.Enter
        Pon_Tool()
    End Sub

    Private Sub chkVizSubNivel_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkVizSubNivel.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                If TVArticulos.Enabled = False Then
                    msgArticulos.Focus()
                Else
                    TVArticulos.Focus()
                End If
        End Select
    End Sub

    Private Sub CmbRefrescar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbRefrescar.Click
        msgArticulos.Clear()
        Encabezado()
        Refrescar()
        '    TVArticulos.SetFocus
        chkVizSubNivel.CheckState = System.Windows.Forms.CheckState.Unchecked
        DatosGenerales()
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        If dbcSucursal.Text = "" Then
            msgArticulos.Clear()
            Encabezado()
            Exit Sub
        End If

        '        On Local Error GoTo MErr
        '        Dim lStrSql As String

        '        If FueraChange Then Exit Sub

        '        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.text) & "%'"
        '        ModDCombo.DCChange lStrSql, tecla, dbcSucursal

        '            '''If Trim(dbcSucursal.text) = "" Then mintCodSucursal = 0
        '        If dbcSucursal.SelectedItem <> 0 Then MostrarDatosSucursal()

        'MErr:
        '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcSucursal_Click(ByVal eventSender As System.Object, ByVal eventArgs As EventArgs) Handles dbcSucursal.Click
        'If dbcSucursal.SelectedItem <> "" Then
        MostrarDatosSucursal()
        'End If
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' "
        ModDCombo.DCGotFocus(gStrSql, dbcSucursal)
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursal.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.Close()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        Dim I As Integer
        Dim Aux As Integer

        If FueraChange Then Exit Sub
        If dbcSucursal.Text = dbcSucursal.Tag Then Exit Sub

        'If ActiveControl.Name <> "" Then
        '    Exit Sub
        'Else
        If Trim(Me.dbcSucursal.Text) = "" Or Trim(Me.dbcSucursal.Text) = C_TODAS Then Exit Sub
        'End If

        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        Aux = mintCodSucursal
        mintCodSucursal = 0
        ModDCombo.DCLostFocus((Me.dbcSucursal), gStrSql, mintCodSucursal)
        dbcSucursal.Refresh()
        Me.Refresh()

        System.Windows.Forms.Application.DoEvents()
        FueraChange = True
        Nuevo()
        Encabezado()
        'Load(frmBarraDesplazamiento)
        'Dim frmBarraDesplazamiento As New frmBarraDesplazamiento()
        'frmBarraDesplazamiento.InitializeComponent()
        frmBarraDesplazamiento.Text = "Análisis Comparativo"
        frmBarraDesplazamiento.Tag = Me.Name
        frmBarraDesplazamiento.Show()
        frmBarraDesplazamiento.BringToFront()
        frmBarraDesplazamiento.Refresh()
        'System.Windows.Forms.Application.DoEvents()
        frmBarraDesplazamiento.PrgBarra.Value = 20
        frmBarraDesplazamiento.BringToFront()
        frmBarraDesplazamiento.Refresh()
        'System.Windows.Forms.Application.DoEvents()
        Refrescar()
        'System.Windows.Forms.Application.DoEvents()
        frmBarraDesplazamiento.PrgBarra.Value = 40
        frmBarraDesplazamiento.BringToFront()
        frmBarraDesplazamiento.Refresh()
        'System.Windows.Forms.Application.DoEvents()
        frmBarraDesplazamiento.PrgBarra.Value = 60
        frmBarraDesplazamiento.BringToFront()
        frmBarraDesplazamiento.Refresh()
        'System.Windows.Forms.Application.DoEvents()
        DatosGenerales()
        'System.Windows.Forms.Application.DoEvents()
        frmBarraDesplazamiento.PrgBarra.Value = 80
        frmBarraDesplazamiento.BringToFront()
        frmBarraDesplazamiento.Refresh()
        'System.Windows.Forms.Application.DoEvents()
        frmBarraDesplazamiento.PrgBarra.Value = 100
        frmBarraDesplazamiento.BringToFront()
        frmBarraDesplazamiento.Refresh()
        'System.Windows.Forms.Application.DoEvents()
        frmBarraDesplazamiento.Close()
        FueraChange = False
        msgArticulos.Col = 0
        msgArticulos.Row = 1
        msgArticulos.Focus()

        'FueraChange = True
        'Nuevo()
        'Encabezado()
        'Refrescar()
        'DatosGenerales()
        'FueraChange = False
        'msgArticulos.SetFocus

        'Código insertado del inicio de la aplicación
        '    FueraChange = True
        'Load frmBarraDesplazamiento
        '    frmBarraDesplazamiento.Caption = "Análisis Comparativo"
        'frmBarraDesplazamiento.Tag = Me.Name
        'frmBarraDesplazamiento.Show
        'frmBarraDesplazamiento.ZOrder
        'frmBarraDesplazamiento.Refresh
        'Refrescar()
        'frmBarraDesplazamiento.PrgBarra.Value = 20
        'DoEvents
        'frmBarraDesplazamiento.PrgBarra.Value = 40
        'TVArticulos.Refresh
        'frmBarraDesplazamiento.ZOrder
        'frmBarraDesplazamiento.PrgBarra.Value = 60
        'Nuevo()
        'Encabezado()
        'frmBarraDesplazamiento.ZOrder
        'DoEvents
        'frmBarraDesplazamiento.PrgBarra.Value = 80
        'DatosGenerales()
        'frmBarraDesplazamiento.ZOrder
        'DoEvents
        'frmBarraDesplazamiento.PrgBarra.Value = 100
        'DoEvents
        'frmBarraDesplazamiento.ZOrder
        'Unload frmBarraDesplazamiento
        '    FueraChange = False
        'msgArticulos.SetFocus
    End Sub

    'Private Sub dbcSucursal_Change()
    '    On Local Error GoTo MErr
    '    Dim lStrSql As String
    '
    '    If mblnFueraChange Then Exit Sub
    '    If dbcSucursal.text = "" Then
    '       msgArticulos.Clear
    '       Encabezado
    '       Exit Sub
    '    End If
    '    lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.text) & "%'"
    '    ModDCombo.DCChange lStrSql, Tecla, dbcSucursal
    '
    '    If Trim(Me.dbcSucursal.text) = "" Then
    '        gintCodAlmacen = 0
    '        dbcSucursal_LostFocus
    '    End If
    '    Exit Sub
    'MErr:
    '    ModEstandar.MostrarError
    'End Sub
    '
    'Private Sub dbcSucursal_GotFocus()
    '    Pon_Tool
    '    gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen WHERE TipoAlmacen = 'P'"
    '    ModDCombo.DCGotFocus gStrSql, dbcSucursal
    'End Sub
    '
    'Private Sub dbcSucursal_LostFocus()
    '    Dim i As Integer
    '    Dim Aux As Integer
    '    If Screen.ActiveForm.Name <> Me.Name Then
    '        Exit Sub
    '    Else
    '        If Trim(Me.dbcSucursal.text) = "" Or Trim(Me.dbcSucursal.text) = C_TODAS Then Exit Sub
    '    End If
    '    gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.text) & "%'"
    '    Aux = gintCodAlmacen
    '    gintCodAlmacen = 0
    '    ModDCombo.DCLostFocus Me.dbcSucursal, gStrSql, gintCodAlmacen
    '
    '    'Código insertado del inicio de la aplicación
    '    Load frmBarraDesplazamiento
    '    frmBarraDesplazamiento.Caption = "Análisis Comparativo"
    '    frmBarraDesplazamiento.Tag = Me.Name
    '    frmBarraDesplazamiento.Show
    '    frmBarraDesplazamiento.ZOrder
    '    frmBarraDesplazamiento.Refresh
    '    Refrescar
    '    frmBarraDesplazamiento.PrgBarra.Value = 20
    '    DoEvents
    '    frmBarraDesplazamiento.PrgBarra.Value = 40
    '    TVArticulos.Refresh
    '    frmBarraDesplazamiento.ZOrder
    '    frmBarraDesplazamiento.PrgBarra.Value = 60
    '    Nuevo
    '    Encabezado
    '    frmBarraDesplazamiento.ZOrder
    '    DoEvents
    '    frmBarraDesplazamiento.PrgBarra.Value = 80
    '    DatosGenerales
    '    frmBarraDesplazamiento.ZOrder
    '    DoEvents
    '    frmBarraDesplazamiento.PrgBarra.Value = 100
    '    DoEvents
    '    frmBarraDesplazamiento.ZOrder
    '    Unload frmBarraDesplazamiento
    'End Sub
    '
    'Private Sub dbcSucursal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Dim Aux As String
    '    Aux = Trim(Me.dbcSucursal.text)
    '    If Me.dbcSucursal.SelectedItem <> 0 Then
    '        dbcSucursal_LostFocus
    '    End If
    '    Me.dbcSucursal.text = Aux
    'End Sub

    Private Sub frmInvAnalisisComparativo_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        frmBarraDesplazamiento.Close()
        'Me.ZOrder
        'CmbRefrescar_Click()
        'DatosGenerales()
        'Refrescar()
    End Sub

    Private Sub frmInvAnalisisComparativo_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        'Desactivar todas las opciones del Menu
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmInvAnalisisComparativo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmInvAnalisisComparativo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmInvAnalisisComparativo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        isload = True
        Dim lCont As Integer
        '   mstrProceso = "FRMINVANALISISCOMPARATIVO"
        '   mstrProcesoInv = "Análisis Comparativo"
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        '    Load frmBarraDesplazamiento
        '    frmBarraDesplazamiento.Caption = "Análisis Comparativo"
        '    frmBarraDesplazamiento.Tag = Me.Name
        '    frmBarraDesplazamiento.Show
        '    frmBarraDesplazamiento.ZOrder
        '    DoEvents
        '    frmBarraDesplazamiento.PrgBarra.Value = 20
        '    CargaDatosGrupos
        '    frmBarraDesplazamiento.ZOrder
        '    DoEvents
        '    frmBarraDesplazamiento.PrgBarra.Value = 40
        '    TVArticulos.Refresh
        '    frmBarraDesplazamiento.ZOrder
        '    frmBarraDesplazamiento.PrgBarra.Value = 60
        Nuevo()
        Encabezado()
        '    frmBarraDesplazamiento.ZOrder
        '    DoEvents
        '    frmBarraDesplazamiento.PrgBarra.Value = 80
        '    DatosGenerales
        '    frmBarraDesplazamiento.ZOrder
        '    DoEvents
        '    frmBarraDesplazamiento.PrgBarra.Value = 100
        '    DoEvents
        '    frmBarraDesplazamiento.ZOrder
        '    Do While lCont <= 6000000
        '       lCont = lCont + 1
        '    Loop
    End Sub

    'Private Sub Form_Load()
    '    '                              N uevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
    '    ModEstandar.ActivaMenu C_DESACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO
    '    Icono Me, MenuPrincipal
    '    ModEstandar.CentrarForma Me
    '    CargaDatosGrupos()

    '    Nuevo()
    '    Encabezado()
    '    DatosGenerales()
    'End Sub

    Private Sub frmInvAnalisisComparativo_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSalir Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        'Else 'Se quiere salir con escape
        '    mblnSalir = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
        '        Case MsgBoxResult.Yes
        '            Cancel = 0 'Sale de la Captura, Con 1: Sigue en la captura
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Public Sub CargaDatosGrupos()
        On Error GoTo Error_Renamed
        Dim LetraLlave As String ''Identifica cual será la letra par ala Llave, esto para identificar Relojeria y VArios, Dependiendo del Grupo, se Pondrá una letra diferente
        With TVArticulos
            '    JOYERIA
            '    Familia -F
            '    LINEA -L
            '    sublinea -S
            '    RELOJERIA
            '    MARCA -M
            '    MODELO -D
            '    VARIOS
            '    FAMILIA -I
            '    LINA -N


            'BUSCAR TODOS LOS GRUPOS EXISTENTES EN LA TABLA DE HOJA DE CONTROL, PARA MOSTRARLOS.
            Sql = "SELECT   DISTINCT " & "         RIGHT('0000' + RTRIM(LTRIM(STR(dbo.InvHojaControl.CodGrupo))), 4) AS CodGrupo, LTRIM(RTRIM(dbo.CatGrupos.DescGrupo)) AS Descripcion " & "FROM     dbo.InvHojaControl INNER JOIN " & "         dbo.CatGrupos ON dbo.InvHojaControl.CodGrupo = dbo.CatGrupos.CodGrupo " & "WHERE    dbo.InvHojaControl.CodAlmacen =  " & mintCodSucursal & " " & "And      ((ExistenciaTeorica+ExistenciaFisica) > 0) "

            BorraCmd()
            Cmd.CommandText = "Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
            rec = Cmd.Execute
            If rec.RecordCount > 0 Then
                rec.MoveFirst()
                For mdl = 1 To rec.RecordCount
                    Nodo = .Nodes.Add(VB.Right("0000" & RTrim(LTrim(Str(rec.Fields("CodGrupo").Value))), 4) & "G", Trim(rec.Fields("Descripcion").Value), "Cerrada", "Abierta")
                    '.Nodes.Item(mdl).ForeColor = System.Drawing.ColorTranslator.FromOle(&H800000)
                    '.Nodes.Item(mdl).ExpandedImage = "Abierta"
                    Nodo.Expand()
                    rec.MoveNext()
                Next mdl
            End If

            ''Aqui se buscan Las Familias   y Marcas Que existan  en los Grupos
            For pry = 1 To .Nodes.Count
                If CDbl(Mid(.Nodes.Item(pry).Name, 1, 4)) = gCODJOYERIA Then
                    LetraLlave = "F"
                ElseIf CDbl(Mid(.Nodes.Item(pry).Name, 1, 4)) = gCODVARIOS Then
                    LetraLlave = "I"
                End If

                Sql = "SELECT   DISTINCT " & "         RIGHT('0000' + RTRIM(LTRIM(STR(dbo.InvHojaControl.CodFamilia))), 4) AS CodFamilia, LTRIM(RTRIM(dbo.CatFamilias.DescFamilia)) " & "         AS Descripcion " & "FROM     dbo.InvHojaControl INNER JOIN " & "         dbo.CatGrupos ON dbo.InvHojaControl.CodGrupo = dbo.CatGrupos.CodGrupo INNER JOIN " & "         dbo.CatFamilias ON dbo.InvHojaControl.CodGrupo = dbo.CatFamilias.CodGrupo AND " & "         dbo.InvHojaControl.CodFamilia = dbo.CatFamilias.CodFamilia " & "Where    (dbo.InvHojaControl.CodGrupo = " & Mid(.Nodes.Item(pry).Name, 1, 4) & ") " & "And      (dbo.InvHojaControl.CodAlmacen = " & mintCodSucursal & ") " & "And      ((ExistenciaTeorica+ExistenciaFisica) > 0) "

                ModEstandar.BorraCmd()
                Cmd.CommandText = "Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                rec = Cmd.Execute
                For mdl = 1 To rec.RecordCount
                    Nodo = TVArticulos.Nodes.Insert(pry, Mid(.Nodes.Item(pry).Name, 1, 4) & rec.Fields("CodFamilia").Value & LetraLlave, Trim(rec.Fields("Descripcion").Value), "Cerrada", "Abierta")
                    'Nodo.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80)
                    'Nodo.ExpandedImage = "Abierta"
                    Nodo.Expand()
                    rec.MoveNext()
                Next mdl
            Next pry

            For pry = 1 To .Nodes.Count

                Sql = "SELECT    DISTINCT " & "          RIGHT('0000' + RTRIM(LTRIM(STR(dbo.InvHojaControl.CodMarca))), 4) AS CodMarca, LTRIM(RTRIM(dbo.CatMarcas.DescMarca)) " & "          AS Descripcion " & "FROM      dbo.InvHojaControl INNER JOIN " & "          dbo.CatMarcas ON dbo.InvHojaControl.CodGrupo = dbo.CatMarcas.CodGrupo AND " & "          dbo.InvHojaControl.CodMarca = dbo.CatMarcas.CodMarca " & "Where     (dbo.InvHojaControl.CodGrupo = " & Mid(.Nodes.Item(pry).Name, 1, 4) & ") " & "And       (dbo.InvHojaControl.CodAlmacen = " & mintCodSucursal & ") " & "And       ((ExistenciaTeorica+ExistenciaFisica) > 0) "

                ModEstandar.BorraCmd()
                Cmd.CommandText = "Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                rec = Cmd.Execute
                For mdl = 1 To rec.RecordCount
                    Nodo = TVArticulos.Nodes.Insert(pry, Mid(.Nodes.Item(pry).Name, 1, 4) & rec.Fields("CodMArca").Value & "M", Trim(rec.Fields("Descripcion").Value), "Cerrada", "Abierta")
                    'Nodo.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80)
                    'Nodo.ExpandedImage = "Abierta"
                    Nodo.Expand()
                    rec.MoveNext()
                Next mdl

            Next pry

            ''A continuacion se CArgan las Lineas para las Familias y Los Modelos  para las Marcas
            For pry = 1 To .Nodes.Count
                If VB.Right(.Nodes.Item(pry).Name, 1) = "F" Or VB.Right(.Nodes.Item(pry).Name, 1) = "I" Then
                    If CDbl(Mid(.Nodes.Item(pry).Name, 1, 4)) = gCODJOYERIA Then
                        LetraLlave = "L"
                    ElseIf CDbl(Mid(.Nodes.Item(pry).Name, 1, 4)) = gCODVARIOS Then
                        LetraLlave = "N"
                    End If

                    Sql = "SELECT    DISTINCT RIGHT('0000' + RTRIM(LTRIM(STR(dbo.InvHojaControl.CodLinea))), 4) AS CodLinea, dbo.CatLineas.DescLinea as Descripcion " & "FROM      dbo.InvHojaControl INNER JOIN " & "          dbo.CatLineas ON dbo.InvHojaControl.CodGrupo = dbo.CatLineas.CodGrupo AND dbo.InvHojaControl.CodFamilia = dbo.CatLineas.CodFamilia AND " & "          dbo.InvHojaControl.codLinea = dbo.CatLineas.codLinea " & "Where     dbo.InvHojaControl.CodGrupo = " & CInt(Mid(.Nodes.Item(pry).Name, 1, 4)) & " And dbo.InvHojaControl.CodFamilia = " & CInt(Mid(.Nodes.Item(pry).Name, 5, 4)) & "  " & "And       dbo.InvHojaControl.CodAlmacen = " & mintCodSucursal & " " & "And       ((ExistenciaTeorica+ExistenciaFisica) > 0) "

                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                    rec = Cmd.Execute
                    For prc = 1 To rec.RecordCount
                        Nodo = TVArticulos.Nodes.Insert(pry, Mid(.Nodes.Item(pry).Name, 1, 4) & Mid(.Nodes.Item(pry).Name, 5, 4) & rec.Fields("COdLinea").Value & LetraLlave, Trim(rec.Fields("Descripcion").Value), "Cerrada", "Abierta")
                        'Nodo.ForeColor = System.Drawing.ColorTranslator.FromOle(&H808000)
                        'Nodo.SelectedImage = "VerCarpeta"
                        'Nodo.ExpandedImage = "Abierta"
                        Nodo.Expand()
                        rec.MoveNext()
                    Next prc
                End If
            Next pry

            For pry = 1 To .Nodes.Count
                If VB.Right(.Nodes.Item(pry).Name, 1) = "M" Then

                    Sql = "SELECT    DISTINCT " & "          RIGHT('0000' + RTRIM(LTRIM(STR(dbo.InvHojaControl.CodModelo))), 4) AS CodModelo, LTRIM(RTRIM(dbo.CatModelos.DescModelo)) " & "          AS Descripcion " & "FROM      dbo.InvHojaControl INNER JOIN " & "          dbo.CatModelos ON dbo.InvHojaControl.CodGrupo = dbo.CatModelos.CodGrupo AND dbo.InvHojaControl.CodMarca = dbo.CatModelos.CodMarca AND " & "          dbo.InvHojaControl.CodModelo = dbo.CatModelos.CodModelo " & "Where     (dbo.InvHojaControl.CodGrupo = " & CInt(Mid(.Nodes.Item(pry).Name, 1, 4)) & ") And (dbo.InvHojaControl.CodMarca = " & CInt(Mid(.Nodes.Item(pry).Name, 5, 4)) & " ) " & "And       (dbo.InvHojaControl.CodAlmacen = " & mintCodSucursal & ") " & "And       ((ExistenciaTeorica+ExistenciaFisica) > 0) "

                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                    rec = Cmd.Execute
                    For prc = 1 To rec.RecordCount
                        Nodo = TVArticulos.Nodes.Insert(pry, Mid(.Nodes.Item(pry).Name, 1, 4) & Mid(.Nodes.Item(pry).Name, 5, 4) & rec.Fields("CodModelo").Value & "D", Trim(rec.Fields("Descripcion").Value), "Cerrada", "Abierta")
                        'Nodo.ForeColor = System.Drawing.ColorTranslator.FromOle(&H808000)
                        'Nodo.ExpandedImage = "Abierta"
                        Nodo.Expand()
                        rec.MoveNext()
                    Next prc
                End If
            Next pry

            ''en esta seccion se cargan las SubLineas Para las Lineas.
            For pry = 1 To .Nodes.Count
                If VB.Right(.Nodes.Item(pry).Name, 1) = "L" Then
                    Sql = "SELECT    DISTINCT " & "          RIGHT('0000' + RTRIM(LTRIM(STR(dbo.InvHojaControl.CodSubLinea))), 4) AS CodSubLinea, LTRIM(RTRIM(dbo.CatSubLineas.DescSubLinea)) " & "          AS Descripcion " & "FROM      dbo.InvHojaControl INNER JOIN " & "          dbo.CatSubLineas ON dbo.InvHojaControl.CodGrupo = dbo.CatSubLineas.CodGrupo AND " & "          dbo.InvHojaControl.CodFamilia = dbo.CatSubLineas.CodFamilia AND dbo.InvHojaControl.CodLinea = dbo.CatSubLineas.CodLinea AND " & "          dbo.InvHojaControl.CodSubLinea = dbo.CatSubLineas.CodSubLinea " & "Where     dbo.InvHojaControl.CodGrupo = " & CInt(Mid(.Nodes.Item(pry).Name, 1, 4)) & "And       dbo.InvHojaControl.CodFamilia = " & CInt(Mid(.Nodes.Item(pry).Name, 5, 4)) & "And       dbo.InvHojaControl.codLinea = " & CInt(Mid(.Nodes.Item(pry).Name, 9, 4)) & "And       (dbo.InvHojaControl.CodAlmacen = " & mintCodSucursal & ") " & "And       ((ExistenciaTeorica+ExistenciaFisica) > 0) "

                    ModEstandar.BorraCmd()
                    Cmd.CommandText = "Up_Select_Datos"
                    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                    Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                    rec = Cmd.Execute
                    For prg = 1 To rec.RecordCount
                        Nodo = TVArticulos.Nodes.Insert(pry, Mid(.Nodes.Item(pry).Name, 1, 4) & Mid(.Nodes.Item(pry).Name, 5, 4) & Mid(.Nodes.Item(pry).Name, 9, 4) & rec.Fields("CodSubLinea").Value & "S", Trim(rec.Fields("Descripcion").Value), "Cerrada", "Abierta")
                        'Nodo.ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
                        rec.MoveNext()
                    Next prg
                End If
            Next pry
        End With

Error_Renamed:
        If Err.Number <> 0 Then ModErrores.Errores()
    End Sub

    Public Sub Refrescar()
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        With TVArticulos
            .Nodes.Clear()
            .Visible = False
            CargaDatosGrupos()
            ReDim Preserve NodosExp(.Nodes.Count)
            ExpandirNodos()
            .Visible = True
        End With
        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    Public Sub ExpandirNodos()
        Dim I As Integer
        With TVArticulos
            For I = 1 To .Nodes.Count
                If NodosExp(I) > 0 Then
                    .Nodes.Item(NodosExp(I)).Expand()
                End If
            Next I
        End With
    End Sub

    Sub Encabezado()
        'Genera el encabezao del Grid, asigna el tamaño y número de columas y centra el texto dentro de ellas
        Dim LnContador As Integer

        With msgArticulos
            .FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusHeavy 'flexFocusLight 'flexFocusNone
            .WordWrap = True
            .FixedRows = 1
            .FixedCols = 0
            .set_RowHeight(0, 500)
            .set_ColWidth(C_COLCODIGO, 0, 1000)
            .set_ColWidth(C_COLDESCRIPCION, 0, 3000)
            .set_ColWidth(C_ColCODIGOANT, 0, 1000)
            .set_ColWidth(C_COLUNIDAD, 0, 0) '1000
            .set_ColWidth(C_ColTEORICO, 0, 1000) '1200
            .set_ColWidth(C_ColFISICO, 0, 1000) '1200
            .set_ColWidth(C_ColGRUPO, 0, 0) '1100
            .set_ColWidth(C_ColFAMILIA, 0, 0)
            .set_ColWidth(C_ColLINEA, 0, 0) '910
            .set_ColWidth(C_ColSUBLINEA, 0, 0)
            .set_ColWidth(C_ColMARCA, 0, 0)
            .set_ColWidth(C_ColMODELO, 0, 0)
            .set_ColWidth(C_ColAJUSTE, 0, 0)
            .set_ColWidth(C_ColCOSTOUNITARIO, 0, 1200)
            .set_ColWidth(C_ColIMPORTE, 0, 1200)

            .set_TextMatrix(0, C_COLCODIGO, "CODIGO")
            .set_TextMatrix(0, C_COLDESCRIPCION, "DESCRIPCION")
            .set_TextMatrix(0, C_COLUNIDAD, "UNIDAD")
            .set_TextMatrix(0, C_ColTEORICO, "TEORICO")
            .set_TextMatrix(0, C_ColFISICO, "FISICO")
            .set_TextMatrix(0, C_ColGRUPO, "GRUPO")
            .set_TextMatrix(0, C_ColFAMILIA, "FAMILIA")
            .set_TextMatrix(0, C_ColLINEA, "LINEA")
            .set_TextMatrix(0, C_ColSUBLINEA, "SUBLINEA")
            .set_TextMatrix(0, C_ColMARCA, "MARCA")
            .set_TextMatrix(0, C_ColMODELO, "MODELO")
            .set_TextMatrix(0, C_ColAJUSTE, "AJUSTE")
            .set_TextMatrix(0, C_ColCOSTOUNITARIO, "C.UNITARIO")
            .set_TextMatrix(0, C_ColIMPORTE, "IMPORTE")
            .set_TextMatrix(0, C_ColCODIGOANT, "ANTERIOR")

            .Row = 0
            For LnContador = 0 To 13 - C_ColIMPORTE
                .Col = LnContador
                .CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
                .CellFontBold = True
            Next LnContador
            .Row = 1
            .Col = C_COLCODIGO
            .WordWrap = False 'Hacer esto , para que no se puedan escribir dos o mal lineas de texto en una  sola fila, solo se usa para el encabezado
        End With
    End Sub

    Function MostrarArticulos(ByRef CodGrupo As Integer, ByRef CodFamilia As Integer, ByRef COdLinea As Integer, ByRef CodSubLinea As Integer, ByRef CodMArca As Integer, ByRef CodModelo As Integer) As Decimal
        On Error GoTo Merr
        Dim C_ORDERBY As String
        Dim C_WHERE As String
        C_ORDERBY = ""
        C_WHERE = ""
        If chkOrdenarCodAnt.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            C_ORDERBY = " Order by I.CodArticulo"
        Else
            C_ORDERBY = " Order by A.CodigoAnt"
        End If
        If optImpSoloDiferencias.Checked = True Then
            C_WHERE = " And I.Ajuste <> 0"
        End If
        ModEstandar.BorraCmd()
        'If chkOrdenarCodAnt.Value = vbUnchecked Then
        Sql = "SELECT     I.CodAlmacen, Al.DescAlmacen, I.CodArticulo , A.DescArticulo, A.CodAlmacenOrigen , O.DescAlmacenOrigen, A.CodGrupo, Gr.DescGrupo, isnull(A.CodFamilia,0) " & "as COdFamilia,  Isnull(A.CodLinea,0) as CodLinea,  Isnull(A.CodSubLinea,0)  as CodSubLinea,  Isnull(A.CodMarca,0) as CodMarca, Isnull(A.CodModelo,0) as CodModelo,  A.CodUnidad , " & "U.DescUnidad, I.ExistenciaTeorica, IsNull(I.ExistenciaFisica,0) as ExistenciaFisica , I.Ajuste, I.CostoUnitario, I.CostoUnitario * I.Ajuste AS Importe  , Case A.CodigoAnt " & "When 0 Then '' Else  cast(A.OrigenAnt as nvarchar) + '-' + right('00000'+  Cast(A.CodigoAnt as varchar),5) " & "End  as CodigoAnterior , A.CodigoAnt " & "FROM         dbo.InvHojaControl I INNER JOIN dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN " & "dbo.CatUnidades U ON A.CodUnidad = U.CodUnidad INNER JOIN dbo.CatOrigen O ON A.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN " & "dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen INNER JOIN " & "dbo.CatGrupos Gr ON A.CodGrupo = Gr.CodGrupo LEFT OUTER JOIN " & "dbo.CatFamilias Fa ON A.CodGrupo = Fa.CodGrupo AND A.CodFamilia = Fa.CodFamilia AND Gr.CodGrupo = Fa.CodGrupo LEFT OUTER JOIN " & "dbo.CatLineas Li ON A.CodGrupo = Li.CodGrupo AND A.CodFamilia = Li.CodFamilia AND A.CodLinea = Li.CodLinea AND Fa.CodGrupo = Li.CodGrupo AND " & "Fa.CodFamilia = Li.CodFamilia LEFT OUTER JOIN  dbo.CatSubLineas su ON A.CodGrupo = su.CodGrupo AND A.CodFamilia = su.CodFamilia AND A.CodLinea = su.CodLinea AND " & "A.CodSubLinea = su.CodSubLinea AND Li.CodGrupo = su.CodGrupo AND Li.CodFamilia = su.CodFamilia AND " & "Li.CodLinea = su.CodLinea LEFT OUTER JOIN " & "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca AND Gr.CodGrupo = Ma.CodGrupo LEFT OUTER JOIN " & "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND A.CodMarca = Mo.CodMarca AND A.CodModelo = Mo.CodModelo AND " & "Ma.CodGrupo = Mo.CodGrupo AND Ma.CodMarca = Mo.CodMarca " & "WHERE I.CodAlmacen = " & mintCodSucursal & " And " & IIf((CodGrupo = 0), "I.CodGrupo in (1,2,3)", "I.CodGrupo  = " & CodGrupo & " ") & IIf((CodMArca = 0), " ", " AND I.CodMarca = " & CodMArca & " ") & IIf((CodModelo = 0), " ", " AND I.CodModelo = " & CodModelo & " ") & IIf((CodFamilia = 0), " ", " AND I.CodFamilia = " & CodFamilia & " ") & IIf((COdLinea = 0), " ", " AND I.CodLinea = " & COdLinea & " ") & IIf((CodSubLinea = 0), " ", " AND I.CodSubLinea = " & CodSubLinea & " ") & C_WHERE & " And (IsNull(I.ExistenciaTeorica,0) + IsNull(I.ExistenciaFisica,0)) > 0 " & " GROUP BY I.CodAlmacen, I.CodArticulo, A.CodArticulo, A.DescArticulo, A.CodAlmacenOrigen, A.CodGrupo, A.CodFamilia, A.CodLinea, A.CodSubLinea, A.CodMarca, " & "A.CodModelo, A.CodUnidad, U.DescUnidad,  I.ExistenciaTeorica, I.ExistenciaFisica, I.Ajuste, " & "I.CostoUnitario,Al.DescAlmacen,O.DescAlmacenOrigen,Gr.DescGrupo, A.CodigoAnt,a.OrigenAnt " & C_ORDERBY
        'Else


        'Sql = "SELECT     I.CodAlmacen, Al.DescAlmacen, I.CodArticulo , A.DescArticulo , I.CodAlmacenOrigen , O.DescAlmacenOrigen, A.CodGrupo, Gr.DescGrupo, isnull(A.CodFamilia,0) " &
        '        "as COdFamilia,  Isnull(A.CodLinea,0) as CodLinea,  Isnull(A.CodSubLinea,0)  as CodSubLinea,  Isnull(A.CodMarca,0) as CodMarca, Isnull(A.CodModelo,0) as CodModelo,  A.CodUnidad , " &
        '        "U.DescUnidad, I.ExistenciaTeorica, IsNull(I.ExistenciaFisica,0) as ExistenciaFisica , I.Ajuste, I.CostoUnitario, I.CostoUnitario * I.Ajuste AS Importe   Case A.CodigoAnt " &
        '        "When 0 Then '' Else  cast(A.OrigenAnt as nvarchar) + '-' + right('00000'+  Cast(A.CodigoAnt as varchar),5) " &
        '        "End  as CodigoAnterior , A.CodigoAnt " &
        '        "FROM         dbo.InvHojaControl I INNER JOIN dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN " &
        '        "dbo.CatUnidades U ON A.CodUnidad = U.CodUnidad INNER JOIN dbo.CatOrigen O ON A.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN " &
        '        "dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen INNER JOIN " &
        '        "dbo.CatGrupos Gr ON A.CodGrupo = Gr.CodGrupo LEFT OUTER JOIN " &
        '        "dbo.CatFamilias Fa ON A.CodGrupo = Fa.CodGrupo AND A.CodFamilia = Fa.CodFamilia AND Gr.CodGrupo = Fa.CodGrupo LEFT OUTER JOIN " &
        '        "dbo.CatLineas Li ON A.CodGrupo = Li.CodGrupo AND A.CodFamilia = Li.CodFamilia AND A.CodLinea = Li.CodLinea AND Fa.CodGrupo = Li.CodGrupo AND " &
        '        "Fa.CodFamilia = Li.CodFamilia LEFT OUTER JOIN  dbo.CatSubLineas su ON A.CodGrupo = su.CodGrupo AND A.CodFamilia = su.CodFamilia AND A.CodLinea = su.CodLinea AND " &
        '        "A.CodSubLinea = su.CodSubLinea AND Li.CodGrupo = su.CodGrupo AND Li.CodFamilia = su.CodFamilia AND " &
        '        "Li.CodLinea = su.CodLinea LEFT OUTER JOIN " &
        '        "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca AND Gr.CodGrupo = Ma.CodGrupo LEFT OUTER JOIN " &
        '        "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND A.CodMarca = Mo.CodMarca AND A.CodModelo = Mo.CodModelo AND " &
        '        "Ma.CodGrupo = Mo.CodGrupo AND Ma.CodMarca = Mo.CodMarca " &
        '        "WHERE I.CodAlmacen = " & gintCodAlmacen & " And " & IIf((CodGrupo = 0), "I.CodGrupo in (1,2,3)", "I.CodGrupo  = " & CodGrupo & " ") & IIf((CodMArca = 0), " ", " AND I.CodMarca = " & CodMArca & " ") &
        '        IIf((CodModelo = 0), " ", " AND I.CodModelo = " & CodModelo & " ") &
        '        IIf((CodFamilia = 0), " ", " AND I.CodFamilia = " & CodFamilia & " ") &
        '        IIf((COdLinea = 0), " ", " AND I.CodLinea = " & COdLinea & " ") &
        '        IIf((CodSubLinea = 0), " ", " AND I.CodSubLinea = " & CodSubLinea & " ") &
        '        "GROUP BY I.CodAlmacen, I.CodArticulo, A.CodArticulo, A.DescArticulo, I.CodAlmacenOrigen, A.CodGrupo, A.CodFamilia, A.CodLinea, A.CodSubLinea, A.CodMarca, " &
        '        "A.CodModelo, A.CodUnidad, U.DescUnidad,  I.ExistenciaTeorica, I.ExistenciaFisica, I.Ajuste, " &
        '        "I.CostoUnitario,Al.DescAlmacen,O.DescAlmacenOrigen,Gr.DescGrupo order by A.CodigoAnt"
        'End If

        Cmd.CommandText = "Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
        rec = Cmd.Execute

        With msgArticulos
            If rec.RecordCount > .Rows - 1 Then
                .Rows = rec.RecordCount + 1
            End If
            For I = 1 To rec.RecordCount
                If Not rec.EOF Then
                    .set_TextMatrix(I, C_COLCODIGO, rec.Fields("CodArticulo").Value)
                    .set_TextMatrix(I, C_ColCODIGOANT, rec.Fields("CodigoAnterior").Value)
                    .set_TextMatrix(I, C_COLDESCRIPCION, Trim(rec.Fields("DescArticulo").Value))
                    .set_TextMatrix(I, C_COLUNIDAD, rec.Fields("DescUnidad").Value)
                    .set_TextMatrix(I, C_ColTEORICO, rec.Fields("ExistenciaTeorica").Value)
                    .set_TextMatrix(I, C_ColGRUPO, rec.Fields("CodGrupo").Value)
                    .set_TextMatrix(I, C_ColFAMILIA, rec.Fields("CodFamilia").Value)
                    .set_TextMatrix(I, C_ColLINEA, rec.Fields("COdLinea").Value)
                    .set_TextMatrix(I, C_ColSUBLINEA, rec.Fields("CodSubLinea").Value)
                    .set_TextMatrix(I, C_ColMARCA, rec.Fields("CodMArca").Value)
                    .set_TextMatrix(I, C_ColMODELO, rec.Fields("CodModelo").Value)
                    .set_TextMatrix(I, C_ColFISICO, rec.Fields("ExistenciaFisica").Value)
                    .set_TextMatrix(I, C_ColAJUSTE, rec.Fields("Ajuste").Value)
                    .set_TextMatrix(I, C_ColCOSTOUNITARIO, VB6.Format(rec.Fields("CostoUnitario").Value, gstrFormatoCantidad))
                    .set_TextMatrix(I, C_ColIMPORTE, VB6.Format(rec.Fields("importe").Value, gstrFormatoCantidad))
                End If
                rec.MoveNext()
            Next
        End With
        Exit Function

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub DatosGenerales()
        On Error GoTo Merr
        ''en este procedimiento se cargan los datos registrados para cada nivel
        'antes de Modificar de nuevo el Grid, se Checa si hubo cambios , para Guardar.
        If TVArticulos.Nodes.Count = 0 Then Exit Sub
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        msgArticulos.Clear()
        Encabezado()
        If chkVizSubNivel.CheckState = System.Windows.Forms.CheckState.Checked Then
            With TVArticulos

                If (TVArticulos.SelectedNode Is Nothing) Then
                    Me.Refrescar()
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
                End If

                If TVArticulos.SelectedNode.GetNodeCount(False) = 0 Then
                    ''Si el nodo tiene Hijos, no se muestran los datos
                    Select Case VB.Right(.SelectedNode.Name, 1)
                        Case "M"
                            'D -- Relojeria -- Marca
                            MostrarArticulos(CInt(Mid(.SelectedNode.Name, 1, 4)), 0, 0, 0, CInt(Mid(.SelectedNode.Name, 5, 4)), 0)
                            msgArticulos_EnterCell(msgArticulos, New System.EventArgs())
                        Case "D"
                            'D -- Relojeria-- Marca -- Modelo
                            MostrarArticulos(CInt(Mid(.SelectedNode.Name, 1, 4)), 0, 0, 0, CInt(Mid(.SelectedNode.Name, 5, 4)), CInt(Mid(.SelectedNode.Name, 9, 4)))
                            msgArticulos_EnterCell(msgArticulos, New System.EventArgs())

                        Case "F"
                            ''Joyeria -- Familia
                            MostrarArticulos(CInt(Mid(.SelectedNode.Name, 1, 4)), CInt(Mid(.SelectedNode.Name, 5, 4)), 0, 0, 0, 0)
                            msgArticulos_EnterCell(msgArticulos, New System.EventArgs())
                        Case "L"
                            ''Joyeria -- Familia -Linea
                            MostrarArticulos(CInt(Mid(.SelectedNode.Name, 1, 4)), CInt(Mid(.SelectedNode.Name, 5, 4)), CInt(Mid(.SelectedNode.Name, 9, 4)), 0, 0, 0)
                            msgArticulos_EnterCell(msgArticulos, New System.EventArgs())
                        Case "S"
                            ''Joyeria -- Familia -Linea -- SubLinwea
                            MostrarArticulos(CInt(Mid(.SelectedNode.Name, 1, 4)), CInt(Mid(.SelectedNode.Name, 5, 4)), CInt(Mid(.SelectedNode.Name, 9, 4)), CInt(Mid(.SelectedNode.Name, 13, 4)), 0, 0)
                            msgArticulos_EnterCell(msgArticulos, New System.EventArgs())
                        Case "I"
                            ''VArios  ---- Familia
                            MostrarArticulos(CInt(Mid(.SelectedNode.Name, 1, 4)), CInt(Mid(.SelectedNode.Name, 5, 4)), 0, 0, 0, 0)
                            msgArticulos_EnterCell(msgArticulos, New System.EventArgs())
                        Case "N"
                            ''VArios ---- Familia -Linea
                            MostrarArticulos(CInt(Mid(.SelectedNode.Name, 1, 4)), CInt(Mid(.SelectedNode.Name, 5, 4)), CInt(Mid(.SelectedNode.Name, 9, 4)), 0, 0, 0)
                            msgArticulos_EnterCell(msgArticulos, New System.EventArgs())
                        Case Else
                            txtDesArticulo.Text = ""
                    End Select
                End If
            End With
        Else
            MostrarArticulos(0, 0, 0, 0, 0, 0)
            msgArticulos_EnterCell(msgArticulos, New System.EventArgs())
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub

Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MostrarError("Ocurrió un error al intentar mostrar los datos generales")
    End Sub

    Sub MostrarDatosAlmacen()
        On Error GoTo Merr
        gStrSql = "SELECT DISTINCT dbo.InvHojaControl.CodAlmacen, dbo.CatAlmacen.DescAlmacen, dbo.InvHojaControl.CodAlmacenOrigen, dbo.CatOrigen.DescAlmacenOrigen " & "FROM dbo.InvHojaControl INNER JOIN " & "dbo.CatAlmacen ON dbo.InvHojaControl.CodAlmacen = dbo.CatAlmacen.CodAlmacen INNER JOIN " & "dbo.CatOrigen ON dbo.InvHojaControl.CodAlmacenOrigen = dbo.CatOrigen.CodAlmacenOrigen " & "Where (dbo.InvHojaControl.CodAlmacen =" & mintCodSucursal & ")"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            lblAlmacen.Text = Trim(RsGral.Fields("DescAlmacen").Value)
            Select Case RsGral.RecordCount
                Case 1
                    lblOrigen.Text = Trim(RsGral.Fields("DescAlmacenorigen").Value)
                Case 2
                    lblOrigen.Text = ""
                    For I = 1 To RsGral.RecordCount
                        lblOrigen.Text = lblOrigen.Text & Trim(RsGral.Fields("DescAlmacenorigen").Value) & "  "
                    Next
                Case Is >= 3
                    lblOrigen.Text = "[T O D O S]"
            End Select
        End If
        Exit Sub
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub frmInvAnalisisComparativo_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '   mstrProceso = ""
        '   mstrProcesoInv = ""
    End Sub

    Private Sub msgArticulos_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles msgArticulos.KeyDownEvent
        Select Case eventArgs.keyCode
            Case System.Windows.Forms.Keys.Escape
                If TVArticulos.Enabled = False Then
                    chkVizSubNivel.Focus()
                Else
                    TVArticulos.Focus()
                End If
        End Select
    End Sub

    Private Sub msgArticulos_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgArticulos.Leave
        msgArticulos.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
    End Sub
    Private Sub optImpSoloDiferencias_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optImpSoloDiferencias.CheckedChanged
        If eventSender.Checked Then
            DatosGenerales()
            With msgArticulos
                .Col = 0
                .Row = 1
                .Focus()
            End With
        End If
    End Sub

    Private Sub optImpSoloDiferencias_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optImpSoloDiferencias.Enter
        Pon_Tool()
    End Sub

    Private Sub optImpTodo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optImpTodo.CheckedChanged
        If (isload = False) Then
            Exit Sub
        End If

        If eventSender.Checked Then
            DatosGenerales()
            With msgArticulos
                .Col = 0
                .Row = 1
                .Focus()
            End With
        End If
    End Sub

    Private Sub optImpTodo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optImpTodo.Enter
        Pon_Tool()
    End Sub

    Private Sub TVArticulos_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TVArticulos.Click
        'If TVArticulos.Nodes.Count > 0 Then
        DatosGenerales()
        'End If
    End Sub

    Private Sub TVArticulos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TVArticulos.Enter
        Pon_Tool()
    End Sub

    Private Sub TVArticulos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TVArticulos.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Escape
                'If TVArticulos.Enabled = False Then
                chkVizSubNivel.Focus()
                'Else
                'TVArticulos.SetFocus
                'End If
        End Select
    End Sub


    Private Sub TVArticulos_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TVArticulos.KeyUp
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        'If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
        'If TVArticulos.Nodes.Count > 0 Then
        DatosGenerales()
        'End If
    End Sub

    Private Sub msgArticulos_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgArticulos.EnterCell
        'Aqui poner la descripcion del articulo cuando se este moviendo entre las filas del grid.
        'Poner la descripcion del articulo seleccionado, o dle que tenga la fila seleccionada.
        ' en el Textbox de Descripcion completa de abajo.
        With msgArticulos
            txtDesArticulo.Text = Trim(.get_TextMatrix(.Row, C_COLDESCRIPCION))
        End With
    End Sub

    Private Sub msgArticulos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles msgArticulos.Enter
        msgArticulos.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
        msgArticulos.Col = C_ColFISICO
        Pon_Tool()
        msgArticulos.Col = 0
        msgArticulos.Row = 1
    End Sub

    Public Sub Imprime()
        Imprimir()
    End Sub

    Sub Imprimir()
        Dim rptInvAnalisisComparativoSinGpo_Ubicacion As New rptInvAnalisisComparativoSinGpo_Ubicacion
        Dim rptInvAnalisisComparativo_Ubicacion As New rptInvAnalisisComparativo_Ubicacion
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr
        Dim aParam(3) As Object
        Dim aValues(3) As Object
        Dim Encabezado As String
        Dim ConsultaReporte As String
        Dim cWHERE As String
        Dim CodArticulo As Integer
        Dim CodGrupo As Integer
        Dim CodFamilia As Integer
        Dim COdLinea As Integer
        Dim CodSubLinea As Integer
        Dim CodMArca As Integer
        Dim CodModelo As Integer
        Dim C_ORDERBY As String
        C_ORDERBY = ""

        Encabezado = "Análisis Comparativo de Inventario Fisico"
        If chkOrdenarCodAnt.CheckState = System.Windows.Forms.CheckState.Checked Then
            C_ORDERBY = " Order By A.CodigoAnt "
        Else
            C_ORDERBY = " Order By I.CodArticulo "
        End If

        If chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked Then
            ConsultaReporte = "SELECT     I.CodAlmacen, Al.DescAlmacen, I.CodArticulo, A.DescArticulo, A.CodAlmacenOrigen, O.DescAlmacenOrigen, I.CodGrupo, Gr.DescGrupo, I.CodFamilia, LTRIM(RTRIM(Fa.DescFamilia)) AS DescFamilia, I.CodLinea, LTRIM(RTRIM(Li.DescLinea)) AS DescLinea, I.CodSubLinea, " & "LTRIM(RTRIM(su.DescSubLinea)) AS DescSubLinea, I.CodMarca, LTRIM(RTRIM(Ma.DescMarca)) AS DescMarca, I.CodModelo, LTRIM(RTRIM(Mo.DescModelo)) AS DescModelo, A.CodUnidad, U.DescUnidad, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) " & "AS NombreEmpresa, I.ExistenciaTeorica, IsNull(I.ExistenciaFisica,0) as ExistenciaFisica , I.Ajuste, I.CostoUnitario, I.CostoUnitario * I.Ajuste AS Importe   , Case A.CodigoAnt " & "When 0 Then '' Else  cast(A.OrigenAnt as nvarchar) + '-' + right('00000'+  Cast(A.CodigoAnt as varchar),5) " & "End  as CodigoAnterior , A.CodigoAnt, I.Ubicacion " & "FROM         dbo.InvHojaControl I INNER JOIN " & "dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN " & "dbo.CatUnidades U ON A.CodUnidad = U.CodUnidad INNER JOIN " & "dbo.CatOrigen O ON A.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN " & "dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen INNER JOIN " & "dbo.CatGrupos Gr ON A.CodGrupo = Gr.CodGrupo LEFT OUTER JOIN " & "dbo.CatFamilias Fa ON A.CodGrupo = Fa.CodGrupo AND A.CodFamilia = Fa.CodFamilia AND Gr.CodGrupo = Fa.CodGrupo LEFT OUTER JOIN " & "dbo.CatLineas Li ON A.CodGrupo = Li.CodGrupo AND A.CodFamilia = Li.CodFamilia AND A.CodLinea = Li.CodLinea AND Fa.CodGrupo = Li.CodGrupo AND " & "Fa.CodFamilia = Li.CodFamilia LEFT OUTER JOIN " & "dbo.CatSubLineas su ON A.CodGrupo = su.CodGrupo AND A.CodFamilia = su.CodFamilia AND A.CodLinea = su.CodLinea AND " & "A.CodSubLinea = su.CodSubLinea AND Li.CodGrupo = su.CodGrupo AND Li.CodFamilia = su.CodFamilia AND " & "Li.CodLinea = su.CodLinea LEFT OUTER JOIN " & "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca AND Gr.CodGrupo = Ma.CodGrupo LEFT OUTER JOIN " & "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND A.CodMarca = Mo.CodMarca AND A.CodModelo = Mo.CodModelo AND " & "Ma.CodGrupo = Mo.CodGrupo AND Ma.CodMarca = Mo.CodMarca CROSS JOIN " & "dbo.ConfiguracionGeneral " & "GROUP BY I.CodAlmacen, I.CodArticulo, A.CodArticulo, A.DescArticulo, A.CodAlmacenOrigen, I.CodGrupo, I.CodFamilia, I.CodLinea, I.CodSubLinea, I.CodMarca, " & "I.CodModelo, a.CodUnidad, u.DescUnidad, O.DescAlmacenOrigen, Al.DescAlmacen, Gr.DescGrupo, Fa.DescFamilia, Li.DescLinea, su.DescSubLinea, " & "Ma.DescMarca, Mo.DescModelo, U.DescUnidad, dbo.ConfiguracionGeneral.NombreEmp, I.ExistenciaTeorica, I.ExistenciaFisica, I.Ajuste, " & "I.CostoUnitario, A.CodigoAnt, A.OrigenAnt, I.Ubicacion Having ( I.CodAlmacen= " & mintCodSucursal & " And (IsNull(I.ExistenciaTeorica,0) + IsNull(I.ExistenciaFisica,0) ) > 0 ) "
        Else
            ConsultaReporte = "SELECT     I.CodAlmacen, Al.DescAlmacen, I.CodArticulo, A.DescArticulo, A.CodAlmacenOrigen, O.DescAlmacenOrigen, I.CodFamilia, LTRIM(RTRIM(Fa.DescFamilia)) AS DescFamilia, I.CodLinea, LTRIM(RTRIM(Li.DescLinea)) AS DescLinea, I.CodSubLinea, " & "LTRIM(RTRIM(su.DescSubLinea)) AS DescSubLinea, I.CodMarca, LTRIM(RTRIM(Ma.DescMarca)) AS DescMarca, I.CodModelo, LTRIM(RTRIM(Mo.DescModelo)) AS DescModelo, A.CodUnidad, U.DescUnidad, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) " & "AS NombreEmpresa, I.ExistenciaTeorica, IsNull(I.ExistenciaFisica,0) as ExistenciaFisica , I.Ajuste, I.CostoUnitario, I.CostoUnitario * I.Ajuste AS Importe, Case A.CodigoAnt " & "When 0 Then '' Else  cast(A.OrigenAnt as nvarchar) + '-' + right('00000'+  Cast(A.CodigoAnt as varchar),5) " & "End  as CodigoAnterior , A.CodigoAnt, I.Ubicacion " & "FROM         dbo.InvHojaControl I INNER JOIN " & "dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN " & "dbo.CatUnidades U ON A.CodUnidad = U.CodUnidad INNER JOIN " & "dbo.CatOrigen O ON A.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN " & "dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen INNER JOIN " & "dbo.CatGrupos Gr ON A.CodGrupo = Gr.CodGrupo LEFT OUTER JOIN " & "dbo.CatFamilias Fa ON A.CodGrupo = Fa.CodGrupo AND A.CodFamilia = Fa.CodFamilia AND Gr.CodGrupo = Fa.CodGrupo LEFT OUTER JOIN " & "dbo.CatLineas Li ON A.CodGrupo = Li.CodGrupo AND A.CodFamilia = Li.CodFamilia AND A.CodLinea = Li.CodLinea AND Fa.CodGrupo = Li.CodGrupo AND " & "Fa.CodFamilia = Li.CodFamilia LEFT OUTER JOIN " & "dbo.CatSubLineas su ON A.CodGrupo = su.CodGrupo AND A.CodFamilia = su.CodFamilia AND A.CodLinea = su.CodLinea AND " & "A.CodSubLinea = su.CodSubLinea AND Li.CodGrupo = su.CodGrupo AND Li.CodFamilia = su.CodFamilia AND " & "Li.CodLinea = su.CodLinea LEFT OUTER JOIN " & "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca AND Gr.CodGrupo = Ma.CodGrupo LEFT OUTER JOIN " & "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND A.CodMarca = Mo.CodMarca AND A.CodModelo = Mo.CodModelo AND " & "Ma.CodGrupo = Mo.CodGrupo AND Ma.CodMarca = Mo.CodMarca CROSS JOIN " & "dbo.ConfiguracionGeneral " & "GROUP BY I.CodAlmacen, I.CodArticulo, A.CodArticulo, A.DescArticulo, A.CodAlmacenOrigen, I.CodFamilia, I.CodLinea, I.CodSubLinea, I.CodMarca, " & "I.CodModelo, a.CodUnidad, u.DescUnidad, O.DescAlmacenOrigen, Al.DescAlmacen, Gr.DescGrupo, Fa.DescFamilia, Li.DescLinea, su.DescSubLinea, " & "Ma.DescMarca, Mo.DescModelo, U.DescUnidad, dbo.ConfiguracionGeneral.NombreEmp, I.ExistenciaTeorica, I.ExistenciaFisica, I.Ajuste, " & "I.CostoUnitario, A.CodigoAnt, A.OrigenAnt, I.Ubicacion Having (I.CodAlmacen= " & mintCodSucursal & " And (IsNull(I.ExistenciaTeorica,0) + IsNull(I.ExistenciaFisica,0) ) > 0 ) "
        End If

        With msgArticulos
            If Trim(.get_TextMatrix(1, C_COLCODIGO)) <> "" Then
                CodGrupo = CInt(.get_TextMatrix(1, C_ColGRUPO))
                CodFamilia = CInt(.get_TextMatrix(1, C_ColFAMILIA))
                COdLinea = CInt(.get_TextMatrix(1, C_ColLINEA))
                CodSubLinea = CInt(.get_TextMatrix(1, C_ColSUBLINEA))
                CodMArca = CInt(.get_TextMatrix(1, C_ColMARCA))
                CodModelo = CInt(.get_TextMatrix(1, C_ColMODELO))
            End If
        End With
        If optImpTodo.Checked = False Then
            'cWHERE = " And I.CodGrupo = " & CodGrupo & IIf((CodMArca = 0), " ", " AND I.CodMarca = " & CodMArca & " ") &
            '            IIf((CodModelo = 0), " ", " AND I.CodModelo = " & CodModelo & " ") &
            '            IIf((CodFamilia = 0), " ", " AND I.CodFamilia = " & CodFamilia & " ") &
            '            IIf((COdLinea = 0), " ", " AND I.CodLinea = " & COdLinea & " ") &
            '            IIf((CodSubLinea = 0), " ", " AND I.CodSubLinea = " & CodSubLinea & " ")
            'cWHERE = " ANd I.Ajuste <>0 "
        Else
            cWHERE = ""
        End If

        gStrSql = ConsultaReporte & cWHERE & C_ORDERBY

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existe que reportar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            Exit Sub
        Else
            rptInvAnalisisComparativo_Ubicacion.SetDataSource(frmReportes.rsReport)
            rptInvAnalisisComparativoSinGpo_Ubicacion.SetDataSource(frmReportes.rsReport)
        End If

        If (chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked) Then

            If (Encabezado <> Nothing) Then
                pdvNum.Value = Encabezado : pvNum.Add(pdvNum)
                rptInvAnalisisComparativo_Ubicacion.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvAnalisisComparativo_Ubicacion.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
            End If

            'If ("C.UNITARIO" <> Nothing) Then
            '    pdvNum.Value = "C.UNITARIO" : pvNum.Add(pdvNum)
            '    rptInvAnalisisComparativo_Ubicacion.DataDefinition.ParameterFields("EncabImporte").ApplyCurrentValues(pvNum)
            'Else
            '    pdvNum.Value = "" : pvNum.Add(pdvNum)
            '    rptInvAnalisisComparativo_Ubicacion.DataDefinition.ParameterFields("EncabImporte").ApplyCurrentValues(pvNum)
            'End If

            'If (Trim(lblOrigen.Text) <> Nothing) Then
            '    pdvNum.Value = Trim(lblOrigen.Text) : pvNum.Add(pdvNum)
            '    rptInvAnalisisComparativo_Ubicacion.DataDefinition.ParameterFields("Origen").ApplyCurrentValues(pvNum)
            'Else
            '    pdvNum.Value = "" : pvNum.Add(pdvNum)
            '    rptInvAnalisisComparativo_Ubicacion.DataDefinition.ParameterFields("Origen").ApplyCurrentValues(pvNum)
            'End If

            frmReportes.reporteActual = rptInvAnalisisComparativo_Ubicacion
            frmReportes.Show()

        Else

            If (Encabezado <> Nothing) Then
                pdvNum.Value = Encabezado : pvNum.Add(pdvNum)
                rptInvAnalisisComparativoSinGpo_Ubicacion.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvAnalisisComparativoSinGpo_Ubicacion.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
            End If

            'If ("C.UNITARIO" <> Nothing) Then
            '    pdvNum.Value = "C.UNITARIO" : pvNum.Add(pdvNum)
            '    rptInvAnalisisComparativoSinGpo_Ubicacion.DataDefinition.ParameterFields("EncabImporte").ApplyCurrentValues(pvNum)
            'Else
            '    pdvNum.Value = "" : pvNum.Add(pdvNum)
            '    rptInvAnalisisComparativoSinGpo_Ubicacion.DataDefinition.ParameterFields("EncabImporte").ApplyCurrentValues(pvNum)
            'End If

            'If (Trim(lblOrigen.Text) <> Nothing) Then
            '    pdvNum.Value = Trim(lblOrigen.Text) : pvNum.Add(pdvNum)
            '    rptInvAnalisisComparativoSinGpo_Ubicacion.DataDefinition.ParameterFields("Origen").ApplyCurrentValues(pvNum)
            'Else
            '    pdvNum.Value = "" : pvNum.Add(pdvNum)
            '    rptInvAnalisisComparativoSinGpo_Ubicacion.DataDefinition.ParameterFields("Origen").ApplyCurrentValues(pvNum)
            'End If

            frmReportes.reporteActual = rptInvAnalisisComparativoSinGpo_Ubicacion
            frmReportes.Show()
        End If

        'aParam(1) = "EncabezadoReporte"
        'aValues(1) = Encabezado
        'aParam(2) = "EncabImporte"
        'aValues(2) = "C.UNITARIO"
        'aParam(3) = "Origen"
        'aValues(3) = Trim(lblOrigen.Text)
        ''Set frmReportes.Report = IIf((chkOrdenarporGrupo.Value = vbChecked), rptInvAnalisisComparativo, rptInvAnalisisComparativosinGrupo)     'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Report = IIf((chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked), rptInvAnalisisComparativo_Ubicacion, rptInvAnalisisComparativoSinGpo_Ubicacion) 'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Imprime(Me.Text, aParam, aValues)

        'Exit Sub

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    '    Function ExisteInfoInventario() As Boolean
    '        On Local Error GoTo MErr
    '        gStrSql = "SELECT * " & _
    '        '            "FROM dbo.InvHojaControl " & _
    '        '            "Where (dbo.InvHojaControl.CodAlmacen =" & gintCodAlmacen & ")"
    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.Up_Select_Datos"
    '        Cmd.CommandType = adCmdStoredProc
    '        Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
    '        Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
    '        Set RsGral = Cmd.Execute
    '        If RsGral.RecordCount > 0 Then
    '            ExisteInfoInventario = True
    '        Else
    '            'No existe informacion
    '            MsgBox "No existe información almacenada sobre inventarios. " + vbNewLine + "Verifique Por Favor", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '            ExisteInfoInventario = False
    '        End If
    '        Exit Function
    'MErr:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '        '''
    '    End Function

    Private Sub MostrarDatosSucursal()
        Dim I As Integer
        Dim Aux As Integer

        If ActiveControl.Name <> "" Then
            Exit Sub
        Else
            If Trim(Me.dbcSucursal.Text) = "" Or Trim(Me.dbcSucursal.Text) = C_TODAS Then Exit Sub
        End If

        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        Aux = mintCodSucursal
        mintCodSucursal = 0
        ModDCombo.DCLostFocus((Me.dbcSucursal), gStrSql, mintCodSucursal)
        If mintCodSucursal > 0 Then
            dbcSucursal.Tag = dbcSucursal.Text
            dbcSucursal.Refresh()
            Me.Refresh()

            'System.Windows.Forms.Application.DoEvents()
            Nuevo()
            Encabezado()
            'frmBarraDesplazamiento.InitializeComponent()
            frmBarraDesplazamiento.Text = "Análisis Comparativo"
            frmBarraDesplazamiento.Tag = Me.Name
            frmBarraDesplazamiento.Show()
            frmBarraDesplazamiento.BringToFront()
            frmBarraDesplazamiento.Refresh()
            'System.Windows.Forms.Application.DoEvents()
            FueraChange = True
            frmBarraDesplazamiento.PrgBarra.Value = 20
            frmBarraDesplazamiento.BringToFront()
            frmBarraDesplazamiento.Refresh()
            'System.Windows.Forms.Application.DoEvents()
            Refrescar()
            'System.Windows.Forms.Application.DoEvents()
            frmBarraDesplazamiento.PrgBarra.Value = 40
            frmBarraDesplazamiento.BringToFront()
            frmBarraDesplazamiento.Refresh()
            'System.Windows.Forms.Application.DoEvents()
            frmBarraDesplazamiento.PrgBarra.Value = 60
            frmBarraDesplazamiento.BringToFront()
            frmBarraDesplazamiento.Refresh()
            'System.Windows.Forms.Application.DoEvents()
            DatosGenerales()
            'System.Windows.Forms.Application.DoEvents()
            frmBarraDesplazamiento.PrgBarra.Value = 80
            frmBarraDesplazamiento.BringToFront()
            frmBarraDesplazamiento.Refresh()
            'System.Windows.Forms.Application.DoEvents()
            frmBarraDesplazamiento.PrgBarra.Value = 100
            frmBarraDesplazamiento.BringToFront()
            frmBarraDesplazamiento.Refresh()
            'System.Windows.Forms.Application.DoEvents()
            frmBarraDesplazamiento.Close()
            'FueraChange = False
            msgArticulos.Col = 0
            msgArticulos.Row = 1
            msgArticulos.TopRow = 1
            'msgArticulos.SetFocus
            FueraChange = False
        End If
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInvAnalisisComparativo))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtDesArticulo = New System.Windows.Forms.Label()
        Me.chkOrdenarporGrupo = New System.Windows.Forms.CheckBox()
        Me.optImpTodo = New System.Windows.Forms.RadioButton()
        Me.optImpSoloDiferencias = New System.Windows.Forms.RadioButton()
        Me.chkOrdenarCodAnt = New System.Windows.Forms.CheckBox()
        Me.chkVizSubNivel = New System.Windows.Forms.CheckBox()
        Me.CmbRefrescar = New System.Windows.Forms.Button()
        Me.TVArticulos = New System.Windows.Forms.TreeView()
        Me.Imagenes = New System.Windows.Forms.ImageList(Me.components)
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.msgArticulos = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblOrigen = New System.Windows.Forms.Label()
        Me._Label_2 = New System.Windows.Forms.Label()
        Me.lblAlmacen = New System.Windows.Forms.Label()
        Me._Label_7 = New System.Windows.Forms.Label()
        Me.frame1 = New System.Windows.Forms.GroupBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me._Label_12 = New System.Windows.Forms.Label()
        Me._Label_13 = New System.Windows.Forms.Label()
        Me._Label_14 = New System.Windows.Forms.Label()
        Me._Label_15 = New System.Windows.Forms.Label()
        Me._Label_16 = New System.Windows.Forms.Label()
        Me._Label_17 = New System.Windows.Forms.Label()
        Me._Label_18 = New System.Windows.Forms.Label()
        Me._Label_19 = New System.Windows.Forms.Label()
        Me.Label = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        CType(Me.msgArticulos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.frame1.SuspendLayout()
        Me.Frame3.SuspendLayout()
        CType(Me.Label, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtDesArticulo
        '
        Me.txtDesArticulo.BackColor = System.Drawing.SystemColors.Info
        Me.txtDesArticulo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDesArticulo.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDesArticulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtDesArticulo.Location = New System.Drawing.Point(13, 342)
        Me.txtDesArticulo.Name = "txtDesArticulo"
        Me.txtDesArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDesArticulo.Size = New System.Drawing.Size(650, 21)
        Me.txtDesArticulo.TabIndex = 26
        Me.txtDesArticulo.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolTip1.SetToolTip(Me.txtDesArticulo, "Descripción de Artículos")
        '
        'chkOrdenarporGrupo
        '
        Me.chkOrdenarporGrupo.BackColor = System.Drawing.SystemColors.Control
        Me.chkOrdenarporGrupo.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkOrdenarporGrupo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkOrdenarporGrupo.Location = New System.Drawing.Point(10, 20)
        Me.chkOrdenarporGrupo.Name = "chkOrdenarporGrupo"
        Me.chkOrdenarporGrupo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkOrdenarporGrupo.Size = New System.Drawing.Size(140, 17)
        Me.chkOrdenarporGrupo.TabIndex = 13
        Me.chkOrdenarporGrupo.Text = "Ordenado por grupo"
        Me.ToolTip1.SetToolTip(Me.chkOrdenarporGrupo, "Mostrar órden por grupo ...")
        Me.chkOrdenarporGrupo.UseVisualStyleBackColor = False
        '
        'optImpTodo
        '
        Me.optImpTodo.BackColor = System.Drawing.SystemColors.Control
        Me.optImpTodo.Checked = True
        Me.optImpTodo.Cursor = System.Windows.Forms.Cursors.Default
        Me.optImpTodo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optImpTodo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.optImpTodo.Location = New System.Drawing.Point(27, 42)
        Me.optImpTodo.Name = "optImpTodo"
        Me.optImpTodo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optImpTodo.Size = New System.Drawing.Size(146, 17)
        Me.optImpTodo.TabIndex = 14
        Me.optImpTodo.TabStop = True
        Me.optImpTodo.Text = "Toda la información"
        Me.ToolTip1.SetToolTip(Me.optImpTodo, "Mostrar toda la información ...")
        Me.optImpTodo.UseVisualStyleBackColor = False
        '
        'optImpSoloDiferencias
        '
        Me.optImpSoloDiferencias.BackColor = System.Drawing.SystemColors.Control
        Me.optImpSoloDiferencias.Cursor = System.Windows.Forms.Cursors.Default
        Me.optImpSoloDiferencias.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optImpSoloDiferencias.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.optImpSoloDiferencias.Location = New System.Drawing.Point(27, 61)
        Me.optImpSoloDiferencias.Name = "optImpSoloDiferencias"
        Me.optImpSoloDiferencias.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optImpSoloDiferencias.Size = New System.Drawing.Size(123, 21)
        Me.optImpSoloDiferencias.TabIndex = 15
        Me.optImpSoloDiferencias.TabStop = True
        Me.optImpSoloDiferencias.Text = "Sólo diferencias"
        Me.ToolTip1.SetToolTip(Me.optImpSoloDiferencias, "Mostrar sólo diferencias ...")
        Me.optImpSoloDiferencias.UseVisualStyleBackColor = False
        '
        'chkOrdenarCodAnt
        '
        Me.chkOrdenarCodAnt.BackColor = System.Drawing.SystemColors.Control
        Me.chkOrdenarCodAnt.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkOrdenarCodAnt.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkOrdenarCodAnt.Location = New System.Drawing.Point(9, 66)
        Me.chkOrdenarCodAnt.Name = "chkOrdenarCodAnt"
        Me.chkOrdenarCodAnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkOrdenarCodAnt.Size = New System.Drawing.Size(185, 17)
        Me.chkOrdenarCodAnt.TabIndex = 10
        Me.chkOrdenarCodAnt.Text = "&Ordenar por código anterior"
        Me.ToolTip1.SetToolTip(Me.chkOrdenarCodAnt, "Ordenar por código anterior ...")
        Me.chkOrdenarCodAnt.UseVisualStyleBackColor = False
        '
        'chkVizSubNivel
        '
        Me.chkVizSubNivel.BackColor = System.Drawing.SystemColors.Control
        Me.chkVizSubNivel.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVizSubNivel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkVizSubNivel.Location = New System.Drawing.Point(10, 46)
        Me.chkVizSubNivel.Name = "chkVizSubNivel"
        Me.chkVizSubNivel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVizSubNivel.Size = New System.Drawing.Size(185, 17)
        Me.chkVizSubNivel.TabIndex = 9
        Me.chkVizSubNivel.Text = "&Visualizar por SubNivel"
        Me.ToolTip1.SetToolTip(Me.chkVizSubNivel, "Visualizar por subniveles ...")
        Me.chkVizSubNivel.UseVisualStyleBackColor = False
        '
        'CmbRefrescar
        '
        Me.CmbRefrescar.BackColor = System.Drawing.SystemColors.Control
        Me.CmbRefrescar.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmbRefrescar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmbRefrescar.Location = New System.Drawing.Point(139, 13)
        Me.CmbRefrescar.Name = "CmbRefrescar"
        Me.CmbRefrescar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmbRefrescar.Size = New System.Drawing.Size(71, 32)
        Me.CmbRefrescar.TabIndex = 16
        Me.CmbRefrescar.Text = "Actuali&zar"
        Me.ToolTip1.SetToolTip(Me.CmbRefrescar, "Actualizar proyectos (Alt+Z)")
        Me.CmbRefrescar.UseVisualStyleBackColor = False
        '
        'TVArticulos
        '
        Me.TVArticulos.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TVArticulos.Enabled = False
        Me.TVArticulos.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TVArticulos.ImageIndex = 0
        Me.TVArticulos.ImageList = Me.Imagenes
        Me.TVArticulos.Indent = 6
        Me.TVArticulos.Location = New System.Drawing.Point(10, 86)
        Me.TVArticulos.Name = "TVArticulos"
        Me.TVArticulos.SelectedImageIndex = 0
        Me.TVArticulos.Size = New System.Drawing.Size(193, 191)
        Me.TVArticulos.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.TVArticulos, "Subniveles ...")
        '
        'Imagenes
        '
        Me.Imagenes.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.Imagenes.ImageSize = New System.Drawing.Size(16, 16)
        Me.Imagenes.TransparentColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.msgArticulos)
        Me.Frame2.Controls.Add(Me.dbcSucursal)
        Me.Frame2.Controls.Add(Me._lblVentas_0)
        Me.Frame2.Controls.Add(Me.Label1)
        Me.Frame2.Controls.Add(Me.txtDesArticulo)
        Me.Frame2.Controls.Add(Me.lblOrigen)
        Me.Frame2.Controls.Add(Me._Label_2)
        Me.Frame2.Controls.Add(Me.lblAlmacen)
        Me.Frame2.Controls.Add(Me._Label_7)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(235, 4)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(682, 376)
        Me.Frame2.TabIndex = 17
        Me.Frame2.TabStop = False
        '
        'msgArticulos
        '
        Me.msgArticulos.DataSource = Nothing
        Me.msgArticulos.Location = New System.Drawing.Point(12, 104)
        Me.msgArticulos.Name = "msgArticulos"
        Me.msgArticulos.OcxState = CType(resources.GetObject("msgArticulos.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msgArticulos.Size = New System.Drawing.Size(650, 226)
        Me.msgArticulos.TabIndex = 25
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(68, 24)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(169, 21)
        Me.dbcSucursal.TabIndex = 19
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_0, CType(0, Short))
        Me._lblVentas_0.Location = New System.Drawing.Point(10, 28)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(51, 13)
        Me._lblVentas_0.TabIndex = 18
        Me._lblVentas_0.Text = "Almacén:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(12, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(648, 17)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "EXISTENCIA EN ALMACÉN"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblOrigen
        '
        Me.lblOrigen.BackColor = System.Drawing.Color.White
        Me.lblOrigen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOrigen.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOrigen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblOrigen.Location = New System.Drawing.Point(68, 52)
        Me.lblOrigen.Name = "lblOrigen"
        Me.lblOrigen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOrigen.Size = New System.Drawing.Size(365, 19)
        Me.lblOrigen.TabIndex = 23
        '
        '_Label_2
        '
        Me._Label_2.AutoSize = True
        Me._Label_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_2, CType(2, Short))
        Me._Label_2.Location = New System.Drawing.Point(265, 28)
        Me._Label_2.Name = "_Label_2"
        Me._Label_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_2.Size = New System.Drawing.Size(51, 13)
        Me._Label_2.TabIndex = 20
        Me._Label_2.Text = "Almacén:"
        Me._Label_2.Visible = False
        '
        'lblAlmacen
        '
        Me.lblAlmacen.BackColor = System.Drawing.Color.White
        Me.lblAlmacen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAlmacen.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAlmacen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblAlmacen.Location = New System.Drawing.Point(325, 24)
        Me.lblAlmacen.Name = "lblAlmacen"
        Me.lblAlmacen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAlmacen.Size = New System.Drawing.Size(338, 19)
        Me.lblAlmacen.TabIndex = 21
        Me.lblAlmacen.Visible = False
        '
        '_Label_7
        '
        Me._Label_7.AutoSize = True
        Me._Label_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_7, CType(7, Short))
        Me._Label_7.Location = New System.Drawing.Point(8, 56)
        Me._Label_7.Name = "_Label_7"
        Me._Label_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_7.Size = New System.Drawing.Size(41, 13)
        Me._Label_7.TabIndex = 22
        Me._Label_7.Text = "Origen:"
        '
        'frame1
        '
        Me.frame1.BackColor = System.Drawing.SystemColors.Control
        Me.frame1.Controls.Add(Me.Frame3)
        Me.frame1.Controls.Add(Me.chkOrdenarCodAnt)
        Me.frame1.Controls.Add(Me.chkVizSubNivel)
        Me.frame1.Controls.Add(Me.CmbRefrescar)
        Me.frame1.Controls.Add(Me.TVArticulos)
        Me.frame1.Controls.Add(Me._Label_12)
        Me.frame1.Controls.Add(Me._Label_13)
        Me.frame1.Controls.Add(Me._Label_14)
        Me.frame1.Controls.Add(Me._Label_15)
        Me.frame1.Controls.Add(Me._Label_16)
        Me.frame1.Controls.Add(Me._Label_17)
        Me.frame1.Controls.Add(Me._Label_18)
        Me.frame1.Controls.Add(Me._Label_19)
        Me.frame1.ForeColor = System.Drawing.SystemColors.Highlight
        Me.frame1.Location = New System.Drawing.Point(9, 4)
        Me.frame1.Name = "frame1"
        Me.frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frame1.Size = New System.Drawing.Size(215, 392)
        Me.frame1.TabIndex = 0
        Me.frame1.TabStop = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.chkOrdenarporGrupo)
        Me.Frame3.Controls.Add(Me.optImpTodo)
        Me.Frame3.Controls.Add(Me.optImpSoloDiferencias)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(10, 281)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(193, 95)
        Me.Frame3.TabIndex = 12
        Me.Frame3.TabStop = False
        Me.Frame3.Text = " Mostrar ....."
        '
        '_Label_12
        '
        Me._Label_12.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._Label_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_12.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label.SetIndex(Me._Label_12, CType(12, Short))
        Me._Label_12.Location = New System.Drawing.Point(8, 15)
        Me._Label_12.Name = "_Label_12"
        Me._Label_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_12.Size = New System.Drawing.Size(9, 9)
        Me._Label_12.TabIndex = 1
        '
        '_Label_13
        '
        Me._Label_13.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me._Label_13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_13.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label.SetIndex(Me._Label_13, CType(13, Short))
        Me._Label_13.Location = New System.Drawing.Point(8, 29)
        Me._Label_13.Name = "_Label_13"
        Me._Label_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_13.Size = New System.Drawing.Size(9, 9)
        Me._Label_13.TabIndex = 5
        '
        '_Label_14
        '
        Me._Label_14.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._Label_14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_14.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label.SetIndex(Me._Label_14, CType(14, Short))
        Me._Label_14.Location = New System.Drawing.Point(72, 15)
        Me._Label_14.Name = "_Label_14"
        Me._Label_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_14.Size = New System.Drawing.Size(9, 9)
        Me._Label_14.TabIndex = 3
        '
        '_Label_15
        '
        Me._Label_15.BackColor = System.Drawing.Color.Black
        Me._Label_15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_15.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label.SetIndex(Me._Label_15, CType(15, Short))
        Me._Label_15.Location = New System.Drawing.Point(72, 29)
        Me._Label_15.Name = "_Label_15"
        Me._Label_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_15.Size = New System.Drawing.Size(9, 9)
        Me._Label_15.TabIndex = 7
        '
        '_Label_16
        '
        Me._Label_16.AutoSize = True
        Me._Label_16.BackColor = System.Drawing.Color.Transparent
        Me._Label_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_16, CType(16, Short))
        Me._Label_16.Location = New System.Drawing.Point(24, 13)
        Me._Label_16.Name = "_Label_16"
        Me._Label_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_16.Size = New System.Drawing.Size(41, 13)
        Me._Label_16.TabIndex = 2
        Me._Label_16.Text = "Grupos"
        '
        '_Label_17
        '
        Me._Label_17.AutoSize = True
        Me._Label_17.BackColor = System.Drawing.Color.Transparent
        Me._Label_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_17, CType(17, Short))
        Me._Label_17.Location = New System.Drawing.Point(24, 27)
        Me._Label_17.Name = "_Label_17"
        Me._Label_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_17.Size = New System.Drawing.Size(40, 13)
        Me._Label_17.TabIndex = 6
        Me._Label_17.Text = "Nivel 1"
        '
        '_Label_18
        '
        Me._Label_18.AutoSize = True
        Me._Label_18.BackColor = System.Drawing.Color.Transparent
        Me._Label_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_18, CType(18, Short))
        Me._Label_18.Location = New System.Drawing.Point(88, 13)
        Me._Label_18.Name = "_Label_18"
        Me._Label_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_18.Size = New System.Drawing.Size(40, 13)
        Me._Label_18.TabIndex = 4
        Me._Label_18.Text = "Nivel 2"
        '
        '_Label_19
        '
        Me._Label_19.AutoSize = True
        Me._Label_19.BackColor = System.Drawing.Color.Transparent
        Me._Label_19.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label_19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label.SetIndex(Me._Label_19, CType(19, Short))
        Me._Label_19.Location = New System.Drawing.Point(88, 27)
        Me._Label_19.Name = "_Label_19"
        Me._Label_19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label_19.Size = New System.Drawing.Size(40, 13)
        Me._Label_19.TabIndex = 8
        Me._Label_19.Text = "Nivel 3"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(127, 417)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 143
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(12, 417)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 142
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmInvAnalisisComparativo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(924, 465)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(51, 181)
        Me.MaximizeBox = False
        Me.Name = "frmInvAnalisisComparativo"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Análisis Comparativo de Inventario Teórico - Físico"
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        CType(Me.msgArticulos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.frame1.ResumeLayout(False)
        Me.frame1.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        CType(Me.Label, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub
End Class